import json
import queue
import threading
import time
from datetime import datetime
import pyaudio
from vosk import Model, KaldiRecognizer, SetLogLevel
SetLogLevel(-1)  # Отключаем логирование Vosk
import os
import win32com.client
import sys
import re
import random
import ctypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import zipfile
import shutil
from tqdm import tqdm
import requests
import webbrowser
import psutil
import subprocess

# Конфигурационные параметры
MODELS_DIR = "models"  # Папка для хранения моделей распознавания речи
MODEL_NAME = "vosk-model-small-ru-0.22"  # Название модели Vosk
MODEL_URL = "https://alphacephei.com/vosk/models/vosk-model-small-ru-0.22.zip"  # URL для скачивания модели
SAMPLE_RATE = 16000  # Частота дискретизации аудио
CHUNK_SIZE = 4000  # Размер чанка для аудиопотока
KEYWORDS = ["квант", "кван", "ван"]  # Ключевые слова для активации
ACTIVE_TIMEOUT = 7  # Таймаут неактивности в секундах

# Пути к программам и браузерам
BROWSER_PATH = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s'
PROGRAM_PATHS = {
    'paint': r'C:\Windows\System32\mspaint.exe',
    'telegram': r'C:\Users\User\AppData\Roaming\Telegram Desktop\Telegram.exe',
    'yandex': r'C:\Users\User\AppData\Local\Yandex\YandexBrowser\Application\browser.exe'
}
browser = webbrowser.get(BROWSER_PATH)

def print_with_time(message, color=None):
    """Выводит сообщение с текущим временем и опциональным цветом"""
    current_time = datetime.now().strftime("%H:%M:%S")
    colored_message = f"[{current_time}] {message}"
    
    if color == "green":
        colored_message = f"\033[32m{colored_message}\033[0m"
    elif color == "bold_green":
        colored_message = f"\033[1;32m{colored_message}\033[0m"
    
    print(colored_message)

def download_file(url, filename):
    """Скачивает файл с отображением прогресса через tqdm"""
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    
    with open(filename, 'wb') as f, tqdm(
        desc=filename,
        total=total_size,
        unit='iB',
        unit_scale=True,
        unit_divisor=1024,
    ) as bar:
        for data in response.iter_content(chunk_size=1024):
            size = f.write(data)
            bar.update(size)

def ensure_model_exists():
    """Проверяет наличие модели речи и скачивает при необходимости"""
    model_path = os.path.join(MODELS_DIR, MODEL_NAME)
    os.makedirs(MODELS_DIR, exist_ok=True)
    
    if os.path.exists(model_path):
        print_with_time("Модель для распознавания речи готова", color="green")
        return True
    
    print_with_time(f"Модель {MODEL_NAME} не найдена, начинаю загрузку...", color="green")
    
    try:
        zip_path = os.path.join(MODELS_DIR, "temp_model.zip")
        download_file(MODEL_URL, zip_path)
        
        # Распаковываем архив с моделью
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            for file in tqdm(file_list, desc="Распаковка"):
                zip_ref.extract(file, MODELS_DIR)
        
        os.remove(zip_path)
        
        # Проверяем корректность распаковки
        if not os.path.exists(model_path):
            extracted_dir = os.path.join(MODELS_DIR, file_list[0].split('/')[0])
            if os.path.exists(extracted_dir):
                shutil.move(extracted_dir, model_path)
        
        print_with_time("Модель для распознавания готова", color="green")
        return True
    
    except Exception as e:
        print_with_time(f"Ошибка при загрузке модели: {e}")
        return False

class VolumeController:
    """Контроллер громкости системы через Windows API"""
    def __init__(self):
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))
        
    def get_volume(self):
        """Возвращает текущую громкость в процентах"""
        return int(self.volume.GetMasterVolumeLevelScalar() * 100)
    
    def set_volume(self, percent):
        """Устанавливает громкость (0-100%)"""
        percent = max(0, min(100, percent))
        self.volume.SetMasterVolumeLevelScalar(percent / 100.0, None)
        return percent

class VoiceAssistant:
    """Основной класс голосового ассистента"""
    def __init__(self):
        self.start_time = datetime.now()
        print_with_time("Запуск ассистента", color="bold_green")
        
        # Проверяем наличие модели распознавания речи
        if not ensure_model_exists():
            raise Exception("Не удалось загрузить модель распознавания")
        
        # Очередь для аудиоданных между потоками
        self.audio_queue = queue.Queue(maxsize=20)
        self.is_running = True  # Флаг работы основного цикла
        self.is_active = False  # Флаг активного режима (после ключевого слова)
        self.last_activity = 0  # Время последней активности
        self.last_command_time = 0  # Время последней команды
        self.min_command_interval = 1.0  # Минимальный интервал между командами
        
        # Инициализация голосового движка (SAPI)
        self.voice_engine_ready = False
        self.speaker_lock = threading.Lock()  # Блокировка для синхронизации речи
        self.init_voice_engine()
        
        self.volume_controller = VolumeController()
        
        # Словарь для преобразования слов в числа
        self.number_words = {
            'ноль': 0, 'один': 1, 'два': 2, 'три': 3, 'четыре': 4,
            'пять': 5, 'шесть': 6, 'семь': 7, 'восемь': 8, 'девять': 9,
            'десять': 10, 'одиннадцать': 11, 'двенадцать': 12, 'тринадцать': 13,
            'четырнадцать': 14, 'пятнадцать': 15, 'шестнадцать': 16,
            'семнадцать': 17, 'восемнадцать': 18, 'девятнадцать': 19,
            'двадцать': 20, 'тридцать': 30, 'сорок': 40,
            'пятьдесят': 50, 'шестьдесят': 60, 'семьдесят': 70,
            'восемьдесят': 80, 'девяносто': 90, 'сто': 100
        }
        
        # Инициализация распознавания речи
        self.init_asr()
        
        self.welcome_message = f"Готов. Скажите '{KEYWORDS[0]}'..."
        
        # Регулярные выражения для распознавания команд
        self.patterns = {
            'greeting': re.compile(r'(привет|здравствуй|добрый день)'),
            'time': re.compile(r'(время|час|который час|сколько времени)'),
            'weather': re.compile(r'(погода|погоду|прогноз погоды)'),
            'thanks': re.compile(r'(спасибо|благодарю|пасиб|спс)'),
            'restart': re.compile(r'(перезапуск|перезагрузись|обновись|рестарт)'),
            'help': re.compile(r'(помощь|помоги|что ты умеешь|команды|возможности)'),
            'volume_set': re.compile(r'(громкость на|установи громкость|поставь громкость|громкость)'),
            'search': re.compile(r'(поиск|найди|найти)'),
            'open_paint': re.compile(r'(paint|рисовать)'),
            'open_telegram': re.compile(r'(telegram|телеграм|телега)'),
            'open_yandex': re.compile(r'(яндекс|браузер)'),
            'deepseek_search': re.compile(r'(нейронка|нейросеть)'),
            'system_status': re.compile(r'(состояние системы|загрузка системы|диск|диски)')
        }

    def init_asr(self):
        """Инициализация системы распознавания речи (ASR)"""
        model_path = os.path.join(MODELS_DIR, MODEL_NAME)
        if not os.path.exists(model_path):
            raise FileNotFoundError(f"Модель не найдена: {model_path}")
        
        self.model = Model(model_path)
        self.recognizer = KaldiRecognizer(self.model, SAMPLE_RATE)
        self.recognizer.SetWords(True)  # Включаем распознавание отдельных слов

    def init_voice_engine(self):
        """Инициализация голосового движка (SAPI) с несколькими попытками"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                time.sleep(1)  # Даем время на инициализацию
                self.voice_engine_ready = True
                print_with_time("Голосовой движок готов", color="green")
                return
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2)
        
        self.voice_engine_ready = False
        print_with_time("Ошибка инициализации голосового движка")

    def speak(self, text, interrupt=False):
        """Произносит текст с возможностью прерывания текущей речи"""
        if not self.voice_engine_ready:
            return
            
        print_with_time(f"Ответ: {text}")
        self.last_command_time = time.time()  # Обновляем время последней команды
        
        try:
            with self.speaker_lock:  # Блокируем для потокобезопасности
                if interrupt and self.speaker.Status.RunningState == 2:  # 2 = speaking
                    self.speaker.Speak("", 2)  # Прерываем текущую речь
                
                # Флаг 1 - асинхронное произношение, 0 - синхронное
                self.speaker.Speak(text, 1 if interrupt else 0)
        except Exception as e:
            # При ошибке пытаемся переинициализировать движок
            self.init_voice_engine()

    def deactivate(self, silent=False):
        """Переводит ассистента в режим ожидания"""
        if self.is_active:
            self.is_active = False
            if not silent:
                self.speak("Режим ожидания", interrupt=True)
            self.audio_queue.queue.clear()  # Очищаем очередь аудио

    def audio_capture(self):
        """Поток для захвата аудио с микрофона"""
        p = pyaudio.PyAudio()
        stream = p.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=SAMPLE_RATE,
            input=True,
            frames_per_buffer=CHUNK_SIZE,
            input_device_index=None,
            stream_callback=None,
            start=False
        )
        
        print_with_time(self.welcome_message, color="bold_green")
        print("-" * 40)
        self.speak("Готов")
        stream.start_stream()
        
        try:
            while self.is_running:
                try:
                    # Читаем данные с микрофона без блокировки при переполнении
                    data = stream.read(CHUNK_SIZE, exception_on_overflow=False)
                    
                    # Очищаем очередь, если в режиме ожидания и очередь переполнена
                    if not self.is_active and self.audio_queue.qsize() > 5:
                        self.audio_queue.queue.clear()
                        
                    self.audio_queue.put(data, timeout=0.1)
                except queue.Full:
                    continue
                except Exception as e:
                    break
        finally:
            # Гарантированно останавливаем поток
            stream.stop_stream()
            stream.close()
            p.terminate()

    def process_audio(self):
        """Поток для обработки аудио и распознавания команд"""
        last_processed_text = ""
        last_processed_time = 0
        
        while self.is_running:
            try:
                data = self.audio_queue.get(timeout=0.1)
                
                if self.recognizer.AcceptWaveform(data):
                    result = json.loads(self.recognizer.Result())
                    text = result.get("text", "").strip().lower()
                    
                    # Обрабатываем только новые команды (защита от дублирования)
                    if text and (time.time() - last_processed_time > 0.5 or text != last_processed_text):
                        keyword_detected = any(keyword in text for keyword in KEYWORDS)
                        
                        # Выводим в консоль только если есть ключевое слово или в активном режиме
                        if keyword_detected or self.is_active:
                            print_with_time(f"Распознано: {text}")
                        
                        self.handle_command(text)
                        last_processed_text = text
                        last_processed_time = time.time()
                
                self.audio_queue.task_done()
                
            except queue.Empty:
                continue
            except Exception as e:
                print_with_time(f"Ошибка обработки аудио: {e}")

    def handle_command(self, text):
        """Определяет и выполняет команды из распознанного текста"""
        text_lower = text.lower()
        keyword_detected = any(keyword in text_lower for keyword in KEYWORDS)
        
        # Активация по ключевому слову
        if keyword_detected:
            if not self.is_active:
                self.activate()
            # Удаляем ключевое слово из команды
            command = re.sub('|'.join(KEYWORDS), '', text_lower).strip()
            if command:
                self.process_user_input(command)
            return
        
        # Обработка команд в активном режиме
        if self.is_active:
            self.last_activity = time.time()  # Обновляем время активности
            self.process_user_input(text_lower)

    def activate(self):
        """Активирует режим ожидания команд"""
        if not self.is_active:
            self.is_active = True
            self.last_activity = time.time()
            self.audio_queue.queue.clear()  # Очищаем накопившиеся данные

    def restart(self):
        """Перезапускает ассистента"""
        self.is_running = False
        python = sys.executable
        print("-" * 40)
        os.execl(python, python, *sys.argv)


    def open_program(self, program_name):
        """Открывает указанную программу"""
        try:
            if program_name in PROGRAM_PATHS:
                path = PROGRAM_PATHS[program_name]
                subprocess.Popen(path)
                return True
        except Exception as e:
            print_with_time(f"Ошибка при открытии программы {program_name}: {e}")
        return False

    def get_system_status(self):
        """Возвращает информацию только о загруженности дисков в формате 'C: 88% D: 90%'"""
        disk_status = []
        for partition in psutil.disk_partitions(all=False):
            if partition.fstype and 'cdrom' not in partition.opts:
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    drive_letter = partition.mountpoint[0]
                    percent = round(usage.percent)  # Округляем до целого числа
                    disk_status.append(f"{drive_letter}: {percent}%")
                except Exception:
                    continue
        
        return f"Диски: {' '.join(disk_status)}"

    def process_user_input(self, text):
        """Обрабатывает распознанный текст и формирует ответ"""
        responses = {
            'greeting': ["Приветствую!", "Здравствуйте!", "Привет!"],
            'time': [datetime.now().strftime('%H:%M')],
            'weather': [
                "Посмотрите погоду в приложении",
                "Не могу узнать погоду, проверьте сами",
                "Погоду лучше уточнить в интернете"
            ],
            'thanks': ["Всегда пожалуйста!", "Не за что!", "Рад помочь!"],
            'help': [
                "Команды: время, погода, перезагрузка. Скажите 'квант' перед командой.",
                "Просто скажите 'квант' и что вам нужно: время, погода, громкость и т.д."
            ],
            'volume_set': ["{}"],  # Шаблон для ответа о громкости
            'search': ["Ищу информацию в интернете"],
            'open_paint': ["Открываю Paint"],
            'open_telegram': ["Открываю Telegram"],
            'open_yandex': ["Открываю Яндекс Браузер"],
            'deepseek_search': ["Открываю DeepSeek в браузере"],
            'system_status': ["{}"],
            'restart': ["Выполняю перезапуск"],  
            'default': ["Не понял", "Повторите, пожалуйста"]
        }

        # Обработка команд громкости
        if self.patterns['volume_set'].search(text):
            # Пытаемся найти цифры в команде
            numbers = re.findall(r'\d+', text)
            if numbers:
                vol = int(numbers[0])
            else:
                # Ищем числительные в тексте
                words = text.split()
                vol = None
                for word in words:
                    if word in self.number_words:
                        vol = self.number_words[word]
                        break
                
                # Если не нашли число - сообщаем текущую громкость
                if vol is None:
                    current_vol = self.volume_controller.get_volume()
                    response = f"Текущая громкость {current_vol}%"
                    self.speak(response, interrupt=True)
                    self.deactivate(silent=True)
                    return
            
            # Устанавливаем новую громкость
            new_vol = self.volume_controller.set_volume(vol)
            response = random.choice(responses['volume_set']).format(f"Громкость установлена на {new_vol}%")
        
        # Поиск информации в интернете
        elif self.patterns['search'].search(text):
            query = re.sub(r'(поиск|найди|найти)\s*', '', text).strip()
            if query:
                browser.open(f"https://www.google.com/search?q={query}")
                response = random.choice(responses['search'])
            else:
                response = "Что нужно найти?"
        
        # Открытие Paint
        elif self.patterns['open_paint'].search(text):
            if self.open_program('paint'):
                response = random.choice(responses['open_paint'])
            else:
                response = "Не удалось открыть Paint"
        
        # Открытие Telegram
        elif self.patterns['open_telegram'].search(text):
            if self.open_program('telegram'):
                response = random.choice(responses['open_telegram'])
            else:
                response = "Не удалось открыть Telegram"
        
        # Открытие Яндекс Браузера
        elif self.patterns['open_yandex'].search(text):
            if self.open_program('yandex'):
                response = random.choice(responses['open_yandex'])
            else:
                response = "Не удалось открыть Яндекс Браузер"
        
        # Поиск в DeepSeek
        elif self.patterns['deepseek_search'].search(text):
            browser.open("https://www.deepseek.com")
            response = random.choice(responses['deepseek_search'])
        
        # Состояние системы (включая диски)
        elif self.patterns['system_status'].search(text):
            status = self.get_system_status()
            response = random.choice(responses['system_status']).format(status)
        
        # Обработка остальных команд
        elif self.patterns['greeting'].search(text):
            response = random.choice(responses['greeting'])
        elif self.patterns['time'].search(text):
            response = random.choice(responses['time'])
        elif self.patterns['weather'].search(text):
            response = random.choice(responses['weather'])
        elif self.patterns['thanks'].search(text):
            response = random.choice(responses['thanks'])
        elif self.patterns['restart'].search(text):
            response = random.choice(responses['restart'])
            self.speak(response, interrupt=True)
            self.restart()
            return
        elif self.patterns['help'].search(text):
            response = random.choice(responses['help'])
        else:
            response = random.choice(responses['default'])

        self.speak(response, interrupt=True)
        self.deactivate(silent=True)

    def run(self):
        """Основной цикл работы ассистента"""
        try:
            # Запускаем потоки для захвата и обработки аудио
            audio_thread = threading.Thread(target=self.audio_capture)
            process_thread = threading.Thread(target=self.process_audio)
            
            audio_thread.daemon = True
            process_thread.daemon = True
            
            # Повышаем приоритет на Windows
            if sys.platform == 'win32':
                try:
                    import win32api, win32process, win32con
                    win32process.SetPriorityClass(win32api.GetCurrentProcess(), win32process.HIGH_PRIORITY_CLASS)
                    win32process.SetThreadPriority(win32api.GetCurrentThread(), win32process.THREAD_PRIORITY_HIGHEST)
                except:
                    pass
            
            audio_thread.start()
            process_thread.start()
            
            # Основной цикл управления
            while self.is_running:
                time.sleep(0.1)
                
                # Автоматическое отключение после таймаута неактивности
                if self.is_active and (time.time() - self.last_activity) > ACTIVE_TIMEOUT:
                    self.deactivate()
                
        except KeyboardInterrupt:
            print_with_time("\nЗавершение работы...")
        finally:
            self.is_running = False
            os._exit(0)

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run()