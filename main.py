import json
import queue
import threading
import time
from datetime import datetime
import pyaudio
from vosk import Model, KaldiRecognizer
import os
import win32com.client
import sys
import re
import random
import ctypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Конфигурация
MODEL_VOSK = "vosk-model-small-ru-0.22"
SAMPLE_RATE = 16000
CHUNK_SIZE = 4000
KEYWORDS = ["квант", "кван", "ван"]
ACTIVE_TIMEOUT = 7

class VolumeController:
    def __init__(self):
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))
        
    def get_volume(self):
        return int(self.volume.GetMasterVolumeLevelScalar() * 100)
    
    def set_volume(self, percent):
        percent = max(0, min(100, percent))
        self.volume.SetMasterVolumeLevelScalar(percent / 100.0, None)
        return percent


class VoiceAssistant:
    def __init__(self):
        self.start_time = datetime.now()
        self.print_with_time(f"Запуск")
        
        self.audio_queue = queue.Queue(maxsize=20)
        self.is_running = True
        self.is_active = False
        self.last_activity = 0
        self.last_command_time = 0
        self.min_command_interval = 1.0
        
        self.voice_engine_ready = False
        self.speaker_lock = threading.Lock()
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
        
        self.init_asr()
        
        self.welcome_message = f"Готов. Скажите '{KEYWORDS[0]}'..."
        
        self.patterns = {
            'greeting': re.compile(r'(привет|здравствуй|добрый день)'),
            'time': re.compile(r'(время|час|который час|сколько времени)'),
            'weather': re.compile(r'(погода|погоду|прогноз погоды)'),
            'thanks': re.compile(r'(спасибо|благодарю|пасиб|спс)'),
            'restart': re.compile(r'(перезапуск|перезагрузись|обновись|рестарт)'),
            'help': re.compile(r'(помощь|помоги|что ты умеешь|команды|возможности)'),
            'volume_set': re.compile(r'(громкость на|установи громкость|поставь громкость|громкость)')
        }

    def init_voice_engine(self):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                time.sleep(1)
                self.voice_engine_ready = True
                self.print_with_time("Голос готов")
                return
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2)
        
        self.voice_engine_ready = False
        self.print_with_time("Ошибка голоса")

    def print_with_time(self, message):
        current_time = datetime.now().strftime("%H:%M:%S")
        print(f"[{current_time}] {message}")

    def init_asr(self):
        if not os.path.exists(MODEL_VOSK):
            raise FileNotFoundError(f"Модель не найдена: {MODEL_VOSK}")
        
        self.model = Model(MODEL_VOSK)
        self.recognizer = KaldiRecognizer(self.model, SAMPLE_RATE)
        self.recognizer.AcceptWaveform(b'\x00\x00' * 2000)

    def speak(self, text, interrupt=False):
        if not self.voice_engine_ready:
            return
            
        self.print_with_time(f"Ответ: {text}")
        self.last_command_time = time.time()
        
        try:
            with self.speaker_lock:
                if interrupt and self.speaker.Status.RunningState == 2:
                    self.speaker.Speak("", 2)
                
                self.speaker.Speak(text, 1 if interrupt else 0)
        except Exception as e:
            self.init_voice_engine()

    def deactivate(self, silent=False):
        if self.is_active:
            self.is_active = False
            if not silent:
                self.speak("Режим ожидания", interrupt=True)
            self.audio_queue.queue.clear()

    def audio_capture(self):
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
        
        self.print_with_time(self.welcome_message)
        stream.start_stream()
        
        try:
            while self.is_running:
                try:
                    data = stream.read(CHUNK_SIZE, exception_on_overflow=False)
                    
                    if not self.is_active and self.audio_queue.qsize() > 5:
                        self.audio_queue.queue.clear()
                        
                    self.audio_queue.put(data, timeout=0.1)
                except queue.Full:
                    continue
                except Exception as e:
                    break
        finally:
            stream.stop_stream()
            stream.close()
            p.terminate()

    def process_audio(self):
        last_processed_text = ""
        last_processed_time = 0
        
        while self.is_running:
            try:
                data = self.audio_queue.get(timeout=0.1)
                
                if self.recognizer.AcceptWaveform(data):
                    result = json.loads(self.recognizer.Result())
                    text = result.get("text", "").strip().lower()
                    
                    if text and (time.time() - last_processed_time > 0.5 or text != last_processed_text):
                        self.print_with_time(f"Распознано: {text}")
                        self.handle_command(text)
                        last_processed_text = text
                        last_processed_time = time.time()
                
                self.audio_queue.task_done()
                
            except queue.Empty:
                continue
            except Exception as e:
                self.print_with_time(f"Ошибка: {e}")

    def handle_command(self, text):
        text_lower = text.lower()
        keyword_detected = any(keyword in text_lower for keyword in KEYWORDS)
        
        if keyword_detected:
            if not self.is_active:
                self.activate()
            command = re.sub('|'.join(KEYWORDS), '', text_lower).strip()
            if command:
                self.process_user_input(command)
            return
        
        if self.is_active:
            self.last_activity = time.time()
            self.process_user_input(text_lower)

    def activate(self):
        if not self.is_active:
            self.is_active = True
            self.last_activity = time.time()
            self.audio_queue.queue.clear()

    def restart(self):
        self.print_with_time("Перезапуск...")
        self.speak("Рестарт", interrupt=True)
        self.is_running = False
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def process_user_input(self, text):
        responses = {
            'greeting': [
                "Приветствую!", 
                "Здравствуйте!", 
                "Привет!"
            ],
            'time': [
                datetime.now().strftime('%H:%M'),
            ],
            'weather': [
                "Посмотрите погоду в приложении",
                "Не могу узнать погоду, проверьте сами",
                "Погоду лучше уточнить в интернете"
            ],
            'thanks': [
                "Всегда пожалуйста!",
                "Не за что!",
                "Рад помочь!"
            ],
            'help': [
                "Команды: время, погода, перезагрузка. Управление громкостью: 'громкость на 50' или 'громкость на десять'. Скажите 'квант' перед командой.",
                "Просто скажите 'квант' и что вам нужно: время, погода, громкость и т.д."
            ],
            'volume_set': [
                "{}"
            ],
            'default': [
                "Не понял",
                "Повторите, пожалуйста"
            ]
        }

        # Обработка команд громкости
        if self.patterns['volume_set'].search(text):
            # Ищем цифры
            numbers = re.findall(r'\d+', text)
            if numbers:
                vol = int(numbers[0])
            else:
                # Ищем слова-числа
                words = text.split()
                vol = None
                for word in words:
                    if word in self.number_words:
                        vol = self.number_words[word]
                        break
                
                if vol is None:
                    current_vol = self.volume_controller.get_volume()
                    response = f"Текущая {current_vol}%"
                    self.speak(response, interrupt=True)
                    self.deactivate(silent=True)
                    return
            
            # Устанавливаем громкость
            new_vol = self.volume_controller.set_volume(vol)
            response = random.choice(responses['volume_set']).format(new_vol)
        
        
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
        try:
            audio_thread = threading.Thread(target=self.audio_capture)
            process_thread = threading.Thread(target=self.process_audio)
            
            audio_thread.daemon = True
            process_thread.daemon = True
            
            if sys.platform == 'win32':
                try:
                    import win32api, win32process, win32con
                    win32process.SetPriorityClass(win32api.GetCurrentProcess(), win32process.HIGH_PRIORITY_CLASS)
                    win32process.SetThreadPriority(win32api.GetCurrentThread(), win32process.THREAD_PRIORITY_HIGHEST)
                except:
                    pass
            
            audio_thread.start()
            process_thread.start()
            
            while self.is_running:
                time.sleep(0.1)
                
                # Автоматическое отключение после периода неактивности
                if self.is_active and (time.time() - self.last_activity) > ACTIVE_TIMEOUT:
                    self.deactivate()
                
        except KeyboardInterrupt:
            self.print_with_time("\nВыход...")
        finally:
            self.is_running = False
            os._exit(0)

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run()