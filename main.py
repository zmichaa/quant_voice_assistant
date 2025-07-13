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

# Конфигурация
MODEL_VOSK = "vosk-model-small-ru-0.22"
SAMPLE_RATE = 16000
CHUNK_SIZE = 4000
KEYWORDS = ["квант"]
ACTIVE_TIMEOUT = 7

class VoiceAssistant:
    def __init__(self):
        self.start_time = datetime.now()
        self.print_with_time(f"Ассистент запущен")
        
        self.audio_queue = queue.Queue()
        self.is_running = True
        self.is_active = False
        self.last_activity = 0
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.init_asr()
        
        self.welcome_message = f"Система готова. Скажите '{KEYWORDS[0]}' для активации..."

    def print_with_time(self, message):
        current_time = datetime.now().strftime("%H:%M:%S")
        print(f"[{current_time}] {message}")

    def init_asr(self):
        if not os.path.exists(MODEL_VOSK):
            raise FileNotFoundError(f"Модель Vosk не найдена: {MODEL_VOSK}")
        self.model = Model(MODEL_VOSK)
        self.recognizer = KaldiRecognizer(self.model, SAMPLE_RATE)

    def speak(self, text, priority=False):
        try:
            self.print_with_time(f"Ответ: {text}")
            if priority:
                self.speaker.Speak(text, 1)
            else:
                while self.speaker.Status.RunningState == 2:
                    time.sleep(0.1)
                self.speaker.Speak(text, 1)
        except Exception as e:
            self.print_with_time(f"Ошибка синтеза речи: {e}")

    def audio_capture(self):
        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16,
                       channels=1,
                       rate=SAMPLE_RATE,
                       input=True,
                       frames_per_buffer=CHUNK_SIZE)
        
        self.print_with_time(self.welcome_message)
        
        while self.is_running:
            try:
                data = stream.read(CHUNK_SIZE, exception_on_overflow=False)
                self.audio_queue.put(data)
            except Exception as e:
                self.print_with_time(f"Ошибка записи аудио: {e}")
                break
        
        stream.stop_stream()
        stream.close()
        p.terminate()

    def process_audio(self):
        while self.is_running:
            data = self.audio_queue.get()
            
            if self.recognizer.AcceptWaveform(data):
                result = json.loads(self.recognizer.Result())
                text = result.get("text", "").strip().lower()
                
                if text:
                    self.print_with_time(f"Распознано: {text}")
                    self.handle_command(text)
            
            self.audio_queue.task_done()

    def handle_command(self, text):
        # Проверяем, содержит ли текст ключевое слово
        for keyword in KEYWORDS:
            if keyword in text:
                # Активируем и обрабатываем всю фразу целиком
                if not self.is_active:
                    self.activate()
                # Удаляем ключевое слово из текста для обработки
                command = text.replace(keyword, "").strip()
                if command:
                    self.process_user_input(command)
                else:
                    self.speak("Слушаю вас", priority=True)
                return
        
        # Если ключевого слова нет, но ассистент активен - обрабатываем команду
        if self.is_active:
            self.last_activity = time.time()
            self.process_user_input(text)

    def activate(self):
        self.is_active = True
        self.last_activity = time.time()
        self.print_with_time("Активирован")

    def deactivate(self, silent=False):
        if self.is_active:
            self.is_active = False
            self.print_with_time("Деактивирован")
            if not silent:
                self.speak("Перехожу в режим ожидания", priority=True)

    def restart(self):
        self.print_with_time("Перезапуск системы...")
        self.speak("Перезагружаюсь", priority=True)
        self.is_running = False
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def process_user_input(self, text):
        patterns = {
            'greeting': r'(привет|здравствуй|добрый день)',
            'time': r'(время|час|который час|сколько времени)',
            'weather': r'(погода|погоду|прогноз погоды)',
            'thanks': r'(спасибо|благодарю|пасиб|спс)',
            'deactivate': r'(стоп|хватит|замолчи|выключись|отстань)',
            'restart': r'(перезапуск|перезагрузись|обновись|рестарт)',
            'help': r'(помощь|помоги|что ты умеешь|команды|возможности)'
        }

        response = None
        
        if re.search(patterns['greeting'], text):
            response = "Приветствую! Чем могу помочь?"
        elif re.search(patterns['time'], text):
            response = f"Сейчас {datetime.now().strftime('%H:%M')}"
        elif re.search(patterns['weather'], text):
            response = "Рекомендую посмотреть погоду в приложении на телефоне"
        elif re.search(patterns['thanks'], text):
            response = "Всегда рад помочь!"
        elif re.search(patterns['deactivate'], text):
            self.deactivate()
            return
        elif re.search(patterns['restart'], text):
            self.restart()
            return
        elif re.search(patterns['help'], text):
            response = ("Я могу: сказать текущее время, "
                      "подсказать про погоду, перезагрузиться. "
                      "Просто скажите 'квант' и свою просьбу.")
        else:
            response = "Не совсем понял. Попробуйте сказать 'помощь' для списка команд"

        if response:
            self.speak(response)

    def run(self):
        try:
            audio_thread = threading.Thread(target=self.audio_capture)
            process_thread = threading.Thread(target=self.process_audio)
            
            audio_thread.daemon = True
            process_thread.daemon = True
            
            audio_thread.start()
            process_thread.start()
            
            while self.is_running:
                if self.is_active and (time.time() - self.last_activity) > ACTIVE_TIMEOUT:
                    self.deactivate(silent=True)
                time.sleep(1)
                
        except KeyboardInterrupt:
            self.print_with_time("\nЗавершение работы...")
        finally:
            self.is_running = False

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run()