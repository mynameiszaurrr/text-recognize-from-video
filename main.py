import speech_recognition as sr
from moviepy.editor import *
from pydub import AudioSegment
from pathlib import Path
import shutil
import os
import docx


videos_pool = []
dir_path = input("Введите путь к папке с видеофайлами: ")
for file_name in os.listdir(dir_path):
    if file_name.lower().endswith((".mp4", ".avi", ".mov", ".mkv")):
        videos_pool.append(os.path.join(dir_path, file_name))
videos_pool = sorted(videos_pool)


def extract_audio(klip):
    print('Открываем видеофайл...')
    video = VideoFileClip(klip)
    print('Извлекаем аудиодорожку...')
    audio = video.audio
    folder_path = 'audios'
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    audio.write_audiofile(os.path.join('audios', os.path.splitext(klip.split('/')[-1])[0] + ".wav"))
    print('Аудиодорожка извлечена успешно...')
    print('*'*50)


def audios_to_shorts():
    path = Path("audios")
    files = os.listdir(path)
    sorted_files = sorted(files)
    for file in sorted_files:
        print(f'Обрезаю - {file}...')
        Path("short audios").mkdir(exist_ok=True)
        for file_path in sorted(Path("audios").iterdir()):
            if file_path.is_file() and file_path.suffix == ".wav":
                audio = AudioSegment.from_file(file_path)
                chunks = list(audio[::119 * 1000])
                for i, chunk in enumerate(chunks):
                    chunk.export(f"short audios/{file_path.stem}_{i}.wav", format="wav")
                file_path.unlink()
    shutil.rmtree(os.path.join(os.getcwd(), 'audios'))


def recognize_and_save():
    r = sr.Recognizer()
    doc = docx.Document()
    for filename in sorted(os.listdir("short audios")):
        if filename.endswith(".wav"):
            print(f'Начинаю обрабатывать файл - {filename}...')
            with sr.AudioFile("short audios/" + filename) as source:
                audio = r.record(source)  # записываем аудиофайл в объект AudioData
                text = r.recognize_google(audio, language="ru-RU")  # распознаем текст
                doc.add_paragraph(text)  # добавляем распознанный текст к остальному
    doc.save("all_text.docx")
    shutil.rmtree(os.path.join(os.getcwd(), 'short audios'))


try:
    for video_in_pool in videos_pool:
        extract_audio(video_in_pool)
    print('Все аудиофайлы обработаны...')
    print('Начинаю обрезать файлы на короткие клипы...')
    audios_to_shorts()
    print('Обрезка файлов окончена успешно...')
    print('Анализирую текст из клипов и создаю текстовой документ...')
    recognize_and_save()
    print('Текстовой документ создан успешно! Приятного дня!')

except Exception as e:
    print(f'Ошибка - {e}...')
