import subprocess
import sys
import requests
from io import BytesIO
import threading
from comtypes import CoInitialize, CoUninitialize
import tkinter as tk
from PIL import Image, ImageTk
import pygame
import os
import time

# Функция для установки библиотек
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


# Попытка импорта необходимых библиотек; если не удается, библиотека устанавливается
try:
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
    from comtypes import CLSCTX_ALL
except ImportError:
    print("pycaw или comtypes не установлены. Устанавливаем...")
    install('pycaw')
    install('comtypes')
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
    from comtypes import CLSCTX_ALL

try:
    import tkinter as tk
    from PIL import Image, ImageTk
except ImportError:
    print("tkinter или Pillow не установлены. Устанавливаем...")
    install('Pillow')
    from PIL import Image, ImageTk

try:
    import requests
except ImportError:
    print("requests не установлена. Устанавливаем...")
    install('requests')
    import requests

try:
    import pygame
except ImportError:
    print("pygame не установлена. Устанавливаем...")
    install('pygame')
    import pygame


# Функция для отключения всех звуков и сохранения текущих уровней громкости
def mute_system():
    CoInitialize()  # Инициализация COM
    sessions = AudioUtilities.GetAllSessions()
    volume_levels = {}
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        volume_levels[session.ProcessId] = volume.GetMasterVolume()
        volume.SetMasterVolume(0, None)  # Отключаем звук
    return volume_levels


# Функция для восстановления звуков до сохраненных уровней громкости
def unmute_system(volume_levels):
    CoInitialize()  # Инициализация COM
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.ProcessId in volume_levels:
            volume.SetMasterVolume(volume_levels[session.ProcessId], None)  # Восстанавливаем громкость
    CoUninitialize()  # Освобождение ресурсов COM


# Функция для загрузки GIF из интернета
def download_gif(url):
    response = requests.get(url)
    if response.status_code == 200:  # Убедимся, что файл успешно загружен
        content_type = response.headers.get('Content-Type')
        if 'image' in content_type:  # Проверим, что загружен именно файл изображения
            print(f"Тип содержимого: {content_type}")
            return BytesIO(response.content)
        else:
            raise ValueError(f"URL не ведет к изображению, получен тип: {content_type}")
    else:
        raise ValueError(f"Не удалось загрузить файл. Статус: {response.status_code}")


# Преобразование ссылки Google Диска для скачивания
def get_google_drive_download_link(google_drive_url):
    file_id = google_drive_url.split('/')[-2]
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    return download_url


# Функция для загрузки звукового файла из Google Диска
def download_sound_from_google_drive(google_drive_url, output_path):
    download_url = get_google_drive_download_link(google_drive_url)
    response = requests.get(download_url)
    if response.status_code == 200:
        with open(output_path, 'wb') as f:
            f.write(response.content)
        print(f"Файл загружен: {output_path}")
    else:
        raise ValueError(f"Не удалось загрузить файл. Статус: {response.status_code}")


# Функция для отображения GIF на весь экран
def show_fullscreen_gif(gif_data):
    # Создаем окно
    root = tk.Tk()
    root.attributes('-fullscreen', True)  # Полноэкранный режим
    root.configure(background='black')  # Фон черный

    # Загружаем GIF
    gif_image = Image.open(gif_data)
    frames = []

    # Сохраняем все кадры GIF
    try:
        while True:
            frames.append(ImageTk.PhotoImage(gif_image.copy()))
            gif_image.seek(len(frames))  # Переходим к следующему кадру
    except EOFError:
        pass

    label = tk.Label(root, bg="black")
    label.pack(expand=True)

    # Функция для анимации GIF
    def update_frame(frame_idx):
        label.config(image=frames[frame_idx])
        root.after(100, update_frame, (frame_idx + 1) % len(frames))

    # Закрытие окна по нажатию клавиши Escape
    root.bind("<Escape>", lambda e: root.destroy())

    # Запуск анимации
    root.after(0, update_frame, 0)
    root.mainloop()


# Функция для воспроизведения звука через pygame
def play_sound_pygame(sound_file):
    try:
        pygame.mixer.init()
    except pygame.error as e:
        print(f"Ошибка инициализации pygame.mixer: {e}")
        return

    # Проверим, что микшер инициализирован
    if not pygame.mixer.get_init():
        print("Ошибка: pygame.mixer не был инициализирован.")
        return

    try:
        pygame.mixer.music.load(sound_file)
        pygame.mixer.music.play()

        # Ожидание завершения воспроизведения звука
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)

    except pygame.error as e:
        print(f"Ошибка воспроизведения звука: {e}")

    finally:
        # Завершаем работу микшера
        pygame.mixer.quit()


if __name__ == "__main__":
    # Укажите URL на ваш GIF-файл
    gif_url = 'https://steamuserimages-a.akamaihd.net/ugc/23962585513184780/470A97ED5BA40555916947A584EAB22E61090B40/?imw=637&imh=358&ima=fit&impolicy=Letterbox&imcolor=%23000000&letterbox=true'

    # Укажите ссылку на аудиофайл в Google Диске
    google_drive_url = 'https://drive.google.com/file/d/1_mDhRQws1z-ecm1jA72Tgsqhwc7Lc1Ow/view?usp=drive_link'

    # Временный файл для сохранения аудиофайла
    sound_file = 'downloaded_sound.mp3'

    print("Отключаем все звуки...")
    saved_volume_levels = mute_system()  # Отключаем звуки и сохраняем уровни громкости

    try:
        # Загружаем аудиофайл из Google Диска
        download_sound_from_google_drive(google_drive_url, sound_file)

        # Загружаем GIF из интернета
        gif_data = download_gif(gif_url)

        # Запускаем отображение GIF в отдельном потоке
        gif_thread = threading.Thread(target=show_fullscreen_gif, args=(gif_data,))
        gif_thread.start()

        # Запускаем воспроизведение загруженного аудиофайла
        sound_thread = threading.Thread(target=play_sound_pygame, args=(sound_file,))
        sound_thread.start()

        # Основная программа (например, ожидание завершения)
        print("Программа работает. Нажмите Enter для завершения...")
        input()

    finally:
        # Останавливаем воспроизведение звука перед удалением файла
        pygame.mixer.music.stop()
        pygame.mixer.quit()

        # Удаляем временный аудиофайл
        if os.path.exists(sound_file):
            os.remove(sound_file)

        # Восстанавливаем звуки при завершении программы
        print("Восстанавливаем все звуки...")
        unmute_system(saved_volume_levels)
        print("Звуки восстановлены.")
