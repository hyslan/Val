"""Módulo de gif de bloqueio do sistema"""
import threading
import time
import tkinter
import tkinter as tk
from collections.abc import Generator
from typing import Any

import cv2
import imageio
import pygame
import simpleaudio as sa
from PIL import Image, ImageTk
from PIL.ImageTk import PhotoImage
from pydub import AudioSegment
from simpleaudio import PlayObject

stop_audio = False


def play_audio(audio_path) -> None:
    audio: AudioSegment | Generator[Any, Any, None] = AudioSegment.from_file(audio_path, format="m4a")
    play_obj: PlayObject = sa.play_buffer(audio.raw_data, num_channels=audio.channels,
                                          bytes_per_sample=audio.sample_width, sample_rate=audio.frame_rate,
                                          )
    while play_obj.is_playing() and not stop_audio:
        time.sleep(0.1)
    play_obj.stop()


def video():
    global stop_audio
    video_path: str = "media/gandalf.mp4"
    audio_path: str = "media/gandalf_audio.mp3"
    # Create object
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print("Error opening video stream or file")
        exit()

    # Get the frame rate
    fps: float = cap.get(cv2.CAP_PROP_FPS)
    frame_delay: float = 1 / fps

    # Create thread to play audio
    audio_thread: threading.Thread = threading.Thread(target=play_audio, args=(audio_path,))
    audio_thread.start()

    while cap.isOpened():
        start_time: float = time.time()
        # Capture frame-by-frame
        ret, frame = cap.read()
        if not ret:
            print("Can't receive frame (stream end?). Exiting ...")
            break

        # Show frame
        cv2.imshow("Voce nao vai passar!", frame)

        # press 'q' to exit
        if cv2.waitKey(1) & 0xFF == ord("q"):
            stop_audio = True
            break

        # Wait for the next frame
        time.sleep(max(0, frame_delay - (time.time() - start_time)))

    # Release object e close all windows
    cap.release()
    cv2.destroyAllWindows()
    stop_audio = True
    audio_thread.join()


def gif():
    root: tkinter.Tk = tk.Tk()

    # Configuração do GIF
    gif_path: str = "media/gandalf.gif"
    gif_reader: imageio.core.Format.Reader = imageio.get_reader(gif_path)
    frame_cnt: int = gif_reader.get_length()
    frames: list[PhotoImage] = [ImageTk.PhotoImage(Image.fromarray(gif_reader.get_data(i)))
                                for i in range(frame_cnt)]

    # Configuração do áudio
    pygame.mixer.init()
    audio_path: str = "media/gandalf_shallnotpass.mp3"
    pygame.mixer.music.load(audio_path)

    def update(ind):
        """Update dos frames"""
        frame: PhotoImage = frames[ind]
        label.configure(image=frame)
        ind += 1

        if ind == 1:  # Inicia o áudio quando o primeiro quadro é exibido
            pygame.mixer.music.play()

        if ind == frame_cnt:
            ind = 0
            pygame.mixer.music.rewind()  # Reinicia o áudio no final da animação

        root.after(100, update, ind)

    label: tk.Label = tk.Label(root, text="Você não vai passar!",
                     background="black",
                     )
    label.pack()

    root.after(0, update, 0)
    root.title("Você não vai passar!")
    root.mainloop()


def you_cant_pass(mode):
    """Function do bloqueio"""
    if mode == "gif":
        gif()
    if mode == "video":
        video()
