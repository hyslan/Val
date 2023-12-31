'''Módulo de gif de bloqueio do sistema'''
import tkinter as tk
from PIL import Image, ImageTk
import imageio
import pygame


def you_cant_pass():
    '''Function do bloqueio'''
    root = tk.Tk()

    # Configuração do GIF
    gif_path = 'gandalf.gif'
    gif = imageio.get_reader(gif_path)
    frameCnt = gif.get_length()
    frames = [ImageTk.PhotoImage(Image.fromarray(gif.get_data(i)))
              for i in range(frameCnt)]

    # Configuração do áudio
    pygame.mixer.init()
    audio_path = 'gandalf_shallnotpass.mp3'
    pygame.mixer.music.load(audio_path)

    def update(ind):
        '''Update dos frames'''
        frame = frames[ind]
        label.configure(image=frame)
        ind += 1

        if ind == 1:  # Inicia o áudio quando o primeiro quadro é exibido
            pygame.mixer.music.play()

        if ind == frameCnt:
            ind = 0
            pygame.mixer.music.rewind()  # Reinicia o áudio no final da animação

        root.after(100, update, ind)

    label = tk.Label(root)
    label.pack()

    root.after(0, update, 0)
    root.mainloop()
