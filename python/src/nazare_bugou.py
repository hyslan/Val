"""Módulo de gif de bloqueio do sistema."""
import tkinter as tk

import imageio
from PIL import Image, ImageTk


def oxe() -> None:
    """Gif da Nazaré pra quando da erro na aplicação."""
    root = tk.Tk()

    # Configuração do GIF
    gif_path = "media/nazare.gif"
    gif = imageio.get_reader(gif_path)
    frame_cnt = gif.get_length()
    frames = [ImageTk.PhotoImage(Image.fromarray(gif.get_data(i)))
              for i in range(frame_cnt)]

    def update(ind) -> None:
        """Update dos frames."""
        frame = frames[ind]
        label.configure(image=frame)
        ind += 1

        if ind == frame_cnt:
            ind = 0

        root.after(100, update, ind)

    label = tk.Label(root, text="Ué!",
                     background="black",
                     )
    label.pack()

    root.after(0, update, 0)
    root.title("Ué!")
    root.mainloop()
