import win32com.client
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk, ImageSequence
from itertools import cycle


def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    voices = speaker.GetVoices()

    speaker.Voice = voices.Item(1)
    speaker.Speak(text)


def on_speak(event=None):
    text = entry.get()
    speak(text)
    entry.delete(0, tk.END)


def animate_gif():
    global current_frame
    try:
        current_frame = next(frames)
        gif_label.config(image=current_frame)
        root.after(100, animate_gif)
    except StopIteration:
        pass


def q():
    root.destroy()


root = tk.Tk()
root.title("Robo Speaker 1.1")
root.geometry("600x800")
root.configure(bg="#1a1a2e")

title_frame = tk.Frame(root, bg="#16213e", padx=10, pady=10)
title_frame.pack(fill=tk.X)

title_label = tk.Label(
    title_frame, text="🤖 Robo Speaker", font=("Helvetica", 28, "bold"), fg="#e94560", bg="#16213e"
)
title_label.pack()

animation_frame = tk.Frame(root, bg="#1a1a2e")
animation_frame.pack(pady=10)

gif_image = Image.open("robot.webp")
gif_frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(gif_image)]
frames = cycle(gif_frames)

gif_label = tk.Label(animation_frame, bg="#1a1a2e")
gif_label.pack()
current_frame = next(frames)
root.after(0, animate_gif)

input_frame = tk.Frame(root, bg="#1a1a2e", pady=20)
input_frame.pack()

entry_label = tk.Label(
    input_frame, text="Enter text for Robo to speak:", font=("Helvetica", 16), fg="#f5f6fa", bg="#1a1a2e"
)
entry_label.pack()

entry = ttk.Entry(input_frame, font=("Helvetica", 14), width=40)
entry.pack(pady=10)
entry.bind("<Return>", on_speak)

button_frame = tk.Frame(root, bg="#1a1a2e")
button_frame.pack(pady=20)

speak_button = tk.Button(
    button_frame, text="🎤 Speak", command=on_speak, font=("Helvetica", 14), bg="#0f3460", fg="#f5f6fa", padx=20, pady=10
)
speak_button.grid(row=0, column=0, padx=10)

quit_button = tk.Button(
    button_frame, text="❌ Quit", command=q, font=("Helvetica", 14), bg="#e94560", fg="#f5f6fa", padx=20, pady=10
)
quit_button.grid(row=0, column=1, padx=10)

footer_label = tk.Label(
    root, text="Developed by Farhan Ali", font=("Helvetica", 10), fg="#7f8c8d", bg="#1a1a2e", pady=10
)
footer_label.pack(side=tk.BOTTOM)

root.mainloop()
