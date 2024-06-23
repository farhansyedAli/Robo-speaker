import win32com.client
import tkinter as tk
from PIL import Image, ImageTk, ImageSequence
from itertools import cycle


def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)


def on_speak(event=None):
    text = entry.get()
    if text.lower() == "q":
        speak("Good Bye Friend.")
        root.destroy()
    else:
        speak(text)
        entry.delete(0, tk.END)


def animate_gif():
    global current_frame
    try:
        current_frame = next(frames)
        gif_label.config(image=current_frame)
        root.after(100, animate_gif)  # Adjust the delay of animation
    except StopIteration:
        pass


def q():
    speak("Good bye friend")
    root.destroy()


root = tk.Tk()
root.title("Robo Speaker 1.1")

root.geometry("600x900")
root.configure(bg="black")

title_label = tk.Label(root, text="Robo Speaker", font=("Helvetica", 20, "bold"), fg="cyan", bg="black")
title_label.pack(pady=10)

gif_path = "robot.webp"
gif_image = Image.open(gif_path)

gif_frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(gif_image)]
frames = cycle(gif_frames)

gif_label = tk.Label(root, bg="black")
gif_label.pack(pady=10)

current_frame = next(frames)
root.after(0, animate_gif)

entry_label = tk.Label(root, text="Enter what you want to speak:", font=("Helvetica", 12), fg="white", bg="black")
entry_label.pack(pady=5)

entry = tk.Entry(root, width=50, font=("Helvetica", 12))
entry.pack(pady=10)
entry.bind("<Return>", on_speak)

quit_button = tk.Button(root, text="Quit", command=q, font=("Helvetica", 12), bg="red", fg="white")
quit_button.pack(pady=10)

root.mainloop()
