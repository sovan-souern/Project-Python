import os
import tkinter as tk

Mywindow = tk.Tk()
Mywindow.title("Shutdown pc")
Mywindow.geometry("500x500")
frame = tk.Frame(Mywindow, width=500, height=500)
frame.pack()
def shutdown_computer():
    os.system("shutdown /s /t 1") 

def restart_computer():
    os.system("shutdown /r /t 1")
    


Button_shutdown = tk.Button(frame, text="Shutdown", font=("Arail",10), command=shutdown_computer, bg="red",fg="white")
Button_shutdown.place(x=100,y=100, width=100, height=50)

Button_restart = tk.Button(frame, text="Restart", font=("Arail",10), command=restart_computer, bg="green",fg="white")
Button_restart.place(x=300,y=100, width=100, height=50)


Mywindow.mainloop()