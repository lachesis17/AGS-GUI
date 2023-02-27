import tkinter as tk
from PIL import ImageTk, Image 
import customtkinter as ct

# after_cancel tasks: https://stackoverflow.com/questions/63628566/how-to-handle-invalid-command-name-error-while-executing-after-script-in-tk

class SplashWin(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        if screen_height == 1440:
            wid = int(screen_width / 2.5)
            hei = int(screen_height / 2.35)
        elif screen_height == 2160:
            wid = int(screen_width / 2)
            hei = int(screen_height / 1.9)
        elif screen_height == 1080:
            wid = int(screen_width / 2.75)
            hei = int(screen_height / 2.6)
        else:
            wid = int(screen_width / 3)
            hei = int(screen_height / 3)

        self.geometry(f"1000x1000+{wid}+{hei}")

        self.wm_overrideredirect(True)
        self.lift()
        self.config(bg = '#141a23')
        self.wm_attributes("-transparentcolor", "#141a23")
        splash_logo = Image.open('common/images/GQ_AGS.PNG')
        splash_img = ImageTk.PhotoImage(splash_logo)
        splash_label = tk.Label(image=splash_img,bg="#141a23")
        splash_label.image = splash_img
        splash_label.place(x=0,y=0)
        self.resizable(False,False)

        self.splash_init()

    def splash_init(self):
        self.after(3000,lambda:self.del_splash())
        self.attributes("-topmost", True)
        
    def del_splash(self):
        for after_id in self.tk.eval('after info').split():
            self.after_cancel(after_id)
        self.destroy()

if __name__ == '__main__' or __name__ == 'common.assets.splash':
    splash = SplashWin()
    splash.mainloop()
