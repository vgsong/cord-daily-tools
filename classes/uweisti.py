import subprocess
import time
import tkinter as tk
import pyautogui as pygui
from os import startfile
from subprocess import Popen



class Uweisti(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent['bg'] = '#7d7d7d'
        self.parent.buttonMain = tk.Button(self.parent, text='UWEISTI', width = 25, command=self.startpygui)

        # self.parent.buttonDeltek = tk.Button(self.parent, text='DELTEK', width=25, command=self.open_deltek)
        self.parent.main_title = tk.Label(self.parent, text='????????', bg='#7d7d7d')

        self.parent.main_title.pack()
        self.parent.buttonMain.pack()



    def startpygui(self):
        print('Opening BOXFIN DIR')
        a = Popen(['explorer', "C:\\Users\\V Song\\Box\\Invoicing\\_FINANCE"])
        time.sleep(3)
        print(a.communicate())

        return None
