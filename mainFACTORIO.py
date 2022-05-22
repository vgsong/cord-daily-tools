import tkinter as tk
from classes.Atalhos import Atalhos
from classes.PyOutlook import PyOutlook
from classes.PDFReader import PDFReader


class MainGui(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.geometry('250x550')
        self.parent.title('MAIN TOOLS')
        self.parent['bg'] = '#7d7d7d'
        self.atalhos = Atalhos(self)
        self.pyout = PyOutlook(self)
        self.pdfreader = PDFReader(self)

        self.atalhos.pack()
        self.pyout.pack()
        self.pdfreader.pack()


if __name__ == '__main__':
    root = tk.Tk()
    MainGui(root).pack(side='top')
    root.mainloop()
