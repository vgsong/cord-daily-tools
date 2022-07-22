import tkinter as tk
from classes.Atalhos import Atalhos
from classes.PyOutlook import PyOutlook
from classes.PDFReader import PDFReader
from classes.uweisti import Uweisti
from classes.DTek import DTek
from classes.InvoiceAppender import InvoiceAppender


class MainGui(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.geometry('300x1000')
        self.parent.title('工作機械 - MAIN TOOLS')
        self.parent['bg'] = '#7d7d7d'
        # self.atalhos = Atalhos(self)
        self.pyout = PyOutlook(self)
        self.pdfreader = PDFReader(self)
        self.dtek = DTek(self)
        # self.invappender = InvoiceAppender(self)
        # self.uweisti = Uweisti(self)
        self.parent.buttonQuit = tk.Button(self.parent, text='QUIT', width = 25, command=self.parent.quit, bg='#C33838')
        # self.atalhos.pack()

        self.pyout.pack()
        self.pdfreader.pack()
        self.dtek.pack()
        # self.uweisti.pack()
        # self.invappender.pack()
        self.parent.buttonQuit.pack()



if __name__ == '__main__':
    root = tk.Tk()
    MainGui(root).pack()
    root.mainloop()
