import tkinter as tk
from os import startfile
from subprocess import Popen


class Atalhos(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent['bg'] = '#7d7d7d'
        self.parent.buttonQuit = tk.Button(self.parent, text='QUIT', width = 25, command=self.parent.quit)
        self.parent.buttonInvControl = tk.Button(self.parent, text='INV CONTROL', width=25, command=self.open_control)
        self.parent.buttonInvFolder = tk.Button(self.parent, text='INV DIR', width=25, command=self.open_invfolder)
        self.parent.buttonReaderFolder = tk.Button(self.parent, text='PDFREADER DIR', width=25, command=self.open_readerfolder)
        self.parent.buttonFromHost = tk.Button(self.parent, text='FROMHOST', width=25,command=self.open_fromhost)
        self.parent.buttonDownloads = tk.Button(self.parent, text='DOWNLOADS', width=25, command=self.open_downloads)
        self.parent.buttonProjList = tk.Button(self.parent, text='PROJLIST', width=25, command=self.open_projlist)
        self.parent.buttonEmpList = tk.Button(self.parent, text='EMPLIST', width=25, command=self.open_emplist)

        # self.parent.buttonDeltek = tk.Button(self.parent, text='DELTEK', width=25, command=self.open_deltek)
        self.parent.main_title = tk.Label(self.parent, text='quickaccess', bg='#7d7d7d')

        self.parent.main_title.pack()

        self.parent.buttonInvControl.pack()
        self.parent.buttonInvFolder.pack()
        self.parent.buttonReaderFolder.pack()
        self.parent.buttonFromHost.pack()
        self.parent.buttonDownloads.pack()
        self.parent.buttonProjList.pack()
        self.parent.buttonEmpList.pack()
        # self.parent.buttonDeltek.pack()
        self.parent.buttonQuit.pack()

    def open_control(self):
        print('Opening INVCONTROL file')
        return startfile("C:\\Users\\V Song\\Documents\\FromHost\\1INVFROMDTEK\\_INV_LOGGER_MASTER_VS.xlsm")

    def open_invfolder(self):
        print('Opening INV DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Documents\\FromHost\\1INVFROMDTEK\\INV"])

    def open_readerfolder(self):
        print('Opening PDFREADER DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\PyP\\pdfReader\\_INVOICES"])

    def open_fromhost(self):
        print('Opening FromHost DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Documents\\FromHost"])

    def open_downloads(self):
        print('Opening Downloads DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Downloads"])

    def open_projlist(self):
        # uses an exe file to open a file in dir
        print('Opening PROJ List')
        return Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\PROJECT_LIST_EXPORT.txt'])

    def open_emplist(self):
        print('Opening Emp List')
        return Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\EMPLIST_DETAIL.txt'])
