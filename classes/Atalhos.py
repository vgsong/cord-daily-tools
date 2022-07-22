import tkinter as tk
from os import startfile
from subprocess import Popen


class Atalhos(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent['bg'] = '#7d7d7d'

        self.parent.buttonQuit = tk.Button(self.parent, text='QUIT', width = 25, command=self.parent.quit, bg='#C33838')
        self.parent.buttonInvControl = tk.Button(self.parent, text='INV CONTROL', width=25, command=self.open_control,
                                                 bg='#00994C')
        self.parent.buttonInvFolder = tk.Button(self.parent, text='INV DIR', width=25, command=self.open_invfolder)
        self.parent.buttonReaderFolder = tk.Button(self.parent, text='PDFREADER DIR', width=25, command=self.open_readerfolder)
        self.parent.buttonFromHost = tk.Button(self.parent, text='FROMHOST', width=25,command=self.open_fromhost)
        # self.parent.buttonAdHoc = tk.Button(self.parent, text='ADHOC DIR', width=25, command=self.open_adhoc)
        self.parent.buttonDownloads = tk.Button(self.parent, text='DOWNLOADS', width=25, command=self.open_downloads)
        self.parent.buttonBoxFinance = tk.Button(self.parent, text='BOXFINANCE', width=25, command=self.open_boxfinance)
        self.parent.buttonProjList = tk.Button(self.parent, text='PROJLIST', width=25, command=self.open_projlist,
                                               bg='#004C99')
        self.parent.buttonEmpList = tk.Button(self.parent, text='EMPLIST', width=25, command=self.open_emplist,
                                              bg='#FF8000')
        self.parent.main_title = tk.Label(self.parent, text='早道', bg='#7d7d7d')

        self.pack_list = [
            self.parent.main_title,
            self.parent.buttonInvControl,
            self.parent.buttonInvFolder,
            self.parent.buttonReaderFolder,
            self.parent.buttonFromHost,
            self.parent.buttonDownloads,
            self.parent.buttonBoxFinance,
            self.parent.buttonProjList,
            self.parent.buttonEmpList,
            self.parent.buttonQuit
        ]

        # self.parent.buttonDeltek = tk.Button(self.parent, text='DELTEK', width=25, command=self.open_deltek)

        self.excelexe_fpath = r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
        self.fromhost_fpath = r'C:\Users\V Song\Documents\FromHost'

        self.pack_all_buttons()

    def pack_all_buttons(self):
        for x in range(0, len(self.pack_list)):
            self.pack_list[x].pack()
        return None


    def open_control(self):
        print('Opening INVCONTROL file')
        return startfile(fr'{self.fromhost_fpath}\1INVFROMDTEK\_INV_LOGGER_MASTER_VS.xlsm')

    def open_invfolder(self):
        print('Opening INV DIR')
        return Popen(['explorer', fr'{self.fromhost_fpath}\1INVFROMDTEK\INV'])

    def open_readerfolder(self):
        print('Opening PDFREADER DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\PyP\\pdfReader\\_INVOICES"])

    def open_fromhost(self):
        print('Opening FromHost DIR')
        return Popen(['explorer', fr'{self.fromhost_fpath}'])

    def open_downloads(self):
        print('Opening Downloads DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Downloads"])

    def open_boxfinance(self):
        print('Opening BOXFIN DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Box\\Invoicing\\_FINANCE"])

    def open_projlist(self):
        # uses an exe file to open a file in dir
        print('Opening PROJ List')
        return Popen([self.excelexe_fpath, fr'{self.fromhost_fpath}\_FROMOUTLOOK\PROJECT_LIST_EXPORT.txt'])

    def open_emplist(self):
        print('Opening Emp List')
        return Popen([self.excelexe_fpath, fr'{self.fromhost_fpath}\_FROMOUTLOOK\EMPLIST_DETAIL.txt'])

    def open_adhoc(self):
        print('Opening ADHOC DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Documents\\FromHost\\_ADHOC"])