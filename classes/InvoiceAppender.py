import tkinter as tk
import time
import PyPDF2
import glob
import datetime
import pandas as pd
import os



class InvoiceAppender(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.main_label = tk.Label(self.parent, text='APPEND INVS!')





        self.parent.main_label.pack()
        