import os
import sys
import PyPDF2
import re

from PyQt5.QtWidgets import *
from PyQt5 import QtCore

from subprocess import Popen
from datetime import datetime
from shutil import copyfile
from time import sleep
from glob import glob
from csv import writer
from os import startfile

import pandas as pd


class Window(QWidget):
    def __init__(self, parent=None):
        def pack_atalhos():
            for button, cb in zip(self.atalhos_button_obj, self.atalhos_callback_list):
                button.setStyleSheet('height: 2em')
                self.atalhos_layout.addWidget(button)
                button.clicked.connect(cb)

        def pack_pdfreader():
            for button, cb in zip(self.pdfreader_button_obj, self.pdfreader_callback_list):
                button.setStyleSheet('height: 2em')
                self.pdfreader_layout.addWidget(button)
                button.clicked.connect(cb)

            # packing radio button for final/draft switch for pdfreader
            self.pdfreader_true_button = QRadioButton('FINAL')
            self.pdfreader_false_button = QRadioButton('DRAFT')
            self.pdfreader_layout.addWidget(self.pdfreader_true_button)
            self.pdfreader_layout.addWidget(self.pdfreader_false_button)

        def pack_itd():
            for button, cb in zip(self.itd_button_obj, self.itd_callback_list):
                button.setStyleSheet('height: 2em')
                self.itd_layout.addWidget(button)
                button.clicked.connect(cb)

        def pack_outlook():
            for button, cb in zip(self.outlook_button_obj, self.outlook_callback_list):
                button.setStyleSheet('height: 2em')
                self.outlook_layout.addWidget(button)
                button.clicked.connect(cb)

        # ---------------------------------------------------------------------------------------------------------- #
        # parent layout settings
        super(Window, self).__init__(parent)
        self.resize(300, 400)
        self.setWindowTitle('MAIN TOOLS')

        # ---------------------------------------------------------------------------------------------------------- #
        # MAIN PATHS
        self.excelexe_fpath = r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
        self.fromhost_fpath = r'C:\Users\V Song\Documents\FromHost'

        self.pdfreader_main_fpath = r'C:\Users\V Song\PyP\pdfReader\_INVOICES'
        self.pdfreader_pdfout_fpath = r'C:\Users\V Song\PyP\pdfReader\_BATCHSUMMARY'

        self.dl_fpath = r'C:\Users\V Song\Downloads'
        self.pdfreader_input_fpath = r'C:\Users\V Song\PyP\pdfReader\_INVOICES\_INPUT'
        self.temp_fpath = r'C:\Users\V Song\Documents\FromHost\TEMP'

        self.itd_pomaster_fpath = r'C:\Users\V Song\OneDrive - Cordoba Corp\SHARED'
        self.itd_mastercontrol_fpath = r'C:\Users\V Song\Box\Invoicing\_FINANCE\ITD_VS.csv'

        # ---------------------------------------------------------------------------------------------------------- #
        # ---------------------------------------------------------------------------------------------------------- #
        # parent layout :
        self.horizont_layout = QHBoxLayout()
        # self.horizont_layout.addWidget(QLabel(f'Today is {datetime.now()}'))

        # ---------------------------------------------------------------------------------------------------------- #
        # horizont_layout > child atalhos section
        self.atalhos_layout = QVBoxLayout()
        self.atalhos_layout.addWidget(QLabel('早道'))
        self.atalhos_layout.setSpacing(5)
        self.atalhos_layout.setContentsMargins(0, 0, 0, 0)

        self.atalhos_button_list = (
            'INV CONTROL',
            'INVDIR',
            'PDFReader DIR',
            'FromHost DIR',
            'DOWNLOAD DIR',
            'BOX DIR',
            'PROJLIST',
            'EMPLIST',
            'QUIT'
        )

        self.atalhos_callback_list = (
            self.atalhos_open_control,
            self.atalhos_open_invfolder,
            self.atalhos_open_readerfolder,
            self.atalhos_open_fromhost,
            self.atalhos_open_downloads,
            self.atalhos_open_boxfinance,
            self.atalhos_open_projlist,
            self.atalhos_open_emplist,
            self.atalhos_button_quit
        )

        # ---------------------------------------------------------------------------------------------------------- #
        # horizont_layout > child pdfreader section
        self.pdfreader_layout = QVBoxLayout()
        self.pdfreader_layout.addWidget(QLabel('PDFReader Tools'))

        self.pdfreader_button_list = (
            'OPEN PDFOUT',
            'SEND PDF TO INPUT',
            'PURGE DL/TEMP DIR',
            'RUN PDFREADER',
            'PDFOUT to CLIPBOARD'
        )

        self.pdfreader_callback_list = (
            self.pdfreader_open_pdfout,
            self.pdfreader_move_pdf,
            self.pdfreader_clean_dltemp,
            self.pdfreader_run,
            self.pdfreader_toclipboard
        )

        # ---------------------------------------------------------------------------------------------------------- #
        # horizont_layout > child ITD TOOLS section
        self.itd_layout = QVBoxLayout()
        self.itd_layout.addWidget(QLabel('ITD Tools'))

        self.itd_button_list = (
            'PRINT ITD',

        )

        self.itd_callback_list = (
            self.itd_get,

        )
        # ---------------------------------------------------------------------------------------------------------- #
        # horizont_layout > child outlook section
        self.outlook_layout = QVBoxLayout()
        self.outlook_layout.addWidget(QLabel('OUTL Tools'))

        self.outlook_button_list = (
            'PRINT ITD',

        )

        self.outlook_callback_list = (
            self.itd_get,

        )
        # ---------------------------------------------------------------------------------------------------------- #
        # ---------------------------------------------------------------------------------------------------------- #
        # creates button object based on self.button_list
        self.atalhos_button_obj = [QPushButton(name) for name in self.atalhos_button_list]
        self.pdfreader_button_obj = [QPushButton(name) for name in self.pdfreader_button_list]
        self.itd_button_obj = [QPushButton(name) for name in self.itd_button_list]
        self.outlook_button_obj = [QPushButton(name) for name in self.outlook_button_list]

        pack_pdfreader()
        pack_atalhos()
        pack_itd()
        pack_outlook()

        # ---------------------------------------------------------------------------------------------------------- #
        # layout structure
        # self  parent
        # atalhos and pdfreader rolls into horizont layout
        self.horizont_layout.addLayout(self.atalhos_layout)
        self.horizont_layout.addLayout(self.pdfreader_layout)
        self.horizont_layout.addLayout(self.itd_layout)

        self.setLayout(self.horizont_layout)

    # ---------------------------------------------------------------------------------------------------------- #
    # atalhos callbacks
    def atalhos_open_control(self):
        print('Opening INVCONTROL file')
        return startfile(fr'{self.fromhost_fpath}\1INVFROMDTEK\_INV_LOGGER_MASTER_VS.xlsm')

    def atalhos_open_invfolder(self):
        print('Opening INV DIR')
        return Popen(['explorer', fr'{self.fromhost_fpath}\1INVFROMDTEK\INV'])

    def atalhos_open_readerfolder(self):
        print('Opening PDFREADER DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\PyP\\pdfReader\\_INVOICES"])

    def atalhos_open_fromhost(self):
        print('Opening FromHost DIR')
        return Popen(['explorer', fr'{self.fromhost_fpath}'])

    def atalhos_open_downloads(self):
        print('Opening Downloads DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Downloads"])

    def atalhos_open_boxfinance(self):
        print('Opening BOXFIN DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Box\\Invoicing\\_FINANCE"])

    def atalhos_open_projlist(self):
        # uses an exe file to open a file in dir
        print('Opening PROJ List')
        return Popen([self.excelexe_fpath, fr'{self.fromhost_fpath}\_FROMOUTLOOK\PROJECT_LIST_EXPORT.txt'])

    def atalhos_open_emplist(self):
        print('Opening Emp List')
        return Popen([self.excelexe_fpath, fr'{self.fromhost_fpath}\_FROMOUTLOOK\EMPLIST_DETAIL.txt'])

    def atalhos_button_quit(self):
        return QtCore.QCoreApplication.instance().quit()

    # ---------------------------------------------------------------------------------------------------------- #
    # pdf reader callbacks
    def pdfreader_open_pdfout(self):
        print('Opening PDFOUT file')
        return os.startfile(fr'{self.pdfreader_pdfout_fpath}\pdfReader_OUTPUT.csv')

    def pdfreader_move_pdf(self):

        # for d in zip(glob(os.path.join(self.dl_fpath, '*')), glob(os.path.join(self.temp_fpath, '*'))):
        #     print(d[0])
        #     print(d[1])
        #
        for i in glob(os.path.join(self.dl_fpath, '*')):
            fname, fext = os.path.splitext(os.path.basename(i))
            if fext == '.pdf':
                copyfile(i, fr'{self.pdfreader_input_fpath}\{fname}{fext}')
                print(f'file {fname} copied to input folder')
                sleep(0.2)

        for i in glob(os.path.join(self.temp_fpath, '*')):
            fname, fext = os.path.splitext(os.path.basename(i))
            if fext == '.pdf':
                copyfile(i, fr'{self.pdfreader_input_fpath}\{fname}{fext}')
                print(f'file {fname} copied to input folder')
                sleep(0.2)

        sleep(0.1)
        print('Job Done!')
        return None

    def pdfreader_clean_dltemp(self):

        print('Cleaning temp and dl dir...')

        for x in os.listdir(self.dl_fpath):
            os.remove(os.path.join(self.dl_fpath, x))

        for y in os.listdir(self.temp_fpath):
            os.remove(os.path.join(self.temp_fpath, y))

        print('Job Done!')
        return None

    def pdfreader_run(self):

        def pdf_to_str(apdf):
            # gets opened pdf obj from passed argument
            # uses PyPDF2 to parse strings
            # and loop pdfdoc into a reader to get str values
            # returns result
            pdfdoc = PyPDF2.PdfFileReader(apdf)
            result = ''

            for i in range(pdfdoc.numPages // 2):
                current_page = pdfdoc.getPage(i)
                result += current_page.extractText()

            return result

        def remove_single_quote(alist):
            # findall returns a list of tuples
            # the list comphrehension below transforms the tuples into lists
            # the with a nested loop I count how many times I need to use the list method
            # list.remove() to remove the empty values in each list
            if len(totalFound) == 0:
                alist = 'no amt found'
                return alist
            else:
                alist = [list(x) for x in totalFound]
                for x in alist:
                    for y in range(x.count('')):
                        x.remove('')
                return alist

        def join_total_lists(alist):
            # after remove_single_quote the list is passed into join_total_lists
            # the transformed list come as multiple strings inside list
            # the function concatenate each value inside list
            # then convert str to float
            temp_list = []
            result = []

            temp_list.append([''.join(x) for x in alist])
            result = temp_list.pop()  # pop() used because temp_list output is [[]]
            result = [float(x.replace(',', '')) for x in result]
            return result

        print(f'Initializing pdfReader...')

        # fpath CONST where invoices are dropped
        FPATH = self.pdfreader_input_fpath

        # List of values obtained from start_pdfReader
        totalList = []
        invList = []
        projList = []
        timeList = []
        invCount = 0

        # RegEx patterns used when reading the invoice in pdf file
        projPattern = '\d{4}.\d{3}.\d{3}|\d{4}.\d{3}.00R|\d{4}.00E.\d{3}'
        invPattern = 'Invoice No:\s*\d{7}'
        perPattern = 'Professional Services from.*202\d'

        # uncomment below if the new regex pattern do not work
        # totalAmountPattern = 'Total this Invoice\n\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])|\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])\nTotal this Invoice|\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])Total this Invoice'
        # totalAmountPattern = '\n\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])|\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])\n'

        # totalAmountPattern = '1115'

        # Setting up the os.chdir into FPATH CONST, where glob generator is created
        os.chdir(FPATH)

        # used to know when Invoice was created/pdfREADER ran
        timeStamp = datetime.now().strftime('%y%m%d')

        # Loops from os.getcwd fpath, which contain the invoices to be processed into excel
        for fname in glob('*.pdf'):

            sample_pdf = open(fname, mode='rb')
            fStr = pdf_to_str(sample_pdf)

            # ----------------------------uncomment this for debugging------------------------------
            print(fStr)
            print(fname)

            # regex search for proj number
            projFound = re.search(projPattern, fStr).group(0)
            periodFound = re.search(perPattern, fStr).group(0)

            # regex for invoice number
            try:
                invFound = re.search(invPattern, fStr).group(0)[-7:]
            except:
                invFound = 'DRAFT no number'

            print(invFound)
            print(periodFound)

            # retired regex to find the total amount in the invoice, the regex used below was not a catch all:
            # totalFound = re.sub('\nTotal this Invoice','',re.search(totalAmountPattern,fStr).group(0))
            # totalFound = re.sub('Total this Invoice','',totalFound)

            # However, I noted a pattern that it is always the highest number of all totals that has a dollar ($) sign
            # so the regex captures every string that has a dollar sign and returns the one with the highest value

            totalFound = re.findall('\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)
            totalFound += re.findall('Total Billings\n\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])',
                                     fStr)
            totalFound += re.findall('Total Billings\n([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)
            totalFound += re.findall('Totals\n([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)

            remove_single_quote(totalFound)
            final_list = join_total_lists(totalFound)
            total_amount = max(final_list)

            # if invnumber found is draft then the filename convention has to change with a counter from loop
            if invFound == 'DRAFT no number':
                copyfile(fr'{self.pdfreader_main_fpath}\_INPUT\{fname}',
                         fr'{self.pdfreader_main_fpath}\{invCount} (DRAFT).pdf')
                invCount += 1
            else:
                if self.pdfreader_true_button.isChecked():
                    copyfile(fr'{self.pdfreader_main_fpath}\_INPUT\{fname}',
                             fr'{self.pdfreader_main_fpath}\{invFound}.pdf')
                elif self.pdfreader_false_button.isChecked():
                    copyfile(fr'{self.pdfreader_main_fpath}\_INPUT\{fname}',
                             fr'{self.pdfreader_main_fpath}\{invFound} (DRAFT).pdf')

            projList.append(projFound)
            invList.append(invFound)
            totalList.append(total_amount)
            timeList.append(timeStamp)

            sample_pdf.close()

        # decided the hardcode the filename, maybe have something to append into a database
        csvFname = f'pdfReader_OUTPUT.csv'

        with open(fr'{self.pdfreader_pdfout_fpath}\{csvFname}', 'w', newline='') as f:
            csv_writer = writer(f)
            for p, to, i, ti in zip(projList, totalList, invList, timeList):
                csv_writer.writerow([p, to, i, ti])

        for invoice in os.scandir(FPATH):
            os.remove(invoice.path)

        print(f'Success!')
        print(fr'CSV OUTPUT SAVED AT: {self.pdfreader_pdfout_fpath}\{csvFname}')

        return None

    def pdfreader_toclipboard(self):
        pdf_out = pd.read_csv(fr'{self.pdfreader_pdfout_fpath}\pdfReader_OUTPUT.csv', header=None)
        del pdf_out[3]
        pdf_out.to_clipboard(excel=True, index=False, header=None)
        print(f'Copied pdfOUTPUT into clipboard!\n')
        return None

    # ---------------------------------------------------------------------------------------------------------- #
    # itd callbacks
    def itd_get(self):
        def get_po_info():
            def get_latest_po_filepath():
                file_list = glob(os.path.join(self.itd_pomaster_fpath, '*'))
                max_result = max(file_list, key=os.path.getctime)
                return max_result

            latest_file_path = get_latest_po_filepath()
            raw_df = pd.read_excel(latest_file_path, skiprows=1)

            result = pd.DataFrame()
            col_index = (
                1,
                2,
                4,
                8,
                10,
                11,
                12,
                13,
                15
            )

            for x in col_index:
                result[raw_df.columns[x]] = raw_df[raw_df.columns[x]]

            return result

        def get_itd_info():
            itd_df_columns = (
                'DTEKPROJNUM',
                'INVAMT',
                'INVNUM',
                'PO',
                'M_PER',
                'Y_PER'
            )

            result = pd.read_csv(self.itd_mastercontrol_fpath, header=None, names=itd_df_columns)

            return result.drop(result.columns[[3, 4, 5]], axis=1)

        po_df = get_po_info()
        itd_df = get_itd_info()

        itd_df = itd_df.groupby(['DTEKPROJNUM'], as_index=False).sum()
        # print(po_df.loc[po_df['DTEKPROJNUM'] == '1111.000.001'])
        print(itd_df)

        return None

    # ---------------------------------------------------------------------------------------------------------- #
    def open_adhoc(self):
        print('Opening ADHOC DIR')
        return Popen(['explorer', "C:\\Users\\V Song\\Documents\\FromHost\\_ADHOC"])


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = Window()
    ex.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
