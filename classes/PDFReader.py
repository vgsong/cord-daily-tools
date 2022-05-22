import tkinter as tk

import PyPDF2

import re
import os

from glob import glob
from csv import writer
from datetime import datetime, timedelta
from shutil import copyfile


class PDFReader(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.main_title = tk.Label(self.parent, text='PDF Reader', bg='#7d7d7d')
        self.parent.buttonPdfOut = tk.Button(self.parent, width=25, text='OPEN PDFOUT', command=self.open_pdfout)
        self.parent.buttonFinal = tk.Button(self.parent, width=25, text='RUN FINAL', command=self.run_final)
        self.parent.buttonDraft = tk.Button(self.parent, width=25, text='RUN DRAFT', command=self.run_draft)
        self.parent.main_title.pack()
        self.parent.buttonPdfOut.pack()
        self.parent.buttonDraft.pack()
        self.parent.buttonFinal.pack()

    def open_pdfout(self):
        return os.startfile("C:\\Users\\V Song\\PyP\\pdfReader\\_BATCHSUMMARY\\pdfReader_OUTPUT.csv")

    def run_final(self):
        self.start_pdfReader()

    def run_draft(self):
        self.start_pdfReader('DRAFT')

    def start_pdfReader(self, astage='FINAL'):

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
        FPATH = r'C:\Users\V Song\PyP\pdfReader\_INVOICES\_INPUT'

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

            # retired regex to find the total amount in the invoice, the regex was not a catch all
            # totalFound = re.sub('\nTotal this Invoice','',re.search(totalAmountPattern,fStr).group(0))
            # totalFound = re.sub('Total this Invoice','',totalFound)

            # instead I noted a pattern that it is always the highest number of all totals that has a dollar ($) sign
            # so the regex captures every string that has a dollar sign and returns the one with the highest value

            totalFound = re.findall('\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)
            totalFound += re.findall('Total Billings\n\$([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)
            totalFound += re.findall('Total Billings\n([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)
            totalFound += re.findall('Totals\n([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])', fStr)

            remove_single_quote(totalFound)
            final_list = join_total_lists(totalFound)
            total_amount = max(final_list)

            # if invnumber found is draft then the filename convention has to change with a counter from loop
            if invFound == 'DRAFT no number':
                copyfile(fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\_INPUT\{fname}',
                         fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\{invCount} (DRAFT).pdf')
                invCount += 1
            else:
                if astage == 'FINAL':
                    copyfile(fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\_INPUT\{fname}',
                             fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\{invFound}.pdf')
                elif astage == 'DRAFT':
                    copyfile(fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\_INPUT\{fname}',
                             fr'C:\Users\V Song\PyP\pdfReader\_INVOICES\{invFound} (DRAFT).pdf')

            projList.append(projFound)
            invList.append(invFound)
            totalList.append(total_amount)
            timeList.append(timeStamp)

            sample_pdf.close()

        # decided the hardcode the filename, maybe have something to append into a database
        csvFname = f'pdfReader_OUTPUT.csv'

        with open(fr'C:\Users\V Song\PyP\pdfReader\_BATCHSUMMARY\{csvFname}', 'w', newline='') as f:
            csv_writer = writer(f)
            for p, to, i, ti in zip(projList, totalList, invList, timeList):
                csv_writer.writerow([p, to, i, ti])

        for invoice in os.scandir(FPATH):
            os.remove(invoice.path)

        print(f'Success!')
        print(rf'CSV OUTPUT SAVED AT: C:\Users\V Song\PyP\pdfReader\_BATCHSUMMARY\{csvFname}')

        return