import tkinter as tk
import pandas as pd

from os import walk
from os.path import join
from win32com.client import Dispatch
from datetime import datetime, timedelta


class PyOutlook(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent['bg'] = '#7d7d7d'
        self.parent.main_title = tk.Label(self.parent, text='Py Outlook', bg='#7d7d7d')
        self.parent.buttonBotherSam = tk.Button(self.parent, text='Bother Sam', width=25, command=self.out_bothersam)
        self.parent.textProjectCode = tk.Text(self.parent, height=2, width=22)
        self.parent.buttonEleventwelveSend = tk.Button(self.parent, text='1112 PROJSEND', width=25, command=self.send_eleventwelve)
        self.parent.buttonSaveReports = tk.Button(self.parent, text='GET DTEK REPORTS', width=25, command=self.get_dtekreports)
        self.parent.buttonSendTaubot = tk.Button(self.parent, text='SEND TAUBOT', width=25, command=self.send_taubot_data)
        self.parent.buttonsZeroRates = tk.Button(self.parent, text='SEND ZERORATE', width=25, command=self.get_unbilled_details)

        self.parent.main_title.pack()
        self.parent.buttonSaveReports.pack()
        self.parent.buttonEleventwelveSend.pack()
        self.parent.buttonSendTaubot.pack()
        self.parent.buttonsZeroRates.pack()
        self.parent.buttonBotherSam.pack()
        self.parent.textProjectCode.pack()

    def out_bothersam(self):
        # Currently used just for laziness practice
        # Sends an email reminder to Sam
        # for changes, use .display instead of .send
        # independent since it does not retrieve the inbox mailObj list
        # it just creates a mail Item

        proj_var = self.parent.textProjectCode.get('1.0','end')

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)

        olmail.To = 'stenorio@cordobacorp.com'
        olmail.CC = 'lmurguia@cordobacorp.com; janel.toregozhina@cordobacorp.com; khiem.ta@cordobacorp.com'

        olmail.Subject = 'DTEK INVOICE APPROVAL'
        olmail.Body = f'Hi Sam -- \n\n' \
                      f'Can you please review/approve invoices for projects {proj_var} in Deltek?\n\n' \
                      f'Let me know if you have any questions,\n\n' \
                      f'Thank you,\n\n' \
                      f'Victor Song\n' \
                      f'Financial Analyst\n\n' \
                      f'Cordoba Corporation | Making a Difference\n' \
                      f'o: (657) 900-8857 ext.5791\n' \
                      f'victor.song@cordobacorp.com | cordobacorp.com\n' \
                      f'LinkedIn | Twitter | Facebook | YouTube | Instagram'

        olmail.display(True)
        # olmail.display

        del olmail
        del olapp

        return None

    def send_eleventwelve(self):
        # Currently used just for laziness practice
        # Sends an email reminder to Sam
        # for changes, use .display instead of .send
        # independent since it does not retrieve the inbox mailObj list
        # it just creates a mail Item

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)
        control_fpath = r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\1112CONTACT_LIST.xlsx'
        contact_sheetname = '1112DISTRLIST'
        to_distr = ''
        cc_distr = ''

        contact_df = pd.read_excel(control_fpath, sheet_name=contact_sheetname)

        for x in contact_df['DISTR']:
            to_distr += f'{x}; '

        for x in contact_df['CC'].dropna():
            cc_distr += f'{x}; '

        olmail.To = to_distr
        olmail.CC = cc_distr

        olmail.Subject = f'1112_SCG Tech Project List (RunDate {datetime.now().strftime("%Y%m%d")})'
        olmail.Body = f'Note, project list is scheduled to be distributed every month as a reference when making project requests to avoid duplicate project requests. \n\nIf there are any discrepancies in the WOA/IO/Billing Contact (SCG PM you support) on the list please submit the updated information to a Cordoba PM lead to validate and submit to timesheet@cordobacorp.com\n\n' \
                      f'Thank you,\n\n' \

        olmail.Attachments.Add(r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\1112_Project List Export.xlsx')

        olmail.display(True)

        del olmail
        del olapp

        return

    def get_dtekreports(self):
        olItems = self.get_mail('victor.song@cordobacorp.com', 12)
        self.run_mail(olItems, self.dtek_reports_saver)
        return

    def get_mail(self, amailaccount, ahours):

        # get_olinbox returns a list object with the outlook items
        ol_items = self.get_olinbox(amailaccount)

        # based on the now() - 8 hours
        result = self.set_olinbox_filters(ol_items, ahours)
        return result

    def get_olinbox(self, amail):
        print(f'Loading {amail} outlook inbox...')
        # DISPATCH Application > NameSpace > Folders (>Folders, >Folders, ...) > Items .Unread, .Subject, .Etc...
        # 'Outlook.Application' > 'MAPI' > 'first.last@company.com' > 'Inbox'

        olapp = Dispatch('Outlook.Application')
        result = olapp.GetNameSpace('MAPI').Folders(amail).Folders('Inbox').Items
        return result

    def set_olinbox_filters(self, aitems, hdiff):
        # Sets the filter for inbox by Received Time
        # uses datetime.now() minus hours (keep the scope to minimum)
        result_list = []
        received_date = datetime.now() - timedelta(hours=hdiff)
        received_date = received_date.strftime('%m/%d/%Y %H:%M %p')
        result = aitems.Restrict("[ReceivedTime] >= '" + received_date + "'")

        for item in result:
            result_list.append(item)

        return result_list

    def dtek_reports_saver(self, aitem):

        # Procedure that saves attachments for Deltek scheduled distribution
        # dict contains email {SUBJECT : Attachment}

        DTEKMAILINFO = {
            'DTEK_DAILY_Unbilled Detail and Aging Report': 'UNBILLED_DETAILS.txt',
            'DTEK_DAILY_Project List Export Report': 'PROJECT_LIST_EXPORT.txt',
            'DTEK_DAILY_Employee Export Report': 'EMPLIST_DETAIL.txt',
            'DTEK_DAILY_Unposted Labor Report': 'UNPOSTED_LABOR_DETAIL.txt',
            '1112_Project List Export Report': '1112_Project List Export.xlsx'
        }

        ATTACHFPATH = r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK'  # local drive
        # ATTACHFPATH2 = r'C:\Users\V Song\Box\Invoicing\_FINANCE\VS\_DTEKREPORTS'  # box network drive remove comment once QA is done

        if aitem.Subject in DTEKMAILINFO:

            # Error handling if I am not logged in the Finance Team Network Drive
            # it tries to save ATTACHFPATH and ATTACHFPATH2, if not logged it saves only on local

            aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH}\{DTEKMAILINFO[aitem.Subject]}')
            # aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH2}\{DTEKMAILINFO[aitem.Subject]}')
            print(f'{DTEKMAILINFO[aitem.Subject]} saved successfully')

        return None

    def send_taubot_data(self):

        inv_dir = r'C:\Users\V Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\PDFINV'

        timeNow = datetime.now().strftime('%H%m%d')

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)

        olmail.To = 'victor.song@cordobacorp.com'
        olmail.Subject = f'TAUBOT_FILES_{timeNow}'

        olmail.Attachments.Add(r'C:\Users\V Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\taupload.csv')

        for path, subfolders, files in walk(inv_dir):
            for names in files:
                olmail.Attachments.Add(join(path, names))

        olmail.Body = f'Please see attached -- \n\n'

        olmail.display(True)

        del olmail
        del olapp

        return None

    def get_unbilled_details(self):
        # fname is the attachment for unbilled details
        fname = 'UNBILLED_DETAILS.txt'

        # fpath of where deltek reports are saved
        fpath = rf'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\{fname}'

        # dict field mapping key (final_report): value (csv)
        unbilled_names = {
            'PROJNUM': 'groupHeader1_GroupColumn',
            'LABORTYPE': 'groupHeader4_GroupColumn',
            'LABORDATE': 'detail_transDate',
            'LABORCODE': 'detail_laborCode',
            'EMPNAME': 'detail_Description',
            'RATE': 'detail_billRate'
        }

        # load deltek report csv file into csv_df
        csv_df = pd.read_csv(fpath, skiprows=3)

        # create pandas dataframe field (cols) using the keys from unbilled_names with empty rows
        final_df = pd.DataFrame(columns=list(unbilled_names.keys()))

        def itemize_dict():

            # moves data from csv to final_df based on unbilled_names key: value mapping
            for k, v in unbilled_names.items():
                final_df[k] = csv_df[v]
            return

        def print_result():
            if final_df.empty:
                print('no Zero Rates! Goo goo g\'joob!')
                return
            else:
                print(f'Unbilled with zero Rates are:\n')
                print(final_df)
                print(f'\nPlease fix it ASAP! お願いします')
                return

        def send_distr():
            # Currently used just for laziness practice
            # Sends an email reminder to Sam
            # for changes, use .display instead of .send
            # independent since it does not retrieve the inbox mailObj list
            # it just creates a mail Item
            olapp = Dispatch('Outlook.Application')
            olmail = olapp.CreateItem(0)

            olmail.To = 'victor.song@cordobacorp.com '
            olmail.CC = 'janel.toregozhina@cordobacorp.com; Khiem.Ta@cordobacorp.com;'

            olmail.Subject = 'DAILY ZERO RATE REPORT'
            olmail.Body = f'Hi -- \n\n' \
                          f'Please see attached -- summary is below:\n\n' \
                          f'{email_body.drop_duplicates(subset="PROJNUM")}\n\n' \
                          f'Let me know if you have any questions,\n\n' \
                          f'Thank you,\n\n' \
                          f'Victor Song\n' \
                          f'Financial Analyst\n\n' \
                          f'Cordoba Corporation | Making a Difference\n' \
                          f'o: (949) 659-2717\n' \
                          f'victor.song@cordobacorp.com | cordobacorp.com\n' \
                          f'LinkedIn | Twitter | Facebook | YouTube | Instagram'

            olmail.Attachments.Add(r'C:\Users\V Song\Documents\FromHost\_OUTBOUND\ZERORATES\DTEK_DAILY_ZERORATES.csv')

            olmail.display(True)

            return

        itemize_dict()

        final_df['PROJNUM'] = final_df['PROJNUM'].str[15:28]
        final_df['LABORTYPE'] = final_df['LABORTYPE'].str[:-1]
        final_df = final_df[final_df['LABORTYPE'] == '   Labor']
        final_df = final_df[final_df['RATE'].isnull()]

        email_body = final_df[['PROJNUM', 'EMPNAME']]

        final_df.to_csv(rf'C:\Users\V Song\Documents\FromHost\_OUTBOUND\ZERORATES\DTEK_DAILY_ZERORATES.csv',
                        index=False
                        )

        print_result()
        send_distr()

        return

    def run_mail(self, aitems, afunc):
        # the main outlook mailItem iterator
        # aitems object is a outlook mail object list
        # it received the mail obj List
        # and modify it
        for eamail in aitems:
            if eamail.Unread == True:
                # print(x.Subject)
                #look for automated daily dtek emails DISTR
                afunc(eamail)
                eamail.Unread = False
        return

