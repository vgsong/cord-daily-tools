import tkinter as tk
import pandas as pd
import re

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
        self.parent.buttonsNewHires = tk.Button(self.parent, text='GET NEWHIRE', width=25,
                                                 command=self.get_newhire)

        self.pack_list = [
            self.parent.main_title,
            self.parent.buttonSaveReports,
            self.parent.buttonEleventwelveSend,
            self.parent.buttonSendTaubot,
            self.parent.buttonsZeroRates,
            self.parent.buttonsNewHires,
            self.parent.buttonBotherSam,
            self.parent.textProjectCode
        ]

        self.pack_all_buttons()

        self.fromhost_fpath = r'C:\Users\V Song\Documents\FromHost'

    def pack_all_buttons(self):
        for x in range(0, len(self.pack_list)):
            self.pack_list[x].pack()
        return None

    def out_bothersam(self):
        # Currently used just for laziness practice
        # Sends an email reminder to Sam
        # for changes, use .display instead of .send
        # independent since it does not retrieve the inbox mailObj list
        # it just creates a mail Item

        proj_var = self.parent.textProjectCode.get('1.0','end-1c')

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)

        olmail.To = 'TO HERE'
        olmail.CC = 'CC HERE'

        olmail.Subject = 'DTEK INVOICE APPROVAL'
        olmail.Body = f'EMAIL MESSAGE BODY HERE' \
'

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
        control_fpath = fr'{self.fromhost_fpath}\_FROMOUTLOOK\1112CONTACT_LIST.xlsx'
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
        olmail.Body = f'Note, project list is scheduled to be distributed every Friday as a reference when making project requests to avoid duplicate project requests. \n\nIf there are any discrepancies in the WOA/IO/Billing Contact (SCG PM you support) on the list please submit the updated information to a Cordoba PM lead to validate and submit to timesheet@cordobacorp.com\n\n' \
                      f'Thank you,\n\n' \

        olmail.Attachments.Add(fr'{self.fromhost_fpath}\_FROMOUTLOOK\1112_Project List Export.xlsx')

        olmail.display(True)

        del olmail
        del olapp

        return

    def get_dtekreports(self):
        olItems = self.get_mail('', 12)
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
            'DTEK_DAILY_Unbilled Detail and Aging Report': ['UNBILLED_DETAILS.txt', 'UNBILLED_DETAILS.csv'],
            'DTEK_DAILY_Project List Export Report': ['PROJECT_LIST_EXPORT.txt', 'PROJECT_LIST_EXPORT.csv'],
            'DTEK_DAILY_Employee Export Report': ['EMPLIST_DETAIL.txt', 'EMPLIST_DETAIL.csv'],
            'DTEK_DAILY_Unposted Labor Report': ['UNPOSTED_LABOR_DETAIL.txt', 'UNPOSTED_LABOR_DETAIL.csv'],
            '1112_Project List Export Report': ['1112_Project List Export.xlsx', '1112_Project List Export.xlsx'],
            '1168_Project List Export Report': ['1168_Project List Export.xlsx', '1168_Project List Export.xlsx']
        }

        ATTACHFPATH = fr'{self.fromhost_fpath}\_FROMOUTLOOK'  # local drive
        # ATTACHFPATH2 = r'C:\Users\V Song\Box\Invoicing\_FINANCE\VS\_DTEKREPORTS'  # box network drive remove comment once QA is done

        if aitem.Subject in DTEKMAILINFO:
            # Error handling if I am not logged in the Finance Team Network Drive
            # it tries to save ATTACHFPATH and ATTACHFPATH2, if not logged it saves only on local
            # print(f'{DTEKMAILINFO[aitem.Subject][0]}')
            aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH}\{DTEKMAILINFO[aitem.Subject][0]}')
            aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH}\{DTEKMAILINFO[aitem.Subject][1]}')
            # aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH2}\{DTEKMAILINFO[aitem.Subject]}')
            print(f'{DTEKMAILINFO[aitem.Subject]} saved successfully')

        return None

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

    def send_taubot_data(self):

        inv_dir = r'C:\Users\V Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\PDFINV'

        timeNow = datetime.now().strftime('%H%m%d')

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)

        olmail.To = 'TO HERE'
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
        fpath = fr'{self.fromhost_fpath}\_FROMOUTLOOK\{fname}'

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

            olmail.To = 'TO HERE'
            olmail.CC = 'CC HERE'

            olmail.Subject = 'SUBJECT HERE'
            olmail.Body = f'MSG HERE'

            olmail.Attachments.Add(fr'{self.fromhost_fpath}\_OUTBOUND\ZERORATES\DTEK_DAILY_ZERORATES.csv')

            olmail.display(True)

            return

        itemize_dict()

        final_df['PROJNUM'] = final_df['PROJNUM'].str[15:28]
        final_df['LABORTYPE'] = final_df['LABORTYPE'].str[:-1]
        final_df = final_df[final_df['LABORTYPE'] == '   Labor']
        final_df = final_df[final_df['RATE'].isnull()]

        email_body = final_df[['PROJNUM', 'EMPNAME']]

        final_df.to_csv(rf'{self.fromhost_fpath}\_OUTBOUND\ZERORATES\DTEK_DAILY_ZERORATES.csv',
                        index=False
                        )

        print_result()
        send_distr()

        return

    def get_newhire(self):
        olItems = self.get_mail('Time sheets', 48)
        self.check_newhire(olItems)
        return

    def check_newhire(self, aitem):

        # initiated dictFinal
        # it should match the codes from below which adds cummulative data into dictFinal

        dictFinal = {

            'EMPNAME': [],
            'SECTOR': [],
            'BRANCH': [],
            'PROJECTNAME': [],
            'PRONUM': [],
            'SECTORNUM': [],
            'BRANCHNUM': [],
            'EMPNUM': []

        }

        dtek_upload = {

            'PRONUM': [],
            'SECTORNUM': [],
            'CONCAT': [],
            'EMPNUM': []

        }

        sector_mapp = {

            'Corporate': 10,
            'Education': 15,
            'Energy': 20,
            'Transportation': 25,
            'Water': 30

        }

        org_mapp = {

            'Chatsworth': 10,
            'Los Angeles': 15,
            'Sacramento': 20,
            'San Diego': 25,
            'San Francisco': 30,
            'San Ramon': 35,
            'Santa Ana': 40,
            'Ontario': 45

        }

        pronum_mapp = {

            'Training/Safety': '0000.000.105',
            'Vacation': '0000.000.200',
            'Sick Leave': '0000.000.205',
            'Holiday': '0000.000.210',
            'Bereavement': '0000.000.215',
            'Jury Duty': '0000.000.220',
            'Business Proposal': '0000.000.105',
            'Energy Proposals 2021': '0000.000.105',
            'Business Development': '0000.000.105',
            'Recruiting': '0000.000.105',
            'Emergency-PTO': '0000.000.105',
            'Business Development': '0000.000.105',
            'Admin': '0000.000.105',
            'Admin-WorkCare': '0000.000.105'

        }

        timeStamp = datetime.now().strftime('%y%m%d_%H%M')

        for item in aitem:
            if item.Unread == True and ('FIRST LAST' in item.SenderName or 'Info System' in item.SenderName) and 'RE: New Employee Has Been Created' in item.Subject and item.Attachments.Count > 0:
                arrEmpName = []
                arrEmpNumber = []
                arrSector = []
                arrBranch = []
                arrProject = []

                arrOnBoard = []
                arrOnBoardFinal = []

                regexEmpName = '(?<=Employee Name: )((.|\n)*)(?=Organization:)'
                regexEmpNumber = '(?<=Employee Number: )\d{4}'
                regexOnBoard = '(?<=PTO Designations(.))((.|\n)*)(?=Billable Projects:)'

                # colNames = {
                #     'EMPNAME': 'empName',
                #     'EMPNUM': 'empNumber',
                #     'SECTOR': 'arrSector',
                #     'BRANCH': 'arrBranch',
                #     'PROJECTNAME': 'arrProject'
                # }

                # # print(timeStamp)
                # # print(re.search(regexEmpName,item.Body).group(0))
                # # print(re.search(regexEmpNumber, item.Body).group(0))
                # # print(re.search(regexOnBoard, item.Body).group(0))

                empName = re.search(regexEmpName, item.Body).group(0).strip()
                empNumber = re.search(regexEmpNumber, item.Body).group(0).strip()

                print(empName)

                arrOnBoard.append(re.search(regexOnBoard, item.Body).group(0).strip())
                arrOnBoardFinal = arrOnBoard[0].splitlines()

                print(arrOnBoardFinal)

                for x in arrOnBoardFinal:
                    arrSector.append(x.split('-')[0])
                    arrBranch.append(x.split('-')[1])
                    arrProject.append(x.split('-')[2])

                arrEmpName = [empName] * len(arrSector)
                arrEmpNumber = [empNumber] * len(arrSector)



                # currently hard coded due to inexperience
                # every dict key should be populated
                # make sure it matches with the initial dictFinal initiated above


                dictFinal['EMPNAME'] += arrEmpName
                dictFinal['EMPNUM'] += arrEmpNumber
                dictFinal['SECTOR'] += arrSector
                dictFinal['BRANCH'] += arrBranch
                dictFinal['PROJECTNAME'] += arrProject

                # for debug
                # print(arrEmpName)
                # print(arrEmpNumber)
                # print(arrSector)
                # print(arrBranch)
                # print(arrProject)

        # strip() each value in key
        for k,v in dictFinal.items():
            dictFinal[k] = [x.strip() for x in dictFinal[k]]

        for x in dictFinal['SECTOR']:
            if x in sector_mapp.keys():
                dictFinal['SECTORNUM'].append(sector_mapp[x])
            else:
                dictFinal['SECTORNUM'].append('not found')

        for x in dictFinal['BRANCH']:
            if x in org_mapp.keys():
                dictFinal['BRANCHNUM'].append(org_mapp[x])
            else:
                dictFinal['SECTORNUM'].append('not found')

        for x in dictFinal['PROJECTNAME']:
            if x in pronum_mapp.keys():
                dictFinal['PRONUM'].append(pronum_mapp[x])
            else:
                dictFinal['SECTORNUM'].append('not found')


        dtek_upload['PRONUM'] = dictFinal['PRONUM']
        dtek_upload['SECTORNUM'] = dictFinal['SECTORNUM']
        dtek_upload['CONCAT'] = [''.join([str(x[0][-3:]),str(x[1]),str(x[2])]) for x in zip(dictFinal['PRONUM'],dictFinal['SECTORNUM'],dictFinal['BRANCHNUM'])]
        dtek_upload['EMPNUM'] = dictFinal['EMPNUM']

        email_message = fr'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\NEWHIRERECORDS\_EMAIL_MESSAGE.csv'


        # the CSV file/fpath  that is used to upload.
        csv_fname = fr'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\NEWHIRERECORDS\CORDOBA_NEWHIRE_{timeStamp}.csv'
        dfFinal = pd.DataFrame.from_dict(dtek_upload)
        dfEmail = pd.DataFrame.from_dict(dictFinal)

        dfEmail.to_csv(email_message, header=True, index=False)

        dfFinal.to_csv(csv_fname, header=False, index=False)
        print(f'NEW EMP CSV OUTPUT SAVED AT: {csv_fname}')

        return