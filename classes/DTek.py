import os
import tkinter as tk
import pandas as pd


class DTek(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.main_title = tk.Label(self.parent, text='DTEK UPLOAD', bg='#7d7d7d')
        self.parent.buttonGetProgram = tk.Button(self.parent, width=25, text='GET Program', command=self.get_prog)
        self.parent.textProgram = tk.Text(self.parent, height=2, width=22)

        # will pack tk widgets in the order of the array below
        self.pack_list = [
            self.parent.main_title,
            self.parent.buttonGetProgram,
            self.parent.textProgram
        ]

        # list of CONST fpath:
        self.report_path = r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK'
        self.outbound_path = r'C:\Users\V Song\Documents\FromHost\_OUTBOUND\_DTEKIMPORT'
        self.box_fpath = r'C:\Users\V Song\Box\Invoicing\_FINANCE\_DELTEKUPLOAD'
        self.project_list = 'PROJECT_LIST_EXPORT.csv'

        self.pack_all_buttons()

    def pack_all_buttons(self):
        for x in range(0, len(self.pack_list)):
            self.pack_list[x].pack()
        return None

    def get_prog(self):

        def show_dfcolumns():
            for x, y in enumerate(rdf.columns):
                print(x, y)
            return None

        # get elementbyID inside textbook values,
        # where 1.0 is the beg of string and end-1c indicates where the str should end
        prog_group = self.parent.textProgram.get('1.0', 'end-1c')
        print(f'Opening {os.path.join(self.report_path, self.project_list)} for {prog_group}...')

        # loads df into rdf (raw data frame)
        rdf = pd.read_csv(os.path.join(self.report_path, self.project_list))

        show_dfcolumns()

        # filters data based on str that contains 1112 (example), project status Active, and does NOT contain .999
        # drop duplicates
        rdf = rdf[(rdf[rdf.columns[0]].str.contains(str(prog_group))) &
                  (rdf[rdf.columns[12]] == 'Active') &
                  (~rdf[rdf.columns[0]].str.contains('.999') &
                  (~rdf[rdf.columns[0]].str.contains('.00E'))
                   )]

        result = rdf[rdf.columns[0]].drop_duplicates()

        print(result)
        result.to_csv(os.path.join(self.outbound_path, f'{prog_group}_EMPUPLOAD.csv'), header=False, index=False)
        result.to_csv(os.path.join(self.box_fpath, f'{prog_group}_EMPUPLOAD.csv'), header=False, index=False)
        print(f'{prog_group} saved at {self.outbound_path}\n and {self.box_fpath}')

        return None
