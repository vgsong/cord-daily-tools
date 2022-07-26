# cord-daily-tools

Daily automation tools used at work for Cordoba finance department and operations. 

Most of these scripts facilitates daily/weekly/monthly SOP processes, which can take a considerable amount of prep time due to its volume/frequency.

I started with tkinter GUI library, and I am now trying to move to PyQT.

mainFACTORIO (currently retired) is the file that inherits all tkinter modules from classes folder. Much less crowded once I learned the OOP approach for tkinter.

Atalhos.py - Contains scripts that launches files, dir shortcuts.

PDFReader.py - PyPDF2 reads pdf files in dir, and transposes strings into excel based on Regex match.

PyOutlook.py - Uses win32com for outlook daily tasks such as newhire file prep, save attached file into dir based on item.subject.


mainATALHOSQ is the same as mainFACTORIO, but I have found that PyQT is more powerful and simpler than tkinter (in a way). I am trying to learn the OOP approach for PyQT, as of now, everything is in one file. (need to re-organize)  

