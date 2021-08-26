import typing
from openpyxl import load_workbook
from unidecode import unidecode
# importing all files  from tkinter
from tkinter import * 
from tkinter import ttk
from enum import Enum
class colNames(Enum):
    NUMAR_FACTURA = 'Numar factura'
    DATA = 'Data'
    PERSOANA_JURIDICA = 'Persoana juridica'
    CUI = 'CNP / CUI'
    COD_EXTERN = 'Cod extern'
    PRODUS = 'Comenzi'
    COD_PRODUS = 'Cod produs'
    CANTITATE = 'Cantitate'
    PRET = 'Valoare fara TVA'


  
# import only asksaveasfile from filedialog
# which is used to save file in any extension
from tkinter.filedialog import asksaveasfilename, askopenfilename

root = Tk()
root.geometry('200x200')

# function to call when user press
# the save button, a filedialog will
# open and ask to save file
NoneType = type(None)
def texting(text_change: typing.Union[str, NoneType] = None):
    global text
    text.delete('1.0', END)
    if text_change is not None:
        text.insert('1.0', text_change)

collumn_names = dict()

def save():
    files = [('All Files', '*.*'),
             ('Text Document', '*.txt')]
    file = asksaveasfilename(filetypes = files, defaultextension = files)
    print(file)

def getFromRow(row, col_enum):
    return row[collumn_names[col_enum.value]].value

lastNumber = 0 # will be modified further
externalCodes = dict()
companiesToQuery = list()

def addCompany(CUI, externalCode):
    pass

def iterOnRows(rows):
    second_row = rows[2]
    last_doc = getFromRow(second_row, colNames.NUMAR_FACTURA)


def processing(file):
    global inputtxt
    global lastNumber
    lastCode = inputtxt.get("1.0", END)
    print(lastCode)
    lastNumber = int(lastCode[1:].lstrip("0"))
    print(lastNumber)
    currentFile = load_workbook(file)
    currentSheet = currentFile.active
    first_row = currentSheet[1]
    for index, collumn in enumerate(first_row):
       collumn_names[collumn.value] = index
    iterOnRows(currentSheet)
    


def open():
    files = [('All Files', '*.*'), 
             ('excel', '*.xlsx')]
    file = askopenfilename(filetypes = files, defaultextension = files)
    processing(file)
    print(file)

inputtxt = Text(root,
                   height = 1,
                   width = 20)

inputtxt.pack(side = TOP, pady = 10)

btnSave = ttk.Button(root, text = 'Save', command = lambda : save())
btnOpen = ttk.Button(root, text = 'Open', command = lambda : open())
btnOpen.pack(side = TOP, pady = 10)
btnSave.pack(side = TOP, pady = 10)
text = Text(root)
text.pack(side = TOP, pady = 15)
texting("open excel")
inputtxt.insert('1.0', "C005186")


mainloop()