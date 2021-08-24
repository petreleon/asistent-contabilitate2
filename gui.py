import typing
from openpyxl import load_workbook
from unidecode import unidecode
# importing all files  from tkinter
from tkinter import * 
from tkinter import ttk
  
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



def save():
    files = [('All Files', '*.*'),
             ('Text Document', '*.txt')]
    file = asksaveasfilename(filetypes = files, defaultextension = files)
    print(file)


def processing(file):
    global inputtxt
    collumn_names = dict()
    lastCode = inputtxt.get("1.0", END)
    print(lastCode)
    lastNumber = int(lastCode[1:].lstrip("0"))
    print(lastNumber)
    currentFile = load_workbook(file)
    currentSheet = currentFile.active
    first_row = currentSheet[1]
    for index, collumn in enumerate(first_row):
       collumn_names[collumn.value] = index
    



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