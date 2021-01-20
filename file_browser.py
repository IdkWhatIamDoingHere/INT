import tkinter
from tkinter import filedialog

def set_Adonis():
    
    Entry_Adonis.insert(0, filedialog.askopenfilename(initialdir="/", title="Select"))
    
def set_BWise():
    
    Entry_BWise.insert(0, filedialog.askopenfilename(initialdir="/", title="Select"))
    
def set_Schlüsselkontrollen():
    
    Entry_Schlüsselkontrollen.insert(0, filedialog.askopenfilename(initialdir="/", title="Select"))
    
def set_Users():
    
    Entry_Users.insert(0, filedialog.askopenfilename(initialdir="/", title="Select"))
    
root = tkinter.Tk()

root.title("test")
root.geometry("800x400")

Adonis_Label = tkinter.Label(root, text="Adonis Export")
Adonis_Label.grid(row=0,column=0)
Select_Adonis = tkinter.Button(root, text="Suchen", command=set_Adonis)
Select_Adonis.grid(row=0,column=1)
Entry_Adonis= tkinter.Entry(root, width= 50)
Entry_Adonis.grid(row=0, column=2, pady=25)

BWise_Label = tkinter.Label(root, text="BWise Export")
BWise_Label.grid(row=1, column=0)
Select_BWise = tkinter.Button(root, text="Suchen", command=set_BWise)
Select_BWise.grid(row=1, column=1)
Entry_BWise = tkinter.Entry(root, width=50)
Entry_BWise.grid(row=1, column=2, pady=25)

Schlüsselkontrollen_Label = tkinter.Label(root, text="Alle Schlüsselkontrollen")
Schlüsselkontrollen_Label.grid(row=2, column=0)
Select_Schlüsselkontrollen = tkinter.Button(root, text="Suchen", command=set_Schlüsselkontrollen)
Select_Schlüsselkontrollen.grid(row=2, column=1)
Entry_Schlüsselkontrollen = tkinter.Entry(root, width=50)
Entry_Schlüsselkontrollen.grid(row=2, column=2, pady=25)

Users_Label = tkinter.Label(root, text="Alle User")
Users_Label.grid(row=3, column=0)
Select_Users = tkinter.Button(root, text="Suchen", command=set_Users)
Select_Users.grid(row=3, column=1)
Entry_Users = tkinter.Entry(root, width=50)
Entry_Users.grid(row=3, column=2, pady=25)

start = tkinter.Button(root, text="Abgleich starten").grid(row=4, column=0, pady=25)


root.mainloop()