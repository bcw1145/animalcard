from  tkinter import *
from tkinter import filedialog

root = Tk()
root.filename =  filedialog.askopenfilename(initialdir = "E:/Images",title = "choose your file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
print (root.filename)