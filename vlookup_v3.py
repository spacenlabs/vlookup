from tkinter import *
#from tkinter.filedialog import askopenfile
from tkinter.filedialog import *
from time import sleep
import pandas as pd 
import numpy as np

ws = Tk()
ws['bg'] = '#00cc66'
ws.title('Excel Vlookup')
ws.geometry('470x400') 

label = Label(ws, text='Excel Vlookup', font=('bold',20), bg='#00cc66', fg='white')
label.grid(row=0,column=0, columnspan=2,padx = 150, pady=(35,0))

def open_file1():
    file_path1 = askopenfile(mode='r', filetypes=[('Excel Files', ['*.csv','*.xls'])])
    if file_path1 is not None:
        global df1
        df1 = pd.read_csv(file_path1.name, dtype=np.object_, encoding = 'ISO-8859-1')
        OPTIONS1 = list(df1.columns)
        variable1 = StringVar(ws)
        variable1.set('Select Column') # default value

        w = OptionMenu(ws, variable1, *OPTIONS1)
        w.grid(row=2, column=0, padx = 50, pady=15)
        df1['Duplicate'] = None
        def ok1():
            global var1
            var1 = variable1.get()
            if df1[var1].duplicated().sum():
                indexes = df1[df1[var1].duplicated(keep=False)].index
                for idx in indexes:
                    df1['Duplicate'][idx] = 'Yes'
            if df1[var1].duplicated().sum():
                dup1 = Label(ws, text=str(df1[var1].duplicated(keep=False).sum())+' Duplicates', font=('bold',13), bg='#00cc66', fg='white')
                dup1.grid(row=4,column=0, padx = 50, pady=(35,0))
            else:
                dup1 = Label(ws, text='No Duplicates', font=('bold',13), bg='#00cc66', fg='white')
                dup1.grid(row=4,column=0, padx = 50, pady=(35,0))

        button1 = Button(ws, text="OK", command=ok1, bg='white', fg='#2d8659')
        button1.grid(row=3, column=0)

def open_file2():
    file_path2 = askopenfile(mode='r', filetypes=[('Excel Files', ['*.csv','*.xls'])])
    if file_path2 is not None:
        global df2
        df2 = pd.read_csv(file_path2.name, dtype=np.object_, encoding = 'ISO-8859-1')

        OPTIONS2 = list(df2.columns)
        variable2 = StringVar(ws)
        variable2.set('Select Column') # default value

        y = OptionMenu(ws, variable2, *OPTIONS2)
        y.grid(row=2, column=1, padx = 50, pady=15)
        df2['Duplicate'] = None
        def ok2():
            global var2
            var2 = variable2.get()
            if df2[var2].duplicated().sum():
                indexes1 = df2[df2[var2].duplicated(keep=False)].index
                for idx1 in indexes1:
                    df2['Duplicate'][idx1] = 'Yes'
            if df2[var2].duplicated().sum():
                dup1 = Label(ws, text=str(df2[var2].duplicated(keep=False).sum())+' Duplicates', font=('bold',13), bg='#00cc66', fg='white')
                dup1.grid(row=4,column=1, padx = 50, pady=(35,0))
            else:
                dup1 = Label(ws, text='No Duplicates', font=('bold',13), bg='#00cc66', fg='white')
                dup1.grid(row=4,column=1, padx = 50, pady=(35,0))

        button2 = Button(ws, text="OK", command=ok2, bg='white', fg='#2d8659')
        button2.grid(row=3, column=1)

    def output():
        df3 = pd.merge(df1, df2, left_on=var1, right_on=var2, how='left')
        filpath = asksaveasfilename(defaultextension='.*', initialdir='C:/', title='Save File', filetypes=(("Text File",".txt"),("All Files","*.")))
        if filpath:
            df3.to_csv(filpath, index=False)

    btn2 = Button(ws, text ='Get Output', command = lambda:output(), font=('bold',13), bg='white', fg='#2d8659') 
    btn2.grid(row=5, column=0, columnspan=2, padx = 30, pady=35)

btn1 = Button(ws, text ='Select File One', command = lambda:open_file1(), font=('bold',13), bg='white', fg='#2d8659') 
btn1.grid(row=1, column=0, padx = 30, pady=35)

btn2 = Button(ws, text ='Select File Two', command = lambda:open_file2(), font=('bold',13), bg='white', fg='#2d8659') 
btn2.grid(row=1, column=1, padx = 30, pady=35)

ws.mainloop()