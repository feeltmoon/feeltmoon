import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import pandas as pd

window = tk.Tk()


window.geometry("670x300")
window.title("Hello!")

style = ttk.Style()
style.theme_use("clam")

def rgb_hack(rgb):
    return "#%02x%02x%02x" % rgb 
window.config(bg=rgb_hack((255, 255, 255)))


def open_file():
    file_types = [('excel file','*.xlsx')]
    filename = fd.askopenfilename(title='Open',initialdir='/',filetypes=file_types)
    showinfo(title='Selected File',message=filename)
    txt.set(filename)


def open_file2():
    file_types = [('excel file','*.xlsx')]
    filename = fd.askopenfilename(title='Open',initialdir='/',filetypes=file_types)
    showinfo(title='Selected File',message=filename)
    txt1.set(filename)
    
    
def get_Entry():
    get_data = entry.get()
    return get_data
    
def get_Entry1():
    get_data = entry1.get()
    return get_data
    

def save_file_and_run():
    # Dynamic Path
    savePath = fd.askdirectory()
    showinfo(title='Selected File',message='HELLO')
    getEntryData = get_Entry()
    getEntryData1 = get_Entry1()
    showinfo(title='GetEntryData',message=getEntryData)
    showinfo(title='GetEntryData1',message=getEntryData1)

    df = pd.read_excel(txt_field.get(),sheet_name='Fields')

    df1 = pd.read_excel(txt_field1.get(),sheet_name='Fields') 

    if getEntryData != '' and getEntryData1 == '':
        #showinfo(title='prompt',message='Forms are not empty')
        entry_lst = []
        #entry_lst.append(getEntryData.split(,))
        entry_lst = getEntryData.split(',')
        #showinfo(title='prompt',message=entry_lst[0])
        for formOID in entry_lst:
            showinfo(title='prompt',message=formOID)
            df_Form = df.loc[df['FormOID']==formOID,:]
            df_Form.to_excel(savePath + '\\' + formOID + '.xlsx', index=False)
        


    
    
#label_01
myfont = ('Arial', 10)
label = tk.Label(window,text="ALS_First:",width=13,height=2,font=myfont,bg='White')
label.grid(row=1,column=1)
#text_field: to show path    
txt = tk.StringVar(None)    
txt_field = tk.Entry(window, textvariable=txt,bg='Silver')
txt_field.grid(row=1,columns=1)
txt_field.place(x=95, y=33, width=525, height=25)
#button: openFile
browse_text = tk.StringVar()
button = tk.Button(window, text='Open',command=open_file,width=8,height=1,bg="cornflower blue", fg="White")
button.grid(row=1,column=1)
button.place(x=23, y=31)

#label_02
myfont1 = ('Arial', 10)
label1 = tk.Label(window,text="ALS_Second:",width=14,height=1,font=myfont1,bg='White')
label1.grid(row=1,column=1)
label1.place(x=3,y=61)
#text_field: to show path    
txt1 = tk.StringVar(None)    
txt_field1 = tk.Entry(window, textvariable=txt1,bg='Light Gray')
txt_field1.grid(row=1,columns=1)
txt_field1.place(x=95, y=87, width=525, height=25)
#button: openFile
browse_text1 = tk.StringVar()
button1 = tk.Button(window, text='Open',command=open_file2,width=8,height=1,bg="cornflower blue", fg="White")
button1.grid(row=1,column=1)
button1.place(x=23, y=84)


#entry_01
entry = tk.Entry(window,bg='Silver')
entry.grid(row=1,columns=1)
entry.place(x=25, y=137, width=595, height=25)
#entry_01_label
entry_01_label_myfont = ('Arial', 10)
entry_01_label = tk.Label(window,text="Form(s):",width=11,height=1,font=entry_01_label_myfont,bg='White')
entry_01_label.grid(row=1,column=1)
entry_01_label.place(x=3,y=115)


#entry_02
entry1 = tk.Entry(window,bg='Silver')
entry1.grid(row=1,columns=1)
entry1.place(x=25, y=187, width=595, height=25)
#entry_01_label
entry_02_label_myfont = ('Arial', 10)
entry_02_label = tk.Label(window,text="Fields(s):",width=11,height=1,font=entry_02_label_myfont,bg='White')
entry_02_label.grid(row=1,column=1)
entry_02_label.place(x=3.5,y=165)


#button: Generate
gen_button = tk.Button(window, text='Choose a folder to run and save files', command=save_file_and_run, bg="cornflower blue", fg="White")
gen_button.grid(row=1,column=1)
gen_button.place(x=25, y=225)

# ending
window.mainloop()