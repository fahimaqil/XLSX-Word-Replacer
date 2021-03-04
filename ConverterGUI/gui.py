#design inspired from https://github.com/KeithGalli/GUI/blob/master/WeatherApp.py
#time taken 2.568
import tkinter as tk

from tkinter import ttk
from tkscrolledframe import ScrolledFrame
from tkinter import filedialog,messagebox
import os
from extractGUI import *
from tkinter.messagebox import showerror
import time

HEIGHT = 900
WIDTH = 1400
PATH="tets"
listSheet=[]
listCol=[]
root = tk.Tk()
general_checkbuttons={}
hasPress=False
style = ttk.Style()
 
style.configure('Button', foreground='Green')

def UploadAction(event=None):
    try:
       # print("yes")
        filename = filedialog.askopenfilename(title = "Select .xlsx file",filetypes = [("XLSX Files",".xlsx")])
        global PATH

        PATH=filename
        print(PATH)
        canvas2.pack_forget()
        label.config(text="File: {}".format(PATH))
        global listSheet
        listSheet=return_sheet(PATH)
        mycomboSheet.config(values=listSheet)
        mycomboSheet.current(0)
        canvas.pack()
    except(OSError,FileNotFoundError):
        showerror("Error", message=root.filename)
        root.destroy()



        print(f'Unable to find or open <{root.filename}>')
    except Exception as error:
        showerror("Error", message="Error occured! Please choose a file next time")
        root.destroy()

   
def on_selectSheet(event=None):
    global listColumn
    listColumn=return_column(PATH,mycomboSheet.get())
    show_frame = tk.Frame(canvas, bg='#6DD5ED', bd=10)
    show_frame.place(relx=0.5, rely=0.25, relwidth=0.8, relheight=0.12 ,anchor='n')

    labelColumn = tk.Label(show_frame,text="Choose Column",bg='#6DD5ED')
    labelColumn.place(rely=0,relx=0.5,relwidth=1, relheight=0.25,anchor="n")
    
    myComboColumn=ttk.Combobox(show_frame,state="readonly",values=listColumn)
    myComboColumn.place(relwidth=1, relheight=0.45,rely=0.4)
    myComboColumn.current(0)
    # print(myComboColumn.get())

    myComboColumn.bind('<<ComboboxSelected>>', on_selectColumn)

def on_selectColumn(eventObject):
    global value
    value=return_value(PATH,mycomboSheet.get(),eventObject.widget.get())
    global chosenSheet
    chosenSheet=mycomboSheet.get()
    global chosenColumn
    chosenColumn=eventObject.widget.get()
    value=list(value)
    create_column(value)
def callback(valueVar):
    print(valueVar.get())
def valueSearch(valStr):
    # var=tk.StringVar()
    # var.set(valStr)
    # var.trace("w", callback(var))
    
    print(valStr)

    
    
def create_column(currentValue):
    show_scroll=tk.Label(canvas,text="Choose word to replace" ,bg='#6DD5ED', bd=10)
    show_scroll.place(relx=0.5, rely=0.4, relwidth=0.8, relheight=0.05, anchor='n')
    sf = ScrolledFrame(canvas)
    sf.place(relx=0.5, rely=0.45, relwidth=0.8, relheight=0.35, anchor='n')

    # Bind the arrow keys and scroll wheel
    sf.bind_arrow_keys(canvas)
    sf.bind_scroll_wheel(canvas)

    # Create a frame within the ScrolledFrame
    inner_frame = sf.display_widget(tk.Frame)
    global general_checkbuttons
    general_checkbuttons = {}
    col=3
    counterX=0
    counterY=0
    arrayAns=[]
   # print(currentValue)
    # currentValue=["A","B","C","D","A","B","C","D","A","B","C","D","A","B","C","D"]
    root.grid_columnconfigure(4, minsize=50)
    for i in range(len(currentValue)):
        for y in range(col):
            if counterX % col!= 0 or i==0:
                cal=i%col

                cb = tk.Checkbutton(inner_frame, font=(None, 12),variable=currentValue[i],text=currentValue[i], wraplength=250)
                cb.grid(row=counterY, column=cal, sticky="w", pady=1, padx=1)
                general_checkbuttons[i] = cb
                break
            elif counterX %col ==0:
                cal=i%col


                counterY+=1
                cb = tk.Checkbutton(inner_frame, font=(None, 12),variable=currentValue[i],text=currentValue[i], wraplength=250)

                cb.grid(row=counterY, column=cal, sticky="w", pady=1, padx=1)
                general_checkbuttons[i] = cb
                break

        counterX+=1
    frameSearch=tk.Frame(canvas,bd=5,bg="#6DD5ED")
    frameSearch.place(rely=0.80,relx=0.1,relwidth=0.74, relheight=0.04)
    entrySearch=tk.Entry(frameSearch)
    entrySearch.place(relwidth=1)
    entrySearch.insert(0, 'Search for words')
    entrySearch.bind("<FocusIn>", lambda args: entrySearch.delete('0', 'end'))

    entrySearch.bind("<KeyPress>", lambda args: valueSearch(entrySearch.get(),currentValue))


    
    


    buttonWord = tk.Button(canvas, text='Choose Word', command=lambda:extractChecked(general_checkbuttons,currentValue,True))
    buttonWord.place(rely=0.80,relx=0.84)

def extractChecked(check,valueArray,pressed):
    array=[]
    for i in range(len(valueArray)):
        cb = check[i]
        varname = cb.cget("variable")
        value = canvas.getvar(varname)
        if value=="1":
            array.append(valueArray[i])
    frame3 = tk.Frame(canvas, bg='#6DD5ED',bd=10)
    frame3.place(relwidth=0.8, relheight=0.1,rely=0.86,relx=0.1)
    label= tk.Label(frame3,text="Editing column in file: {}".format(PATH),bg='#6DD5ED')
    label.place(rely=0.06,relx=0.5,anchor="center")
    entry1 = tk.Entry (frame3) 
    entry1.place(relwidth=1, relheight=0.6,rely=0.38)
    global hasPress
    hasPress=pressed
    buttonSubmit = tk.Button(canvas, text='Submit',command=lambda: confirm(array,entry1))
    buttonSubmit.place(rely=0.95,relx=0.86)
def clear():

    for i in general_checkbuttons:
        general_checkbuttons[i].deselect()   # forget checkbutton
        # i.destroy()   
def confirm(array,entry1):
    check =entry1.get() if entry1.get()!="" else ""
    global hasPress
    #print("hasPress:", hasPress)
    if hasPress==True and len(array)>0 :
        message=messagebox.askquestion(title=None, message="Do you wish to submit {}".format( check ), icon='question')
        if message =="yes":
            # print("Yesssss",array)
            start = time.process_time()

            createDict(PATH,check,array,chosenColumn,chosenSheet)
            clear()
            value=list(return_value(PATH,chosenSheet,chosenColumn))
            # print("Value: ",value)
            create_column(value)
            hasPress=False
            print(time.process_time() - start)

            messagebox.showinfo("showinfo", "Successful!") 
            
        else:
            # print("No")
            hasPress=False
    else:
        showerror("Error", message="Please reselect word and press 'Choose Word' ")
        hasPress=False

    


def getChecked(value):
   # print(cb.cget("text"))
    for i in range(value):
        cb = general_checkbuttons[i]
        
        varname = cb.cget("variable")
        value = canvas.getvar(varname)
       # print(f"{i}: {value}")

def page1():
    message=messagebox.askquestion(title=None, message="Do you wish to exit?", icon='question')
    if message =="yes":
        print("Yesssss")
        root.destroy()
    else:
        return



def page2():
    canvas2.pack_forget()
    canvas.pack()


canvas2 = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas2.pack()



main_frame = tk.Frame(canvas2,bg='#6DD5ED')
main_frame.place(relwidth=1, relheight=1)

labelMain=tk.Label(main_frame,text="Word Converter",bg="#6DD5ED",font=("Helvetica", 48))
labelMain.place(relx=0.5,rely=0.4,anchor="center")

buttonFile = ttk.Button(main_frame, text='Choose File', command= UploadAction)
buttonFile.place(rely=0.5,relx=0.5,anchor="center")

#### First Canvas ##################################


canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack_forget()

# background_image = tk.PhotoImage(file='Cool Blues.png')
background_label = tk.Label(canvas, bg='#40C8E7')
background_label.place(relwidth=1, relheight=1)


############## top frame #################
frame = tk.Frame(canvas,bg='#6DD5ED', bd=5)
frame.place(relx=0.5, rely=0.02, relwidth=0.8, relheight=0.05, anchor='n')
label= tk.Label(frame,text="Editing column in file: {}".format(PATH),bg='#6DD5ED')
label.place(relwidth=1, relheight=1)
############## end top frame #################


###############show sheet ##############################

show_sheet= tk.Frame(canvas, bg='#6DD5ED', bd=10)
show_sheet.place(relx=0.5, rely=0.1, relwidth=0.8, relheight=0.12, anchor='n')

labelSheet = tk.Label(show_sheet,text="Choose Sheet",bg='#6DD5ED')
labelSheet.place(rely=0,relx=0.5,relwidth=1, relheight=0.25,anchor="n")

mycomboSheet=ttk.Combobox(show_sheet,state="readonly",values=listSheet)
mycomboSheet.place(relwidth=1, relheight=0.4,rely=0.4)
mycomboSheet.bind('<<ComboboxSelected>>', on_selectSheet)




buttonClose=tk.Button(canvas, text='Exit',command=page1)
buttonClose.place(anchor="nw")

root.mainloop()
