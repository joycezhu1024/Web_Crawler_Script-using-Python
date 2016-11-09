from Tkinter import *
import tkMessageBox
import TokyoMOU
import os
import India

if not os.path.exists('database'): 
        os.mkdir('database')

def sel(): 
    if(var.get()!=2):
        e1.delete(0,END)
        e2.delete(0,END)
        e1.insert(0,'23.10.2015')
        e2.insert(0,'24.10.2015')
    else:
        label_1['text']='Year:'
        label_2['text']='Month:'
        e1.delete(0,END)
        e2.delete(0,END)
        e1.insert(0,'2015')
        e2.insert(0,'10')
def ChooseFunction():

    selection =var.get()
    start =  e1.get()
    end =  e2.get()
    print e1.get()
    print e2.get()
    
    if(var.get()==1):
        tkMessageBox.showinfo('Info','You can see the processes in the py.exe which is in a black window.')
        TokyoMOU.LoadIP()
        TokyoMOU.Prepare(start,end)
        print TokyoMOU.total
        TokyoMOU.GetAllInspId()
        TokyoMOU.GetDetailInfo()
        TokyoMOU.ProcessDetail()
        tkMessageBox.showinfo('Info','All work are done.')
    elif(var.get()==2):
        tkMessageBox.showinfo('Info','You can see the processes in the py.exe which is in a black window.')
        import ParisMOU,ProcessParisBasicInfo
        ParisMOU.LoadIP()
        ParisMOU.GetAllInspId(start,end)
        ParisMOU.GetDetailInfo()
        ProcessParisBasicInfo.ProcessAllBasicInfo(start,end)
        tkMessageBox.showinfo('Info','All work are done.')

    if(var.get()==3):
        tkMessageBox.showinfo('Info','You can see the processes in the py.exe which is in a black window.')
        start_date=str(start)
        end_date=str(end)
        trans=India.Getlist(start_date, end_date)
        print "All lists are downloaded"
        India.Getinfo(start_date, end_date,trans)
        print 'All work done.'
        tkMessageBox.showinfo('Info','All work are done.')
        
root = Tk()
root.title("Spider For MOU")
root.geometry('350x250')

Label(root, text='   ',font=("Arial", 12)).grid(row=1)
Label(root, text='   ',font=("Arial", 12)).grid(row=2)
Label(root, text='   ',font=("Arial", 12)).grid(row=7)
Label(root, text='   ',font=("Arial", 12)).grid(row=8)


label_1 = Label(root, text='start time:',font=("Arial", 12))
label_1.grid(row=3,column=1) 
label_2 = Label(root, text='end time:',font=("Arial", 12))
label_2.grid(row=5,column=1)
e1 = Entry(root)  
e2 = Entry(root)
e1.grid(row=3, column=2)
e2.grid(row=5, column=2) 



#Choose MOU
label_3 = Label(root, text='MOU',font=("Arial", 12))
label_3.grid(row=6,column=0)

var = IntVar()
R1 = Radiobutton(root, text="Tokyo", variable=var, value=1,command=sel)
R1.grid(row=6,column=1)
R2 = Radiobutton(root, text="Paris", variable=var, value=2,command=sel)
R2.grid(row=6,column=2)
R3 = Radiobutton(root,text="India", variable=var, value=3,command=sel)
R3.grid(row=6,column=3)
var.set(1)


#Make Sure 
Sure = Button(root, text='OK',command = ChooseFunction)
Sure.grid(row=9,column=2)
e1.insert(0,'23.10.2015')
e2.insert(0,'24.10.2015')
root.mainloop()
