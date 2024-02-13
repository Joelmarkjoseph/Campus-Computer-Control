import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import cx_Oracle    
from tkinter import messagebox
import win32com.client as win32 # pip isntall pywin32
import pandas as pd
import pandas as pd
from sqlalchemy import create_engine
import datetime
from PIL import ImageTk,Image

def Entry1():
    appl=tk.Tk()
    appl.title("PBR VITS: Application")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='Day-wise entry form:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=650,y=20)

    label1 = tk.Label(appl, text='Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=100)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=620, y=100)

    label2 = tk.Label(appl, text='Name of the lab/Department(30):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=170)
    e2 = tk.Entry(appl, font=("Arial Rounded MT Bold", 20),width=20)
    e2.place(x=620, y=170)

    label3 = tk.Label(appl, text='Asset-code(15):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=130,y=240)
    e3 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e3.place(x=620, y=240)
    
    label4 = tk.Label(appl, text='Nature of complaint(50):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=130,y=310)
    e4 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e4.place(x=620, y=310)

    label1 = tk.Label(appl, text='Rectified-Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=380)
    er = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    er.place(x=620, y=380)

    label5 = tk.Label(appl, text='Action of work done(100):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label5.place(x=130,y=460)
    t = tk.Text(appl, width=45, height=7)
    t.place(x=620,y=460)

    label6 = tk.Label(appl, text='Status(30):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label6.place(x=130,y=600)
    e5 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e5.place(x=620, y=600)

    label7 = tk.Label(appl, text='Remarks(100):', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label7.place(x=130,y=660)
    t2 = tk.Text(appl, width=45, height=7)
    t2.place(x=620,y=660)

    b1=tk.Button(appl, command=lambda: submit(e1.get_date().strftime('%Y-%m-%d'),e2.get(),e3.get(),e4.get(),er.get_date().strftime('%Y-%m-%d'),t.get("1.0", "end-1c"),e5.get(),t2.get("1.0", "end-1c"),appl), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=800,y=785)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    appl.mainloop()
    pass

def getdata(fromm,to,fname,appl,n):
    try:
        print(fromm,to,n)
        start_date = fromm
        end_date = to
        print("Excel sheet function")
        engine = create_engine('oracle://system:joel@localhost:1521/xe')
        query = f"SELECT * FROM daywmain WHERE dt BETWEEN TO_DATE('{start_date}', 'YYYY-MM-DD') AND TO_DATE('{end_date}', 'YYYY-MM-DD') ORDER BY tl"
        dataframe = pd.read_sql(query, engine)
        dataframe['dt'] = dataframe['dt'].dt.date 
        dataframe['rd'] = dataframe['rd'].dt.date 

        new_column_names = {
        'dt': 'Date',
        'notld': 'Name_of_the_lab/Department',
        'asscd': 'Asset_Code',
        'noc': 'Nature_of_complaint',
        'rd': 'Rectified_date',
        'aowd': 'Action_of_workdone',
        'status': 'Status',
        'remarks': 'Remarks',
        'tl':'Sno.'}
        if fname=='':
            messagebox.showinfo("Error", "Specify a name to your file!",parent=appl)
        else:
            dataframe.rename(columns=new_column_names, inplace=True) # Keep only the date part
            current_date = datetime.datetime.now().date()
            output_excel_file = r"D:\\VITSCHOOL\\Shabbir sir project\\{}.xlsx".format(fname)
            dataframe.to_excel(output_excel_file, index=False, engine='openpyxl')
            msg_box = tk.messagebox.askquestion('Download Success', 'Excel Sheet Generated Successfully!\nDo you want to Generate Report again?',icon='info')
            if msg_box == 'no':
                appl.destroy()
            else:
                backf(appl)
                getreport()
                
            pass
    except Exception as e:
        messagebox.showinfo("Oops!!!", "Something went wrong!!!\nTry again!",parent=appl)
        print(e)
def backf(window):
        window.destroy()

def submit(dt,notld,asscd,noc,rd,aowd,status,remarks,win):
    try:
        print("Submit Function")
        print(dt,notld,asscd,noc,rd,aowd,status,remarks)       
        forma='yyyy-mm-dd'
        con = cx_Oracle.connect('system/joel@localhost:1521/xe')
        cur = con.cursor()
        
        
        cur.execute('select * from daywmain')
        li=[x for x in cur]
        tls=[x[-1] for x in li]
        print(tls)
        tl=max(tls)+1
        cur.execute('insert into daywmain values(to_date(\'{}\',\'{}\'),\'{}\',\'{}\',\'{}\',to_date(\'{}\',\'{}\'),\'{}\',\'{}\',\'{}\',{})'.format(dt,forma,notld,asscd,noc,rd,forma,aowd,status,remarks,tl))
        con.commit()
        print("Details inserted as a row")
        msg_box = tk.messagebox.askquestion('Insertion Success', 'Insertion success!\nDo you want to enter data again?',icon='info')
        if msg_box == 'no':
            win.destroy()
        else:
            backf(win)
            Entry1()
            pass
    except Exception as e:
        messagebox.showinfo("Oops!!!", "Something went wrong!!!\nTry again!",parent=win)
        print(e)
def getreport():
    appl=tk.Tk()
    appl.title("PBR VITS: Reports")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='REPORTS', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=600,y=20)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    
    label1 = tk.Label(appl, text='Enter Dates to get data b/w them', compound='center', font=("sans serif", 20),fg='white',bg='#1D5D9B')
    label1.place(x=470,y=100)

    label2 = tk.Label(appl, text='From Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=200)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=350, y=200)

    label3 = tk.Label(appl, text='To Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=730,y=200)
    e2 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e2.place(x=950, y=200)

    label3 = tk.Label(appl, text='Save as:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=790,y=500)
    e3=tk.Entry(appl,font=("Sans serif",20),width=12)
    e3.place(x=950,y=500)

    label3 = tk.Label(appl, text='.xlsx(Excel)', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=1100,y=500)

    src='daywmain'
    
    b1=tk.Button(appl, command=lambda:getdata(e1.get_date(),e2.get_date(),e3.get(),appl,src), text=" Generate Report! ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=700,y=785)



    pass

def mainf():
    main=tk.Tk()
    main.title("PBR VITS: COMPUTER CENTER")
    main.state('zoomed')
    main.configure(bg=('#1D5D9B'))
    label1 = tk.Label(main, text='PBR VITS: COMPUTER CENTER', compound='center', font=("Algerian", 35),fg='white',bg='#1D5D9B')
    label1.place(x=430,y=50)
    label2 = tk.Label(main, text='MAINTAINANCE DATA ', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=150)
    b1=tk.Button(command=Entry1,text=" DAY WISE ENTRY",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=130,y=230)
    
    label3 = tk.Label(main, text='Reports', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label3.place(x=680,y=150)
    b1=tk.Button(command=getreport,text=" Get Reports ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=670,y=230)
    main.mainloop()

#main function
mainf()