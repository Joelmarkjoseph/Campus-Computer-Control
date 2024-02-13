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

# ((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((( SHABBIR SIR PROJECT )))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))

def asset_entry(atype):
    appl=tk.Tk()
    appl.title("Asset Entry")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label1 = tk.Label(appl, text='Asset Type: \"{}\"'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=600,y=20)

    # l1=tk.Label(appl, text='Sno:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # l1.place(x=70,y=70)
    # e1= tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    # e1.place(x=350, y=70)

    # l2=tk.Label(appl, text='Asset Code:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # l2.place(x=70,y=140)
    # e2 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    # e2.place(x=350, y=140)

    l3=tk.Label(appl, text='Date of purchase:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l3.place(x=70,y=210)
    e3 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e3.place(x=350, y=210)

    l4=tk.Label(appl, text='Processor Type:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l4.place(x=70,y=280)
    e4 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)
    e4.place(x=350, y=280)

    l5=tk.Label(appl, text='Ram:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l5.place(x=70,y=350)
    e5 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    e5.place(x=350, y=350)

    l6=tk.Label(appl, text='Hard Disk:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l6.place(x=70,y=420)
    e6 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    e6.place(x=350, y=420)

    l7=tk.Label(appl, text='Motherboard:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l7.place(x=70,y=490)
    e7 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    e7.place(x=350, y=490)

    l8=tk.Label(appl, text='Supplier Name:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l8.place(x=70,y=560)
    e8 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    e8.place(x=350, y=560)

    l9=tk.Label(appl, text='Bill no:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l9.place(x=70,y=630)
    e9 = tk.Entry(appl, font=("Arial Rounded MT ", 25), width=10)  # Use DateEntry widget
    e9.place(x=350, y=630)

    l10=tk.Label(appl, text='Bill Date:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l10.place(x=750,y=140)
    e10 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e10.place(x=960, y=140)
    
    l11=tk.Label(appl, text='Quantity:'.format(atype), compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    l11.place(x=750,y=210)
    e11 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e11.place(x=960, y=210)

    tbl=''
    if atype=='CPU':
        tbl='cpudet'
        ac='PVCPU'
    if atype=='Monitors':
        tbl='mondet'    
        ac='PVMLD'
    if atype=='Printers':
        tbl='pridet'
        ac='PVPRD'
    if atype=='UPS':
        tbl='upsdet'
        ac='PVUPS'

    b1=tk.Button(appl, command=lambda:insert_asset(appl,tbl,e3.get_date(),e4.get(),e5.get(),e6.get(),e7.get(),e8.get(),int(e9.get()),e10.get_date(),int(e11.get()),ac), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=800,y=785)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    pass

def insert_asset(win, tbl, dt1, pt, ram, hdd, mb, sname, bno, bdt, qn, ac):
    forma = 'yyyy-mm-dd'
    yr = str(bdt)[2:4]
    con = cx_Oracle.connect('system/joel@localhost:1521/xe')
    cur = con.cursor()
    cur.execute("select * from {}".format(tbl))
    li = [x for x in cur]
    snos = [x[0] for x in li]
    acs = [x[1] for x in li]
    acs.sort()
    az = acs[-1]
    num = int(az[8:11])
    sno = max(snos)
    for i in range(qn):
        num += 1
        if num < 10:
            at = "00" + str(num)
        elif num < 100 and num >= 10:
            at = "0" + str(num)
        elif num < 1000 and num >= 100:
            at = str(num)
        acode = (ac + yr + '-' + at)
        sno += 1
        print(sno, acode)
        cur.execute('insert into {} values({},\'{}\',to_date(\'{}\',\'{}\'),\'{}\',\'{}\',\'{}\',\'{}\',\'{}\',{},to_date(\'{}\',\'{}\'))'.format(tbl, sno, acode, dt1, forma, pt, ram, hdd, mb, sname, bno, bdt, forma))
        con.commit()
        print("Details inserted as a row")
    backf(win)


def Entry1():
    appl=tk.Tk()
    appl.title("PBR VITS: Application")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='DPR - I ENTRY FORM:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=650,y=20)

    label1 = tk.Label(appl, text='Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=100)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=250, y=100)

    label2 = tk.Label(appl, text='Name & Address of the supplier:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=170)
    e2 = tk.Entry(appl, font=("Arial Rounded MT Bold", 20),width=20)
    e2.place(x=620, y=170)

    label3 = tk.Label(appl, text='Purchase order no.:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=130,y=240)
    e3 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e3.place(x=450, y=240)
    
    label4 = tk.Label(appl, text='Invoice/Bill No & Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=130,y=310)
    e4 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e4.place(x=450, y=310)
    indt = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    indt.place(x=680, y=310)

    label5 = tk.Label(appl, text='Description of material:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label5.place(x=130,y=380)
    t = tk.Text(appl, width=45, height=7)
    t.place(x=470,y=380)

    label6 = tk.Label(appl, text='Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label6.place(x=130,y=520)
    e5 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e5.place(x=450, y=520)

    label7 = tk.Label(appl, text='Quantity:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label7.place(x=130,y=590)
    e6 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e6.place(x=450, y=590)
    
    # label8 = tk.Label(appl, text='Total Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label8.place(x=130,y=660)
    # e7 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e7.place(x=450, y=660)
    
    label9 = tk.Label(appl, text='Other charges:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label9.place(x=130,y=730)
    e8 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e8.place(x=450, y=730)
    
    # label10 = tk.Label(appl, text='Total Amount:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label10.place(x=130,y=785)
    # e9 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e9.place(x=450, y=785)

    b1=tk.Button(appl, command=lambda: submit(e1.get_date().strftime('%Y-%m-%d'),e2.get(),e3.get(),e4.get(),t.get("1.0", "end-1c"),e5.get(),e6.get(),e8.get(),indt.get_date().strftime('%Y-%m-%d'),appl,1), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=800,y=785)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    appl.mainloop()
    pass

def Entry1C2():
    appl=tk.Tk()
    appl.title("PBR VITS Campus -2 : Application")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='DPR - I ENTRY FORM:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=650,y=20)

    label1 = tk.Label(appl, text='Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=100)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=250, y=100)

    label2 = tk.Label(appl, text='Name & Address of the supplier:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=170)
    e2 = tk.Entry(appl, font=("Arial Rounded MT Bold", 20),width=20)
    e2.place(x=620, y=170)

    label3 = tk.Label(appl, text='Purchase order no.:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=130,y=240)
    e3 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e3.place(x=450, y=240)
    
    label4 = tk.Label(appl, text='Invoice/Bill No & Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=130,y=310)
    e4 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e4.place(x=450, y=310)
    indt = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    indt.place(x=680, y=310)

    label5 = tk.Label(appl, text='Description of material:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label5.place(x=130,y=380)
    t = tk.Text(appl, width=45, height=7)
    t.place(x=470,y=380)

    label6 = tk.Label(appl, text='Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label6.place(x=130,y=520)
    e5 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e5.place(x=450, y=520)

    label7 = tk.Label(appl, text='Quantity:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label7.place(x=130,y=590)
    e6 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e6.place(x=450, y=590)
    
    # label8 = tk.Label(appl, text='Total Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label8.place(x=130,y=660)
    # e7 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e7.place(x=450, y=660)
    
    label9 = tk.Label(appl, text='Other charges:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label9.place(x=130,y=730)
    e8 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e8.place(x=450, y=730)
    
    # label10 = tk.Label(appl, text='Total Amount:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label10.place(x=130,y=785)
    # e9 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e9.place(x=450, y=785)

    b1=tk.Button(appl, command=lambda: submit(e1.get_date().strftime('%Y-%m-%d'),e2.get(),e3.get(),e4.get(),t.get("1.0", "end-1c"),e5.get(),e6.get(),e8.get(),indt.get_date().strftime('%Y-%m-%d'),appl,3), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=800,y=785)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    appl.mainloop()
    pass

def Entry2():
    appl=tk.Tk()
    appl.title("PBR VITS: Application")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='DPR - II ENTRY FORM:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=650,y=20)

    label1 = tk.Label(appl, text='Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=100)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=250, y=100)

    label2 = tk.Label(appl, text='Name & Address of the supplier:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=170)
    e2 = tk.Entry(appl, font=("Arial Rounded MT Bold", 20),width=20)
    e2.place(x=620, y=170)

    label3 = tk.Label(appl, text='Purchase order no.:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=130,y=240)
    e3 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e3.place(x=450, y=240)
    
    label4 = tk.Label(appl, text='Invoice/Bill No:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=130,y=310)
    e4 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e4.place(x=450, y=310)
    indt = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    indt.place(x=680, y=310)

    label5 = tk.Label(appl, text='Description of material:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label5.place(x=130,y=380)
    t = tk.Text(appl, width=45, height=7)
    t.place(x=470,y=380)

    label6 = tk.Label(appl, text='Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label6.place(x=130,y=520)
    e5 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e5.place(x=450, y=520)

    label7 = tk.Label(appl, text='Quantity:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label7.place(x=130,y=590)
    e6 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e6.place(x=450, y=590)
    
    # label8 = tk.Label(appl, text='Total Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label8.place(x=130,y=660)
    # e7 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e7.place(x=450, y=660)
    
    label9 = tk.Label(appl, text='Other charges:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label9.place(x=130,y=730)
    e8 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e8.place(x=450, y=730)
    
    # label10 = tk.Label(appl, text='Total Amount:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label10.place(x=130,y=785)
    # e9 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e9.place(x=450, y=785)

    b1=tk.Button(appl, command=lambda: submit(e1.get_date().strftime('%Y-%m-%d'),e2.get(),e3.get(),e4.get(),t.get("1.0", "end-1c"),e5.get(),e6.get(),e8.get(),indt.get_date().strftime('%Y-%m-%d'),appl,2), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=800,y=785)
    b2=tk.Button(appl, command=lambda:backf(appl),  text=" <-Back ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b2.place(x=50,y=25)
    appl.mainloop()
    pass

def Entry2C2():
    appl=tk.Tk()
    appl.title("PBR VITS: Application")
    appl.state('zoomed')
    appl.configure(bg=('#1D5D9B'))
    label0 = tk.Label(appl, text='DPR - II ENTRY FORM:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label0.place(x=650,y=20)

    label1 = tk.Label(appl, text='Date:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label1.place(x=130,y=100)
    e1 = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    e1.place(x=250, y=100)

    label2 = tk.Label(appl, text='Name & Address of the supplier:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=170)
    e2 = tk.Entry(appl, font=("Arial Rounded MT Bold", 20),width=20)
    e2.place(x=620, y=170)

    label3 = tk.Label(appl, text='Purchase order no.:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label3.place(x=130,y=240)
    e3 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e3.place(x=450, y=240)
    
    label4 = tk.Label(appl, text='Invoice/Bill No:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=130,y=310)
    e4 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e4.place(x=450, y=310)
    indt = DateEntry(appl, font=("Arial Rounded MT Bold", 25), width=10)  # Use DateEntry widget
    indt.place(x=680, y=310)

    label5 = tk.Label(appl, text='Description of material:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label5.place(x=130,y=380)
    t = tk.Text(appl, width=45, height=7)
    t.place(x=470,y=380)

    label6 = tk.Label(appl, text='Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label6.place(x=130,y=520)
    e5 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e5.place(x=450, y=520)

    label7 = tk.Label(appl, text='Quantity:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label7.place(x=130,y=590)
    e6 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e6.place(x=450, y=590)
    
    # label8 = tk.Label(appl, text='Total Price:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label8.place(x=130,y=660)
    # e7 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e7.place(x=450, y=660)
    
    label9 = tk.Label(appl, text='Other charges:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label9.place(x=130,y=730)
    e8 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    e8.place(x=450, y=730)
    
    # label10 = tk.Label(appl, text='Total Amount:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    # label10.place(x=130,y=785)
    # e9 = tk.Entry(appl, font=("Arial Rounded MT Bold", 25),width=10)
    # e9.place(x=450, y=785)

    b1=tk.Button(appl, command=lambda: submit(e1.get_date().strftime('%Y-%m-%d'),e2.get(),e3.get(),e4.get(),t.get("1.0", "end-1c"),e5.get(),e6.get(),e8.get(),indt.get_date().strftime('%Y-%m-%d'),appl,4), text=" Submit ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
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
        if n=='Dpr_1':
            query = f"SELECT * FROM dpra WHERE dt BETWEEN TO_DATE('{start_date}', 'YYYY-MM-DD') AND TO_DATE('{end_date}', 'YYYY-MM-DD') ORDER BY tl"
            dataframe = pd.read_sql(query, engine)
            dataframe['dt'] = dataframe['dt'].dt.date 
        elif n=='Dpr_2':
            query = f"SELECT * FROM dprb WHERE dt BETWEEN TO_DATE('{start_date}', 'YYYY-MM-DD') AND TO_DATE('{end_date}', 'YYYY-MM-DD') ORDER BY tl"
            dataframe = pd.read_sql(query, engine)
            dataframe['dt'] = dataframe['dt'].dt.date 
        else:
            query1 = f"SELECT * FROM dpra WHERE dt BETWEEN TO_DATE('{start_date}', 'YYYY-MM-DD') AND TO_DATE('{end_date}', 'YYYY-MM-DD') ORDER BY tl"
            query2 = f"SELECT * FROM dprb WHERE dt BETWEEN TO_DATE('{start_date}', 'YYYY-MM-DD') AND TO_DATE('{end_date}', 'YYYY-MM-DD') ORDER BY tl"
            df1 = pd.read_sql(query1, engine) 
            df1['dt'] = df1['dt'].dt.date 
            df1['indt'] = df1['indt'].dt.date 
            df2 = pd.read_sql(query2, engine) 
            df2['dt'] = df2['dt'].dt.date 
            df2['indt'] = df2['indt'].dt.date 
            frames = [df1, df2] 
            dataframe = pd.concat(frames) 
            dataframe = dataframe.sort_values(by=['tl'], ascending=True)

        new_column_names = {
        'dt': 'Date',
        'na': 'Name_&_Address_of_the_supplier',
        'pn': 'Purchase_order_no',
        'bn': 'Invoice/Bill_no',
        'desc': 'Description',
        'price': 'Price',
        'qn': 'Quantity',
        'tp': 'Total_Price',
        'oc': 'Other_Charges',
        'ta': 'Total_Amount',
        'tl':'Sno.',
        'indt':'Invoice_Date'}
        if fname=='':
            messagebox.showinfo("Error", "Specify a name to your file!",parent=appl)
        else:
            dataframe.rename(columns=new_column_names, inplace=True) # Keep only the date part
            # dataframe['Date']=dataframe['Date'][::].strftime('%Y-%m-%d')
            current_date = datetime.datetime.now().date()
            output_excel_file = r"D:\\VITSCHOOL\\Shabbir sir project\{}.xlsx".format(fname)
            dataframe.to_excel(output_excel_file, index=False, engine='openpyxl')
            msg_box = tk.messagebox.askquestion('Download Success', 'Excel Sheet Generated Successfully!\nDo you want to Generate Report again?',icon='info')
            if msg_box == 'no':
                appl.destroy()
            else:
                backf(appl)
                getreport()
    except Exception as e:
        messagebox.showinfo("Oops!!!", "Something went wrong!!!\nTry again!",parent=appl)
        print(e)
def backf(window):
        window.destroy()

def submit(dt, na, pn, bn, desc, pr, qn, oc, indt, win, dpr):
    try:
        print("Submit Function")
        print(dt, na, pn, bn, desc, pr, qn, oc, indt)
        dt = dt  # Date is already in the correct format ('YYYY-MM-DD')
        na = na  
        pn = pn  
        bn = int(bn)  # Convert to INTEGER
        desc = desc  # Keep it as a string
        pr = float(pr)  # Convert to FLOAT
        qn = int(qn)  # Convert to INTEGER
        tp = qn * pr  # Convert to FLOAT
        oc = float(oc)  # Convert to FLOAT
        ta = tp + oc  # Convert to FLOAT
        print(dt)
        forma = 'yyyy-mm-dd'
        con = cx_Oracle.connect('system/joel@localhost:1521/xe')
        cur = con.cursor()

        if dpr == 1:
            dprname = 'dpra'
            dprn = 'DPR - I'
            cur.execute('select * from dpra')
        elif dpr == 2:
            dprname = 'dprb'
            dprn = 'DPR - II'
            cur.execute('select * from dprb')
        elif dpr == 3:
            dprname = 'dprac2'
            dprn = 'DPR - I'
            cur.execute('select * from dprac2')
        elif dpr == 4:
            dprname = 'dprbc2'
            dprn = 'DPR - II'
            cur.execute('select * from dprbc2')
        print(dprname)
        li = [x for x in cur]
        tls = [x[-2] for x in li]
        print(tls)
        tl = max(tls) + 1
        print("Tl= ",tl, "Now inserting...")
        cur.execute('insert into {} values(to_date(\'{}\',\'{}\'),\'{}\',\'{}\',{},\'{}\',{},{},{},{},{},{},to_date(\'{}\',\'{}\'))'.format(dprname,dt,forma,na,pn,bn,desc,pr,qn,tp,oc,ta,tl,indt,forma))
        con.commit()
        print("Details inserted as a row")
        msg_box = tk.messagebox.askquestion('Insertion Success', 'Insertion success!\nDo you want to enter data again?',icon='info')
        if msg_box == 'no':
            win.destroy()
        else:
            backf(win)
            if dpr == 1:
                Entry1()
            else:
                Entry2()
    except Exception as e:
        messagebox.showinfo("Oops!!!", "Something went wrong!!!\nTry again!", parent=win)
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

    label4 = tk.Label(appl, text='Select source:', compound='center', font=("sans serif", 25),fg='white',bg='#1D5D9B')
    label4.place(x=200,y=500)
    n=tk.StringVar()
    dprv = ttk.Combobox(appl, width = 10, textvariable = n,font=("sans serif",25))
    dprv['values'] = ('Dpr_1','Dpr_2','Both')
    dprv.place(x=420,y=510)
    dprv.current(0)

    b1=tk.Button(appl, command=lambda:getdata(e1.get_date(),e2.get_date(),e3.get(),appl,dprv.get()), text=" Generate Report! ", font=("Arial Rounded MT Bold",15), fg='blue', bg='white')
    b1.place(x=700,y=785)



    pass

def mainf():
    main=tk.Tk()
    main.title("PBR VITS: COMPUTER CENTER")
    main.state('zoomed')
    main.configure(bg=('#1D5D9B'))
    label1 = tk.Label(main, text='PBR VITS: COMPUTER PURCHASE CENTER', compound='center', font=("Algerian", 35),fg='white',bg='#1D5D9B')
    label1.place(x=430,y=30)
    label2 = tk.Label(main, text='DPR REGISTERS', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label2.place(x=110,y=85)
    label2 = tk.Label(main, text='CAMPUS - 1', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label2.place(x=130,y=150)
    b1=tk.Button(command=Entry1,text=" DPR - I ENTRY ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=130,y=230)
    b2=tk.Button(command=Entry2,text=" DPR - II ENTRY ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=130,y=300)
    
    label2N = tk.Label(main, text='CAMPUS - 2', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label2N.place(x=130,y=400)
    b1=tk.Button(command=Entry1C2,text=" DPR - I ENTRY ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=130,y=480)
    b2=tk.Button(command=Entry2C2,text=" DPR - II ENTRY ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=130,y=550)

    label3 = tk.Label(main, text='Reports', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label3.place(x=680,y=150)
    b1=tk.Button(command=getreport,text=" Get Reports ",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=670,y=230)

    label3 = tk.Label(main, text='Assets Data', compound='center', font=("lucida", 25),fg='white',bg='#1D5D9B')
    label3.place(x=1150,y=150)

    b2=tk.Button(command=lambda:asset_entry("CPU"),text="CPU",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=1200,y=230)
    b3=tk.Button(command=lambda:asset_entry("Monitors"),text="Monitors",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=1200,y=300)
    b4=tk.Button(command=lambda:asset_entry("Printers"),text="Printers",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=1200,y=370)
    b5=tk.Button(command=lambda:asset_entry("UPS"),text="UPS",font=("Arial Rounded MT Bold",15),fg='blue',bg='white').place(x=1200,y=440)
    main.mainloop()
#main function
mainf()