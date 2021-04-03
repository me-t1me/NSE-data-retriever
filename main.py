from datetime import date
from nsepy import get_history
from openpyxl import Workbook
import datetime
import pandas as pd 
from tkinter import *
import sys 


# * GUi 
window = Tk()
title = window.title("NSE data retriever")

# * labels
l1 = Label(window , text = "Start Date")
l1.grid(row = 0 , column = 0)

l1 = Label(window , text = "End Date")
l1.grid(row = 1 , column = 0)

l1 = Label(window , text = "Type")
l1.grid(row = 2 , column = 0)

l1 = Label(window , text = "Name of file")
l1.grid(row = 3 , column = 0)

l1 = Label(window , text = "Company Name")
l1.grid(row = 4 , column = 0)

l1 = Label(window , text = "(YYYY-MM-DD)")
l1.grid(row = 0 , column = 2)

l1 = Label(window , text = "(YYYY-MM-DD)")
l1.grid(row = 1, column = 2)

l1 = Label(window , text = "0 = current month, 1 = near month, 2=far month")
l1.grid(row = 2 , column = 2)

l1 = Label(window , text = "ex - current")
l1.grid(row = 3 , column = 2)

l1 = Label(window , text = "ex - BOSCHLTD")
l1.grid(row = 4 , column = 2)

l1 = Label(window , text = "close program to generate file" )
l1.grid(row = 5 , column = 1)

l1 = Label(window , text = "File will be created in folder where NSE data retriever.exe file is located")
l1.grid(row = 6 , column = 0)
# * lebals ends

# * entry field
l1_text = StringVar()
el1 = Entry(window , textvariable=l1_text)
el1.grid(row=0 , column = 1)

l2_text = StringVar()
el2 = Entry(window , textvariable=l2_text)
el2.grid(row=1 , column = 1)

l3_text = IntVar()
el3 = Entry(window , textvariable=l3_text)
el3.grid(row=2 , column = 1)

l4_text = StringVar()
el4 = Entry(window , textvariable=l4_text)
el4.grid(row=3 , column = 1)

l5_text = StringVar()
el5 = Entry(window , textvariable=l5_text)
el5.grid(row=4 , column = 1)
# * entry field ends

window.mainloop()
# * Gui ends



# * input 
type =l3_text.get()
startdate = date(int(l1_text.get()[0:4]),int(l1_text.get()[5:7]),int(l1_text.get()[8:10]))
lastmon = int(l2_text.get()[5:7])
name = l4_text.get()
companyname = l5_text.get()
# * input end 



# * main function
workbook = Workbook()
sheet = workbook.active

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    return next_month - datetime.timedelta(days=next_month.day)

def exdate(month,year):
    if month >= 13 :
            month = month - 12
            year = year + 1
    today = last_day_of_month(date(year,month,1))
    idx = today.isoweekday() - 4
    if idx < 0 :
      idx = idx + 7
    return today - datetime.timedelta(days=idx)



def datestart (month,year):
    if month == 3: 
      stt = startdate 
    else :
        if month == 1 :
            month = 12
            year = year - 1
        stt = exdate(month - 1,year) + datetime.timedelta(days=1)
    return stt

total = get_history(
                   symbol=companyname,
                   start=datestart(startdate.month,startdate.year),
                   end=exdate(startdate.month,startdate.year),
                   futures=True,
                   expiry_date = exdate(startdate.month + type,startdate.year)
                   )
if type == 0:
    for i in range(startdate.month + 1, 13 + lastmon):
        sbin1 = get_history(symbol=companyname,
                   start=datestart(i,startdate.year),
                   end=exdate(i,startdate.year),
                   futures=True,
                   expiry_date = exdate(i+type,startdate.year))
        total = pd.concat([total , sbin1])
elif type == 1:
    for i in range(startdate.month + 1 , 12 + lastmon):
        sbin1 = get_history(symbol=companyname,
                   start=datestart(i,startdate.year),
                   end=exdate(i,startdate.year),
                   futures=True,
                   expiry_date = exdate(i+type,startdate.year))
        total = pd.concat([total , sbin1])
elif type == 2:
    for i in range(startdate.month + 1, 11 + lastmon):
        sbin1 = get_history(symbol=companyname,
                   start=datestart(i,startdate.year),
                   end=exdate(i,startdate.year),
                   futures=True,
                   expiry_date = exdate(i+type,startdate.year))
        total = pd.concat([total , sbin1])

total.to_excel(name + ".xlsx")

# * main function ends

sys.exit()
