import xlrd,xlutils,xlwt
from xlutils.copy import copy
from xlwt import Workbook
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

def compare(o):
    l=0
    w=xlrd.open_workbook("Attendance_mark.xlsx")
    ksheet=w.sheet_by_index(0)
    w1=copy(w)
    hsheet=w1.get_sheet(0)
    p=(ksheet.nrows)
    for i in range(1,ksheet.nrows):
        k=ksheet.cell_value(i,1)
        if(k==o):
            print("Already marked present")
            speak.Speak("Hello"+o+"!!!"+"You are already marked present!!")
            l=1
    if(l==0):
        speak.Speak("Hello"+o+"!!!"+"You are marked present!!")
        print(o)
        hsheet.write(p+1,1,o)
        hsheet.write(p+1,2,(datetime.date()))
        hsheet.write(p+1,3,(datetime.datetime.now().time))