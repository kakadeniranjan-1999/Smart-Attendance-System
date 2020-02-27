import cv2
import numpy as np
import os 
import speech_recognition as sr
import win32com.client as wincl
import time
import xlrd,xlutils,xlwt
from xlutils.copy import copy
from xlwt import Workbook
import threading
import datetime

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

def new_entry():
    wb=xlrd.open_workbook("q.xls")
    sheet=wb.sheet_by_index(0)
    wb1=copy(wb)
    vsheet=wb1.get_sheet(0)
    i=(sheet.nrows)
    speak.Speak("Please speak your roll no:")
    q=r.listen(source)
    roll=r.recognize_google(q)
    print("You said :{}".format(roll))
    speak.Speak("Please speak your name:")
    w=r.listen(source)
    name=r.recognize_google(w)
    print("You said :{}".format(name))
    speak.Speak("The roll no you spoke is"+roll)
    speak.Speak("The name you spoke is"+name)
    speak.Speak("Are your details correct? please say yes or no...")
    t=r.listen(source)
    s=r.recognize_google(t)
    print(s)
    if(s=="yes"):
        vsheet.write(i,0,roll)
        vsheet.write(i,1,name)
        names.append(name)
        wb1.save("q.xls")
        face_dataset.photo_entry(roll)
        speak.Speak("your details are successfully added.")
    else:
        speak.Speak("Please try again!!!")
        new_entry()
def start_camera():
    
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.read('trainer/trainer.yml')
    cascadePath = "haarcascade_frontalface_default.xml"
    faceCascade = cv2.CascadeClassifier(cascadePath);

    font = cv2.FONT_HERSHEY_SIMPLEX
    

    names = ['niranjan', 'swapnil', 'aadesh','prathamesh']

    # Initialize and start realtime video capture
    speak.Speak("Please look towards the camera")
    cam = cv2.VideoCapture(0)
    cam.set(3, 320) # set video widht   
    cam.set(4, 240) # set video height

    # Define min window size to be recognized as a face
    minW = 0.1*cam.get(3)
    minH = 0.1*cam.get(4)

    while True:

        ret,img =cam.read()

        cv2.imshow('camera1',img)

        gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
        

        faces = faceCascade.detectMultiScale(gray,scaleFactor = 1.2,minNeighbors = 5,minSize = (int(minW), int(minH)),)

        for(x,y,w,h) in faces:

            cv2.rectangle(img, (x,y), (x+w,y+h), (0,255,0), 2)

            id, confidence = recognizer.predict(gray[y:y+h,x:x+w])
            # Check if confidence is less them 100 ==> "0" is perfect match 
            d=names[id]
            if (50<confidence<100):
                compare(d)
            else:
                d = "unknown"
                confidence = "  {0}%".format(round(100 - confidence))
                print('unknown')
                speak.Speak("Unknown face detected")
            cv2.putText(img, str(d), (x+5,y-5), font, 1, (255,255,255), 2)
            cv2.putText(img, str(confidence), (x+5,y+h-5), font, 1, (255,255,0), 1)  
        cv2.imshow('camera',img)
        #y=threading.Thread(target=v(img))
        #y.start()
       
        k = cv2.waitKey(10) & 0xff # Press 'ESC' for exiting video
        if k == 27:
            break
        time.sleep(2)
    
    # Do a bit of cleanup
    print("\n [INFO] Exiting Program and cleanup stuff")
    cam.release()
    cv2.destroyAllWindows()
    speak.Speak("your attendance is marked"+"thank you"+"visit again")
    

r = sr.Recognizer()
with sr.Microphone() as source:
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak("Welcome to my world!!!")
    time.sleep(0.5)
    print("01.New student entry or 02.Start attendance")
    speak.Speak("Say 01 for new student entry or 02 for starting attendance")
    print("Listening...")
    audio = r.listen(source)
    try:
        text = r.recognize_google(audio)
        print("You said :{}".format(text))
        if(text=="01"):
            speak.Speak("Welcome to new student entry")
            new_entry()
        elif(text=="02"):
            speak.Speak("Please wait till your camera is starting...")
            start_camera()
        else:
            print("Oops!!!Wrong entry!!!")
            speak.Speak("Oops!!!Wrong entry!!!")
    except:
        print("Sorry could not recognize what you said")