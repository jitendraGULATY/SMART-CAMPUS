import pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import webbrowser
import os
import smtplib
import serial
import xlsxwriter
import xlrd 
import pandas as pd
engine = pyttsx3.init()
loc="C:/Users/jitendraprakash/Desktop/Bo12.xlsx"
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

workbook = xlsxwriter.Workbook('abc1.xlsx')
worksheet = workbook.add_worksheet("Sheet1")
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0)



Arduino_Serial = serial.Serial('com10',9600)
print(Arduino_Serial.readline())


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am S R M Connect. Please tell me how may I help you")
def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

        try:
            print("Recognizing...")    
            query = r.recognize_google(audio)
            print("User said:{}".format(query))

        except:
            #print(e)    
            print("Say that again please...")  
            return "None"
        return query
def takeAttendance():
    s=[]
    for i in range(sheet.nrows): 
                speak(sheet.cell_value(i, 0))  
                r = sr.Recognizer()
                
                with sr.Microphone() as source:
                    print("Listening...")
                    r.pause_threshold = 1
                    audio = r.listen(source)
                    
                    try:
                        print("Recognizing...")    
                        roll = r.recognize_google(audio)
                        print("User said:{}".format(roll))
                        s.append(roll)
                        print(s)
                        

                    
                       
                                                
                    except:
                        return 0 
    row=0
    col=0
    for i in range(sheet.nrows): 
    
            worksheet.write(row, col, sheet.cell_value(i,0)) 
            worksheet.write(row, col + 1, s[i])
            row+=1
    workbook.close()        
    

if __name__ == "__main__":
    wishMe()
    
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'lights on' in query:
            speak('Turninglights on...')
            id="1"
            Arduino_Serial.write(id.encode())
        if 'lights off' in query:
            speak('Turninglights off...')
            id1="0"
            Arduino_Serial.write(id1.encode())
        if 'fan on' in query:
            speak('Turningfans on...')
            id2="2"
            Arduino_Serial.write(id2.encode())
        if 'fan off' in query:
            speak('Turningfans off...')
            id3="3"
            Arduino_Serial.write(id3.encode())
        
        if 'take attendance' in query:
            speak('TakingAttendance...')
            takeAttendance()
                        

       

            #query = query.replace("wikipedia", "")
            #results = wikipedia.summary(query, sentences=2)
            #speak("According to Wikipedia")
            #print(results)
            #speak(results)


