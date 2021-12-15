import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import subprocess
import ctypes
import requests
import json
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    
    
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good Evening!") 
        
    speak(" I am Ruby sir. Please tell me how may I help you")       

def takeCommand():
    r= sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold =1
        audio = r.listen(source)
        
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query} \n)")
    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('shuvro171115071@gmail.com', '@Shuvrodas13')
    server.sendmail('shuvro171115071@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
        wishMe()
        while True:
        #if 1:
            query = takeCommand().lower()
            
            if 'wikipedia' in query:
                speak('Searching Wikipedia...')
                query= query.replace("wikipedia", "")
                #results = wikipedia.summary(query, sentences=5)
                results = wikipedia.summary(f'{query}', sentences=5 )
                speak("According to wikipedia...")
                print(results)
                speak(results)
          
            elif 'open youtube' in query:
                webbrowser.open("youtube.com")
                
            elif 'open google' in query:
                webbrowser.open("google.com")
                
            elif 'open stackoverflow' in query:
                webbrowser.open("stackoverflow.com")
            
            elif 'play music' in query:
                music_dir = 'D:\\music\\songs'
                songs = os.listdir(music_dir)
                print(songs)
                os.startfile(os.path.join(music_dir, songs[0]))
            
            elif 'the time' in query:
                strTime = datetime.datetime.now().strftime("%H PM %M minit %S")
                speak(f"sir, the time is{strTime}")
            
            elif 'open code' in query:
                codePath = "C:\\Users\\Hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                os.startfile(codePath)
            
            elif 'email' in query:
                try:
                    speak("what should i say?")
                    content = takeCommand()
                    to = "shuvro6589@gmail.com"
                    sendEmail(to, content)
                    speak("Email has been sent")
                except Exception as e:
                    print(e)
                    speak("Sorry my friend Shuvro bhai. I am not able to send this mail")
                

            elif 'send a mail' in query:
                try:
                   speak("What should I say?")
                   content = takeCommand()
                   speak("whome should i send")
                   to = input()   
                   sendEmail(to, content)
                   speak("Email has been sent !")
                except Exception as e:
                 print(e)
                speak("I am not able to send this email")
            
            elif 'how are you' in query:
                 speak("I am fine, Thank you")
                 speak("How are you, Sir")
            
                
            elif "where is" in query:
                query = query.replace("where is", "")
                location = query
                speak("User asked to Locate")
                speak(location)
                webbrowser.open("https://www.google.nl/maps/place/" + location + "")
                
            elif "weather" in query:
                 
                # Google Open weather website
                # to get API of Open weather
                api_key = "ac004c60fedd947e7a665fb8ea1de302"
                base_url = "https://api.openweathermap.org/data/2.5/weather?"
                
                speak(" City name ")
                print("City name : ")
                city_name = takeCommand()
                complete_url = base_url + "appid=" + api_key + "&q=" + city_name
                response = requests.get(complete_url)
                x = response.json()
                
                if x["cod"] != "404":
                    y = x["main"]
                    current_temperature = y["temp"]
                    current_pressure = y["pressure"]
                    current_humidiy = y["humidity"]
                    z = x["weather"]
                    weather_description = z[0]["description"]
                    print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                    speak(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                else:
                    speak(" City Not Found ")
            elif 'lock window' in query:
                    speak("locking the device")
                    ctypes.windll.user32.LockWorkStation()
                    
            elif 'project presentation' in query:
                    speak("opening project presentation")
                    power = r"C:\\Users\\Hp\\Desktop\\voc (1).pptx"
                    os.startfile(power)
                    
                
                

