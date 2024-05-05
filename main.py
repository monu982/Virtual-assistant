import win32com.client
import os
import speech_recognition as sr
import webbrowser
from requests import get
import datetime
import wikipedia            # for searching 
import pywhatkit as kit      
import smtplib
import sys
import pyjokes;

# to speak the text
def say(text):
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    print(text)
    speaker.Speak(text)

# to take input from user 
def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        audio=r.listen(source)
        try:
            query=r.recognize_google(audio,language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

# to wish
def wish():
    hour =int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        say("Good morning")
    elif hour>12 and hour<18:
        say("Good afternoon")
    else:
        say("Good evening")
    say("I am jarvis sir, please tell me how can i help you")

# to send email
def sendEmail(to,content):
    server=smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login('manishsinghb004@gmail.com','Manish@12345')
    # server.sendmail('manishsinghb004@gmail.com',to,content)
    server.close()


if __name__ == '__main__':
    wish()
    while True:
        print("Listening...")
        query=takeCommand().lower()

        sites=[["youtube","https://youtube.com"]]
        for site in sites:
            if f"open {site[0]}" in query:
                say(f"Opening {site[0]} sir..")
                webbrowser.open(site[1])
            # TODO : add more website
        
    
        if "favourite music" in query or "favourite song" in query:
            musicPath="C:/Users/manis.LAPTOP-MANI/Desktop/projects/jarvis/music/sweet.mp3"
            os.system(musicPath)

        if "notepad" in query:
            # open notepad
            os.system("notepad")

        if "camera" in query:
            # open camera
            os.system("start microsoft.windows.camera:")

        if "time" in query:
            # need update
            strftime=datetime.datetime.now().strftime("%H:%M")
            say(f"Sir the time is {strftime}")

        if "ip address" in query:
            # Find my ip
            ip=get('https://api.ipify.org').text
            print(ip)
            say(f"your IP address is {ip}")

        if "wikipedia" in query:
            # to search in wikipedia
            say("searching wikipedia...")
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query,sentences=2)
            print(results)
            say("according to wikipedia")
            say(results)

        if "open search" in query:
            # open search
            say("sir,what should i search on google")
            search=takeCommand().lower()
            webbrowser.open(f"{search}")

        if "send message" in query:
            # send whatsapp message
            kit.sendwhatmsg("+919599323982","this is testing protocol",2,3)
            # TODO : NEED TO UPDATE WITH MORE NEW TRICK 

        if "play song" in query or "play music":
            # kit.playonyt("see you again")
            say("sir,what should i play on youtube")
            search=takeCommand().lower()
            kit.playonyt(search)
            # TODO : need to more update with more trick

        if "send email" in query:
            # send email with my email
            try:
                say("what should i say?")
                content=takeCommand().lower()
                to="manisharmy004@gmail.com"
                sendEmail(to,content)
                say("Email has been sent to manish")

            except Exception as e:
                print(e)
                say("sorry sir, i am not able to sent this mail to manish")

        if "turn off" in query:
            # to turn off 
            say("thanks for using me sir, have a good day.")
            sys.exit()

        # to close any application
        if "close notepad" in query:
            say("okay sir,closing notepad")
            os.system("taskkill /f /im notepad.exe")

        # to find a joke
        if "tell me a joke" in query:
            joke=pyjokes.get_joke(language="en")
            say(joke)

        if "shutdown" in query:
            os.system("shutdown /s /t 5")

        if "restart" in query:
            os.system("shutdown /r /t 5")
        
        if "sleep" in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")