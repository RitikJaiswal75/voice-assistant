from datetime import datetime  # import is used for reading date time from system
import pyttsx3  # used for speaking
import requests  # requesting url's
import json  # processes json files returned by api
import geocoder  # used to return location of latitude and longitude of the user
from dateutil import tz  # used for conversion of uct to ist
import speech_recognition as sr  # speech recognition module
import wikipedia  # search and return wikipedia results
import webbrowser  # opening something in web browser
import os  # for functions related to operating systems
import ezgmail  # for sending and reading mails
import random   # generating random numbers
from bs4 import BeautifulSoup  # parsing html pages that are sent in emails
import wolframalpha

# converts uct to ist as the weather api returns uct
def tc(t):
    t = int(t)
    utc = datetime.fromtimestamp(t)
    itc_time = utc.astimezone(tz.gettz('ITC'))
    return itc_time


# converts kelvin to celcius
def tempConverter(temp):
    temp = int(temp)
    temp = temp-273.15
    return temp


# function to tell weather
def weather():
    g = geocoder.ip('me')
    lat = g.latlng[0]
    lon = g.latlng[1]
    api_key = "Your_OWM_API_key"
    base_url = "https://api.openweathermap.org/data/2.5/onecall?"
    url = base_url + "lat="+str(lat)+"&lon="+str(lon) + \
        "&exclude=hourly,daily&appid=" + api_key
    response = requests.get(url)
    response.raise_for_status()
    w = json.loads(response.text)
    print('date = '+str(tc(w['current']['dt'])))
    print("sunrise = "+str(tc(w['current']['sunrise'])))
    print("sunset = "+str(tc(w['current']['sunset'])))
    print("Temperature = "+str(round(tempConverter(w['current']['temp']), 2)))
    print("feels like = " +
          str(round(tempConverter(w['current']['feels_like']), 2)))
    speak("Temperature is " +
          str(round(tempConverter(w['current']['temp']), 2))+" Degree Celcius")
    speak("feels like " +
          str(round(tempConverter(w['current']['feels_like']), 2)))


# Function to authorize gmail
# changing webbrouser to chrome to open links in chrome
try:
    webbrowser.register('chrome',
                        None,
                        webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
except:
    pass


# to set up the voice engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty("rate", 175)


# to make the assistant speak
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


# to greet the user when he starts to use the AI
def wishMe():
    hour = int(datetime.now().hour)
    if hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak(" I am Ishi sir. Please tell me how may I help you")


# to take speach input and then convert it into text query
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("User Input: ", query)
    except:
        print("sorry... I couldn't catch that. please can you Say it again...")
        speak("sorry... I couldn't catch that. please can you Say it again...")
        return 'None'
    return query


# to tell the time to the user
def time():
    now = datetime.now()
    h = int(now.hour)
    m = int(now.minute)
    ans = "The Time is "+str(h)+" : "+str(m)
    print(ans)
    speak(ans)


# to search on wikipedia
def wikiSearch(query):
    print("Searching Wikipedia...")
    speak("Searching Wikipedia...")
    query.replace("Wikipedia", "")
    results = wikipedia.summary(query, sentences=2)
    print("According to Wikipedia")
    speak("According to Wikipedia")
    print(results)
    speak(results)


# to open something in webpage thats not installed by adding '.com' to it
def webAccess(query):
    if ".com" not in query:
        query = query.split()
        l = query.index("open")+1
        query[l] = query[l]+".com"
    l = query.index("open")+1
    print("Opening "+query[l])
    speak("Opening "+query[l])
    try:
        webbrowser.get('chrome').open(query[l])
    except:
        webbrowser.open(query[l])


# to open different apps
def app(query):
    pathApp = "C:\\Program Files\\Microsoft Office\\root\\Office16"
    if "access" in query:
        print("Opening Microsoft Access")
        speak("Opening Microsoft Access")
        os.startfile(os.path.join(pathApp, "MSACCESS.EXE"))
    elif "excel" in query:
        os.startfile(os.path.join(pathApp, "EXCEL.EXE"))
        print("Opening Microsoft Excel")
        speak("Opening Microsoft Excel")
    elif "one note" in query:
        os.startfile(os.path.join(pathApp, "ONENOTE.EXE"))
    elif "outlook" in query:
        print("Opening Outlook")
        speak("Opening Outlook")
        os.startfile(os.path.join(pathApp, "OUTLOOK.EXE"))
    elif "powerpoint" in query:
        print("Opening Microsoft PowerPoint")
        speak("Opening Microsoft PowerPoint")
        os.startfile(os.path.join(pathApp, "POWERPNT.EXE"))
    elif "publisher" in query:
        print("Opening MicroSoft Publisher")
        speak("Opening MicroSoft Publisher")
        os.startfile(os.path.join(pathApp, "MSPUB.EXE"))
    elif "word" in query:
        print("Opening MicroSoft Word")
        speak("Opening MicroSoft Word")
        os.startfile(os.path.join(pathApp, "WINWORD.EXE"))
    elif "chrome" in query:
        os.startfile(
            "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")
    elif "edge" in query:
        print("Opening Edge")
        speak("Opening Edge")
        os.startfile(
            "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
    elif "settings" in query:
        print("Opening Settings")
        speak("Opening Settings")
        os.system("start ms-settings:")
        os.system("exit")
    elif "notepad" in query:
        print("Opening Notepad")
        speak("Opening Notepad")
        os.system("notepad")
        os.system("exit")
    elif "control" in query:
        print("Opening Control Panel")
        speak("Opening Control Panel")
        os.system("control")
        os.system("exit")
    else:
        print("I did not find an app like that one. trying for it's website...")
        speak("I did not find an app like that one. trying for it's website...")
        webAccess(query)


# to play music
def music():
    print("This function is under development")
    speak("This function is under development")


# to send emails from authorized mail id
def sendEmail():
    try:
        speak("Please type in the email of recipitant")
        to = input("Enter the email: ")
        if to == "self":
            to = "ritik12103032000@gmail.com"
        print('Tell me the subject of the mail')
        speak('Tell me the subject of the mail')
        subject = takeCommand()
        print('Tell me the body of the mail')
        speak('Tell me the body of the mail')
        body = takeCommand()
        ezgmail.send(to, subject, body)
        print("Email Sent")
        speak("Email Sent")
    except:
        print("I could not get authorized please give me permission")
        speak("I could not get authorized please give me permission")


# to read any 1 random email from your gmail
def readEmail():
    unreadThreads = ezgmail.unread()
    t = random.randint(0, len(unreadThreads)-1)
    print(str(unreadThreads[t].messages[0].sender))
    speak(str(unreadThreads[t].messages[0].sender))
    print(str(unreadThreads[t].messages[0].subject))
    speak(str(unreadThreads[t].messages[0].subject))
    try:
        soup = BeautifulSoup(unreadThreads[t].messages[0].body[:200], "lxml")
        print(soup.get_text())
        speak(soup.get_text())
    except:
        print(str(unreadThreads[t].messages[0].body))
        speak(str(unreadThreads[t].messages[0].body))


# function to beatbox
def BeatBox():
    engine.setProperty("rate", 250)
    l = ['nope ', "i can't "]
    print("Nope i can't")
    for s in range(5):
        str = ''
        speak("Nope i can't")
        for k in range(random.randint(1, 4)):
            engine.setProperty("rate", random.randint(300, 500))
            i = random.randint(0, 1)
            str = str+l[i]
            speak(str)
    engine.setProperty("rate", 175)

    
def helping(query):
    app_id='YOUR_WOLFRAM_ALPHA_APP_ID'
    client = wolframalpha.Client(app_id) 
    res = client.query(query)
    try:
        answer = next(res.results).text 
    except:
        answer="Could not find"
    print(answer)
    speak(answer)

# Main function
if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()
        if "what" in query and "time" in query:
            time()
        elif "wikipedia" in query:
            wikiSearch(query)
        elif "open" in query:
            try:
                app(query)
            except:
                webAccess(query)
        elif "email" in query:
            if "send" in query:
                sendEmail()
            elif "read" in query:
                readEmail()
        elif "weather" in query:
            weather()
        elif "exit" in query or "bye" in query:
            print("Thank you for using me! Have a nice time.")
            speak("Thank you for using me! Have a nice time.")
            exit()
        else: 
            helping(query)
        '''
        elif "beatbox" in query and 'me' in query:
            BeatBox()
        elif "play" in query and "music" in query:
            music()
        '''