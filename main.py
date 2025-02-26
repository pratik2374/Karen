from datetime import datetime
import win32com.client
import speech_recognition as sr
import webbrowser
import os


speaker = win32com.client.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()
speaker.Voice = voices[1]

# setup to run youtube
from my_package import tubePlay,tubeSearch,tubeStart
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
DriverPATHChrom = r"C:\Program Files (x86)\chromedriver.exe"
service = Service(DriverPATHChrom)
options = Options()
options.add_argument(r"user-data-dir=C:\Users\ABC\AppData\Local\Google\Chrome\User Data")
options.add_argument("--profile-directory=Profile 8")
driver = webdriver.Chrome(service=service, options=options)

#Groq modules
from my_package.Extractors.extract import extractnum,extractsearch,classify,ytclassify

#Quotes gen
from my_package.quotes.quote import quote_gen

def speak(text):
    speaker.Speak(text)

def is_command_in_query(command, query):
    return command.lower() in query.lower()

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio)
            print(f"User said: {query}")
            return query
        except Exception as e:
            print(f"Error: {e}")
            return "Some Error Occurred"

speak(
   "You know we could Die anytime and there's some million possibilities, still....., Sir! I'm KAREN, always ready "
   "to assist to in Best way I can..")

def youtube(query) :
    option = ytclassify(query=query)
    # optionA :user wants just open youtube
    # optionB:user wants just search something on youtube
    # optionC:user wants to open youtube and serach something
    # optionD:user wants to play a video
    if is_command_in_query("a", option) :
        tubeStart(driver=driver)
    if is_command_in_query("b" , option) :
        search=extractsearch(query=query)
        tubeSearch(driver=driver, search_query=search)
    if is_command_in_query("c", option) :
        search=extractsearch(query=query)
        tubeStart(driver=driver)
        tubeSearch(driver=driver, search_query=search)
    if is_command_in_query("d" , option) :
        number= extractnum(query=query)
        tubePlay(driver=driver, num=number)

def openGapps(query) :
    sites = [
        ["google", "https://google.com"],
        ["classroom", "https://classroom.google.com/"]
    ]

    for site in sites:
        if is_command_in_query(f"open {site[0]}", query):
            speak(f"Opening {site[0]} master")
            webbrowser.open(site[1])

def openSapps(query) :
    if "arduino" in query.lower():
        os.startfile(r"C:\Program Files\Arduino IDE\Arduino IDE.exe")
    elif "open music" in query.lower():
        path = r"D:\Self\Songs\Naruto-Blue-Bird.mp3"
        os.startfile(path)

def time() :
    current_time = datetime.now()
    Hours = current_time.hour % 12 or 12
    min = current_time.strftime("%M")
    period = "AM" if current_time.hour < 12 else "PM"
    speak(f"The time is {Hours} {min} {period}")

def quotes() :
    speak(quote_gen())

while True:
    query = takeCommand()
    """
    option1:user wants to play any video or song on youtube
    option2:user wants to open google or classroom appliaction
    option3:user wamts to open another application
    option4:user wants to know time
    option5:user wants to know something just an query
    option6:user wants to quit
    option7:user wants to know a quote
    """
    catogory = classify(query=query)

    if (is_command_in_query("1", catogory)) : 
        youtube(query)
    elif (is_command_in_query("2", catogory)) :
        openGapps(query)
    elif (is_command_in_query("3", catogory)) :
        openGapps(query)
    elif (is_command_in_query("4", catogory)) :
        time()
    elif(is_command_in_query("5", catogory)) :
        continue
    elif (is_command_in_query("7", catogory)) :
        quotes()
    elif(is_command_in_query("6", catogory)) :
        speak("Goodbye!, Master")
        break
    else :
        speak("Some Error occured master")

    

        

