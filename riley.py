import pyttsx3  # text to speech
import speech_recognition as sr  # speech recognition
import datetime
import wikipedia  # wikipedia
import webbrowser  # perform web search
import os
import re
import math
import random
import time
from PIL import ImageGrab  # ss
import pygetwindow as gw  # detect window
from PyDictionary import PyDictionary  # dict
import pyjokes  # jokes
import pandas as pd
import requests  # GET and POSTS requests
import bs4  # making a beautiful soup
from ctypes import cast, POINTER  # for volume settings
from comtypes import CLSCTX_ALL  # for volume settings
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # for volume settings

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour in range(6, 12):
        speak("good morning!")

    elif hour in range(12, 17):
        speak("good afternoon!")

    elif hour in range(17, 24):
        speak("good evening!")

    else:
        speak("you shouldn't be awake")
        speak("would you like some help falling asleep?")
        answer = takeCommand().lower()
        if 'yes' in answer:
            webbrowser.open("https://www.youtube.com/watch?v=FjHGZj2IjBk")
            exit()


def takeCommand():
    r = sr.Recognizer()

    while True:
        with sr.Microphone() as source:
            print("\nListening...")
            r.pause_threshold = 1
            r.energy_threshold = 750
            audio = r.listen(source)

        try:
            print("Recognising...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            return query

        except:
            speak("say that again please")


def vol(a):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # get current volume
    currentVolumeDb = volume.GetMasterVolumeLevel()

    if a.startswith('i'):
        volume.SetMasterVolumeLevel(currentVolumeDb + 6.0, None)
        speak("volume has been increased")
    else:
        volume.SetMasterVolumeLevel(currentVolumeDb - 6.0, None)
        speak("volume has been decreased")


def top_rated_movies():
    r = requests.get("https://www.imdb.com/chart/moviemeter")
    soup = bs4.BeautifulSoup(r.text, "lxml")

    titleRatings = dict()

    titles = soup.select(".titleColumn a")
    ratings = soup.select(".imdbRating")

    for i in range(5):
        if ratings[i].text.strip('\n') == '':
            titleRatings[titles[i].text] = 'Unreleased'
        else:
            titleRatings[titles[i].text] = ratings[i].text.strip('\n')

    print(pd.DataFrame.from_dict(titleRatings,
                                 orient='index', columns=['RATINGS']))


def calc():
    while True:
        speak("Which function do you want to access?")
        query = takeCommand().lower()

        # addition
        if "+" in query or "add" in query:
            temp = re.findall(r'\d+', query)
            res = list(map(int, temp))
            if len(res) == 0:
                speak("please tell 2 numbers")
                aga = takeCommand()
                temp1 = re.findall(r'\d+', aga)
                res1 = list(map(int, temp1))
                print(math.fsum(res1))
                speak(math.fsum(res1))
            else:
                print(math.fsum(res))
                speak(math.fsum(res))
            break

        # subtract
        elif "-" in query or "subtract" in query:
            temp = re.findall(r'\d+', query)
            res = list(map(int, temp))
            if len(res) == 0:
                speak("please tell 2 numbers")
                aga = takeCommand()
                temp1 = re.findall(r'\d+', aga)
                res1 = list(map(int, temp1))
                print(res1[0] - res1[1])
                speak(res1[0] - res1[1])
            else:
                print(res[0] - res[1])
                speak(res[0] - res[1])
            break

        # multiply
        elif "x" in query or "multiply" in query or "into" in query or "product" in query:
            temp = re.findall(r'\d+', query)
            res = list(map(int, temp))
            if len(res) == 0:
                speak("please tell 2 numbers")
                aga = takeCommand()
                temp1 = re.findall(r'\d+', aga)
                res1 = list(map(int, temp1))
                print(math.prod(res1))
                speak(math.prod(res1))
            else:
                print(math.prod(res))
                speak(math.prod(res))
            break

        # division
        elif "/" in query or "divide" in query or "by" in query:
            temp = re.findall(r'\d+', query)
            res = list(map(int, temp))
            if len(res) == 0:
                speak("please tell 2 numbers")
                aga = takeCommand()
                temp1 = re.findall(r'\d+', aga)
                res1 = list(map(int, temp1))
                print(res1[0]/res1[1])
                speak(res1[0]/res1[1])
            else:
                print(res[0]/res[1])
                speak(res[0]/res[1])
            break

        # unknown function mentioned
        else:
            speak("unknown function!")


def ss():
    now = datetime.datetime.now()
    dt = now.strftime("%d-%m-%Y %H-%M-%S")

    assist = gw.getWindowsWithTitle('py')[0]
    assist.minimize()

    time.sleep(1)

    im = ImageGrab.grab()
    im.save(f"screenshots\\ss_{dt}.jpg")

    assist.restore()


if __name__ == "__main__":
    # assistant name
    assname = 'Riley'
    speak(f"hello! i am {assname}")

    # greeting
    wishMe()

    while True:

        # take commands
        time.sleep(1)
        speak("what can I do for you today?")
        query = takeCommand().lower()

        # list of all commands
        if 'help' in query:
            commands = {'help': 'List all possible commands',
                        'name': 'Change assistant name',
                        'increase volume': 'Double the current volume',
                        'decrease volume': 'Half the current volume',
                        'date': 'Check the current date',
                        'time': 'Check the current time',
                        'reminder': 'Set a reminder',
                        'wikipedia': 'Search Wikipedia',
                        'where is...': 'Find a specific location',
                        'open youtube': 'To open youtube',
                        'open meet': 'To open google meet',
                        'bored': 'Helps you with your boredom',
                        'top rated movies': 'Shows top 5 rated movies from IMDb',
                        'write a note': 'Takes a note of whatever you say',
                        'show note': 'Shows notes you took that day',
                        'meaning': 'Gives you the meaning of any word',
                        'synonym': 'Gives you the synonym of any word',
                        'antonym': 'Gives you the antonym of any word',
                        'calculator': 'Perform basic calculator functions (+, -, *, /)',
                        'screenshot': 'Take a screenshot of your screen',
                        'toss': 'Performs a coin toss',
                        'joke': 'Tells a joke',
                        'guess': 'Starts number guessing game',
                        'don\'t listen': 'Stops listening for the specified time',
                        'bye or exit': 'Assistant closes',
                        'go to sleep': 'Assistant goes to sleep'}

            speak("here is the list of commands")
            print(pd.DataFrame.from_dict(
                commands, orient='index', columns=['WHAT IT DOES']))

        # ass name
        elif "what is your name" in query:
            print("My name is", assname)
            speak(f"my name is {assname}")

        # volume control
        # increase volume
        elif "increase volume" in query:
            while True:
                vol("increase")
                speak("would you like to increase the volume further?")
                more = takeCommand().lower()
                if more == 'no':
                    speak("okay")
                    break

        # decrease volume
        elif "decrease volume" in query:
            while True:
                vol("decrease")
                speak("would you like to decrease the volume further?")
                more = takeCommand().lower()
                if more == 'no':
                    speak("okay")
                    break

        # tell the current date
        elif 'date' in query:
            strDate = datetime.datetime.now().strftime("%d/%m/%Y")
            print(strDate)
            speak(f"the date today is {strDate}")

        # tell the current time
        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(strTime)
            speak(f"The current time is {strTime}")

        # set a reminder
        elif 'reminder' in query:
            speak("what shall I remind you about?")
            rem = takeCommand()
            speak("in how many minutes?")
            mins = takeCommand()
            rem = datetime.datetime.now()+datetime.timedelta(mins)
            speak(rem)

        # search for something on wikipedia
        elif 'wikipedia' in query:
            speak("searching wikipedia")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("according to Wikipedia")
            speak(results)

        # open youtube
        elif 'open youtube' in query:
            speak("opening youtube")
            webbrowser.open("https://www.youtube.com/")

        # eliminating boredom
        elif 'bored' in query:
            speak("here are some games made specially for you")
            webbrowser.open("https://pythongames.ikathuria.repl.run/")

        # top rated movies
        elif 'top rated movies' in query:
            speak("showing top rated movies from I M D b")
            top_rated_movies()

        # take a note
        elif "write a note" in query or "make a note" in query or "take note" in query:
            speak("what should i write?")
            note = takeCommand()

            dt = datetime.datetime.now()

            with open(f'notes\\{dt.strftime("%d-%m-%Y")}.txt', 'a') as file:
                file.write(dt.strftime("%H:%M:%S"))
                file.write(" - ")
                file.write(note)

            print("Done")

        # show today's notes
        elif "show note" in query or "read note" in query:
            speak("which date's note would you like to see?")

            noteDate = input("Type in the date (dd-mm-yyyy): ")

            if re.search(r"\d\d-\d\d-\d\d\d\d", noteDate):
                with open(f'notes\\{noteDate}.txt', "r") as file:
                    temp = file.read()
                    print(temp)
                    speak(temp)
            else:
                speak("invalid date format.")

        # dictionary commands - meaning, synonym, antonym
        elif 'meaning' in query:
            speak("can you please repeat the word")
            word = takeCommand()
            speak(f"the meaning of {word} is {PyDictionary.meaning(word)}")

        elif 'synonym' in query:
            speak("can you please repeat the word")
            word1 = takeCommand()
            speak(f"the synonym of {word1} is {PyDictionary.synonym(word1)}")

        elif 'antonym' in query:
            speak("can you please repeat the word")
            word2 = takeCommand()
            speak(f"the antonym of {word2} is {PyDictionary.antonym(word2)}")

        # calculator
        elif 'calculator' in query:
            speak("opening calculator")
            calc()

        # take a screenshot of the screen
        elif 'screenshot' in query:
            speak('taking screenshot')
            ss()

        # toss
        elif 'toss' in query:
            heads_or_tails = ['heads', 'tails']
            toss = random.choice(heads_or_tails)
            speak(f"it's {toss}")

        # tell a joke
        elif 'joke' in query:
            temp = pyjokes.get_joke(language='en', category='neutral')
            print(temp)
            speak(temp)

        # easter eggs
        elif 'is santa real' in query:
            speak("of course, I'm surprised you had to ask")
            if datetime.datetime.now().month == 12:
                speak("let's go find santa")
                webbrowser.open("https://santatracker.google.com/")

        elif 'are you real' in query or 'are you human' in query:
            speak("maybe, maybe not.")

        elif 'how old are you' in query:
            speak("i am ageless")

        elif 'when will the world end' in query:
            speak("februray thirtieth of never")

        elif 'what does the fox say' in query:
            foxSayings = ["Ring-ding-ding-ding-dingeringeding!", "Wa-pa-pa-pa-pa-pa-pow!",
                          "Hatee-hatee-hatee-ho!", "Fraka-kaka-kaka-kaka-kow!", "A-hee-ahee ha-hee!"]
            ans = random.choice(foxSayings)
            print(ans)
            speak(ans)

        elif 'i love you' in query:
            speak("i know.")

        elif 'thank you' in query:
            speak("my pleasure.")

        # don't listen for a while
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time do you want me to stop listening, in minutes")
            a = takeCommand()
            while a.isdigit() == False:
                speak("invalid, please try again")
                a = takeCommand()

            a = int(a)
            speak(f"going to sleep for {a} minutes")
            time.sleep(a*60)

        # goodbye, exit
        elif 'bye' in query or 'exit' in query:
            speak("have a nice day, goodbye")
            exit()

        # invalid query
        else:
            speak("not a recognized command!")
            speak("try again or ask for help")
