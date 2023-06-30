import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

# INITIALIZING THE SPEECH ENGINE
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:  # CHECKING IF MORNING TIME
        speak("Good Morning Sir")

    elif hour >= 12 and hour <= 18:  # CHECKING IF AFTERNOON TIME
        speak("Good Afternoon Sir")

    else:  # CHECKING IF EVENING TIME
        speak("Good Evening Sir")

    assistant_name = "Jarvis 1 point 0"
    speak("I'm your assistant")
    speak(assistant_name)


def takeCommand():
    # INITIALIZING MICROPHONE
    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    # IMPLEMENTING ERROR HANDLING

    try:
        print("Recognising...")
        query = r.recognize_google(audio, language='en-in')
        print(f"USer Said: {query}\n")


    except Exception as e:
        print(e)
        print("Unable to Recognize your voice")
        return "none"

    return query


# MAIN FUNCTION

if __name__ == '__main__':

    # INITIALIZING CLEAR SCREEN
    clear = lambda: os.system('cls')

    # TO CLEAR ANY COMMANDS BEFORE EXECUTION OF THIS PYTHON FILE
    clear()
    wishMe()

    while True:
        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)


        elif "open" in query and "youtube" in query:
            speak("Opening Youtube")
            webbrowser.open("youtube.com")
