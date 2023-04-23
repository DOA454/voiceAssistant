import sys
import random
import speech_recognition as sr # convert speech to text
import datetime # for fetching date and time
import wikipedia
import webbrowser
import requests
import playsound # to play saved mp3 file
from gtts import gTTS # google text to speech
import os # to save/open files
import wolframalpha # to calculate strings into formula
from selenium import webdriver # to control browser operations
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Edge
from selenium.webdriver.edge.options import Options
import time
import pyautogui
import win32com.client
import pyttsx3 as tts


class Assistant():
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.speaker = tts.init()

    
    # Capture voice input from user and uses speech_recognition library to capture words spoken as text
    def get_speech(self):
        with sr.Microphone() as source:
            # input.adjust_for_ambient_noise(source)
            audio = self.recognizer.listen(source)
            data = ""
            try:
                data = self.recognizer.recognize_google(audio)
                print("You said, " + data)

            except sr.UnknownValueError:
                self.speaker.say("Cannot recognize speech")

            except sr.RequestError:
                self.speaker.say("Speech recognition failed. Check your internet connection or API key")

        return data

    # -------------------------------------------------------------------------------------------------
    # -------------------------------------------------------------------------------------------------
    # categories of functions that the voice assistant can carry out:
    # -------------------------------------------------------------------------------------------------


    # Email tasks
    # -------------------------------------------------------------------------------------------------
    def email_draft(self):
        
        outlook = win32com.client.Dispatch('Outlook.application')
        mail = outlook.createItem(0)
        
        self.speaker.say("Who would you like to send the email to?")
        recipient = self.get_speech().replace(" ","")
        if recipient == 0:
            self.speaker.say('Please repeat')
            recipient = self.get_speech().replace(" ", "")
        self.speaker.say("What is the subject of your email?")
        subject = self.get_speech()
        if subject is None:
            self.speaker.say('The email will be sent without a subject')
        self.speaker.say("What would you like the email to say?")
        body = self.get_speech()
        if body == "stop":
            self.speaker.say("Aborting process")
            self.run()
        

        mail.subject = subject
        mail.to = recipient
        # mail.CC = "abc@gmail.com"
        mail.body = body
        try:
            mail.save()
        except:
            self.speaker.say('Something went wrong. Lets try this again')
            self.email_draft()
        return mail

    def email_send(self):
        mail = self.email_draft()
        try:
            mail.Send()
            self.speaker.say('Email sent successfully')
        except:
            self.speaker.say("Email did not send, let's start over")
            self.email_send()


    def open_email(self):
        self.speaker.say("Opening Outlook email")
        os.startfile('outlook')
        
    # -------------------------------------------------------------------------------------------------
    # Weather
    # -------------------------------------------------------------------------------------------------
    def weather(self, city):
        api_key = '38af91f97ea3a0243ec6cb45019bfb4d'
        # self.speaker.say("Which city?")
        # city = self.get_speech().lower()

        if city != 0:

            url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"

            response = requests.get(url)

            if response.status_code == 200:
                data = response.json()
                temp = data["main"]["temp"]
                feels_like = data["main"]["feels_like"]
                description = data["weather"][0]["description"]
                self.speaker.say(f"Current weather in {city}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
            else:
                print(f"Error retrieving weather data. Error code: {response.status_code}")

    # Opening applications
    # -------------------------------------------------------------------------------------------------
    def open_word(self):
        os.startfile("WINWORD.EXE")
        self.speaker.say("Opening Microsoft Word")
        self.speaker.runAndWait()
        time.sleep(2)
        self.speaker.say("What do you want to write?")
        self.speaker.runAndWait()
        content = self.get_speech()
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys(content)
        self.save_file()


    def open_notepad(self):
        os.startfile("notepad.exe")
        self.speaker.say('Opening notepad')
        self.speaker.runAndWait()
        time.sleep(2)
        self.speaker.say('what would you like to write')
        self.speaker.runAndWait()
        content = self.get_speech().lower()
        pyautogui.write(content)
        self.save_file()
            
    def save_file(self):
        pyautogui.hotkey('ctrl','s')
        time.sleep(1)
        self.speaker.say('What name would you like to save the file as?')
        self.speaker.runAndWait()
        new_filename =self.get_speech().lower()
        pyautogui.write(new_filename) # enter the file name
        pyautogui.press('enter') # confirm the save
        self.speaker.say(f'file saved as: {new_filename}')
        self.speaker.runAndWait()

    def close_app(self, app_name):    
        if app_name != 0:
            self.speaker.say(f"Closing {app_name}")
            self.speaker.runAndWait()
            try:
                os.system(f"taskkill /f /im {app_name}.exe")
            except:
                self.speaker.say(f"Unable to close {app_name}")
                self.speaker.runAndWait()

        # else:
        #     self.speaker.say(f"Unable to find {app_name} ") 

    # -------------------------------------------------------------------------------------------------
    # Internet search
    # -------------------------------------------------------------------------------------------------
    # Using webbrowser
    def google(self):
        webbrowser.open_new_tab("https://www.google.com")
        self.speaker.say("Google is open")
        self.speaker.runAndWait()
        
    # Using wikipedia library
    def wiki(self, text):
        self.speaker.say('Searching Wikipedia')
        self.speaker.runAndWait()
        text = text.replace("wikipedia", "")
        results = wikipedia.summary(text, sentences = 1)
        self.speaker.say("According to Wikipedia")
        self.speaker.runAndWait()
        print(results)
        self.speaker.say(results) 
        self.speaker.runAndWait()

    # Using selenium and edge webdriver to open webpages
    def get_webpage(self, url, text, sleep=5, retries=3):
        for i in range(1, retries+1):
            time.sleep(sleep * i)

            try:
                options = Options()
                browser = Edge('C://Users/danie/DSProjects/webdrivers/msedgedriver.exe', options=options)
                page = browser.get(url)
                browser.implicitly_wait(3)
                browser.maximize_window()
                self.speaker.say("Opening in youtube")
                self.speaker.runAndWait()
                indx = text.split().index('youtube')
                query = text.split()[indx + 1:]
                search = '+'.join(query)
                browser.get(f"http://www.youtube.com/results?search_query={search}")

                time.sleep(30)
            except TimeoutException:
                print(f"Timeout error on {url}")
                continue
            else:
                break
        return page
    # ----------------------------------------------------------------------------------------------------


    def calculate(self, question):
        app_id="4YVHWL-29XPVVGXQP"
        client = wolframalpha.Client(app_id)
        res = client.query(question)
        answer = next(res.results).text
        self.speaker.say(f"The answer is {answer}")
        self.speaker.runAndWait()

    # System functions
    # ----------------------------------------------------------------------------------------------------
    # Time
    def tell_time(self):
        strTime=datetime.datetime.now().strftime("%H:%M:%S")
        self.speaker.say(f"the time is {strTime}") 
        self.speaker.runAndWait()

    # Shutdown, logout, and restart
    def logout(self):
        self.speaker.say("Logging out of  computer")
        self.speaker.runAndWait()
        os.system("shutdown /l")

    def restart(self):
        self.speaker.say("Restarting computer")
        self.speaker.runAndWait()
        os.system("shutdown /g /t 30")

    def shutdown(self):    
        self.speaker.say("Shutting computer down")
        self.speaker.runAndWait()
        os.system("shutdown /s /t 30")

    # Stop program
    # def stop():
    #     exit()
    def chatbot(self, query):
        self.speaker.say(" Give me a sec")
        self.speaker.runAndWait()
        os.environ["OPENAI_API_KEY"] = "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"
        documents = SimpleDirectoryReader('read').load_data()
        index = GPTSimpleVectorIndex(documents)

        answer = index.query(query)
        self.speaker.say(answer.response)
        self.speaker.runAndWait()  

    def interact(self, text):

        if 'wikipedia' in text:
            self.wiki(text)

        elif 'close' in text:
            app_name = text.replace("close ", "")
            self.close_app(app_name)
            self.speaker.say(f"{app_name} is closed")
            time.sleep(2)

        elif 'time' in text:
            self.tell_time()

        elif 'search'  in text:
            text = text.replace("search ", "")
            webbrowser.open_new_tab(text)
            time.sleep(5)

        elif "calculate" in text:
            self.calculate(text)

        elif 'google' in text:
            self.google()
            time.sleep(5)

        # elif "save file" in text:
        #     self.speaker.say("Saving file")
        #     try:
        #         pyautogui.hotkey('ctrl','s')
        #         time.sleep(1)
            
        #     except Exception as e:
        #         self.speaker.say(f"Error: {e}")

        elif 'youtube' in text:
            url = 'https://www.youtube.com/'
            self.get_webpage(url, text)

        elif "open notepad" in text:
            self.open_notepad()

        elif "open microsoft word" in text:
            self.open_word()
                    
        elif "weather" in text:
            city = text.replace("what's the weather in ","")
            self.weather(city)
                
        elif "shutdown computer " in text:
            self.shutdown()

        elif "restart computer" in text:
            self.restart()

        elif "log out" in text:
            self.logout()
        
        elif "open email" in text:
            self.open_email()

        elif "draft email" in text:
            self.email_draft()

        elif 'send email' in text:
            self.speaker.say('Opening email')
            self.email_send()

        # elif text != 0:
        #     self.chatbot(text)

        else:
            self.speaker.say("I'm not programmed to help with that")
            self.speaker.runAndWait()
        
            