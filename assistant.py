import speech_recognition as sr # convert speech to text
import datetime # for fetching date and time
import wikipedia
import webbrowser
import requests
import os # to save/open files
import wolframalpha # to calculate strings into formula
import time
import pyautogui
import win32com.client
import pyttsx3 as tts
from pyllamacpp.model import Model


class Assistant():
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.speaker = tts.init()

    
    # Capture voice input from user and uses speech_recognition library to capture words spoken as text
    def get_speech(self, source) -> str:
        self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
        audio = self.recognizer.listen(source)
        data = ""
        try:
            data = self.recognizer.recognize_google(audio)
            print("You said, " + data)

        except sr.UnknownValueError:
            print("Cannot recognize speech")            

        except sr.RequestError:
            print("Speech recognition failed. Check your internet connection or API key") 

        return data
    
    # Give the system speaking capabilities
    def say(self, text):
        self.speaker.say(text)
        self.speaker.runAndWait()

    # -------------------------------------------------------------------------------------------------
    # -------------------------------------------------------------------------------------------------
    # categories of functions that the voice assistant can carry out:
    # - Email
    # - Weather
    # - Opening, saving and closing applications
    # - Internet search
    # - Calculations
    # - System fuctions like telling time, shutdown, logout, restart
    # - General conversation
    # -------------------------------------------------------------------------------------------------


    # Email tasks
    # -------------------------------------------------------------------------------------------------
    def email_draft(self, source):
        
        outlook = win32com.client.Dispatch('Outlook.application')
        mail = outlook.createItem(0)
        
        self.say("Who would you like to send the email to?")
        recipient = self.get_speech(source).lower().replace(" ","")
        if recipient == "cancel":
            self.say("Aborting process")
            return
        if recipient == 0:
            self.say('Please repeat')
            recipient = self.get_speech(source).lower().replace(" ", "")
        self.say("What is the subject of your email?")
        subject = self.get_speech(source)
        if subject == "cancel":
            self.say("Aborting process")
            return
        if subject is None:
            self.say('The email will be sent without a subject')
        self.say("What would you like the email to say?")
        body = self.get_speech(source)
        if body == "cancel":
            self.say("Aborting process")
            return
        

        mail.subject = subject
        mail.to = recipient
        # mail.CC = "abc@gmail.com"
        mail.body = body
        try:
            mail.save()
            self.say("Email has been saved")
        except:
            self.say('Something went wrong. Lets try this again')
            self.email_draft(source)
        return mail

    def email_send(self, source):
        mail = self.email_draft(source)
        try:
            mail.Send()
            self.say('Email sent successfully')
        except:
            self.say("Email did not send, let's start over")
            self.email_send(source)


    def open_email(self):
        self.say("Opening Outlook email")
        os.startfile('outlook')
        
    # -------------------------------------------------------------------------------------------------
    # Weather
    # -------------------------------------------------------------------------------------------------
    def weather(self, city):
        api_key = '38af91f97ea3a0243ec6cb45019bfb4d'

        if city != 0:

            url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"

            response = requests.get(url)

            if response.status_code == 200:
                data = response.json()
                temp = data["main"]["temp"]
                feels_like = data["main"]["feels_like"]
                description = data["weather"][0]["description"]
                self.say(f"Current weather in {city}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
            else:
                print(f"Error retrieving weather data. Error code: {response.status_code}")

    # Opening, saving and closing applications
    # -------------------------------------------------------------------------------------------------
    def open_word(self, source):
        # os.startfile("WINWORD.EXE")
        self.say("Opening Microsoft Word")
        time.sleep(2)
        self.say("What do you want to write?")
        content = self.get_speech(source)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys(content)
        self.save_file(source)


    def open_notepad(self, source):
        os.startfile("notepad.exe")
        self.say('Opening notepad')
        time.sleep(2)
        self.say('what would you like to write')
        content = self.get_speech(source).lower()
        pyautogui.write(content)
        self.save_file(source)

    # Saves files in the default save location in the application        
    def save_file(self, source):
        pyautogui.hotkey('ctrl','s')
        time.sleep(1)
        self.say('What name would you like to save the file as?')
        new_filename =self.get_speech(source).lower()
        pyautogui.write(new_filename) # enter the file name
        pyautogui.press('enter') # confirm the save
        self.say(f'file saved as: {new_filename}')

    # Only closes applications with the format {appName}.exe
    def close_app(self, app_name):    
        if app_name != 0:
            self.say(f"Closing {app_name}")
            try:
                os.system(f"taskkill /f /im {app_name}.exe")
            except:
                self.say(f"Unable to close {app_name}")


    # -------------------------------------------------------------------------------------------------
    # Internet search
    # -------------------------------------------------------------------------------------------------
    # Using webbrowser
    def google(self):
        webbrowser.open_new_tab("https://www.google.com")
        self.say("Google is open")

    def youtube(self, url, text):
        indx = text.split().index('youtube')
        query = text.split()[indx + 1:]
        search = '+'.join(query)
        webbrowser.open_new_tab(f"{url}results?search_query={search}")
        self.say("Youtube is open") 

    def web_search(self, text):
        indx = text.split().index('search')
        query = text.split()[indx + 1:]
        search = '+'.join(query)
        self.say(f"Searching for {query} on the interwebs")
        webbrowser.open_new_tab(f"https://www.google.com/search?q={search}")


    # Using wikipedia library
    def wiki(self, text):
        self.say('Searching Wikipedia')
        text = text.replace("wikipedia", "")
        results = wikipedia.summary(text, sentences = 1)
        self.say("According to Wikipedia")
        print(results)
        self.say(results) 

    # Using wolframalpha API
    def calculate(self, question):
        app_id="4YVHWL-29XPVVGXQP"
        client = wolframalpha.Client(app_id)
        res = client.query(question)
        answer = next(res.results).text
        self.say(f"The answer is {answer}")

    # ----------------------------------------------------------------------------------------------------
    # System functions
    # ----------------------------------------------------------------------------------------------------
    # Time
    def tell_time(self):
        strTime=datetime.datetime.now().strftime("%H:%M:%S")
        self.say(f"the time is {strTime}") 

    # Shutdown, logout, and restart
    def logout(self):
        self.say("Logging out of  computer")
        os.system("shutdown /l")

    def restart(self):
        self.say("Restarting computer")
        os.system("shutdown /g /t 30")

    def shutdown(self):    
        self.say("Shutting computer down")
        os.system("shutdown /s /t 30")
    
    # ----------------------------------------------------------------------------------------------------
    # Enabling the assistant with LLM capabilities using a locally available pretrained GPT quantized model

    def GPT(self, prompt):
        model = Model(r'C:\Users\danie\GPT4All\gpt4all-model.bin')
        self.say("Getting the information for you")
        response = model.generate(prompt, n_predict=55, n_threads=10)
        self.say(response)

    # ---------------------------------------------------------------------------------------------------------------
    # Code that governs the type of interaction and execution of commands by the assistant via a series of conditions
    # ---------------------------------------------------------------------------------------------------------------

    def interact(self, text, source):
        # Uses wikipedia API
        if 'wikipedia' in text:
            self.wiki(text)

        elif 'close' in text:
            app_name = text.replace("close ", "")
            self.close_app(app_name)
            self.say(f"{app_name} is closed")

        elif 'time' in text:
            self.tell_time()

        elif 'search'  in text:
            self.web_search(text)
            
        elif "calculate" in text:
            self.calculate(text)

        elif 'google' in text:
            self.google()

        elif 'youtube' in text:
            url = 'https://www.youtube.com/'
            self.youtube(url, text)

        elif "open notepad" in text:
            self.open_notepad(source)

        elif "open microsoft word" in text:
            self.open_word(source)
                    
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
            self.email_draft(source)

        elif 'send email' in text:
            self.say('Opening email')
            self.email_send(source)

        elif text != 0 and len(text) > 20:
            self.GPT(text)

        else:
            self.say("I'm not programmed to help with that")

# End of code       
        
            