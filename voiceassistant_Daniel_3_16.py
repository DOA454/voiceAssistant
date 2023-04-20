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
import pyttsx3 as tts
from selenium import webdriver # to control browser operations
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Edge
from selenium.webdriver.edge.options import Options
import time
import pyautogui
import win32com.client
import threading
#GUI 
from tkinter import *
from tkinter.ttk import*
from IPython.display import Markdown
from llama_index import GPTSimpleVectorIndex, SimpleDirectoryReader


   

class Jarvis:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.speaker = tts.init()

        threading.Thread(target=main).start()
        
        self.window = Tk()



def create_gui(self):
    
    self.window.title('Team 4 Voice Assistant Demo')
   # Set window size
    self.window.geometry('500x500')
    self.window.configure(background='white')


    self.style = Style()
    self.style.configure('W.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'green')

    self.button1 = Button(self.window, text = 'HELLO', command = main , style='W.TButton')
    self.button1.grid(row = 5, column = 3, pady = 10, padx = 100)
    self.button1.place(relx=0.5, rely=0.3, anchor=CENTER)

    self.style.configure('S.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'red')

    self.button2 = Button(self.window, text = 'GOODBYE', command = stop, style='S.TButton')
    self.button2.grid(row = 6, column = 3, pady = 10, padx = 100)
    self.button2.place(relx=0.5, rely=0.5, anchor=CENTER)

    self.label_font = ('Open Sans', 12)
    self.label1 = Label(self.window, text='Virtual Assistant : Jarvis', font = self.label_font)
    self.label2 = Label(self.window, text='Welcome, How can I help you?', font = self.label_font)
    self.label1.pack()
    self.label2.pack()

    self.window.mainloop()




# Capture voice input from user and uses speech_recognition library to capture words spoken as text
def talk(self):
    with sr.Microphone() as source:
        # input.adjust_for_ambient_noise(source)
        audio = self.recognizer.listen(source)
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
    recipient = talk().replace(" ","")
    if recipient == 0:
        self.speaker.say('Please repeat')
        recipient = talk().replace(" ", "")
    self.speaker.say("What is the subject of your email?")
    subject = talk()
    if subject is None:
        self.speaker.say('The email will be sent without a subject')
    self.speaker.say("What would you like the email to say?")
    body = talk()
    if body == "stop":
        self.speaker.say("Aborting process")
        main()
    

    mail.subject = subject
    mail.to = recipient
    # mail.CC = "abc@gmail.com"
    mail.body = body
    try:
        mail.save()
    except:
        self.speaker.say('Something went wrong. Lets try this again')
        email_draft()
    return mail

def email_send(self):
    mail = email_draft()
    try:
        mail.Send()
        self.speaker.say('Email sent successfully')
    except:
        self.speaker.say("Email did not send, let's start over")
        email_send()


def open_email(self):
    self.speaker.say("Opening Outlook email")
    os.startfile('outlook')
    
# -------------------------------------------------------------------------------------------------
# Weather
# -------------------------------------------------------------------------------------------------
def weather(self, city):
    api_key = '38af91f97ea3a0243ec6cb45019bfb4d'
    # self.speaker.say("Which city?")
    # city = talk().lower()

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
    time.sleep(2)
    self.speaker.say("What do you want to write?")
    content = talk()
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys(content)
    save_file()


def open_notepad(self):
	os.startfile("notepad.exe")
	self.speaker.say('Opening notepad')
	time.sleep(2)
	self.speaker.say('what would you like to write')
	content = talk().lower()
	pyautogui.write(content)
	save_file()
        
def save_file(self):
    pyautogui.hotkey('ctrl','s')
    time.sleep(1)
    self.speaker.say('What name would you like to save the file as?')
    new_filename =talk().lower()
    pyautogui.write(new_filename) # enter the file name
    pyautogui.press('enter') # confirm the save
    self.speaker.say(f'file saved as: {new_filename}')

def close_app(self, app_name):    
    if app_name != 0:
        self.speaker.say(f"Closing {app_name}")
        try:
            os.system(f"taskkill /f /im {app_name}.exe")
        except:
            self.speaker.say(f"Unable to close {app_name}")

    # else:
    #     self.speaker.say(f"Unable to find {app_name} ") 

# -------------------------------------------------------------------------------------------------
# Internet search
# -------------------------------------------------------------------------------------------------
# Using webbrowser
def google(self):
    webbrowser.open_new_tab("https://www.google.com")
    self.speaker.say("Google is open")
    
# Using wikipedia library
def wiki(self, text):
    self.speaker.say('Searching Wikipedia')
    text = text.replace("wikipedia", "")
    results = wikipedia.summary(text, sentences = 1)
    self.speaker.say("According to Wikipedia")
    print(results)
    self.speaker.say(results) 

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

# System functions
# ----------------------------------------------------------------------------------------------------
# Time
def tell_time(self):
    strTime=datetime.datetime.now().strftime("%H:%M:%S")
    self.speaker.say(f"the time is {strTime}") 

# Shutdown, logout, and restart
def logout(self):
    self.speaker.say("Logging out of  computer")
    os.system("shutdown /l")

def restart(self):
    self.speaker.say("Restarting computer")
    os.system("shutdown /g /t 30")

def shutdown(self):    
    self.speaker.say("Shutting computer down")
    os.system("shutdown /s /t 30")

# Stop program
def stop():
    exit()

def chatbot(self, query):
    self.speaker.say(" Give me a sec")
    os.environ["OPENAI_API_KEY"] = "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"
    documents = SimpleDirectoryReader('read').load_data()
    index = GPTSimpleVectorIndex(documents)

    answer = index.query(query)
    self.speaker.say(answer.response)

# -------------------------------------------------------------------------------------------------
# Main function
# -------------------------------------------------------------------------------------------------
def main(self):
    greetings = ["I'm Jarvis, your desktop assistant, how can I help today?",
                 "I'm Jarvis, what's up?",
                 "Jarvis at your service",
                 "How can I help?",
                 "What can I do for you?"]
    choice =  random.choice(greetings)
    self.speaker.say(choice)
    text = talk().lower()

    if text is not None:       

        if any(word in text for word in ["stop", "exit", "bye"]):
            self.speaker.say("Jarvis, out!")
            stop()

        elif 'wikipedia' in text:
            wiki(text)

        elif 'close' in text:
            app_name = text.replace("close ", "")
            close_app(app_name)
            self.speaker.say(f"{app_name} is closed")
            time.sleep(2)

        elif 'time' in text:
            tell_time()

        elif 'search'  in text:
            text = text.replace("search ", "")
            webbrowser.open_new_tab(text)
            time.sleep(5)

        elif "calculate" in text:
            calculate(text)

        elif 'google' in text:
            google()
            time.sleep(5)

        elif 'youtube' in text:
            url = 'https://www.youtube.com/'
            get_webpage(url, text)

        elif "open notepad" in text:
            open_notepad()

        elif "open microsoft word" in text:
            open_word()
                    
        elif "weather" in text:
            city = text.replace("what's the weather in ","")
            weather(city)
                
        elif "shutdown computer " in text:
            shutdown()

        elif "restart computer" in text:
            restart()

        elif "log out" in text:
            logout()
        
        elif "open email" in text:
            open_email()

        elif "draft email" in text:
            email_draft()

        elif 'send email' in text:
            self.speaker.say('Opening email')
            email_send()

        elif text != 0:
            chatbot(text)

        else:
            self.speaker.say("Let's start over")
            main()
            # self.speaker.say("I'm not programmed to help with that")
            # time.sleep(2)
            # main()




# Driver code
if __name__=='__main__':
    Jarvis()
  

