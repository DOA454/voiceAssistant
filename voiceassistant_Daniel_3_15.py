import subprocess
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
#GUI 
from tkinter import *
from tkinter.ttk import*

def create_gui():
    root = Tk() # create object
    root.title('Team 4 Voice Assistant Demo')
   # Set window size
    root.geometry('400x400')
    root.configure(background='white')


    style = Style()
    style.configure('W.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'green')

    button1 = Button(root, text = 'HELLO', command = main , style='W.TButton')
    button1.grid(row = 5, column = 3, pady = 10, padx = 100)
    button1.place(relx=0.5, rely=0.3, anchor=CENTER)

    style.configure('S.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'red')

    button2 = Button(root, text = 'GOODBYE', command = stop, style='S.TButton')
    button2.grid(row = 6, column = 3, pady = 10, padx = 100)
    button2.place(relx=0.5, rely=0.5, anchor=CENTER)

    # style.configure('C.TButton', font =
    #             ('calibri', 10, 'bold', 'underline'),
    #                 foreground = 'red')

    # button3 = Button(root, text = 'CLOSE', command = stop,  style='C.TButton')
    # button3.grid(row = 7, column = 3, pady = 10, padx = 100)
    # button3.place(relx=0.5, rely=0.7, anchor=CENTER)

    label_font = ('Arial Black', 12)
    label1 = Label(root, text='Virtual Assistant : Jarvis', font = label_font)
    label2 = Label(root, text='Welcome, How can I help you?', font = label_font)
    label1.pack()
    label2.pack()

    root.mainloop()
   




# Capture voice input from user and uses speech_recognition library to capture words spoken as text
def talk():
    input = sr.Recognizer()
    with sr.Microphone() as source:
        # input.adjust_for_ambient_noise(source)
        audio = input.listen(source)
        data = ""
        try:
            data = input.recognize_google(audio)
            print("You said, " + data)

        except sr.UnknownValueError:
            respond("Cannot recognize speech")

        except sr.RequestError:
            respond("Speech recognition failed. Check your internet connection or API key")

    return data

# Voice assistant responding to user using the playsound library. Writes the speech recognized to a file,
# uses Playsound to read out the file and then deletes the file from the system. 
def respond(output):
    print(output)
    response=gTTS(text=output, lang='en', tld="com.au")
    file = "audio.mp3"
    response.save(file)
    playsound.playsound(file)
    os.remove(file)

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# categories of functions that the voice assistant can carry out:
# -------------------------------------------------------------------------------------------------


# Email tasks
# -------------------------------------------------------------------------------------------------
def email_draft():
    
    outlook = win32com.client.Dispatch('Outlook.application')
    mail = outlook.createItem(0)
    
    respond("Who would you like to send the email to?")
    recipient = talk().replace(" ","")
    if recipient == 0:
        respond('Please repeat')
    respond("What is the subject of your email?")
    subject = talk()
    if subject is None:
        respond('The email will be sent without a subject')
    respond("What would you like the email to say?")
    body = talk()
    

    mail.subject = subject
    mail.to = recipient
    # mail.CC = "abc@gmail.com"
    mail.body = body
    try:
        mail.save()
    except:
        respond('Something went wrong. Lets try this again')
        email_draft()
    return mail

def email_send():
    mail = email_draft()
    try:
        mail.Send()
        respond('Email sent successfully')
    except:
        respond("Email did not send, let's start over")
        email_send()


def open_email():
    respond("Opening Outlook email")
    os.startfile('outlook')
    
# -------------------------------------------------------------------------------------------------
# Weather
# -------------------------------------------------------------------------------------------------
def weather(city):
    api_key = '38af91f97ea3a0243ec6cb45019bfb4d'
    # respond("Which city?")
    # city = talk().lower()

    if city != 0:

        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"

        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            temp = data["main"]["temp"]
            feels_like = data["main"]["feels_like"]
            description = data["weather"][0]["description"]
            respond(f"Current weather in {city}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
        else:
            print(f"Error retrieving weather data. Error code: {response.status_code}")

# Opening applications
# -------------------------------------------------------------------------------------------------
def open_word():
    os.startfile("WINWORD.EXE")
    respond("Opening Microsoft Word")
    time.sleep(2)
    respond("What do you want to write?")
    content = talk()
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys(content)
    save_file()


def open_notepad():
	os.startfile("notepad.exe")
	respond('Opening notepad')
	time.sleep(2)
	respond('what would you like to write')
	content = talk().lower()
	pyautogui.write(content)
	save_file()
        
def save_file():
    pyautogui.hotkey('ctrl','s')
    time.sleep(1)
    respond('What name would you like to put for the file?')
    new_filename =talk().lower()
    pyautogui.write(new_filename) # enter the file name
    pyautogui.press('enter') # confirm the save
    respond(f'file saved as: {new_filename}')

def close_app(app_name):    
    if app_name != 0:
        respond(f"Closing {app_name}")
        os.system(f"taskkill /f /im {app_name}.exe")

    else:
        respond(f"Unable to close {app_name} ") 

# -------------------------------------------------------------------------------------------------
# Internet search
# -------------------------------------------------------------------------------------------------
# Using webbrowser
def google():
    webbrowser.open_new_tab("https://www.google.com")
    respond("Google is open")
    
# Using wikipedia library
def wiki(text):
    respond('Searching Wikipedia')
    text = text.replace("wikipedia", "")
    results = wikipedia.summary(text, sentences = 1)
    respond("According to Wikipedia")
    print(results)
    respond(results) 

# Using selenium and edge webdriver to open webpages
def get_webpage(url, text, sleep=5, retries=3):
    for i in range(1, retries+1):
        time.sleep(sleep * i)

        try:
            options = Options()
            browser = Edge('C://Users/danie/DSProjects/webdrivers/msedgedriver.exe', options=options)
            page = browser.get(url)
            browser.implicitly_wait(3)
            browser.maximize_window()
            respond("Opening in youtube")
            indx = text.split().index('youtube')
            query = text.split()[indx + 1:]
            search = '+'.join(query)
            browser.get(f"http://www.youtube.com/results?search_query={search}")

            time.sleep(15)
        except TimeoutException:
            print(f"Timeout error on {url}")
            continue
        else:
            break
    return page
# ----------------------------------------------------------------------------------------------------


def calculate(question):
    app_id="4YVHWL-29XPVVGXQP"
    client = wolframalpha.Client(app_id)
    res = client.query(question)
    answer = next(res.results).text
    respond(f"The answer is {answer}")

# System functions
# ----------------------------------------------------------------------------------------------------
# Time
def tell_time():
    strTime=datetime.datetime.now().strftime("%H:%M:%S")
    respond(f"the time is {strTime}") 

# Shutdown, logout, and restart
def logout():
    respond("Logging out of  computer")
    os.system("shutdown /l")

def restart():
    respond("Restarting computer")
    os.system("shutdown /g /t 30")

def shutdown():    
    respond("Shutting computer down")
    os.system("shutdown /s /t 30")

# Stop program
def stop():
    exit()

# -------------------------------------------------------------------------------------------------
# Main function
# -------------------------------------------------------------------------------------------------
def main():
    # respond("Hi, I am Jarvis, your personal desktop assistant")
    
    # while(1):
    respond("Hi. How can I help you?")
    text = talk().lower()

    if text != 0:       

        if any(word in text for word in ["stop", "exit", "bye"]):
            respond("Jarvis, out!")
            stop()

        elif 'wikipedia' in text:
            wiki(text)

        elif 'close' in text:
            app_name = text.replace("close ", "")
            close_app(app_name)
            respond(f"{app_name} is closed")
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

        elif "save file" in text:
            respond("Saving file")
            try:
                pyautogui.hotkey('ctrl','s')
                time.sleep(1)
            
            except Exception as e:
                respond(f"Error: {e}")

        elif 'youtube' in text:
            url = 'https://www.youtube.com/'
            get_webpage(url, text)

        elif "open notepad" in text:
            open_notepad()

        elif "open microsoft word" in text:
            open_word()
                    
        elif "weather" in text:
            city = text.replace("what's the weather in ","")
            print(city)
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
            respond('Opening email')
            email_send()

        else:
            respond("I'm not programmed to help with that")
            time.sleep(2)
            main()




# Driver code
if __name__=='__main__':
    create_gui()
  

