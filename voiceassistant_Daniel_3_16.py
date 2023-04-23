import sys
import random
import speech_recognition as sr # convert speech to text
import datetime # for fetching date and time
import wikipedia
import webbrowser
import requests
import os # to save/open files
import wolframalpha # to calculate strings into formula
from selenium import webdriver # to control browser operations
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Edge
from selenium.webdriver.edge.options import Options
import time
import win32com.client
import pyttsx3 as tts
#GUI 
import tkinter as tk
from itertools import count
from PIL import Image, ImageTk
import threading
from IPython.display import Markdown
from assistant import Assistant
from llama_index import GPTSimpleVectorIndex, SimpleDirectoryReader


class Jarvis():
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.speaker = tts.init()

        self.assistant = Assistant()

        self.root = tk.Tk()
        self.label = tk.Label(text='👽', font = ("Arial", 240))
        self.label.pack()
        
        threading.Thread(target=self.run).start()

        self.root.mainloop()     
    
         
    # -------------------------------------------------------------------------------------------------
    # Main function
    # -------------------------------------------------------------------------------------------------
    def run(self):
        while True:
            try:
                text = self.assistant.get_speech().lower()
                if any(word in text for word in ["hey jarvis", "hi jarvis", "hello jarvis"]):
                    # start gui                       
                    self.label.config(fg="blue")
                    greetings = ["How can I help today?",
                                "What's up?",
                                "Jarvis at your service",
                                "How can I help?",
                                "What can I do for you?"]
                    
                    farewell = ["Goodbye", "Bye", "See you later"]
                    choice_greeting =  random.choice(greetings)
                    choice_farewell = random.choice(farewell)

                    self.speaker.say(choice_greeting)
                    self.speaker.runAndWait()
                    text = self.assistant.get_speech().lower()

                    if any(word in text for word in ["stop", "exit", "bye"]):
                            self.speaker.say(choice_farewell)
                            self.speaker.runAndWait()
                            self.speaker.stop()
                            self.root.destroy()
                            sys.exit()

                    else:
                            if text != 0:                        
                                self.label.config(fg="green")      
                                self.assistant.interact(text)
                                self.label.config(fg="black")
            except:    
                self.label.config(fg="black")
                continue



# Driver code
if __name__=='__main__':
    Jarvis()
  

