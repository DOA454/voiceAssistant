import sys
import random
import speech_recognition as sr # convert speech to text
import datetime # for fetching date and time
import wikipedia
import webbrowser
import requests
import os # to save/open files
import wolframalpha # to calculate strings into formula
import time
import win32com.client
import pyttsx3 as tts
#GUI 
import tkinter as tk
import threading
from assistant import Assistant


class Jarvis():
    def __init__(self, root):

        self.assistant = Assistant()

        self.root = root
        self.label = tk.Label(text="👽", font=("Arial", 240, "bold"))
        self.label.pack()
       
        self.running = True

        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()

        self.notepad_button = tk.Button(self.root, text="Readme", command=self.open_notepad_file)
        self.notepad_button.pack()
    
    def stop(self):
        self.running = False

    def open_notepad_file(self):
        file_path = r"C:\Users\danie\Documents\MNSU\Spring 2023\CIS 630\projectcode\Readme.txt"  # Replace with your file path
        os.startfile(file_path)
        
    # -------------------------------------------------------------------------------------------------
    # Running the assistant
    # -------------------------------------------------------------------------------------------------
    def run(self):
        while self.running:
            try:
                with sr.Microphone() as source:

                    text = self.assistant.get_speech(source).lower()
                    if any(word in text for word in ["hey jarvis", "hi jarvis", "hello jarvis"]):                
                        self.label.config(fg="blue")
                        greetings = ["How can I help today?",
                                    "What's up?",
                                    "Jarvis at your service",
                                    "How can I help?",
                                    "What can I do for you?"]
                        
                        farewell = ["Goodbye", "Bye", "See you later"]
                        choice_greeting =  random.choice(greetings)
                        choice_farewell = random.choice(farewell)

                        self.assistant.say(choice_greeting)
                        self.label.config(fg='yellow')
                        text = self.assistant.get_speech(source).lower()
                        self.label.config(fg='red')
                        if any(word in text for word in ["stop", "exit", "bye"]):
                                self.assistant.say(choice_farewell)
                                self.root.destroy()
                                sys.exit()

                        else:
                            if text != 0:                        
                                self.label.config(fg="green")      
                                self.assistant.interact(text, source)
                                self.label.config(fg="black")
                    else:
                        continue
                
            except:    
                print("Closing program")
                self.root.destroy()
                self.stop()
                sys.exit()



# Driver code
if __name__=='__main__':
    root = tk.Tk()
    root.title("Jarvis AI")
    Jarvis(root)
    root.mainloop()

  

