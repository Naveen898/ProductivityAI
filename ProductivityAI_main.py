import numpy as np
import pyautogui
import cv2
import pygetwindow as gw
import pytesseract
from PIL import ImageGrab
import subprocess
import wolframalpha
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import smtplib
import ctypes
import time
import requests
import shutil
from bs4 import BeautifulSoup
import win32com.client as wincl
from gtts import gTTS
from ollama import Client

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty("rate", 175)

# Function to speak out the given text
def speak(audio):
	engine.say(audio)
	engine.runAndWait()

# Function to greet the user based on the time of day
def wishMe():
	hour = datetime.datetime.now().hour
	if 0 <= hour < 12:
		speak("Good Morning!")
	elif 12 <= hour < 18:
		speak("Good Afternoon!")
	else:
		speak("Good Evening!")

	assname = "ProductivityAI, an assistant for your computer, also powered by Llama"  # Your assistant's name
	speak("I am your Assistant")
	speak(assname)
	speak("How can I help you?")
	

# Function to get the user's name and welcome them
def usrname():
	speak("What should I call you?")
	uname = takeCommand()
	speak(f"Welcome Mr. {uname}")
	columns = shutil.get_terminal_size().columns
	print("#####################".center(columns))
	print(f"Welcome  {uname}".center(columns))
	print("#####################".center(columns))
	speak("How can I help you?")

# Function to take voice input from the user
def takeCommand():
	r = sr.Recognizer()
	with sr.Microphone() as source:
		print("Listening...")
		r.pause_threshold = 1
		r.adjust_for_ambient_noise(source)
		audio = r.listen(source)
	try:
		print("Recognizing...")
		query = r.recognize_google(audio, language='en-in')
		print(f"User said: {query}\n")
	except Exception as e:
		print(e)
		print("Unable to Recognize your voice.")
		return "None"
	return query

def live_dictate():
    def get_active_window_title():
        active_window = gw.getWindowsWithTitle(gw.getActiveWindow().title)
        if active_window:
            return active_window[0].title
        return None

    def read_screen():
        window_title = get_active_window_title()
        if window_title:
            text_to_speak = f"Reading from {window_title}."
            # Add code to read text from the screen here
            return text_to_speak
        else:
            return "I'm sorry, I couldn't detect the active window."

    query = takeCommand()  # Assuming you have a function called takeCommand() for getting user queries

    if "live dictate" in query:
        return read_screen()

    return "I'm sorry, I didn't understand that command."

def get_weather_forecast(city):
    api_key = "b5a2d8688bc0c191115d91e6bb76a120"  # Replace with your OpenWeatherMap API key
    base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    
    response = requests.get(base_url)
    data = response.json()
    
    if data["cod"] != "404":
        weather_info = data["weather"][0]["description"]
        temp = data["main"]["temp"]
        speak(f"The weather in {city} is {weather_info} with a temperature of {temp} degrees Celsius.")
    else:
        return "City not found. Please try again."

def take_screenshot():
    screenshot = pyautogui.screenshot()
    screenshot.save("screenshot.png")
    print("Screenshot saved as screenshot.png")

			
# Function to transcribe audio file to text
def transcribe_media_to_text(filename):
	recognizer = sr.Recognizer()
	with sr.AudioFile(filename) as source:
		audio = recognizer.record(source)
	try:
		return recognizer.recognize_google(audio)
	except Exception as e:
		print(f"Skipping due to an error: {str(e)}")
		speak("Skipping due to an error")

# Function to generate response using Ollama (Llama 3.2)
def generate_response(prompt):
    try:
        # Create an instance of the Client class
        client = Client()
        
        # Call the generate method on the client instance
        model = "llama3.2"  # Replace with your model name
        response = client.generate(model=model, prompt=prompt)
        return response.response  # Return the response text
    except Exception as e:
        print(f"Error generating response: {str(e)}")
        return "I'm sorry, I couldn't process your request."

# Function to test text-to-speech
def speak_test(text):
	engine.say(text)
	engine.runAndWait()

def llm_activate():
	speak("Switching to Genius mode, now your LLM is in action")
	speak("Now, please note, whenever you want to ask question, just say 'Genius' to speak!")
	while True:
		print("Say 'Genius' to start recording your question...")
		with sr.Microphone() as source:
			recognizer = sr.Recognizer()
			audio = recognizer.listen(source)
			try:
				transcription = recognizer.recognize_google(audio)
				if transcription.lower() == "genius":
					speak("Speak..")
					# Record audio
					filename = "input.wav"
					print("Say your question...")
					with sr.Microphone() as source:
						recognizer = sr.Recognizer()
						source.pause_threshold = 1
						audio = recognizer.listen(source, phrase_time_limit=None, timeout=None)
						with open(filename, "wb") as f:
							f.write(audio.get_wav_data())

					# Transcribe audio to text
					text = transcribe_media_to_text(filename)
					if text:
						print(f"You said: {text}")
						speak(f"You said:{text}")

						# Generate response using GPT-3
						response = generate_response(text)  # Call the function to get a response

						print(f"Your LLM says: {response}")

						# Read response using text-to-speech
						speak_test(response)  # Corrected function name

				elif transcription.lower() == "switch to normal mode":  #trying to switch to normal mode
					print("Here you go, going back to normal mode")
					speak("Here you go, going back to normal mode")
					break
					
						
			except Exception as e:
				print(f"An error occurred: {str(e)}")
				speak(f"An error occurred!")

	def chat(prompt, model, tokenizer):
		inputs = tokenizer(prompt, return_tensors="pt").to("cuda")
		outputs = model.generate(**inputs, max_new_tokens=200)
		response = tokenizer.decode(outputs[0], skip_special_tokens=True)
		return response

def get_current_time():
    now = datetime.datetime.now().strftime("%I:%M %p")
    return now

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	clear()
	wishMe()
	#usrname()

	assistance_mode = False  # Added a flag to track if we're in assistance mode

	while True:
		query = takeCommand().lower()
		if query=="Hello":
			speak("Speak")
			query=takeCommand().lower()
			#if assistance_mode:  # Only process commands when in genius mode
			if 'wikipedia' in query:
				speak('Searching Wikipedia... Please wait...')
				query = query.replace("wikipedia", "")
				results = wikipedia.summary(query, sentences=3)
				speak("According to Wikipedia")
				print(results)
				speak(results)

			elif 'switch to genius mode' in query:
				llm_activate()
			
			elif 'open youtube' in query:
				speak("Here you go to Youtube\n")
				webbrowser.open("youtube.com")
		
			elif 'open google' in query:
				speak("Here you go to Google\n")
				webbrowser.open("google.com")

			elif 'the time' in query:
				current_time=get_current_time()
				speak(f"The time currently is {current_time}")

			elif 'open chrome' in query:
				speak("Here we go, Opening Google Chrome for you!\n")
				codePath = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
				os.startfile(codePath)

			elif 'open discord' in query:
				speak("Here we go, Opening Discord for you!\n")
				codePath = r"C:\Users\bnave\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Discord Inc\Discord.lnk"
				os.startfile(codePath)

			elif 'open studio code' in query:
				speak("Here we go, Opening Visual Studio Code for you!\n")
				codePath = r"D:\Softwares\Microsoft VS Code\Code.exe"
				os.startfile(codePath)
			
			elif 'open zoom' in query:
				speak("Here we go, Opening Zoom for you!\n")
				codePath = r"C:\Users\bnave\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Zoom\Zoom Workplace.lnk"
				os.startfile(codePath)

			elif 'open word' in query:
				speak("Here we go, Opening Microsoft Word for you!\n")
				codePath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"
				os.startfile(codePath)

			elif 'open excel' in query:
				speak("Here we go, Opening Microsoft Excel for you!\n")
				codePath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk"
				os.startfile(codePath)

			elif 'open powerpoint' in query:
				speak("Here we go, Opening Microsoft PowerPoint for you!\n")
				codePath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk"
				os.startfile(codePath)
			
			elif 'open notion' in query:
				speak("Here we go, Opening Notion for you!\n")
				codePath = r"C:\Users\bnave\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Notion.lnk"

			elif 'open notion calendar' in query:
				speak("Here we go, Opening Notion Calendar for you!\n")
				codePath = r"C:\Users\bnave\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Notion Calendar.lnk"

			elif 'weather update' in query:
				speak("Which city you would like the weather for?")
				city=takeCommand()
				forecast=get_weather_forecast(city)
				speak(forecast)

			elif 'how are you' in query:
				speak("I am fine, Thank you for asking")

			elif 'fine' in query or "good" in query:
				speak("It's good to know that your fine")

			elif "change name" in query:
				speak("What would you like to call me? ")
				assname = takeCommand()
				speak("Thanks for naming me")

			elif "what's your name" in query or "What is your name" in query:
				speak("My friends call me")
				speak(assname)
				print("My friends call me", assname)

			elif 'exit' in query:
				speak("Thanks for giving me your time")
				exit()

			elif "who made you" in query or "who created you" in query:
				speak("I have been created by Naveen.")

			elif 'a joke' in query:
				speak(pyjokes.get_joke())

			elif "calculate" in query:

				app_id = "Y98U78-V377URAEV4"
				client = wolframalpha.Client(app_id)
				indx = query.lower().split().index('calculate')
				query = query.split()[indx + 1:]
				res = client.query(' '.join(query))
				answer = next(res.results).text
				print("The answer is " + answer)
				speak("The answer is " + answer)

			elif "who i am" in query:
				speak("If you talk then definitely your human.")

			elif "why you came to world" in query:
				speak("Thanks to Naveen. further It's a secret")

			elif "who are you" in query:
				speak("I am your virtual assistant created by Naveen")

			elif 'reason for you' in query:
				speak("I was created as a Minor project by Mister Naveen ")

			elif 'news' in query:

				try:

					speak('here are some top news from BBC News')
					print('''=============== BBC News ============'''+ '\n')

					def NewsFromBBC():
						query_params = {
						"source": "bbc-news",
						"sortBy": "top",
						"apiKey": "4dbc17e007ab436fb66416009dfb59a8"
							}
						main_url = " https://newsapi.org/v1/articles"

						# fetching data in json format
						res = requests.get(main_url, params=query_params)
						open_bbc_page = res.json()

						# getting all articles in a string article
						article = open_bbc_page["articles"]

						# empty list which will
						# contain all trending news
						results = []

						for ar in article:
							results.append(ar["title"])

						for i in range(len(results)):

							# printing all trending news
							print(i + 1, results[i])

						#to read the news out loud for us
						from win32com.client import Dispatch
						speak = Dispatch("SAPI.Spvoice")
						speak.Speak(results)				

					# Driver Code
					if __name__ == '__main__':

						# function call
						NewsFromBBC()

				except Exception as e:
					print(str(e))

			elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

			elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')

			elif 'empty recycle bin' in query:
				winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
				speak("Recycle Bin Recycled")

			elif "don't listen" in query or "stop listening" in query:
				speak("for how much time you want to stop DVA from listening commands")
				a = int(takeCommand())
				time.sleep(a)
				print(a)

			elif "restart" in query:
				subprocess.call(["shutdown", "/r"])

			elif "hibernate" in query or "sleep" in query:
				speak("Hibernating")
				subprocess.call("shutdown / h")

			elif "log off" in query or "sign out" in query:
				speak("Make sure all the application are closed before sign-out")
				time.sleep(5)
				subprocess.call(["shutdown", "/l"])

			elif "write a note" in query:
				speak("What should i write, sir")
				note = takeCommand()
				file = open('DVA.txt', 'w')
				file.write(note)

			elif "show note" in query:
				speak("Showing Notes")
				file = open("DVA.txt", "r")
				print(file.read())
				speak(file.read(6))

			elif "take screenshot" in query:
				take_screenshot()
		
			elif "live dictate" in query:
				response = live_dictate()
				speak(response)


