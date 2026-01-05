#import numpy as np
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

#testing
import fnmatch
import winreg
from pathlib import Path

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty("rate", 175)

# Function to speak out the given text
def speak(audio):
    # Using win32com is more stable and avoids conflicts with speech_recognition
    try:
        speaker = wincl.Dispatch("SAPI.SpVoice")
        speaker.Speak(audio)
    except Exception as e:
        print(f"TTS error: {e}")
		

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
        model = "llama3.2:1b"  # Replace with your model name
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

						# Generate response using Ollama
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

#testing
# Sprint-1 : Find and open applications by name on Windows
def find_and_open_app(app_name: str) -> bool:
    """
    Open an app by name on Windows.

    Strategy:
      0) UWP/system URIs (e.g., Settings via ms-settings:)
      1) Known built-in system executables (notepad, calc, etc.)
      2) Shell start (lets Windows resolve registered apps)
      3) PATH lookup (shutil.which)
      4) App Paths registry (HKCU/HKLM)
      5) Search Start Menu + common install folders
    """
    name = (app_name or "").strip().strip('"').strip("'")
    if not name:
        return False

    n = name.lower()

    # 0) UWP/system pages via URI (Settings isn't a normal .exe)
    uri_map = {
        "settings": "ms-settings:",
        "windows settings": "ms-settings:",
        "bluetooth": "ms-settings:bluetooth",
        "wifi": "ms-settings:network-wifi",
        "network": "ms-settings:network",
        "display": "ms-settings:display",
        "apps": "ms-settings:appsfeatures",
        "updates": "ms-settings:windowsupdate",
        "windows update": "ms-settings:windowsupdate",
        "privacy": "ms-settings:privacy",
        "sound": "ms-settings:sound",
    }
    if n in uri_map:
        try:
            os.startfile(uri_map[n])
            return True
        except Exception:
            try:
                subprocess.Popen(f'cmd /c start "" "{uri_map[n]}"', shell=True)
                return True
            except Exception:
                return False

    # 1) Built-in/system executables
    system_map = {
        "notepad": r"C:\Windows\System32\notepad.exe",
        "calc": r"C:\Windows\System32\calc.exe",
        "calculator": r"C:\Windows\System32\calc.exe",
        "paint": r"C:\Windows\System32\mspaint.exe",
        "cmd": r"C:\Windows\System32\cmd.exe",
        "powershell": r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        "control panel": r"C:\Windows\System32\control.exe",
        "task manager": r"C:\Windows\System32\taskmgr.exe",
    }
    if n in system_map and os.path.isfile(system_map[n]):
        subprocess.Popen([system_map[n]], shell=False)
        return True

    # Normalize candidates
    candidates: list[str] = [name]
    if not name.lower().endswith(".exe"):
        candidates.append(f"{name}.exe")

    # 2) Shell start (handles many registered apps / aliases)
    # Note: This may return True even if the app doesn't exist; keep it after known URIs/system.
    try:
        subprocess.Popen(f'cmd /c start "" "{name}"', shell=True)
        return True
    except Exception:
        pass

    # 3) PATH lookup
    for c in candidates:
        try:
            exe = shutil.which(c)
            if exe:
                subprocess.Popen([exe], shell=False)
                return True
        except Exception:
            pass

    # 4) App Paths registry lookup
    def _reg_app_path(exe_name: str) -> str | None:
        subkey = rf"Software\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
        hives = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
        for hive in hives:
            try:
                with winreg.OpenKey(hive, subkey) as k:
                    val, _ = winreg.QueryValueEx(k, "")
                    if val and os.path.isfile(val):
                        return val
            except OSError:
                continue
        return None

    for c in candidates:
        exe_name = c if c.lower().endswith(".exe") else f"{c}.exe"
        reg_path = _reg_app_path(exe_name)
        if reg_path:
            subprocess.Popen([reg_path], shell=False)
            return True

    # 5) Search common locations (shortcuts + install dirs + Windows folders)
    search_dirs = [
        os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
        os.path.join(os.environ.get("PROGRAMDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\WindowsApps"),
        r"C:\Windows\System32",
        r"C:\Windows",
        r"C:\Program Files",
        r"C:\Program Files (x86)",
    ]

    exts = ("*.lnk", "*.exe", "*.appref-ms", "*.url")
    needle = n

    for base in search_dirs:
        if not base or not os.path.isdir(base):
            continue
        for root, _, files in os.walk(base):
            for pattern in exts:
                for fname in fnmatch.filter(files, pattern):
                    if needle in fname.lower():
                        path = os.path.join(root, fname)
                        try:
                            os.startfile(path)
                            return True
                        except Exception:
                            try:
                                subprocess.Popen([path], shell=False)
                                return True
                            except Exception:
                                pass

    return False

#Sprint-1 : Voice Notes
NOTES_DIR = Path(__file__).resolve().parent / "notes"

def save_voice_note(text: str) -> Path:
    NOTES_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    note_path = NOTES_DIR / f"note_{ts}.txt"
    note_path.write_text(text.strip() + "\n", encoding="utf-8")
    return note_path

def get_latest_note_path() -> Path | None:
    if not NOTES_DIR.exists():
        return None
    notes = sorted(NOTES_DIR.glob("note_*.txt"), reverse=True)
    return notes[0] if notes else None

def read_latest_note() -> str | None:
    p = get_latest_note_path()
    if not p:
        return None
    return p.read_text(encoding="utf-8").strip()

#Sprint-1 : Open Websites
# There are 3 ways:
# 1) Hardcoded for common sites (YouTube, Google, etc.)
# 2) Searching the website directly through voice
# 3) If site not available, safely do a web search

WEBSITE_ALIASES = {
    "google": "https://www.google.com",
    "youtube": "https://www.youtube.com",
    "gmail": "https://mail.google.com",
    "github": "https://github.com",
    "stackoverflow": "https://stackoverflow.com",
    # add your frequent ones:
    # "nothing": "https://in.nothing.tech",
}

def _spoken_to_url(text: str) -> str:
    """
    Converts speech-like input to a usable URL.
    Examples:
      'wccftech dot com' -> 'wccftech.com'
      'https colon slash slash example dot com' (rough) -> best-effort
    """
    s = (text or "").strip().lower()

    # Common speech patterns
    s = s.replace(" dot ", ".")
    s = s.replace(" slash ", "/")
    s = s.replace(" colon ", ":")
    s = s.replace(" www ", " www.")
    s = s.replace("www dot ", "www.")
    s = s.strip().strip('"').strip("'")

    return s

def _looks_like_domain_or_url(s: str) -> bool:
    s = (s or "").strip().lower()
    if not s:
        return False
    if s.startswith(("http://", "https://")):
        return True
    # crude but effective: domains usually contain a dot and no spaces
    return ("." in s) and (" " not in s)

def _ensure_scheme(url: str) -> str:
    url = (url or "").strip()
    if url.startswith(("http://", "https://")):
        return url
    return "https://" + url

def open_website_target(target: str) -> bool:
    """
    Opens a website if target is an alias or looks like a domain/url.
    Returns True if opened as website, else False.
    """
    t = _spoken_to_url(target)

    # Alias first (open google/youtube without typing .com)
    if t in WEBSITE_ALIASES:
        webbrowser.open(WEBSITE_ALIASES[t])
        return True

    # Direct domain/url support
    if _looks_like_domain_or_url(t):
        webbrowser.open(_ensure_scheme(t))
        return True

    return False

def open_web_search(query_text: str) -> None:
    q = (query_text or "").strip()
    if not q:
        return
    webbrowser.open("https://www.google.com/search?q=" + requests.utils.quote(q))

#---------------------------------------------
# # # # # # Main Function # # # # # #
#---------------------------------------------

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	clear()
	wishMe()
	#usrname()

	assistance_mode = False  # Added a flag to track if we're in assistance mode

	while True:
		query = takeCommand().lower()
		if query == "jarvis" or query == "hello" or query == "assistant":
			speak("Speak")
			query = takeCommand().lower()
			#if assistance_mode:  # Only process commands when in genius mode
			if 'switch to genius mode' in query:
				llm_activate()
			
			elif 'wikipedia' in query:
				speak('Searching Wikipedia... Please wait...')
				query = query.replace("wikipedia", "")
				results = wikipedia.summary(query, sentences=3)
				speak("According to Wikipedia")
				print(results)
				speak(results)
			
			#elif 'open youtube' in query:
			#	speak("Here you go to Youtube\n")
			#	webbrowser.open("youtube.com")
			
			#elif 'open google' in query:
			#	speak("Here you go to Google\n")
			#	webbrowser.open("google.com")

			elif 'the time' in query:
				current_time=get_current_time()
				speak(f"The time currently is {current_time}")

			#elif 'open powerpoint' in query:
			#	speak("Here we go, Opening Microsoft PowerPoint for you!\n")
			#	codePath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk"
			#	os.startfile(codePath)
				
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
				print(pyjokes.get_joke())

			elif "calculate" in query:

				app_id = "EAW7LV24RK"
				client = wolframalpha.Client(app_id)
				indx = query.lower().split().index('calculate')
				calc_query = ' '.join(query.split()[indx + 1:])
				res = client.query(calc_query)
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

			# Writing and Reading Voice Notes - SPRINT-1
				
			elif "write a note" in query or "voice note" in query:
				speak("What should I write?")
				note_text = takeCommand()
				if not note_text or note_text.lower() == "none":
					speak("I didn't catch that.")
				else:
					path = save_voice_note(note_text)
					speak("Saved.")
					print(f"Saved note to: {path}")
			
			elif "show note" in query or "read note" in query:
				text = read_latest_note()
				if not text:
					speak("You have no notes yet.")
				else:
					print(text)
					# Speak a shorter preview to avoid long TTS
					preview = text if len(text) <= 300 else (text[:300] + "...")
					speak("Here is your latest note.")
					speak(preview)

			elif "take screenshot" in query:
				take_screenshot()
		
			elif "live dictate" in query:
				response = live_dictate()
				speak(response)

			# Open Applications by name - SPRINT-1
			elif query.startswith("open "):
				app_name = query.replace("open ", "", 1).strip()
				if find_and_open_app(app_name):
					speak(f"Opening {app_name} for you.")
				else:
					speak(f"Sorry, I couldn't find an application named {app_name}.")
			
			# Open Websites - SPRINT-1
			elif ("go to a website" in query) or ("go to website" in query) or ("i need to go to a website" in query):
				speak("Which website would you like to open?")
				target=takeCommand().strip()

				if not target or target.lower() == "none":
					speak("I didn't catch that.")
				elif target.lower() in ("cancel", "stop", "nevermind" , "never mind"):
					speak("Okay, cancelling.")
				else:
					if open_website_target(target):
						speak(f"Opening {target} for you.")
					else:
						speak(f"I couldn't identify {target} as a website, so I will search it for you.")
						open_web_search(target)
						speak(f"Here are the search results for {target}.")
			