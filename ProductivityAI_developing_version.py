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
import keyboard
import re
import json
import threading
import winsound
import dateparser 
from deep_translator import GoogleTranslator
from playsound import playsound
import shutil # Ensure shutil is imported

# --- NEW: User Data Management ---
USER_DATA_FILE = Path(__file__).parent / "user_data.json"

def load_user_data():
    if USER_DATA_FILE.exists():
        try:
            with open(USER_DATA_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_user_data(data):
    with open(USER_DATA_FILE, "w") as f:
        json.dump(data, f, indent=4)

# --- NEW: Contextual Memory for Genius Mode ---
CHAT_HISTORY = []  # Stores [{"role": "user", "content": "..."}, {"role": "assistant", "content": "..."}]

def get_chat_history_string():
    """Converts the last few turns of history into a context string."""
    # Keep only last 3 turns to prevent token overflow and confusion
    recent_history = CHAT_HISTORY[-6:] 
    context_str = ""
    for msg in recent_history:
        role = "User" if msg["role"] == "user" else "Assistant"
        context_str += f"{role}: {msg['content']}\n"
    return context_str

# --- NEW: Global variable for the current LLM model ---
CURRENT_LLM_MODEL = "llama3.2:1b"  # Default model

# --- NEW: Dictionary to map simple keywords to full model names ---
LLM_ALIASES = {
    "meta": "llama3.2:1b",
    "micrsosoft": "phi3.5:3.8b",
    "mistral": "mistral-large-3:675b-cloud",
    "google": "gemma3:4b-cloud",
    "coder": "qwen3-coder:480b-cloud"
}

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty("rate", 175)

# Function to speak out the given text
def speak(audio):
    #Using win32com is more stable and avoids conflicts with speech_recognition
    try:
        speaker = wincl.Dispatch("SAPI.SpVoice")
        speaker.Speak(audio)
    except Exception as e:
        print(f"TTS error: {e}")

# --- NEW: Cancellable Speak Function (For LLM only) ---
def speak_cancellable(audio):
    """Interruptible speech by chunking text (press 's' to stop)."""
    try:
        # Split text into chunks (sentences) so we can check for interruption
        parts = re.split(r'(?<=[.!?])\s+', audio)
        
        for part in parts:
            if not part.strip():
                continue

            # Check if user wants to stop
            if keyboard.is_pressed('s'):
                print("\n[Stopped speaking by user]")
                speak("Stopped.")
                break
            
            # Speak the current part synchronously
            speak(part)
            
    except Exception as e:
        print(f"TTS error: {e}")
        # Valid fallback
        speak(audio)

# --- NEW: Clean Text Function ---
def clean_text_for_speech(text):
    """Removes Markdown (*, #, etc) so the assistant speaks clearly."""
    # Remove bold/italic markers
    text = re.sub(r'\*\*|__', '', text) 
    text = re.sub(r'\*|_', '', text)
    # Remove headers
    text = re.sub(r'#+\s', '', text)
    # Remove code blocks
    text = re.sub(r'```[\s\S]*?```', 'Code block omitted.', text)
    return text.strip()

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
    data = load_user_data()
    saved_name = data.get("user_name")

    if saved_name:
        speak(f"Welcome back, {saved_name}")
        columns = shutil.get_terminal_size().columns
        print("#####################".center(columns))
        print(f"Welcome Back {saved_name}".center(columns))
        print("#####################".center(columns))
    else:
        speak("What should I call you?")
        uname = takeCommand()
        if uname and uname != "None":
            speak(f"Welcome Mr. {uname}")
            columns = shutil.get_terminal_size().columns
            print("#####################".center(columns))
            print(f"Welcome {uname}".center(columns))
            print("#####################".center(columns))
            
            # Save the new name
            data["user_name"] = uname
            save_user_data(data)
        else:
            speak("I didn't catch your name, I'll call you User for now.")

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

# --- MODIFIED: Function to generate response using the current global model ---
def generate_response(prompt):
    global CURRENT_LLM_MODEL, CHAT_HISTORY
    try:
        # Create an instance of the Client class
        client = Client()
        
        # 1. Prepare messages for Chat API
        # We start with a system message, then append history and the new prompt
        messages = [
            {"role": "system", "content": "You are a helpful AI assistant named ProductivityAI."}
        ] + CHAT_HISTORY + [{"role": "user", "content": prompt}]

        print(f"--- Generating response using {CURRENT_LLM_MODEL} ---")
        
        # 2. Generate response using the chat API (Streaming/Chat aware)
        response_obj = client.chat(model=CURRENT_LLM_MODEL, messages=messages)
        response_text = response_obj['message']['content']
        
        # 3. Update History
        CHAT_HISTORY.append({"role": "user", "content": prompt})
        CHAT_HISTORY.append({"role": "assistant", "content": response_text})
        
        # Optional: Prune history if it gets too long (e.g., keep last 20 turns)
        if len(CHAT_HISTORY) > 40:
            CHAT_HISTORY = CHAT_HISTORY[-40:]
            
        return response_text 
    except Exception as e:
        print(f"Error generating response: {str(e)}")
        return "I'm sorry, I couldn't process your request."
	
# --- NEW & IMPROVED: Function to switch the active LLM using keywords ---
def switch_llm(model_keyword: str):
    """
    Switches the active LLM based on a keyword.
    """
    global CURRENT_LLM_MODEL, LLM_ALIASES
    
    keyword = model_keyword.lower().strip()
    
    # Check if the keyword exists in our alias dictionary
    if keyword not in LLM_ALIASES:
        speak(f"Sorry, I don't recognize the model keyword '{keyword}'.")
        print(f"Known keywords are: {', '.join(LLM_ALIASES.keys())}")
        return

    target_model_name = LLM_ALIASES[keyword]
    print(f"Attempting to switch to model: {target_model_name}")

    try:
        # Simply switch the model without validating against the local list
        # This is necessary because cloud models won't appear in the local 'ollama list'
        CURRENT_LLM_MODEL = target_model_name
        speak(f"Successfully switched model to {keyword}.")
        print(f"Model switched to: {CURRENT_LLM_MODEL}")

    except Exception as e:
        speak("Sorry, I couldn't connect to Ollama to switch the model.")
        print(f"Error switching LLM: {e}")



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

						# Check for model switching command inside Genius Mode
						if "switch model to" in text.lower() or "change model to" in text.lower() or "switch llm to" in text.lower():
							model_keyword = text.split("to")[-1].strip()
							if model_keyword and model_keyword.lower() != "none":
								switch_llm(model_keyword)
								continue # Skip generating a response for the switch command
							else:
								speak("I didn't catch the model name.")
								continue

						# Generate response using Ollama
						response = generate_response(text)  # Call the function to get a response

						print(f"Your LLM says: {response}")

						# Read response using text-to-speech
						speak_cancellable(clean_text_for_speech(response))

				elif transcription.lower() == "switch to normal mode" or transcription.lower() == "exit genius mode":  #trying to switch to normal mode
					print("Here you go, going back to normal mode")
					speak("Here you go, going back to normal mode")
					break
					
						
			except Exception as e:
				print(f"An error occurred: {str(e)}")
				speak(f"An error occurred!")



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

# --- NEW: Workflow Automation Engine ---
def execute_workflow(mode_name):
    """
    Reads workflows.json and executes the Apps/URLs for the given mode.
    """
    try:
        file_path = Path(__file__).parent / "workflows.json"
        if not file_path.exists():
            speak("I cannot find the workflows configuration file.")
            return

        with open(file_path, "r") as f:
            data = json.load(f)

        # Fuzzy match the mode name (e.g., "work" matches "work", "study" matches "study")
        mode_data = data.get(mode_name.lower())
        
        if not mode_data:
            speak(f"Sorry, I don't have a workflow for {mode_name} mode.")
            return

        # 1. Open URLs
        urls = mode_data.get("urls", [])
        for url in urls:
            webbrowser.open(url)
            time.sleep(1) # Brief pause

        # 2. Open Apps
        apps = mode_data.get("apps", [])
        for app in apps:
            # Reuse our existing find_and_open_app for consistency
            # We don't speak failure for every app to keep it snappy, just print
            if not find_and_open_app(app):
                print(f"Workflow warning: Could not open application '{app}'")

        # 3. Speak Message
        msg = mode_data.get("message", f"{mode_name} mode enabled.")
        speak(msg)

    except Exception as e:
        print(f"Workflow Error: {e}")
        speak("I encountered an error executing that workflow.")

# --- NEW: Language Translation Function ---
def translate_and_speak():
    try:
        # 1. Get Source Text
        speak("What should I translate?")
        print("Listening for text to translate...")
        text = takeCommand()
        
        if text == "None" or not text:
            speak("I didn't catch that.")
            return

        # 2. Get Target Language
        speak("Which language should I translate to? For example, say 'French', 'Hindi', or 'German'.")
        print("Listening for target language...")
        lang_input = takeCommand().lower()

        if lang_input == "None":
            speak("I didn't catch the language.")
            return

        # Map spoken language names to codes (add more as needed)
        lang_map = {
            "hindi": "hi",
            "french": "fr",
            "german": "de",
            "spanish": "es",
            "italian": "it",
            "japanese": "ja",
            "korean": "ko",
            "chinese": "zh-CN",
            "russian": "ru",
            "tamil": "ta",
            "telugu": "te"
        }

        # Simple fuzzy match or direct lookup
        target_code = lang_map.get(lang_input)
        if not target_code:
            # Fallback: try using the input directly if it's a valid code, or fail
            speak(f"Sorry, I don't know the code for {lang_input} yet.")
            return

        # 3. Translate
        translator = GoogleTranslator(source='auto', target=target_code)
        translated_text = translator.translate(text)
        print(f"Translated: {translated_text}")

        # 4. Speak Output (using gTTS for native accent)
        # We save to a temp file because standard 'speak' (SAPI5) only does English well.
        speak(f"In {lang_input}, it is:")
        
        tts = gTTS(text=translated_text, lang=target_code, slow=False)
        temp_file = "translated_audio.mp3"
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        tts.save(temp_file)
        playsound(temp_file)
        os.remove(temp_file)

    except Exception as e:
        print(f"Translation Error: {e}")
        speak("Something went wrong with the translation.")

# --- NEW: Background Alarm System ---
ALARMS = []

def alarm_monitor():
    """Checks for due alarms in the background."""
    while True:
        now = datetime.datetime.now()
        # Iterate over a copy to allow safe removal
        for alarm in ALARMS[:]:
            if now >= alarm['time']:
                reason = alarm.get('reason', 'Alarm')
                print(f"\n[ALARM] Time is up! {reason}")
                
                # Visual/Audio Alert
                winsound.Beep(1000, 500) # Frequency 1000Hz, Duration 500ms
                
                if "Reminder:" in reason:
                     speak(f"Excuse me. {reason}")
                else:
                     speak(f"Time is up! {reason}")
                
                ALARMS.remove(alarm)
        time.sleep(1)

def set_alarm_from_voice():
    """Parses voice input to set an alarm."""
    speak("When should I set the alarm for?")
    print("Listening for time (e.g., 'in 10 minutes', 'at 5 pm')...")
    
    time_query = takeCommand()
    if not time_query or time_query.lower() == "none":
        speak("I didn't hear a time.")
        return

    # Parse time using dateparser
    # settings={'PREFER_DATES_FROM': 'future'} helps with "at 5pm" referring to today's 5pm if it hasn't passed, or tomorrow? 
    # Actually dateparser defaults are usually okay, but let's be safe.
    dt = dateparser.parse(time_query, settings={'PREFER_DATES_FROM': 'future'})
    
    if dt:
        # If the parsed time is in the past (e.g. parsed '5pm' but it is 6pm), dateparser often returns past time by default if not strictly future.
        # We can adjust:
        if dt < datetime.datetime.now():
            # If it's seemingly in the past, maybe they meant tomorrow? 
            # Simple heuristic: add 24 hours if difference is within a day
            dt += datetime.timedelta(days=1)

        ALARMS.append({'time': dt, 'reason': 'User set alarm'})
        confirm_msg = f"Alarm set for {dt.strftime('%I:%M %p')}"
        print(confirm_msg)
        speak(confirm_msg)
    else:
        speak("Sorry, I couldn't understand the time.")

# --- NEW: Specific Reminder Function ---
def set_reminder_from_voice():
    """Parses voice input to set a reminder with a specific message."""
    
    # 1. Ask for the Task
    speak("What should the reminder be about?")
    print("Listening for reminder task...")
    task = takeCommand()
    
    if not task or task.lower() == "none":
        speak("I didn't catch the reminder message.")
        return

    # 2. Ask for the Time
    speak("When should I remind you?")
    print("Listening for time (e.g., 'in 20 minutes', 'at 10 am')...")
    time_query = takeCommand()
    
    if not time_query or time_query.lower() == "none":
        speak("I didn't hear a time.")
        return

    # 3. Parse and Store
    dt = dateparser.parse(time_query, settings={'PREFER_DATES_FROM': 'future'})
    
    if dt:
        if dt < datetime.datetime.now():
            dt += datetime.timedelta(days=1) # Assume tomorrow if time is past

        ALARMS.append({'time': dt, 'reason': f"Reminder: {task}"})
        
        confirm_msg = f"Okay, I will remind you to {task} at {dt.strftime('%I:%M %p')}"
        print(confirm_msg)
        speak(confirm_msg)
    else:
        speak("Sorry, I couldn't understand the time.")

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

	# Load assistant name if exists
	user_data = load_user_data()
	assname = user_data.get("assistant_name", "ProductivityAI")

	wishMe()
	usrname()

	assistance_mode = False  # Added a flag to track if we're in assistance mode

	# Start the alarm monitor thread
	t = threading.Thread(target=alarm_monitor, daemon=True)
	t.start()

	while True:
		query = takeCommand().lower()
		if query == "jarvis" or query == "hello" or query == "assistant":
			speak("Speak")
			query = takeCommand().lower()
			#if assistance_mode:  # Only process commands when in genius mode
			if 'switch to genius mode' in query:
				llm_activate()
			elif "switch model to" in query or "change model to" in query or "switch llm to" in query:
                # Extract the model keyword from the query
				model_keyword = query.split("to")[-1].strip()
				if model_keyword and model_keyword.lower() != "none":
					switch_llm(model_keyword)
				else:
					speak("I didn't catch the model name.")
			
			elif 'wikipedia' in query:
				speak('Searching Wikipedia... Please wait...')
				query = query.replace("wikipedia", "")
				results = wikipedia.summary(query, sentences=3)
				speak("According to Wikipedia")
				print(results)
				speak(results)

			# --- NEW: Workflow Automation Commands ---
			elif "enable" in query and "mode" in query:
				# Example: "enable work mode" -> mode = "work"
				mode = query.replace("enable", "").replace("mode", "").strip()
				execute_workflow(mode)

			elif "start" in query and "workflow" in query:
				# Example: "start study workflow" -> mode = "study"
				mode = query.replace("start", "").replace("workflow", "").strip()
				execute_workflow(mode)

			elif "set" in query and "alarm" in query:
				set_alarm_from_voice()

			# --- NEW: Reminder Command ---
			elif "remind me" in query or "set a reminder" in query:
				set_reminder_from_voice()

			elif "translate" in query or "translation" in query:
				translate_and_speak()
			
			#elif 'open youtube' in query:
			#	speak("Here you go to Youtube\n")
			#	webbrowser.open("youtube.com")
			
			#elif 'open google' in query:
			#	speak("Here you go to Google\n")
			#	webbrowser.open("google.com")

			elif "change name" in query or "change my name" in query:
				speak("Wait, do you want to change MY name, or YOUR name?")
				ans = takeCommand().lower()
				
				if "your" in ans or "assistant" in ans:
					speak("Okay, what is my new name?")
					new_assname = takeCommand()
					if new_assname and new_assname != "None":
						data = load_user_data()
						data["assistant_name"] = new_assname
						save_user_data(data)
						assname = new_assname
						speak(f"Thanks, I like the name {assname}")
				
				else:
					speak("Okay, what should I call you from now on?")
					new_name = takeCommand()
					if new_name and new_name != "None":
						data = load_user_data()
						data["user_name"] = new_name
						save_user_data(data)
						speak(f"Done. I will call you {new_name}.")

			elif 'the time' in query:
				current_time = get_current_time()
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

			# Check if 'articles' exists and has data
			elif 'news' in query:
				try:
					speak('Here are the top headlines from India. Press s to stop reading.')
					print('''=============== Indian News ============'''+ '\n')

					# Updated API params for Indian Top Headlines (NewsData.io)
					query_params = {
						"apikey": "pub_f036840f94a84f65a9eed20f90abbc3c",
						"country": "in",
						"timezone": "Asia/Kolkata",
						"language": "en"
					}
					main_url = "https://newsdata.io/api/1/latest"

					# fetching data in json format
					res = requests.get(main_url, params=query_params)
					news_data = res.json()
					
					# Check if 'results' exists and has data (NewsData.io uses 'results' instead of 'articles')
					if "results" in news_data:
						articles = news_data["results"]
						headlines = []
						
						# Collect top 5 or 10 headlines to avoid reading for too long
						for i, ar in enumerate(articles[:10]):
							title = ar.get("title")
							if title:
								headlines.append(title)
								print(f"{i + 1}. {title}")

						if headlines:
							# Join all headlines into one long string for the cancellable speaker
							full_news_readout = ". Next headline: ".join(headlines)
							speak_cancellable(full_news_readout)
						else:
							speak("No headlines found at the moment.")
					else:
						print(f"API Error: {news_data}")
						speak("I couldn't fetch the news data.")

				except Exception as e:
					print(str(e))
					speak("Sorry, I am not able to fetch news at the moment.")

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
			elif ("search a website" in query) or ("search website" in query) or ("i need to go to a website" in query):
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
						print(f"Could not identify {target} as a website, performing web search instead.")
						speak(f"I couldn't identify {target} as a website, so I will search it for you.")
						open_web_search(target)
						print(f"Here are the search results for {target}.")
						speak(f"Here are the search results for {target}.")
			