import pyttsx3
engine = pyttsx3.init('sapi5')
engine.say("This is a test.")
engine.runAndWait()