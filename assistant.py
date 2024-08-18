import os
import win32com.client
import speech_recognition as sr
import webbrowser
import datetime
import random

def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.pause_threshold = 0.6
        print("Listening...")
        audio = recognizer.listen(source)
        try:
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            speak("Sorry, I did not understand that. Please try again.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            speak("Could not request results. Please check your network connection and try again.")
            return None

def tell_joke():
    jokes = [
        "Why don't scientists trust atoms? Because they make up everything!",
        "Why did the math book look sad? Because it had too many problems.",
        "Why did the scarecrow win an award? Because he was outstanding in his field!"
    ]
    return jokes[random.randint(0, len(jokes) - 1)]

def main():
    speak("How can I help you ?")
    while True:
        query = take_command()
        if query:
            if "exit" in query or "quit" in query or "stop" in query:
                speak("Goodbye, have a nice day!")
                break

            sites = [
                ["youtube", "https://www.youtube.com/"], 
                ["wikipedia", "https://www.wikipedia.org/"], 
                ["google", "https://www.google.com/"],
                ["facebook", "https://www.facebook.com/"],
                ["twitter", "https://www.twitter.com/"],
                ["instagram", "https://www.instagram.com/"],
                ["reddit", "https://www.reddit.com/"],
                ["linkedin", "https://www.linkedin.com/"],
                ["amazon", "https://www.amazon.com/"],
                ["netflix", "https://www.netflix.com/"]
            ]
            opened = False
            for site in sites:
                if f"open {site[0]}" in query:
                    speak(f"Opening {site[0]}, sir.")
                    webbrowser.open(site[1])
                    opened = True
                    break
            if "open music" in query:
                musicpath = r"C:\Users\kiran\Downloads\Apna-Bana-Le(PagalWorld).mp3"
                os.startfile(musicpath)
                speak("Playing music.")
                opened = True
            if "tell me a joke" in query:
                joke = tell_joke()
                speak(joke)
                opened = True
            if "open " in query:
                search_query = query.split("search for")[1].strip()
                webbrowser.open(f"https://www.google.com/search?q={search_query}")
                speak(f"Searching for {search_query}")
                opened = True
            if "time" in query:
                str_time = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"Sir, the time is {str_time}")
                opened = True
            if "date" in query:
                str_date = datetime.datetime.now().strftime("%Y-%m-%d")
                speak(f"Sir, today's date is {str_date}")
                opened = True
            if "open file manager" in query:
                os.startfile("explorer")
                speak("Opening File Manager.")
                opened = True
            if "open vs code" in query:
                vscode_path = r"C:\Users\kiran\OneDrive\Desktop\Visual Studio Code.lnk"  
                os.startfile(vscode_path)
                speak("Opening Visual Studio Code.")
                opened = True
            if not opened:
                speak(f"You said: {query}")

if __name__ == '__main__':
    main()
