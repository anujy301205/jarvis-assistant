import speech_recognition as sr
import asyncio
import edge_tts
import pyautogui
import webbrowser
import time
import os
import random
import wikipedia
from playsound import playsound
from datetime import datetime
import subprocess
import socket
import requests
import urllib.parse
import json
import pygetwindow as gw
import pyperclip
import win32gui
import win32con
import win32process
import win32api
import shutil
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

from tuya_iot import TuyaOpenAPI
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import screen_brightness_control as sbc

# Set working directory
os.chdir(r"C:\Users\Lappy\Jarvis")
pyautogui.FAILSAFE = False

# ============ GLOBAL MEMORY ============
memory_store = []

# ============ CONNECTION CHECK ============
def is_connected():
    try:
        socket.create_connection(("1.1.1.1", 53), timeout=2)
        return True
    except OSError:
        return False

# ============ SPEAK FUNCTION ============
async def speak_async(text):
    print("Jarvis:", text)
    filename = f"jarvis_{datetime.now().strftime('%H%M%S%f')}.mp3"
    communicate = edge_tts.Communicate(text, voice="en-IN-NeerjaNeural")
    await communicate.save(filename)
    playsound(filename)
    os.remove(filename)

def speak(text):
    if not is_connected():
        print(f"[Jarvis is offline] Would have said: {text}")
        return
    asyncio.run(speak_async(text))

# ============ LISTEN FUNCTION ============
def listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎤 Listening...")
        r.adjust_for_ambient_noise(source, duration=0.3)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio)
        print("You said:", command)
        return command.lower()
    except (sr.UnknownValueError, sr.RequestError):
        print("❌ Could not understand or connect.")
        return ""

# ============ JOKES ============
jokes = [
    "Why did the computer go to therapy? Because it had too many bytes of trauma.",
    "I'm not lazy, I'm just on energy-saving mode.",
    "Why don’t programmers like nature? It has too many bugs.",
    "I told my computer I needed a break, and it said 'No problem, I’ll crash!'"
]

# ============ EMOTIONAL REACTIONS ============
def get_reaction():
    moods = [
        " Alrighty!",
        " Got it boss!",
        " Executing that now.",
        " On it, commander.",
        " Hope I do this right.",
        " Let's gooo!"
    ]
    return random.choice(moods)

# ============ Youtube Search ============
def play_youtube_song(query):
    try:
        search_query = query.replace("play", "").replace("on youtube", "").strip()
        url = f"https://www.youtube.com/results?search_query={search_query.replace(' ', '+')}"
        webbrowser.open(url)
        speak(f"Searching YouTube for {search_query}")
        
        time.sleep(4)  # wait for YouTube to load
        pyautogui.moveTo(580, 360)  # adjust this to where the first video is
        pyautogui.click()
        speak("Playing the top result.")
    except Exception as e:
        speak(f"Couldn't play the video. Error: {str(e)}")


# ============ SYSTEM CONTROL ============
def toggle_wifi(turn_on=True):
    if turn_on:
        subprocess.run('netsh wlan connect', shell=True)
        speak("Wi-Fi reconnected.")
    else:
        subprocess.run('netsh wlan disconnect', shell=True)
        speak("Wi-Fi disconnected.")

def open_bluetooth_settings():
    speak("Opening Bluetooth settings.")
    os.system("start ms-settings:bluetooth")
    time.sleep(3)  # allow settings to open


def smart_toggle_bluetooth(desired_state="on"):
    open_bluetooth_settings()

    time.sleep(3)  # Give Settings time to open

    try:
        if desired_state == "on":
            toggle_location = pyautogui.locateCenterOnScreen("bluetooth_off.png", confidence=0.8)
            if toggle_location:
                pyautogui.click(toggle_location)
                speak("Turning on Bluetooth.")
            else:
                speak("Bluetooth is already on.")
        elif desired_state == "off":
            toggle_location = pyautogui.locateCenterOnScreen("bluetooth_on.png", confidence=0.8)
            if toggle_location:
                pyautogui.click(toggle_location)
                speak("Turning off Bluetooth.")
            else:
                speak("Bluetooth is already off.")
        else:
            speak("Unknown Bluetooth command.")
    except Exception as e:
        print("❌ Error toggling Bluetooth:", e)
        speak("Something went wrong while toggling Bluetooth.")

    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')  # Close the Settings window



def turn_on_bluetooth():
    smart_toggle_bluetooth("on")

def turn_off_bluetooth():
    smart_toggle_bluetooth("off")

def set_volume(level):
    if 0 <= level <= 100:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100, None)
        speak(f"Volume set to {level} percent.")
    else:
        speak("Please provide a volume between 0 and 100.")

def set_brightness(level):
    if 0 <= level <= 100:
        sbc.set_brightness(level)
        speak(f"Brightness set to {level} percent.")
    else:
        speak("Please provide a brightness level between 0 and 100.")

# ============ WIKIPEDIA SUMMARY ============
def fetch_wikipedia_summary(query):
    try:
        summary = wikipedia.summary(query, sentences=2)
        speak(f"According to Wikipedia: {summary}")
    except wikipedia.exceptions.DisambiguationError as e:
        first_option = e.options[0]
        speak(f"Did you mean {first_option}? Let me tell you about it.")
        summary = wikipedia.summary(first_option, sentences=2)
        speak(f"According to Wikipedia: {summary}")
    except wikipedia.exceptions.PageError:
        speak("I couldn't find anything on that topic.")
    except Exception as e:
        speak("Something went wrong while fetching information.")

# ============ ChatBot ============

def ask_ai(message):
    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={"model": "mistral", "prompt": message, "stream": True},
            stream=True
        )

        output = ""
        for line in response.iter_lines():
            if line:
                try:
                    data = json.loads(line.decode("utf-8"))
                    token = data.get("response", "")
                    print(token, end="", flush=True)  # shows stream live in terminal
                    output += token
                except:
                    continue
        print()  # newline after full response
        return output.strip()
    except Exception as e:
        print("❌ Error talking to AI:", e)
        return "I couldn't connect to the AI model."

# ============ Full Control ============    

def find_closest_window(title_fragment):
    windows = gw.getWindowsWithTitle(title_fragment)
    return windows[0] if windows else None

def bring_window_to_front(window_title):
    try:
        win = find_closest_window(window_title)
        if win:
            win.restore()
            win.activate()
            return True
        return False
    except:
        return False

def minimize_window(window_title):
    try:
        win = find_closest_window(window_title)
        if win:
            win.minimize()
            return True
        return False
    except:
        return False

def close_app(app_name):
    try:
        os.system(f'taskkill /f /im {app_name}.exe')
        return True
    except:
        return False

def read_clipboard():
    try:
        text = pyperclip.paste()
        speak(f"Clipboard contains: {text}")
    except:
        speak("Couldn't read clipboard.")

def capture_screen():
    try:
        filename = f"screenshot_{datetime.now().strftime('%H%M%S')}.png"
        screenshot = pyautogui.screenshot()
        screenshot.save(filename)
        speak("Screen captured.")
    except:
        speak("Failed to capture screen.")

def type_text(text):
    pyautogui.write(text)
    speak("Typed the given text.")

def press_key(key):
    pyautogui.press(key)
    speak(f"Pressed {key} key.")

 # ============ FILE MANAGEMENT ============
def handle_file_management(command):
    command = command.lower()

    # CREATE FOLDER
    if "create folder" in command:
        folder_name = command.replace("create folder", "").strip()
        if folder_name:
            try:
                os.makedirs(folder_name, exist_ok=True)
                speak(f"Folder '{folder_name}' created successfully.")
            except Exception as e:
                speak(f"Error creating folder: {str(e)}")
        else:
            speak("Please specify the folder name.")

    # DELETE FILE/FOLDER
    elif "delete" in command:
        path = command.replace("delete", "").strip()
        if path:
            if os.path.isdir(path):
                try:
                    shutil.rmtree(path)
                    speak(f"Folder '{path}' deleted successfully.")
                except Exception as e:
                    speak(f"Error deleting folder: {str(e)}")
            elif os.path.isfile(path):
                try:
                    os.remove(path)
                    speak(f"File '{path}' deleted successfully.")
                except Exception as e:
                    speak(f"Error deleting file: {str(e)}")
            else:
                speak("Path not found.")
        else:
            speak("Please specify what to delete.")

    # MOVE FILE/FOLDER
    elif "move" in command and "to" in command:
        try:
            parts = command.split("to")
            src = parts[0].replace("move", "").strip()
            dest = parts[1].strip()
            shutil.move(src, dest)
            speak(f"Moved '{src}' to '{dest}'.")
        except Exception as e:
            speak(f"Error moving: {str(e)}")

    # COPY FILE/FOLDER
    elif "copy" in command and "to" in command:
        try:
            parts = command.split("to")
            src = parts[0].replace("copy", "").strip()
            dest = parts[1].strip()
            if os.path.isdir(src):
                shutil.copytree(src, dest)
            else:
                shutil.copy2(src, dest)
            speak(f"Copied '{src}' to '{dest}'.")
        except Exception as e:
            speak(f"Error copying: {str(e)}")

    # RENAME FILE/FOLDER
    elif "rename" in command and "to" in command:
        try:
            parts = command.split("to")
            old_name = parts[0].replace("rename", "").strip()
            new_name = parts[1].strip()
            os.rename(old_name, new_name)
            speak(f"Renamed '{old_name}' to '{new_name}'.")
        except Exception as e:
            speak(f"Error renaming: {str(e)}")

    # OPEN FOLDER/FILE
    elif "open" in command:
        path = command.replace("open", "").strip()
        if os.path.exists(path):
            try:
                os.startfile(path)
                speak(f"Opening {path}.")
            except Exception as e:
                speak(f"Error opening: {str(e)}")
        else:
            speak("Path not found.")

    # SEARCH FILE/FOLDER
    elif "search for" in command:
        keyword = command.replace("search for", "").strip()
        found = []
        for root, dirs, files in os.walk("."):
            for name in files + dirs:
                if keyword.lower() in name.lower():
                    found.append(os.path.join(root, name))
        if found:
            speak(f"Found {len(found)} result(s). Listing top 3:")
            for path in found[:3]:
                speak(path)
        else:
            speak("No results found.")

# ============ SYSTEM SETTINGS CONTROL ============

def shutdown():
    speak("Shutting down the system.")
    os.system("shutdown /s /t 1")

def restart():
    speak("Restarting the system.")
    os.system("shutdown /r /t 1")

def lock_pc():
    speak("Locking the system.")
    ctypes.windll.user32.LockWorkStation()

def logout():
    speak("Logging out now.")
    os.system("shutdown -l")

def mute_system():
    pyautogui.press("volumemute")
    speak("System muted.")

def unmute_system():
    pyautogui.press("volumemute")
    speak("System unmuted.")

def get_battery_status():
    try:
        import psutil
        battery = psutil.sensors_battery()
        percent = battery.percent
        plugged = battery.power_plugged
        status = "charging" if plugged else "not charging"
        speak(f"Battery is at {percent} percent and it is {status}.")
    except Exception as e:
        speak("I couldn't fetch battery status.")

def get_ip_address():
    try:
        hostname = socket.gethostname()
        ip_address = socket.gethostbyname(hostname)
        speak(f"Your IP address is {ip_address}.")
    except:
        speak("I couldn't fetch the IP address.")


# ============ EXECUTE FUNCTION ============
def execute(command):
    command = command.strip()
    print("✅ Received command:", command)  # Debug print

    # ✅ Check and route file management commands first
    if any(keyword in command for keyword in [
        "create folder", "delete", "move", "copy", "rename", "open", "search for"
    ]):
        handle_file_management(command)
        return

    # ✅ Memory commands
    if command.startswith("remember that"):
        fact = command.replace("remember that", "").strip()
        memory_store.append(fact)
        speak(f"I'll remember that {fact}.")
        return

    elif "what did you remember" in command or "what do you remember" in command:
        if memory_store:
            speak("Here's what I remember:")
            for fact in memory_store:
                speak(f"- {fact}")
        else:
            speak("I don't remember anything yet.")
        return

    # ✅ Fun & info
    elif "tell me a joke" in command:
        speak(random.choice(jokes))
        return

    elif "tell me about" in command:
        topic = command.replace("tell me about", "").strip()
        fetch_wikipedia_summary(topic)
        return

    # ✅ App control
    elif "open notepad" in command:
        speak(get_reaction())
        pyautogui.press('win')
        time.sleep(1)
        pyautogui.write("notepad")
        pyautogui.press('enter')
        speak("Opening Notepad.")

    elif "close notepad" in command:
        speak(get_reaction())
        os.system("taskkill /f /im notepad.exe")
        speak("Closing Notepad.")

    elif "close youtube" in command:
        speak(get_reaction())
        os.system("taskkill /f /im brave.exe")
        speak("Closing YouTube in Brave.")

    elif "play" in command and "youtube" in command:
        play_youtube_song(command)

    # ✅ System toggles
    elif "turn off wi-fi" in command or "disconnect wi-fi" in command:
        toggle_wifi(False)
    elif "turn on wi-fi" in command or "reconnect wi-fi" in command:
        toggle_wifi(True)

    elif "turn on bluetooth" in command:
        turn_on_bluetooth()

    elif "turn off bluetooth" in command:
        turn_off_bluetooth()

    # ✅ Volume / Brightness
    elif "set volume to" in command:
        try:
            level = int(command.split("set volume to")[1].strip().replace("%", ""))
            set_volume(level)
        except:
            speak("I didn't catch the volume level.")

    elif "set brightness to" in command:
        try:
            level = int(command.split("set brightness to")[1].strip().replace("%", ""))
            set_brightness(level)
        except:
            speak("I didn't catch the brightness level.")

    elif "exit" in command:
        speak("Goodbye sir.")
        exit()

    # ✅ Window control
    elif "bring" in command and "front" in command:
        app = command.replace("bring", "").replace("to front", "").strip()
        if bring_window_to_front(app):
            speak(f"Brought {app} to front.")
        else:
            speak(f"Couldn't find window for {app}.")

    elif "minimize" in command:
        app = command.replace("minimize", "").strip()
        if minimize_window(app):
            speak(f"Minimized {app}.")
        else:
            speak(f"Couldn't find window for {app}.")

    elif "close" in command:
        app = command.replace("close", "").strip()
        if close_app(app):
            speak(f"Closed {app}.")
        else:
            speak(f"Couldn't close {app}.")

    elif "switch window" in command:
        windows = [win for win in gw.getAllWindows() if win.title]
        if len(windows) > 1:
            windows[1].activate()
            speak(f"Switched to {windows[1].title}")
        else:
            speak("Not enough windows to switch.")

    elif "read clipboard" in command:
        read_clipboard()

    elif "take screenshot" in command or "capture screen" in command:
        capture_screen()

    elif "type" in command:
        text = command.replace("type", "").strip()
        type_text(text)

    elif "press" in command:
        key = command.replace("press", "").strip()
        press_key(key)

    elif "shutdown" in command:
        shutdown()

    elif "restart" in command:
        restart()

    elif "lock" in command:
        lock_pc()

    elif "log out" in command or "logout" in command:
        logout()

    elif "mute" in command:
        mute_system()

    elif "unmute" in command:
        unmute_system()

    elif "battery" in command:
        get_battery_status()

    elif "ip address" in command or "what is my ip" in command:
        get_ip_address()

    
    # ✅ Fallback to AI if nothing matches
    else:
        speak("Let me think...")
        ai_reply = ask_ai(command)
        speak(ai_reply)

# Assuming these are defined in your Jarvis script
from jarvis import speak, listen, execute

class JarvisGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Jarvis Assistant")
        self.setGeometry(200, 200, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Hello, I am Jarvis.")
        self.label.setFont(QFont("Arial", 18))
        self.label.setAlignment(Qt.AlignCenter)

        self.output_box = QTextEdit()
        self.output_box.setFont(QFont("Consolas", 12))
        self.output_box.setReadOnly(True)

        self.listen_button = QPushButton("🎤 Listen")
        self.listen_button.setFont(QFont("Arial", 14))
        self.listen_button.clicked.connect(self.handle_command)

        layout.addWidget(self.label)
        layout.addWidget(self.output_box)
        layout.addWidget(self.listen_button)

        self.setLayout(layout)

    def handle_command(self):
        threading.Thread(target=self.process_command).start()

    def process_command(self):
        self.append_output("Listening...")
        command = listen()
        if command:
            self.append_output(f"You said: {command}")
            response_text = f"Executing: {command}"
            self.append_output(response_text)
            speak("Working on it.")
            execute(command)
        else:
            self.append_output("Didn't catch that.")

    def append_output(self, text):
        self.output_box.append(text)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    jarvis_gui = JarvisGUI()
    jarvis_gui.show()
    sys.exit(app.exec_())

# ============ START LOOP ============
def start_jarvis():
    speak("Jarvis is now online and listening.")
    while True:
        print("\n\U0001F7E2 Awaiting wake word...")
        wake_command = listen()

        if "jarvis" in wake_command:
            speak(random.choice(["Yes sir?", "How can I help you?", "Listening.", "Ready."]))
            time.sleep(0.5)
            user_command = listen()
            if user_command:
                execute(user_command)

# ============ START ============
start_jarvis()