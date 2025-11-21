This is a Windows desktop voice assistant (GUI + CLI) named Jarvis that listens for voice commands, performs local system control (windows/apps/volume/brightness/network), interacts with web services (Wikipedia, YouTube), runs small automations (file ops, screenshots, typing), can talk to a local LLM endpoint, and speaks back via TTS. It exposes both a Qt GUI and a continuous listen loop.
Models & external services used (current)

Speech-to-text: speech_recognition Python package with Recognizer().recognize_google() — this uses Google Web Speech (online) unless you change backend.

Text-to-speech (TTS): edge_tts.Communicate(..., voice="en-IN-NeerjaNeural") — Microsoft/Edge neural voices (online).

Local LLM: Your ask_ai() posts to http://localhost:11434/api/generate using {"model":"mistral", ...} — implies a locally hosted Mistral-compatible generation API that streams JSON lines.

Wikipedia: wikipedia Python package — online.

System audio control: pycaw (COM interface) to set master volume.

Screen brightness: screen_brightness_control (sbc).

Tuya / IoT: you imported TuyaOpenAPI (not actually used in pasted code but present).

Other OS automation: pyautogui, win32gui, win32api, pygetwindow etc.

Main components & responsibilities

Speech I/O

listen() — captures microphone input and returns lowercase text (uses Google STT).

speak() / speak_async() — synthesizes TTS via edge_tts and plays mp3 via playsound.

Command router

execute(command) — core router that inspects the command string and invokes appropriate handler (file management, system control, app/window control, Wikipedia, fallback to AI).

File management

handle_file_management(command) — create/delete/move/copy/rename/search operations.

System control helpers

volume, brightness, Wi-Fi toggles, bluetooth automation (via screenshots / GUI clicks), process kill, lock/shutdown, clipboard reading, screenshots, typing/keypress simulation.

Local AI fallback

ask_ai(message) — streams from local LLM endpoint and returns combined output.

GUI

JarvisGUI (PyQt5) with Listen button and an output text box. It calls listen() and execute() on a thread.

Main loop

start_jarvis() — terminal/listen loop that waits for wakeword "jarvis" then listens for a command and executes.

*It also Have some bugs which i have noticed but not solved because of too much gap

Circular import / redefinition: You do from jarvis import speak, listen, execute but this script is jarvis — that will cause import issues / circular imports if this file is jarvis.py.

Missing imports / wrong references:

threading is used in JarvisGUI.handle_command() but not imported.

ctypes is used in lock_pc() but not imported (you imported cast and POINTER earlier, but not ctypes).

GUI vs. blocking loop: start_jarvis() runs an infinite loop that blocks; you also start a PyQt GUI in if __name__ == '__main__'. Both can't run simultaneously in the same process/thread without careful threading. Right now the code runs start_jarvis() at the bottom unconditionally — that will start an endless loop and never reach GUI code, or vice versa (order matters).

speak() sync/async mixing: speak() uses asyncio.run(speak_async(text)). That works if no running event loop, but if you call from an already running event loop (e.g., within an async GUI thread), it can raise RuntimeError.

TTS file cleanup: speak_async() writes an mp3 file and deletes it after playsound — playsound blocks but on some systems file may still be locked; consider streaming or in-memory playback.

Edge cases in handle_file_management parsing: parsing using .split("to") is brittle (file/folder names containing ' to ').

Bluetooth automation via image matching: pyautogui.locateCenterOnScreen(..., confidence=0.8) relies on exact screenshots; fragile across DPI / Windows theme.

ask_ai streaming parsing: response.iter_lines() + json.loads(line.decode(...)) expects server to stream pure JSON per line; resilient code should handle partial lines and errors.

Permissions: operations like shutdown, taskkill, brightness control, muting may require elevated privileges.

Network dependence: many functions rely on internet (Google STT, edge_tts, Wikipedia, local LLM might be local but TTS/STT are online). You already have is_connected() check only used in speak() — but listen() will still call Google and error if offline.

Hard-coded paths: os.chdir(r"C:\Users\Lappy\Jarvis") — makes it non-portable.

No error logging: lots of except: without logging the exception object.
