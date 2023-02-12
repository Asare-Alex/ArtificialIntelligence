import math
import speech_recognition as sr
from selenium import webdriver
import webbrowser
import random
import os
import pyttsx3
from defaults import *
from datetime import datetime
import socket
from math import *

"""
    from distutils.core import setup
    setup(name='PC Artificial Intelligence',
    version='1.0.0.2.5.0',
    py_modules=['*'],
)
"""

"""
if __name__ == "__main__":
    from . import build_exe
    build_exe.main()
"""
# from win32com import Dispatch sre_constants

# speak = win32com.client.Dispatch("SAPI.SpVoice")

r = sr.Recognizer()
engine = pyttsx3.init()

engine.say(random.choice(greeting))
engine.runAndWait()

engine.say(random.choice(ask))
engine.runAndWait()

with sr.Microphone() as mic:
    speech = r.listen(mic)
    text = r.recognize_google(speech)
    print(text)
    task = text.lower()
    while True:
        for command in end:
            for on in online:
                for about in about_command:
                    if task == command:
                        engine.say(random.choice(goodbye))
                        engine.runAndWait()

                    elif task == on:
                        engine.say("Hey buddy, I am listening")
                        engine.runAndWait()

                    elif task == about:
                        engine.say(random.choice(excitement))
                        engine.say(random.choice(about_me))
                        engine.runAndWait()

                    else:
                        parts = task.split(" ")
                        for part in parts:
                            if part == "go":
                                engine.say(f"I am opening {part[-1]}")
                                engine.runAndWait()
                                os.startfile(f"D:\{part[-1]}")

                            elif part == "open" or "start":
                                engine.say("Ok, hang on")
                                engine.runAndWait()
                                os.system(f"start {part[-1]}")
                            elif part == "close" or "exit" or "end":
                                engine.say(f"I am closing {part[-1]}")
                                engine.runAndWait()
                                os.system(f"exit {part[-1]}")

                            elif part == "search" or "browse":
                                engine.say(f"I am about to search for {part[-1]} in your favorite browser")
                                engine.runAndWait()
                                webbrowser.open(part[-1])
                            elif part == "direction" or "map":
                                engine.say(f"Hang on, I'm getting you the direction to {part[-1]}")
                                engine.runAndWait()
                                driver = webdriver.Chrome()
                                driver.get("http://google.com")
                                search = driver.find_element_by_id("search")
                                search.send_keys(part[-1])
                                search.submit()

                            elif part == "login" or "log in" or "signin" or "sign in":
                                engine.say(f"Wanna chat with friends? Let me connect you to {part[-1]}")
                                engine.runAndWait()
                                webbrowser.open(part[-1])

                            elif part == "play" or "listen" and "hiplife" or "music":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MUSICZ\HIPLIFE_2020'))
                                songs = library[-1]
                                song = random.choice(songs)
                                os.startfile(f"D:\MUSICZ\HIPLIFE_2020\{song}")

                            elif part == "play" or "listen" and "gosple":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MUSICZ\GOSPEL'))
                                songs = library[-1]
                                song = random.choice(songs)
                                os.startfile(f"D:\MUSICZ\GOSPEL\{song}")

                            elif part == "play" or "listen" and "hilife":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MUSICZ\HILIFE'))
                                songs = library[-1]
                                song = random.choice(songs)
                                os.startfile(f"D:\MUSICZ\HILIFE\{song}")

                            elif part == "play" or "listen" and "foreign":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MUSICZ\FOREIGN'))
                                songs = library[-1]
                                song = random.choice(songs)
                                os.startfile(f"D:\MUSICZ\FOREIGN\{song}")

                            elif part == "play" or "listen" and "reggae":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MUSICZ\REGGAE_INVASION'))
                                songs = library[-1]
                                song = random.choice(songs)
                                os.startfile(f"D:\MUSICZ\REGGAE_INVASION\{song}")

                            elif part == "play" or "watch" and "funny" and "videos":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\Funny_Videos'))
                                videos = library[-1]
                                video = random.choice(videos)
                                os.startfile(f"D:\Funny_Videos\{video}")

                            elif part == "movie" and "watch":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MOVIES'))
                                videos = library[-1]
                                video = random.choice(videos)
                                os.startfile(f"D:\MOVIES\{video}")

                            elif part == "vampire" and "diaries" and "watch" or "play":
                                engine.say(f"I'm getting you {part[-1]}")
                                engine.runAndWait()
                                library = next(os.walk('D:\MOVIES\Money_Heist_Season_5'))
                                videos = library[-1]
                                for video in videos:
                                    os.startfile(f"D:\MOVIES\Money_Heist_season_5\{video}")

                            elif part == "money" and "heist" and "play" or "watch":
                                library = next(os.walk('D:\MOVIES\VAMPIRE_DIARIES'))
                                videos = library[-1]
                                for video in videos:
                                    os.startfile(f"D:\MOVIES\VAMPIRE_DIARIES\{video}")

                            elif part == "what" or "tell" or "today" or "now":
                                engine.say("Today is ")
                                engine.runAndWait()
                                #today = datetime.now()
                                # noinspection PyTypeChecker
                                engine.say(datetime.now())
                                engine.runAndWait()

                            elif part == "ip" and "address" and "what":
                                host_name = socket.gethostname()
                                engine.say(socket.gethostbyname(host_name))
                                engine.runAndWait()

                            elif part == "shutdown":
                                engine.say("I'm shutting down your PC")
                                engine.runAndWait()
                                os.system("shutdown /s")

                            elif part == "restart":
                                engine.say("I'm restarting your PC")
                                engine.runAndWait()
                                os.system("shutdown /r")

                            elif part == "add" or "addition" or "sum":
                                addition = part[-1] + part[-3]
                                engine.say(addition)
                                engine.runAndWait()

                            elif part == "multiply" or "multiplication" or "product":
                                product = part[-1] * part[-3]
                                engine.say(product)
                                engine.runAndWait()

                            elif part == "divide" or "division":
                                division = part[-3] / part[-1]
                                engine.say(division)
                                engine.runAndWait()

                            elif part == "subtract" or "subtraction" "or" "from":
                                subtract = part[-1] - part[-3]
                                engine.say(subtract)
                                engine.runAndWait()
                                
                            elif part == "round":
                                rnd = round(part[-1])
                                engine.say(rnd)
                                engine.runAndWait()

                            elif part == "exponent" or "power":
                                pwr = part[0] ** part[-1]
                                engine.say(pwr)
                                engine.runAndWait()

                            elif part == "tan" or "tangent":
                                tangent = math.tan(part[-1])
                                engine.say(tangent)
                                engine.runAndWait()

                            elif part == "sine":
                                sin = math.sin(part[-1])
                                engine.say(sin)
                                engine.runAndWait()

                            elif part == "cos" or "cosine":
                                cos = math.sin(part[-1])
                                engine.say(cos)
                                engine.runAndWait()

                            elif part == "inverse" and"tan" or "tangent":
                                atangent = math.atan(part[-1])
                                engine.say(atangent)
                                engine.runAndWait()

                            elif part == "sine" and "inverse":
                                asine = math.asin(part[-1])
                                engine.say(asine)
                                engine.runAndWait()

                            elif part == "inverse" and "cos" or "cosine":
                                acosine = math.acos(part[-1])
                                engine.say(acosine)
                                engine.runAndWait()

                            elif part == "log" or "logarithm":
                                logarithm = math.log(part[-5], part[-1])

                            else:
                                engine.say("What you said is not in my command list")
                                engine.runAndWait()
