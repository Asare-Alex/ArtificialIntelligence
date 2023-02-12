import os
import webbrowser
import win32com.client
import pathlib
import random
import selenium
speak = win32com.client.Dispatch("SAPI.SpVoice")


def openfile():
    speak.speak(f"I am opening {part[-1]}")
    os.startfile(f"D:\{part[-1]}")


def start_app():
    speak.speak("Ok, hang on")
    os.system(f"start {part[-1]}")


def exit_app():
    speak.speak(f"I am closing {part[-1]}")
    os.system(f"exit {part[-1]}")


def search():
    speak.speak(f"I am about to search for {part[-1]} in your favorite browser")
    webbrowser.open(part[-1])


def map():
    speak.speak(f"Hang on, I'm getting you the direction to {part[-1]}")
    driver = webdriver.Chrome()
    driver.get("http://google.com")
    search = driver.find_element_by_id("search")
    search.send_keys(part[-1])
    search.submit()


def social_media():
    speak.speak(f"Wanna chat with friends? Let me connect you to {part[-1]}")
    webbrowser.open(part[-1])


def play_music_hiplife():
    library = next(os.walk('D:\MUSICZ\HIPLIFE_2020'))
    songs = library[-1]
    song = random.choice(songs)
    os.startfile(f"D:\MUSICZ\HIPLIFE_2020\{song}")


def play_music_gospel():
    library = next(os.walk('D:\MUSICZ\GOSPEL'))
    songs = library[-1]
    song = random.choice(songs)
    os.startfile(f"D:\MUSICZ\GOSPEL\{song}")


def play_music_foreign():
    library = next(os.walk('D:\MUSICZ\FOREIGN'))
    songs = library[-1]
    song = random.choice(songs)
    os.startfile(f"D:\MUSICZ\FOREIGN\{song}")


def play_music_reggae():
    library = next(os.walk('D:\MUSICZ\REGGAE_INVASION'))
    songs = library[-1]
    song = random.choice(songs)
    os.startfile(f"D:\MUSICZ\REGGAE_INVASION\{song}")


def play_music_hilife():
    library = next(os.walk('D:\MUSICZ\HILIFE'))
    songs = library[-1]
    song = random.choice(songs)
    os.startfile(f"D:\MUSICZ\HILIFE\{song}")


def local_funny_video():
    library = next(os.walk('D:\Funny_Videos'))
    videos = library[-1]
    video = random.choice(videos)
    os.startfile(f"D:\Funny_Videos\{video}")


def funny_video():
    library = next(os.walk('D:\MOVIES'))
    videos = library[-1]
    video = random.choice(videos)
    os.startfile(f"D:\MOVIES\{video}")


def money_heist():
    library = next(os.walk('D:\MOVIES\Money_Heist_Season_5'))
    videos = library[-1]
    for video in videos:
        os.startfile(f"D:\MOVIES\Money_Heist_season_5\{video}")


def vampire_diaries():
    library = next(os.walk('D:\MOVIES\VAMPIRE_DIARIES'))
    videos = library[-1]
    for video in videos:
        os.startfile(f"D:\MOVIES\VAMPIRE_DIARIES\{video}")