import textwrap
import google.generativeai as genai
import speech_recognition as sr
import win32com.client as wincom
import webbrowser
import os
import datetime
import markdown
from config import apikey
import re

GOOGLE_API_KEY = apikey
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-pro")
chatstr=""

def to_markdown(text):
    text = text.replace('•',' *' )
    return markdown.markdown(textwrap.indent(text,"> ",predicate = lambda _  :True))

def chat(prompt):
    global chatstr
    response=model.generate_content(prompt)
    chatstr += f"User :\n{prompt}\nLuna:\n{response.text}"
    print(chatstr)
    speak = to_markdown(response.text)
    say(speak)
    chatstr=""
    print("\n")

speak = wincom.Dispatch("SAPI.SpVoice")

applications = ['notepad', 'cmd', 'excel']
music_list = [['reflections', './demo/reflections.mp3']]
videos = [['demo', './demo/demo.mp4']]
sites = [['youtube', 'https://youtube.com'],['google', 'https://google.com'], ['chess.com', 'https://chess.com'], ['github', 'https://github.com']]
def say(text):
    speak.Speak(text)
def micSpeech():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=0.5
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio,language="en-in")
            #^ Applications
            app_match = re.search(r"start\s+(\w+)", query.lower())
            #^ Music
            music_match = re.search(r"play\s+(\w+)", query.lower())
            
            #^ Video
            video_match = re.search(r"play\s+(\w+)", query.lower())
            
            #^ Websites
            for site in sites:
                #^ 1. Opens sites
                if f"Open {site[0]}".lower() in query.lower():
                    say(f"opening {site[0]}")
                    webbrowser.open(site[1])
                    exit()

            #^ 2. plays music
            if music_match and music_match.group(1) in [song[0].lower() for song in music_list]:
                song_name = music_match.group(1)
                song_path = next(song[1] for song in music_list if song[0].lower() == song_name)
                say(f"playing {song_name}")
                os.system(f'start {song_path}')
                exit()
            
            #^ 3. tell current time
            elif "current time" in query.lower():
                    hour = datetime.datetime.now().strftime("%H")
                    min = datetime.datetime.now().strftime("%M")
                    say(f"Current time is {hour} hours and {min} minutes")

            #^ 4. starts applications
            elif app_match and app_match.group(1) in [app.lower() for app in applications]:
                app_name = app_match.group(1)
                say(f"opening {app_name}")
                os.system(f"start {app_name}")
                exit()

            #^ 5. plays video
            elif video_match and video_match.group(1) in [video[0].lower() for video in videos]:
                video_name = video_match.group(1)
                video_path = next(video[1] for video in videos if video[0].lower() == video_name)
                say(f"playing {video_name}")
                os.system(f'start {video_path}')
                exit()

            #^ 6. says goodbye
            elif "goodbye luna".lower() in query.lower():
                say("Have a nice day")
                exit()
            elif "how are you" in query.lower():
                say("I am good. Thank you!")
            else:
                print("chatting")
                chat(query)
            return query
        except Exception as e:
            say("Some Error Occured, Sorry from Luna.")

def manageItems():
    def update_source_code(list_name, new_list):
        with open(__file__, 'r') as file:
            lines = file.readlines()
        for i, line in enumerate(lines):
            if f"{list_name} = " in line:
                lines[i] = f"{list_name} = {repr(new_list)}\n"
                break
        with open(__file__, 'w') as file:
            file.writelines(lines)

    while True:
        print("\n=== Item Management Menu ===")
        print("1. Manage Applications")
        print("2. Manage Music")
        print("3. Manage Videos")
        print("4. Manage Websites")
        print("5. Exit")
        
        choice = input("\nEnter your choice (1-5): ")
        
        if choice == '5':
            print("Exiting item management...")
            break
            
        if choice not in ['1', '2', '3', '4']:
            print("Invalid choice! Please try again.")
            continue
            
        print("\nWhat would you like to do?")
        print("1. Add item")
        print("2. Remove item")
        print("3. View items")
        action = input("Enter your choice (1-3): ")
        
        if choice == '1':  #^ Applications
            if action == '1':
                app_name = input("Enter application name: ").lower()
                if app_name not in applications:
                    applications.append(app_name)
                    update_source_code('applications', applications)
                    print(f"Added {app_name} to applications.")
                else:
                    print("Application already exists!")
            elif action == '2':
                app_name = input("Enter application name to remove: ").lower()
                if app_name in applications:
                    applications.remove(app_name)
                    update_source_code('applications', applications)
                    print(f"Removed {app_name} from applications.")
                else:
                    print("Application not found!")
            elif action == '3':
                print("\nCurrent Applications:", applications)
                
        elif choice == '2':  #^ Music
            if action == '1':
                music_name = input("Enter music name: ").lower()
                music_path = input("Enter music file path: ")
                if not any(music_name == m[0].lower() for m in music_list):
                    music_list.append([music_name, music_path])
                    update_source_code('music_list', music_list)
                    print(f"Added {music_name} to music list.")
                else:
                    print("Music already exists!")
            elif action == '2':
                music_name = input("Enter music name to remove: ").lower()
                for item in music_list:
                    if item[0].lower() == music_name:
                        music_list.remove(item)
                        update_source_code('music_list', music_list)
                        print(f"Removed {music_name} from music list.")
                        break
                else:
                    print("Music not found!")
            elif action == '3':
                print("\nCurrent Music List:")
                for item in music_list:
                    print(f"Name: {item[0]}, Path: {item[1]}")
                
        elif choice == '3':  #^ Videos
            if action == '1':
                video_name = input("Enter video name: ").lower()
                video_path = input("Enter video file path: ")
                if not any(video_name == v[0].lower() for v in videos):
                    videos.append([video_name, video_path])
                    update_source_code('videos', videos)
                    print(f"Added {video_name} to videos list.")
                else:
                    print("Video already exists!")
            elif action == '2':
                video_name = input("Enter video name to remove: ").lower()
                for item in videos:
                    if item[0].lower() == video_name:
                        videos.remove(item)
                        update_source_code('videos', videos)
                        print(f"Removed {video_name} from videos list.")
                        break
                else:
                    print("Video not found!")
            elif action == '3':
                print("\nCurrent Videos List:")
                for item in videos:
                    print(f"Name: {item[0]}, Path: {item[1]}")
                
        elif choice == '4':  #^ Websites
            if action == '1':
                site_name = input("Enter website name: ").lower()
                site_url = input("Enter website URL: ")
                if not any(site_name == s[0].lower() for s in sites):
                    sites.append([site_name, site_url])
                    update_source_code('sites', sites)
                    print(f"Added {site_name} to sites list.")
                else:
                    print("Website already exists!")
            elif action == '2':
                site_name = input("Enter website name to remove: ").lower()
                for item in sites:
                    if item[0].lower() == site_name:
                        sites.remove(item)
                        update_source_code('sites', sites)
                        print(f"Removed {site_name} from sites list.")
                        break
                else:
                    print("Website not found!")
            elif action == '3':
                print("\nCurrent Websites List:")
                for item in sites:
                    print(f"Name: {item[0]}, URL: {item[1]}")
def Luna():
    say("Hello! Luna this side.")
    print("\n\nInstructions\n")
    print("For opening applications use keyword 'start' followed by application name e.g. : 'start illustrator'.")
    print("For opening websites use keyword 'open' followed by name of website e.g. : 'Open github'.")
    print("For playing music use 'play <name of song>'.")
    print("For playing video use 'play <name of video>'.")
    print("To know the current time use keyword 'current time'.")
    print("To close luna say 'goodbye luna'.\n")
    while True:
        print("listening...")
        text=micSpeech()
if __name__ == '__main__':
    # Luna()
    manageItems()
