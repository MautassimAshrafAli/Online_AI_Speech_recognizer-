import ctypes
import webbrowser
import speech_recognition as sr
from datetime import datetime
import wolframalpha
import os
import subprocess
import getpass
import pyautogui
import random
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
ctypes.windll.kernel32.SetConsoleTitleA("Near Assistant")

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36"
# US english
LANGUAGE = "en-US,en;q=0.5"

name = "Near"
on_m = 0
on_v = 0

def shutdown():
    os.system("shutdown /s /t 1")

def restart():
    os.system("shutdown /r /t 1")

def sleep():

    fs = open("sleep_s.txt", "w")
    fs.write("ss")
    fs.close()

    #os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')

def nearcommand():





    r = sr.Recognizer()
    n1 = 0
    n2 = 0

    with sr.Microphone() as source:

        print(name + " is Listening...")

        while True:



            r.adjust_for_ambient_noise(source,duration=1)

            audio = r.listen(source, phrase_time_limit=5)
            print("Stop.")
            query = ""

            try:
                query = r.recognize_google(audio, language='en-in')
                print('User: ' + query + '\n')

            except sr.UnknownValueError:

                query = nearcommand()

            return query.lower()

ny = 0
nyt = 0

def youtuberc():

    r2 = sr.Recognizer()

    url = 'https://www.youtube.com/results?search_query='

    with sr.Microphone() as s:
        print("search for")

        global ny
        global nyt

        ny +=1
        nyt += 1

        speaker.Speak("search for")

        r2.adjust_for_ambient_noise(s, duration=1)
        query_y = r2.listen(s, phrase_time_limit=5)
        print("Stop.")

        try:
            s_for = r2.recognize_google(query_y)
            print(s_for)
            webbrowser.get().open_new(url + s_for)

            speaker.Speak("search for " + s_for)


        except:

            speaker.Speak("You are too late try again")

ng = 0
ngt = 0

def googlec():

    r3 = sr.Recognizer()

    url = 'https://www.google.com/search?q='

    with sr.Microphone() as s:
        print("search for")

        global ng
        global ngt

        ng += 1
        ngt += 1

        speaker.Speak("search for")

        r3.adjust_for_ambient_noise(s, duration=1)
        query_ = r3.listen(s, phrase_time_limit=5)
        print("Stop.")

        try:
            s_for = r3.recognize_google(query_)
            print(s_for)
            webbrowser.get().open_new(url + s_for)

            speaker.Speak("search for " + s_for)

        except:

            speaker.Speak("You are too late try again")

ns = 0
nst = 0

def soundcloudc():

    r4 = sr.Recognizer()

    url = 'https://soundcloud.com/search?q='

    with sr.Microphone() as s:
        print("search for")

        global ns
        global nst

        ns += 1
        nst += 1

        speaker.Speak("search for")

        r4.adjust_for_ambient_noise(s, duration=1)
        query_ = r4.listen(s, phrase_time_limit=5)
        print("Stop.")

        try:
            s_for = r4.recognize_google(query_)
            print(s_for)
            webbrowser.get().open_new(url + s_for)

            speaker.Speak("search for " + s_for)

        except:

            speaker.Speak("You are too late try again")


nv = 0
nvt = 0

def stackoverflowc():

    r5 = sr.Recognizer()

    url = 'https://stackoverflow.com/search?q='

    with sr.Microphone() as s:
        print("search for")

        global nv
        global nvt

        '''sapi_voice("search for")'''

        nv += 1
        nvt += 1

        speaker.Speak("search for")


        r5.adjust_for_ambient_noise(s, duration=1)
        query_ = r5.listen(s, phrase_time_limit=5)
        print("Stop.")

        try:
            s_for = r5.recognize_google(query_)
            print(s_for)
            webbrowser.get().open_new(url + s_for)

            speaker.Speak("search for " + s_for)

        except:

            speaker.Speak("You are too late try again")

def main():


    c = 0
    #-------------------
    cmdn1 = 0
    #-----------------
    on_cmd_ = 0
    #-----------------
    currentH = datetime.now().hour
    currentS = datetime.now().second
    currentyear = datetime.now().year
    #-----------------------------------
    g_n = 0
    #-------
    y_n = 0
    #------
    s_n = 0
    #------
    v_n = 0
    #------
    m_n = 0
    #-------
    w_n = 0
    #------
    f_n = 0
    #----
    w_s = 0
    #------

    t = str(currentyear) + str(currentH) + str(currentS)


    while True:

        query = nearcommand()

        # near command

#=========================================================== this near =================================================

        if 'hello' in query:

            try:

                print('hello ' + getpass.getuser())

                speaker.Speak("hello " + getpass.getuser())


            except:
                pass

        elif 'how are you' in query:

            try:
                print('how are you')
                stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy', 'good!']

                r = random.choice(stMsgs)

                speaker.Speak(r)

            except:
                pass

        elif 'time' in query:

            try:

                print('time')

                now = datetime.now()
                time = now.strftime("%I:%M:%S")

                speaker.Speak(time)

            except:
                pass

        elif 'date' in query:
            print('date')

            try:

                now = datetime.now()
                date = now.strftime("%m/%d/%Y")

                speaker.Speak(date)

            except:
                pass

        elif 'year' in query:
            print('year')

            try:

                now = datetime.now()
                year = now.strftime("%Y")

                speaker.Speak(year)

            except:
                pass

        elif 'month' in query:
            print('month')

            try:

                now = datetime.now()
                month = now.strftime("%m")

                speaker.Speak(month)

            except:

                pass

        elif 'day' in query:
            print('day')

            try:

                now = datetime.now()
                day = now.strftime("%d")

                speaker.Speak(day)

            except:

                pass



        elif 'your' in query:

            try:
                print(name)

                speaker.Speak("my name is " + name)

            except:
                pass

        elif 'my' in query:

            try:
                print(getpass.getuser())

                speaker.Speak("your name is " + getpass.getuser())

            except:
                pass

        elif 'where are you from' in query:

            try:

                print('where are you from')

                speaker.Speak("i am from egypt")

            except:
                pass

        elif 'mute' in query:

            try:

                speaker.Speak("mute your computer done")

                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(
                    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))

                # Control volume
                # volume.SetMasterVolumeLevel(-0.0, None) #max
                volume.SetMasterVolumeLevel(-28.0, None)  # 0%



            except:
                pass

        elif 'max' in query:

            try:

                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(
                    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))

                # Control volume
                volume.SetMasterVolumeLevel(-0.0, None)  # max
                # volume.SetMasterVolumeLevel(-28.0, None)  # 0%

                speaker.Speak("max volume your computer done")

            except:
                pass


#======================================================end near=========================================================





#======================================================== power ========================================================

        elif 'shutdown' in query:

            shutdown()

        elif 'restart' in query:

            restart()

        elif 'sleep' in query:

            sleep()

# ========================================================end power=====================================================




# ===================================================== cmd & login's ==================================================

        elif 'command prompt' in query:

            try:

                on_cmd_ += 1

                if on_cmd_ == 1:

                  speaker.Speak("The command prompt opening")


                  print("command prompt")
                  os.system("cd\ & start cmd.exe")

                cmdn1 += 1

                speaker.Speak("you are now on command prompt form")

            except:
                pass

# =====================================================end cmd & login's================================================





#==================================================== tack screenshot ==================================================

        elif 'screenshot' in query:

            c += 1

            try:

                print("screenshot")

                speaker.Speak("take an screenshot")

                e = os.path.exists("C:/Users/" + getpass.getuser() + "/Desktop/screenshot" + str(c) + ".png")
                print(e)

                myScreenshot = pyautogui.screenshot()

                if e == True:

                    path = r"C:/Users/" + getpass.getuser() + "/Desktop/screenshot random number" + str(t) + ".png"
                    myScreenshot.save(path)

                else:

                    path = r"C:/Users/" + getpass.getuser() + "/Desktop/screenshot" + str(c) + ".png"
                    myScreenshot.save(path)

                speaker.Speak("done")


            except:
                pass

# ====================================================end tack screenshot===============================================



# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------


        # on web

        elif 'youtube' in query:

            try:

                y_m = ['opening youtube', 'open youtube', 'youtube is open',
                       'done youtube is open now', 'youtube is open now']

                y_n += 1

                r = random.choice(y_m)

                speaker.Speak(r)

                youtuberc()

            except:
                pass

        elif 'google' in query:

            try:

                g_m = ['opening google', 'open google', 'google is open',
                       'done google is open now', 'google is open now']

                g_n += 1

                r = random.choice(g_m)

                speaker.Speak(r)

                googlec()

            except:
                pass

        elif 'soundcloud' in query:

            try:

                s_m = ['opening sound cloud', 'open sound cloud', 'sound cloud is open',
                       'done sound cloud is open now', 'sound cloud is open now']

                s_n += 1

                r = random.choice(s_m)

                speaker.Speak(r)

                soundcloudc()

            except:
                pass

        elif 'stack overflow' in query:

            try:

                s_m = ['opening stackoverflow', 'open stackoverflow', 'stackoverflow is open',
                       'done stackoverflow is open now', 'stackoverflow is open now']


                v_n += 1


                r = random.choice(s_m)

                speaker.Speak(r)

                stackoverflowc()

            except:
              pass

        elif 'messenger' in query:

            try:

                m_m = ['opening messenger', 'open messenger', 'messenger is open',
                       'done messenger is open now', 'messenger is open now']

                m_n += 1

                r = random.choice(m_m)

                speaker.Speak(r)

                url = 'https://www.messenger.com/login.php'

                webbrowser.get().open_new(url)

            except:
                pass

        elif 'whatsapp' in query:

            try:

                w_m = ['opening whatsapp', 'open whatsapp', 'whatsapp is open',
                       'done whatsapp is open now', 'whatsapp is open now']

                w_n += 1

                r = random.choice(w_m)

                speaker.Speak(r)

                url = 'https://web.whatsapp.com/'

                webbrowser.get().open_new(url)

            except:
                pass

        elif 'facebook' in query:

            try:

                f_m = ['opening facebook', 'open facebook', 'facebook is open',
                       'done facebook is open now', 'facebook is open now']

                f_n += 1

                r = random.choice(f_m)

                speaker.Speak(r)

                url = 'https://www.facebook.com/'

                webbrowser.get().open_new(url)

            except:
                pass

        # text note

        elif 'note' in query:

            try:

                date = datetime.now()
                file_name = str(date).replace(":", "-") + "-note.txt"
                subprocess.Popen(["notepad.exe", file_name])


                speaker.Speak("What would you like me to write down?")

                write_down = nearcommand()

                with open(file_name, "w") as f:
                    f.write(write_down)

                while 'exit' in query:
                    query = nearcommand()

                speaker.Speak("I've made a note of that. and you say in the note. " + write_down)


            except:

                pass

        else:

            speaker.Speak("searching for " + query + " in database")

            print("not in query")

            try:



                w_s += 1


                wolfram_id = wolframalpha.Client('HYQ2AG-R6A8J9QJ6K')


                the_r = wolfram_id.query(query)
                output = next(the_r.results).text

                speaker.Speak(output)

                print(output)

            except:
              pass

if __name__ == '__main__':

            try:

                speaker.Speak("Welcome " + getpass.getuser())

                main()
            except:
                pass