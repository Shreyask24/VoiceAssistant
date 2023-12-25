import pyttsx3                      #pip install pyttsx3
import speech_recognition as sr     #pip install speechRecognition
import datetime
import wikipedia                    #pip install wikipedia
import webbrowser
import pyjokes
import pywhatkit
from datetime import date
from docx2pdf import convert
import requests
from bs4 import BeautifulSoup
import os
import operator
from ecapture import ecapture as ec
import subprocess
import ctypes
import time
import winshell

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

import logging



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)



jokes_command = {1: 'boring',
                2: 'tell me more jokes',
                3: 'joke',
                4:'more jokes'}

search={
    # 29 Search Approx
    1:"https://www.bing.com/search?q=what+can+the+new+bing+chat+do%3F&aqs=edge.0.69i64i450l8.77760304j0j1&FORM=ANSPA1&PC=DCTS",
    2:"https://www.bing.com/search?q=bing+homepage+quiz&aqs=edge.1.69i64i450l8.82764390j0j1&FORM=ANSPA1&PC=DCTS"
}


def automatedTask():
    speak("Opening Microsoft Edge!")
    bing = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    os.startfile(bing)

    # Set the path to your Microsoft Edge WebDriver executable
    edge_driver_path = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe'

    # Set the path to the Microsoft Edge binary
    edge_binary_path = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'  # Update with your Edge installation path

    # Create an EdgeOptions instance and set the binary location
    edge_options = Options()
    edge_options.binary_location = edge_binary_path

    # Create a new Edge WebDriver instance with the specified options
    driver = webdriver.Edge(exec(edge_driver_path))

    # Navigate to the search engine (e.g., Google)
    driver.get('https://www.google.com')

    # Find the search input element and enter your search query
    search_box = driver.find_element_by_name('q')  # Adjust this selector based on the search engine's HTML
    search_query = 'Your search query here'
    search_box.send_keys(search_query)
    search_box.send_keys(Keys.RETURN)

    # Wait for some time to see the search results (you can adjust this as needed)
    driver.implicitly_wait(10)

    # Close the browser when done
    driver.quit()

    logging.basicConfig(level=logging.DEBUG)



def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")


    # enter city name
    city = "pune"
    url = "https://www.google.com/search?q=" + "weather" + city

    # requests instance
    html = requests.get(url).content

    # getting raw data
    soup = BeautifulSoup(html, 'html.parser')

    # get the temperature
    temp = soup.find('div', attrs={'class': 'BNeawe iBp4i AP7Wnd'}).text

    speak(f"Hi, I am Jarvis your personal desktop assistant, and  Toady's Temperature in {city} is: "+ temp)
    speak("Please tell me How can I help you?")


def takeCommand():
    #It takes microphone input from the user and returns string output
    query = "Hello"
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)        #Error

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


def word_to_pdf():
    try:
        # set the path of the Word document
        word_file_path = r'C:\Users\Shreyas\PycharmProjects\Jarvis\testing.docx'

        # convert the Word document to PDF
        convert(word_file_path)

        # the converted PDF will be saved in the same folder as the original Word document with the same name
        pdf_file_path = os.path.splitext(word_file_path)[0] + '.pdf'

        speak("Converted word document to pdf successFully")

    except:
        speak("There was an error for converting the file please check the given path of word document")




def get_operator_fn(op):
    return {
        '+' : operator.add,
        '-' : operator.sub,
        'x' : operator.mul,
        '/': operator.__truediv__,
        'divided' :operator.__truediv__,
        'Mod' : operator.mod,
        'mod' : operator.mod,
        '^' : operator.xor,
        }[op]

def eval_binary_expr(op1, oper, op2):
    op1,op2 = int(op1), int(op2)
    return get_operator_fn(oper)(op1, op2)








if __name__ == "__main__":
    # wishMe()

    while (1):
        query = takeCommand().lower()

        if query == 0:
            continue

        if "stop" in str(query) or "exit" in str(query) or "bye" in str(query):
            speak("Ok bye and take care, Thanks for giving me your time")
            break



        if 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
            print(strTime)


        # Logic Commands
        elif 'who' in query or 'what' in query:
            speak("Searching Wikipedia..")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)


        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you?")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")


        elif "tell me about yourself" in query:
            speak("I am Jarvis. Your Personal Assistant")

        elif 'joke' in query:
            speak("Todays joke is ")
            joke = speak(pyjokes.get_joke())
            print(joke)


        elif 'open youtube' in query:
            speak("Opening Youtube..")
            webbrowser.open('youtube.com')


        elif 'search' in query:
            strsearch = query.replace("search", "")
            speak("Searching" + strsearch)
            pywhatkit.search(strsearch)

        elif 'play' in query:
            song = query.replace('play', '')
            speak('Playing' + song)
            pywhatkit.playonyt(song)


        elif 'open google' in query:
            speak("Opening Google...")
            webbrowser.open('google.com')

        elif 'open stack overflow' in query:
            speak("Opening Stackoverflow...")
            webbrowser.open('stackoverflow.com')


        elif 'date' in query:
            strDate = date.today()
            speak(strDate)
            print(strDate)

        elif 'open visual studio' in query:
            speak("Opening Visual studio code")
            codePath = "C:\\Users\\Shreyas\\AppData\\Local\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)


        elif 'open excel' in query:
            speak("Opening Excel sheet")
            excel = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(excel)


        elif 'convert word to pdf' in query:
            speak("Converting Word Document To PDF")
            word_to_pdf()


        elif jokes_command.get(1) in query:
            jokes = pyjokes.get_joke()
            print(jokes)
            speak("Ok!, So the today's joke is" + jokes)


        elif 'shut down' in query or 'shutdown' in query:
            speak("Do you wish to shutdown your computer ? (yes or no): ")
            while True:
                query = takeCommand()

                if "no" in query:
                    speak("Thank u I will not shut down the computer")
                    break

                if "yes" in query:
                    speak("Shutting down the pc within 30 seconds")
                    os.system("shutdown /s /t 30")
                    break
                speak("Say that again!")


        elif "calculate" in query:
            speak("Provide expressions to perform calculations")

            while True:
                query = takeCommand()

                print(eval_binary_expr(*(query.split())))
                speak(f"Your answer is {eval_binary_expr(*(query.split()))}")

                if "stop calculation" in query or "stop calculating" in query:
                    break

            speak("say that again")

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")


        elif 'tickets' in query:
            speak("Opening Book My Show!")
            webbrowser.open("bookmyshow.com")
            speak("Which Movie would you like to watch?")

            while True:

                movieName = takeCommand().lower()

                if 'subedar' in movieName:
                    speak("Ok Selecting Movie")
                    webbrowser.open("https://in.bookmyshow.com/pune/movies/subhedar/ET00354867")

                elif 'theatres' in movieName:
                    speak("These Are The List Of Theaters. Please Select the theatre and select seats")
                    webbrowser.open("https://in.bookmyshow.com/buytickets/subhedar-pune/movie-pune-ET00354867-MT/20230830")



        elif 'microsoft' in query:
            automatedTask()




        elif 'thank you' in query:
            speak("You’re welcome! Good to know that you are happy!")





