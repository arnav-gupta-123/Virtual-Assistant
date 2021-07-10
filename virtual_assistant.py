import wolframalpha
import wikipedia
import PySimpleGUI as sg
import pyttsx3
from win32com.client import Dispatch

engine = pyttsx3.init()
client = wolframalpha.Client('HGYHP3-QXVAAJ3ATG')
speak = Dispatch("SAPI.SpVoice")

sg.theme('DarkBlue')
layout = [[sg.Text("What are you looking for? "), sg.InputText()],[sg.Button('Search'), sg.Button('Cancel')]]
window = sg.Window('Digital Assistant', layout)
speak.Speak("Hello, this is D-A,your digital assistant. What are you looking for? ")

while True:

    event, values = window.read()

    if event in (None, 'Cancel'):
        speak.Speak("See you later")
        break

    try:
        wiki_res = wikipedia.summary(values[0], sentences=2)
        wolfram_res = next(client.query(values[0]).results).text
        engine.say(wolfram_res)
        sg.PopupNonBlocking("Wolfram Result: " + wolfram_res,"Wikipedia Result: " + wiki_res)

    except wikipedia.exceptions.DisambiguationError:
        wolfram_res = next(client.query(values[0]).results).text
        engine.say(wolfram_res)
        sg.PopupNonBlocking(wolfram_res)

    except wikipedia.exceptions.PageError:
        wolfram_res = next(client.query(values[0]).results).text
        engine.say(wolfram_res)
        sg.PopupNonBlocking(wolfram_res)

    except:
        wiki_res = wikipedia.summary(values[0], sentences=2)
        engine.say(wiki_res)
        sg.PopupNonBlocking(wiki_res)

    engine.runAndWait()

window.close()
