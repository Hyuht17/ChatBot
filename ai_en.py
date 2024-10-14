import os
import playsound
import speech_recognition as sr
import time
import sys
import ctypes
import wikipedia
import datetime
import json
import re
import webbrowser
import smtplib
import requests
import urllib
import urllib.request as urllib2
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from time import strftime
from gtts import gTTS
from youtube_search import YoutubeSearch

wikipedia.set_lang('en')
language = 'en'
path = ChromeDriverManager().install()
text = ""

def speak(text):
    print("Bot: {}".format(text))
    tts = gTTS(text=text, lang=language, slow=False)
    tts.save("sound.mp3")
    playsound.playsound("sound.mp3", False)
    os.remove("sound.mp3")

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("You: ", end='')
        audio = r.listen(source, phrase_time_limit=5)
        try:
            text = r.recognize_google(audio, language="en-US")
            print(text)
            return text
        except:
            print("...")
            return 0

def stop():
    speak("See you later!")

def get_text():
    for i in range(3):
        text = get_audio()
        if text:
            return text.lower()
        elif i < 2:
            speak("Bot couldn't hear you clearly. Could you please say that again?")
    time.sleep(10)
    stop()
    return 0

def hello(name):
    day_time = int(strftime('%H'))
    if day_time < 12:
        speak("Good morning, {}. Have a great day!".format(name))
    elif 12 <= day_time < 18:
        speak("Good afternoon, {}. Do you have any plans for the afternoon?".format(name))
    else:
        speak("Good evening, {}. Have you had dinner yet?".format(name))

def get_time(text):
    now = datetime.datetime.now()
    if "time" in text:
        speak('It is now %d:%d' % (now.hour, now.minute))
    elif "date" in text:
        speak("Today is the %dth of %d, %d" %
              (now.day, now.month, now.year))
    else:
        speak("I didn't understand your request. Could you please say that again?")

def open_application(text):
    if "google" in text:
        speak("Opening Google Chrome")
        os.startfile(
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')
    elif "word" in text:
        speak("Opening Microsoft Word")
        os.startfile(
            'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')
    elif "excel" in text:
        speak("Opening Microsoft Excel")
        os.startfile(
            'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE')
    else:
        speak("The application is not installed. Please try again!")

def open_website(text):
    reg_ex = re.search('open (.+)', text)
    if reg_ex:
        domain = reg_ex.group(1)
        url = 'https://www.' + domain
        webbrowser.open(url)
        speak("The website you requested has been opened.")
        return True
    else:
        return False

def open_google_and_search(text):
    search_for = text.split("search", 1)[1]
    speak('Okay!')
    driver = webdriver.Chrome(path)
    driver.get("https://www.google.com")
    que = driver.find_element_by_xpath("//input[@name='q']")
    que.send_keys(str(search_for))
    que.send_keys(Keys.RETURN)

def send_email(text):
    speak('Who do you want to send the email to?')
    recipient = get_text()
    if 'john' in recipient:
        speak('What is the content of the email?')
        content = get_text()
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        mail.login('your_email@gmail.com', 'your_password')
        mail.sendmail('your_email@gmail.com',
                      'recipient_email@gmail.com', content.encode('utf-8'))
        mail.close()
        speak('Your email has been sent. Please check your email!')
    else:
        speak("Bot couldn't understand who you want to send the email to. Please say that again.")

def current_weather():
    speak("Where do you want to check the weather?")
    ow_url = "http://api.openweathermap.org/data/2.5/weather?"
    city = get_text()
    if not city:
        pass
    api_key = "your_api_key"
    call_url = ow_url + "appid=" + api_key + "&q=" + city + "&units=metric"
    response = requests.get(call_url)
    data = response.json()
    if data["cod"] != "404":
        city_res = data["main"]
        current_temperature = city_res["temp"]
        current_pressure = city_res["pressure"]
        current_humidity = city_res["humidity"]
        suntime = data["sys"]
        sunrise = datetime.datetime.fromtimestamp(suntime["sunrise"])
        sunset = datetime.datetime.fromtimestamp(suntime["sunset"])
        wthr = data["weather"]
        weather_description = wthr[0]["description"]
        now = datetime.datetime.now()
        content = """
        Today is the {day}th of {month}, {year}
        The sun rises at {hourrise}:{minrise}
        The sun sets at {hourset}:{minset}
        The average temperature is {temp} degrees Celsius
        The atmospheric pressure is {pressure} hectoPascals
        The humidity is {humidity}%
        The sky is clear today. Scattered showers are forecast in some areas.""".format(day=now.day, month=now.month, year=now.year, hourrise=sunrise.hour, minrise=sunrise.minute,
                                                                                        hourset=sunset.hour, minset=sunset.minute,
                                                                                        temp=current_temperature, pressure=current_pressure, humidity=current_humidity)
        speak(content)
        time.sleep(20)
    else:
        speak("Could not find your location.")

def play_song():
    speak('Please choose a song')
    mysong = get_text()
    while True:
        result = YoutubeSearch(mysong, max_results=10).to_dict()
        if result:
            break
    url = 'https://www.youtube.com' + result[0]['channel_link']
    webbrowser.open(url)
    speak("The song you requested has been opened.")

def change_wallpaper():
    api_key = 'your_unsplash_api_key'
    url = 'https://api.unsplash.com/photos/random?client_id=' + \
        api_key  # pic from unsplash.com
    f = urllib2.urlopen(url)
    json_string = f.read()
    f.close()
    parsed_json = json.loads(json_string)
    photo = parsed_json['urls']['full']
    # Location where we download the image to.
    urllib2.urlretrieve(photo, "C:/Users/Night Fury/Downloads/a.png")
    image = os.path.join("C:/Users/Night Fury/Downloads/a.png")
    ctypes.windll.user32.SystemParametersInfoW(20, 0, image, 3)
    speak('Your desktop wallpaper has been changed.')

def read_news():
    speak("What news do you want to read?")
    

    queue = "Trump"
    params = {
        'apiKey': 'your_news_api_key',
        "q": queue,
    }
    api_result = requests.get('http://newsapi.org/v2/top-headlines?', params)
    api_response = api_result.json()
    print("News")

    for number, result in enumerate(api_response['articles'], start=1):
        print(f"""News {number}:\nTitle: {result['title']}\nDescription: {result['description']}\nLink: {result['url']}
    """)
        if number <= 3:
            webbrowser.open(result['url'])

def tell_me_about():
    try:
        speak("What do you want to know about?")
        text = get_text()
        contents = wikipedia.summary(text).split('\n')
        speak(contents[0])
        time.sleep(20)
        for content in contents[1:]:
            speak("Do you want to hear more?")
            ans = get_text()
            if "yes" not in ans:
                break    
            speak(content)
            time.sleep(20)

        speak('Thank you for listening!!!')
    except:
        speak("Bot couldn't define your term. Please say that again.")

def help_me():
    speak("""Bot can help you with the following commands:
    1. Greetings
    2. Displaying time
    3. Opening websites, applications
    4. Searching on Google
    5. Sending email
    6. Weather forecast
    7. Playing music
    8. Changing desktop wallpaper
    9. Reading today's news
    10. Telling you about the world """)

def assistant():
    speak("Hello, what's your name?")
    time.sleep(3)
    name = get_text()
    if name:
        speak("Hi {}".format(name))
        time.sleep(3)
        speak("What can I help you with today?")
        time.sleep(5)
        while True:
            text = get_text()
            if not text:
                break
            elif "goodbye" in text or "see you later" in text or "stop" in text:
                stop()
                break
            elif "help" in text:
                help_me()
            elif "what's your name" in text:
                speak("I am your assistant. How can I help you today?")
            elif "hello" in text:
                hello(name)
            elif "time" in text or "date" in text:
                get_time(text)
            elif "open" in text:
                if "website" in text:
                    open_website(text)
                else:
                    open_application(text)
            elif "search" in text:
                open_google_and_search(text)
            elif "send email" in text or "send mail" in text:
                send_email(text)
            elif "weather" in text:
                current_weather()
            elif "play music" in text:
                play_song()
            elif "wallpaper" in text:
                change_wallpaper()
            elif "read news" in text:
                read_news()
            elif "tell me about" in text or "what is" in text or "who is" in text:
                tell_me_about()
            else:
                speak("Bot doesn't understand your command. Please say that again.")


assistant()
