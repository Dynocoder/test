from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import smtplib
import pyautogui
import pytz

# set these variables to your gmail address and password.
user_gmail = 'yourmail@gmail.com'
user_gmail_password = 'yourpassword'

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

SCOPES = ['https://www.googleapis.com/auth/calendar']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november',
          'december']
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENSIONS = ['st', 'nd', 'rd', 'th']
GET_DATE_STRINGS = ['what do i have', 'do i have plans', 'what are my plans', 'tell my plans']


def speak(audio):
    '''
    uses the pyttsx3 library to convert text to speech
    :param audio:
    :return:
    '''
    engine.say(audio)
    engine.runAndWait()


def wishme(*args):
    '''
    Wishes me according to the current time
    :return:
    '''
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak('Good Morning Sir')
    elif 12 <= hour < 16:
        speak('Good Afternoon Sir')
    elif 16 <= hour <= 20:
        speak('Good Evening Sir')
    else:
        speak('Good Night Sir')

    speak(str(*args))


def takeCommand():
    '''
    Takes microphone input from the user and returns text output
    :return:
    '''
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Listening...')
        r.pause_threshold = 1
        # inbuilt function that states the amount of non-speaking audio before the audio is considered complete
        audio = r.listen(source=source, timeout=6)

        try:
            print('Recognizing...')
            query = r.recognize_google(audio, language='en-in')
            print(f"user said: {query}")
            return query
        except Exception as e:
            print(e)
            return "None"


def sendMail(to, content):
    '''
    sends mail (gmail)
    :param to:
    :param content:
    :return:
    '''
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(user_gmail, user_gmail_password)
        server.sendmail(user_gmail, to, content)
        server.close()
        speak('mail sent')
    except Exception as e:
        print(e)
        return None


def searchWiki(query):
    '''
    searches wikipedia for any query you make
    replacec 'wikipedia', 'on wikipedia', 'search for','what do you mean by'
    :param query:
    :return:
    '''
    if 'wikipedia' in query:
        query = query.replace('wikipedia', "")
    if 'on wikipedia' in query:
        query = query.replace('on wikipedia', "")
    if 'search for' in query:
        query = query.replace('search for', "")
    if 'what do you mean by' in query:
        query = query.replace('what do you mean by', "")
    print(query)
    result = wikipedia.summary(query, sentences=5)
    speak(f"according to wikipedia")
    speak(result)


def screenshot(name):
    img = pyautogui.screenshot()
    img.save(f'C://Users//saura//OneDrive//Pictures//{name}.png')
    speak('the screenshot is saved in the Pictures folder')


def authenticate_google(scopes):
    '''
    authenticates google(if already authenticated, uses pickle) will return service which contains the calendar
    :return:
    '''
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = scopes

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    # the below four lines convert the date we provide in terms of utctime format
    # If you know what they really mean tell me, please 😅
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(),
                                          timeMax=end_date.isoformat(), singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])




def get_date(text):
    '''
    This fuction will return date value in the form of MM/DD/YYY
    it will convert any phrases passes to a date value if it contains the required data
    :param text:
    :return:
    '''
    text = text.lower()
    today = datetime.date.today()

    # if the query contains 'today' simply return today's date
    if text.count('today') > 0:
        return today
    # if the query contains 'tomorrow' simply return tomorrow's date
    if text.count('tomorrow') > 0:
        return today + datetime.timedelta(1)

    month = -1
    day_of_week = -1
    day = -1
    year = today.year

    # looping over the given phrase
    for word in text.split():
        # if the phrase contains name of month find its value from the list
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        # if the phrase contains day of teh week return day of the week from the list
        if word in DAYS:
            day_of_week = DAYS.index(word)
        # if the phrase contains the digit itself then convert it from str to int
        if word.isdigit():
            day = int(word)
        # again run a loop to get values of DAY_EXTENSIONS and then check if we have them in the word
        for ext in DAY_EXTENSIONS:
            x = word.find(ext)
            if x > 0:
                try:
                    day = int(word[:x])
                except:
                    pass

    # if given month is passed add to the year
    if month < today.month and month != -1:
        year = year+1

    # if only day is given then finding if month is this or the next
    if month == -1 and day != -1:
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    # if only day of the week is provided then find the date
    if day_of_week != -1 and month == -1 and day == -1:
        current_day_of_week = today.weekday()
        diff = day_of_week - current_day_of_week

        if diff < 0:
            diff += 7
            if text.count('next') >= 1:
                diff += 7

        return today + datetime.timedelta(diff)
    if day != -1:
        return datetime.date(month=month, day=day, year=year)


if __name__ == '__main__':
    service = authenticate_google(SCOPES)
    ask = 'how may I help you'
    wishme(ask)
    print(service)
    get_events(3, service)
    while True:
        query = takeCommand().lower()

        if 'quit' in query:
            break

        if 'wikipedia' in query:
            searchWiki(query)
        elif 'send mail' in query:
            to = 'Jasrajjassar775@gmail.com'

            speak('What should I say sir?')
            content = takeCommand()
            sendMail(to, content)
        elif 'screenshot' in query:
            speak('what should I name it?')
            name = takeCommand()
            screenshot(name)
        # looping over the phrase to find if the word has any str that we need to trigger the get_date function
        for phrases in GET_DATE_STRINGS:
            if phrases in query.lower():
                day = get_date(query)
                if day:
                    get_events(day, service)
                else:
                    speak('please try again')