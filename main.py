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
CREATE_EVENT_STRS = ['create an event', 'make an event']


def speak(audio):
    """
    uses the pyttsx3 library to convert text to speech
    :param audio:
    :return:
    """
    engine.say(audio)
    engine.runAndWait()


def wishme(*args):
    """
    Wishes me according to the current time
    :return:
    """
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak('Good Morning Sir')
    elif 12 <= hour < 16:
        speak('Good Afternoon Sir')
    elif 16 <= hour <= 20:
        speak('Good Evening Sir')
    else:
        speak('Good Evening Sir')

    speak(str(*args))


def takeCommand():
    """
    Takes microphone input from the user and returns text output
    :return:
    """
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
            return query.lower()
        except Exception as e:
            print(e)
            return "None"


def sendMail(to, content):
    """
    sends mail (gmail)
    :param to:
    :param content:
    :return:
    """
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
    """
    searches wikipedia for any query you make
    replacec 'wikipedia', 'on wikipedia', 'search for','what do you mean by'
    :param query:
    :return:
    """
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
    """
    authenticates google(if already authenticated, uses pickle) will return service which contains the calendar
    :return:
    """
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


def get_date(text):
    """
    This fuction will return date value in the form of MM/DD/YYY
    it will convert any phrases passes to a date value if it contains the required data
    :param text:
    :return:
    """
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
        year = year + 1

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


def get_events(day, service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                          singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("+")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0]) - 12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            speak(event["summary"] + " at " + start_time)



def get_time(time):
    GMT_OFF = '5:30'
    text_array = time.split(' to ')
    # print(text_array)
    start_hour = text_array[0].split(':')[0].split(' ')[-1]
    # print(start_hour)
    start_minute = text_array[0].split(':')[1].split(' ')[0]
    # print(start_minute)

    end_hour = text_array[1].split(' ')[0].split(':')[0]
    # print(end_hour)
    end_minute = text_array[1].split(' ')[0].split(':')[1]
    # print(type(end_minute))

    day = get_date(time)
    start_ampm = text_array[0].split(' ')[1].split('.')
    if start_ampm[0] == 'a' and int(start_hour) < 12:
        pass
    elif start_ampm[0] == 'p' and int(start_hour) > 12:
        start_hour = int(start_hour) + 12
    start_str = f"{day}T{start_hour}:{start_minute}:00+{GMT_OFF}"
    end_ampm = text_array[1].split(' ')[1].split('.')
    if end_ampm[0] == 'a' and int(end_hour) < 12:
        pass
    elif end_ampm[0] == 'p' and int(end_hour) > 12:
        end_hour = int(end_hour) + 12
    end_str = f"{day}T{end_hour}:{end_minute}:00+{GMT_OFF}"
    return start_str, end_str


def create_events(start, end, content, service):
    EVENT = {'summary': content,
             'start': {'dateTime': start},
             'end': {'dateTime': end}}
    service = service
    ce = service.events().insert(calendarId='primary',
                                 sendNotifications=True,
                                 body=EVENT).execute
    print(ce)




if __name__ == '__main__':
    services = authenticate_google(SCOPES)
    say = 'how may i help you'
    wishme(say)
    x = 0
    print(services)

    while x == 0:
        text = "create an event on monday from 02:40 p.m. to 03:40 p.m."
        print(text)

        if 'quit' in text:
            break
        if 'wikipedia' in text:
            searchWiki(text)
        if 'mail' in text:
            to = 'sauravprashar21@gmail.com'
            content = takeCommand()
            sendMail(to, content)

        if 'screenshot' in text:
            name_of_screenshot = takeCommand()
            screenshot(name_of_screenshot)
        for phrases in GET_DATE_STRINGS:
            if phrases in text:
                day = get_date(text)
                if day:
                    print(get_events(day, services))
                    x += 1
                else:
                    speak('unable to get the events, please re-try')
        if 'create an event' in text:
            str_time, end_time = get_time(text)
            content = 'Main Task'
            create_events(str_time, end_time, content, services)
            break
