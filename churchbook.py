# calendar imports
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime
import time

# webscraping imports
import requests
from bs4 import BeautifulSoup

# data
url_login = "https://newlife010.churchbook.nl/login/Loginform.cfm"
roosterUrl = "https://newlife010.churchbook.nl/locatie/sp_agenda.cfc?method=getEvents"
headers = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.101 Safari/537.36"}
checkPersonen = ["Stefan Poot", "Anne Lucia Snel"]
SCOPES = ['https://www.googleapis.com/auth/calendar']
calendar = True

today = datetime.now().strftime("%Y-%m-%d")
print(datetime.now().strftime("%d-%m-%Y"))

def scraper():
    def ingedeeldChecken(datum, id, title):
        if "Kerkdienst" in title:
            title = "Band"
        elif "Zondagclub" in title:
            title = "Club"

        roosterItemUrl = "https://newlife010.churchbook.nl/locatie/samenkomst.cfm?project_id=" + str(id)
        source = s.get(roosterItemUrl, headers=headers).text
        sourceSoup = BeautifulSoup(source, "lxml")
        roosterGebieden = sourceSoup.find_all("table", class_="table")[-1]
        for roosterGebied in roosterGebieden:
            ingedeeldePersonen = str(roosterGebied)
            print(roosterGebied)
            if not "Planner" in ingedeeldePersonen:
                for persoon in checkPersonen:
                    if persoon in ingedeeldePersonen:
                        print("   -", persoon, "is ingedeeld voor", title)
                        if calendar:
                            addCalendarEvent(title, datum)

    ###############################################################
    with requests.Session() as s:
        print("Verbinding maken met newlife...\n")

        #token ophalen van de website
        inlogpage = s.get(url_login, headers=headers).text
        loginSoup = BeautifulSoup(inlogpage, "lxml")
        loginPage = loginSoup.find_all("div", class_="card-body")
        loginData = loginSoup.find_all("input")

        for data in loginData:
            if "token" in str(data):
                print(str(data))
                token = str(data).split('value="')[-1][:-3]

        formData = {"refer": "", "token": token, "loginnaam": "", "password": ""}

        r = s.post(url_login, data=formData, headers=headers)
        if str(r) == "<Response [200]>":
            print("Login succesvol! Roosters ophalen...\n")
            grootRooster = s.get(roosterUrl, headers=headers).json()

            print("Roosters voor de komende tijd:")
            for roosterItem in grootRooster:
                print(roosterItem)
                datum = roosterItem["start"].split("T")[0]
                if datum > today:
                    if "Kerkdienst" in roosterItem["title"]:
                        print("*", datum)
                        ingedeeldChecken(datum, roosterItem["id"], roosterItem["title"])
                    elif "Zondagclub" in roosterItem["title"]:
                        ingedeeldChecken(datum, roosterItem["id"], roosterItem["title"])

    s.close()

def addCalendarEvent(ingeroosterdeTaak, datum):
    agendaItemBestaatAl = False
    creds = None

    if ingeroosterdeTaak in ["Zanger 1", "Zanger 2", "Piano"]:
        ingeroosterdeTaak = "New Life Band"
    elif ingeroosterdeTaak == "Zondagclub 1":
        ingeroosterdeTaak = "Zondagclub"

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    calendarDateStart = datum + "T10:00:00+01:00"
    calendarDateEnd = datum + "T12:00:00+01:00"

    newEvent = {
      'summary': ingeroosterdeTaak,
      'start': {
        'dateTime': calendarDateStart,
        'timeZone': 'America/Los_Angeles',
      },
      'end': {
        'dateTime': calendarDateEnd,
        'timeZone': 'America/Los_Angeles',
      },
    }

    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=100, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    for event in events:
        if event["summary"] == newEvent["summary"]:
            if event["start"]["dateTime"] == newEvent["start"]["dateTime"]:
                if event["end"]["dateTime"] == newEvent["end"]["dateTime"]:
                    agendaItemBestaatAl = True

    if agendaItemBestaatAl:
        print("      > Er is al een Calendar item voor", ingeroosterdeTaak, "op", datum)
    else:
        newEvent = service.events().insert(calendarId='primary', body=newEvent).execute()
        print('      > Event created: %s' % (newEvent.get('htmlLink')))

scraper()
