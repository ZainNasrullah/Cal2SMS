import httplib2, os, re, pyautogui, time

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime

#part of Google Calendar API
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

#default scope values, saves local authentication
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'

def send_to_vendor(vendors, template):
    a, b = pyautogui.locateCenterOnScreen('ring.PNG')
    pyautogui.click(a, b)
    time.sleep(0.25)

    pyautogui.typewrite(vendors)
    time.sleep(0.25)

    a, b = pyautogui.locateCenterOnScreen('chat.PNG')
    pyautogui.click(a, b)

    pyautogui.typewrite(template)
    time.sleep(0.25)

    #pyautogui.press('enter')

    a, b = pyautogui.locateCenterOnScreen('back.PNG')
    pyautogui.click(a, b)

    for i in range(len(vendors)):
        pyautogui.press('delete')




def convert_time_12h(time):
    time12 = time.split(sep=':')
    if int(time12[0]) >= 13:
        time12[0] = str(int(time12[0]) - 12)
        time12[1]+='PM'
    else:
        time12[1] += 'AM'
    return ':'.join(time12)

def get_credentials():

    #check for credentials
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'calendar-python-quickstart.json')

    #store
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()

    #if credentials were invalid, create new
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    #standard google api for building a calender object
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar','v3',http=http)

    #sets the date range for the current day 12am-11:59pm
    today = datetime.datetime.utcnow().date()
    todaySimple = datetime.datetime(today.year, today.month, today.day)
    now = todaySimple.isoformat() + 'Z'  # 'Z' indicates UTC time
    end = (todaySimple + datetime.timedelta(1)).isoformat() + 'Z'
    print(now)
    print(end)

    #generates a list of calenders and then writes them to a list
    calendarResults = service.calendarList().list().execute()
    calendars = calendarResults.get('items',[])

    timeregex = re.compile(r'\d{4}-\d{2}-\d{2}T(.{5})')
    companyregex = re.compile(r'(.*:|-)\s?(.*)')
    CompanyOrder = {}
    OrderCount = {}

    #iterate through the calendars and write all events to the screen
    for eachcalendar in calendars:
        print('\n')
        print(eachcalendar.get('summary'))
        eventsResult = service.events().list(
            calendarId= eachcalendar.get('id'), timeMin=now, timeMax = end, maxResults=10, singleEvents=True,
            orderBy='startTime').execute()
        #save events to a list
        events = eventsResult.get('items', [])

        #Print out events
        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            time = timeregex.search(start).group(1)
            time12 = convert_time_12h(time)
            company = companyregex.search(event['summary'])
            if company != None:
                vendor = company.group(2)
                CompanyOrder.setdefault(vendor, '')
                OrderCount.setdefault(vendor, 0)
                OrderCount[vendor] += 1
                CompanyOrder[vendor] += str(OrderCount[vendor]) + ') Order for '\
                                        + eachcalendar['summary'] + ', Arrive by ' + time12 + '.\n'
            else:
                print('No upcoming events found.')


    for vendors in CompanyOrder.keys():
        template = 'Hi,\n' \
                   'This is a message from Epicater reminding you about your order(s) today which are:\n' \
                   + CompanyOrder[vendors]  \
                   + 'If you need labels for future orders please let us know. Additionally, please' \
                     ' remember to send us a picture of your order set up. This will help us when promoting' \
                     ' you on social media. Let us know of any questions or concerns you might have. ' \
                     'Hope you have an epic day!'

        print('\n'+ vendors + '\n' + template)


main()
