# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 09:18:38 2015

@author: QN
"""

from datetime import datetime, date, time

import httplib2
import oauth2client
import os
import pytz
from apiclient import discovery
from oauth2client import client, tools

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'pygCal'


def get_service():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.


    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'pygCal.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)

    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    return service


def quick_add(service, calId):
    now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    event = service.events().quickAdd(calendarId=calId, text=now).execute()
    print("quick_add, id:", event['id'], now)


def list_calendarlist(service, initial):
    """return calendar ID if found match intial
    """
    calId = None
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            # print calendar_list_entry['summary'], calendar_list_entry['id']
            if calendar_list_entry['summary'] == initial:
                # calList.append(calendar_list_entry['id'])
                calId = calendar_list_entry['id']
                break
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    if calId is None:
        calId = insert_calendar(service, initial)

    return calId


def insert_calendar(service, summary):
    calendar = {
        'summary': summary,
        'timeZone': 'Asia/Taipei'
    }

    created_calendar = service.calendars().insert(body=calendar).execute()

    return created_calendar['id']


def del_calendar(service, initial=None):
    if initial != None:
        calId = list_calendarlist(service, initial)

    print(initial + ' deleteing.. ')
    service.calendarList().delete(calendarId=calId).execute()


def get_10_event(service):
    now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time

    events = service.events().list(
        calendarId='primary', timeMin=now, maxResults=10, singleEvents=True,
        orderBy='startTime').execute()

    return events


def get_UTC(year=0, month=0, day=0, hour=0, minute=0):
    local = pytz.timezone("Asia/Taipei")

    if month > 12:
        month = 1
        year += 1

    t = datetime.combine(date(year, month, day), time(hour, minute))
    t = local.localize(t, is_dst=None).astimezone(pytz.utc)
    t = datetime.strftime(t, "%Y-%m-%dT%H:%M:%SZ")

    return t


def get_month_event(service, calId, year, month, initial=None):
    if initial is not None:
        calId = list_calendarlist(service, initial)

    start_month = get_UTC(year, month, 1)
    end_month = get_UTC(year, month + 1, 1)

    events = service.events().list(
        calendarId=calId, timeMin=start_month, timeMax=end_month, singleEvents=True,
        orderBy='startTime').execute()

    return events


def get_all_event(service, calId, initial=None):
    if initial is not None:
        calId = list_calendarlist(service, initial)

    page_token = None
    while True:
        events = service.events().list(calendarId=calId, pageToken=page_token).execute()
        for event in events['items']:
            print(event['summary'], event['id'])
        page_token = events.get('nextPageToken')
        if not page_token:
            break

    return events


def print_event(events):
    for event in events['items']:
        print(event['summary'], event['id'])


def del_all_event(service, initial=None):
    if initial is not None:
        calId = list_calendarlist(service, initial)

    print(initial + ' deleteing.. ')
    events = get_all_event(service, calId)
    for event in events['items']:
        del_event(service, calId, event)


def del_month_event(service, calId, year, month):
    events = get_month_event(service, calId, year, month)
    for event in events['items']:
        del_event(service, calId, event)


def del_event(service, calId, event):
    service.events().delete(calendarId=calId, eventId=event['id']).execute()


def insert_event(service, calId, summary, start, end, des=None):
    body = {
        'summary': summary,
        'description': des,
        'start': {
            'dateTime': start,
            'timeZone': 'Asia/Taipei',
        },
        'end': {
            'dateTime': end,
            'timeZone': 'Asia/Taipei',
        },
    }
    service.events().insert(calendarId=calId, body=body).execute()


def main():
    service = get_service()

    members = ['NN', 'HJ', 'BT', 'XP', 'MN', 'EH']
    members += ['BJ', 'FW', 'VB', 'ZD', 'TS', 'GQ']
    members += ['EZ', 'BZ', 'AS', 'EC', 'ZD']
    members += ['FT', 'AD', 'BC', 'JE', 'RK', 'IJ']
    members = sorted(members)

    calId = list_calendarlist(service, 'QN')
    print(calId)


if __name__ == '__main__':
    main()
