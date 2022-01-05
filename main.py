import json
from googleapiclient import discovery
from google.oauth2 import service_account
import dateutil.parser
from dateutil import tz
import requests

credentials = None
service = None
settings = None


def processJsonData(data, settings, spreadsheetid, token):
    global credentials, service
    if credentials is None or service is None:
        credentials = service_account.Credentials.from_service_account_file('service-account.json').with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
        service = discovery.build('sheets', 'v4', credentials=credentials)
    elif credentials.valid is not True:
        credentials = service_account.Credentials.from_service_account_file('service-account.json').with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
        service = discovery.build('sheets', 'v4', credentials=credentials)
    if (
        'event' in data
        and 'uri' in data['payload']
        and 'name' in data['payload']
        and 'email' in data['payload']
    ):
        event = getEvent(data['payload']['event'], token)
        eventuri = data['payload']['event']
        if event and data['event'] == "invitee.created":
            eventtype = event['event_type']
            name = event['tutor']
            if eventtype in settings['event_types']:
                spreadsheetid = findSpreadsheet(settings, name)
                eventcanceled = 'FALSE'
                invitee = data['payload']['name']
                inviteeEmail = data['payload']['email']
                estart = ConvertTimeZone(dateutil.parser.parse(event['start_time']))
                eend = ConvertTimeZone(dateutil.parser.parse(event['end_time']))
                duration = Duration(estart, eend)
                eventstart = estart.strftime("%b %d %Y %H:%M:%S")
                eventend = eend.strftime("%b %d %Y %H:%M:%S")
                row = [name, invitee, inviteeEmail, eventtype, eventstart, eventend, duration, eventcanceled, eventuri]
                if 'questions_and_answers' in data['payload']:
                    for question in data['payload']['questions_and_answers']:
                        if 'answer' in question:
                            row.append(question['answer'])
                    appendData(service, settings, row, spreadsheetid)
                return True
        elif event and data['event'] == "invitee.canceled":
            name = event['tutor']
            eventtype = event['event_type']
            if eventtype in settings['event_types']:
                spreadsheetid = findSpreadsheet(settings, name)
                getIDList(service, settings, eventuri, spreadsheetid)
                return True
    return False


def findSpreadsheet(settings, name):
    for item in settings['custom']:
        if item['name'].lower().strip() == name.lower().strip():
            return item['spreadsheet']
    return settings['spreadsheet']


def upDateCanceled(service, settings, position, canceled, spreadsheetid):
    spreadsheet_id = spreadsheetid
    range_ = settings['cancelCol'] + str(position)
    value_range_body = {"values": [[canceled]]}
    request = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_,
                valueInputOption='USER_ENTERED',
                body=value_range_body)
    request.execute()


def appendData(service, settings, row, spreadsheetid):
    spreadsheet_id = spreadsheetid
    range_ = settings['range']
    value_range_body = {"values": [row]}
    request = service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=range_,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=value_range_body)
    request.execute()


def getIDList(service, settings, eventuri, spreadsheetid):
    sheet = service.spreadsheets()
    spreadsheet_id = spreadsheetid
    id_range = settings['idrange']
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=id_range).execute()
    values = result.get('values', [])
    flat_list = [item for sublist in values for item in sublist]
    if eventuri in flat_list:
        indices = [i for i, x in enumerate(flat_list) if x == eventuri]
        for i in indices:
            upDateCanceled(service, settings, i + 1, "TRUE", spreadsheetid)
        return True
    return False

def Duration(start, end):
    duration = end - start
    return duration.total_seconds()/60

def getEvent(uri, token):
    r = requests.get(uri, headers={"Content-Type":"application/json", "Authorization": "Bearer " + token})
    jsdata = r.json()
    if (
        jsdata 
        and 'resource' in jsdata
        and 'name' in jsdata['resource']
        and 'start_time' in jsdata['resource']
        and 'end_time' in jsdata['resource']
        and 'event_type' in jsdata['resource']
        and 'location' in jsdata['resource']
        and 'event_memberships' in jsdata['resource']
    ):
        for item in jsdata['resource']['event_memberships']:
            if 'user' in item:
                tutorName = getUser(item['user'], token)
        event_type = getEventType (jsdata['resource']['event_type'], token)
        return {"name": jsdata['resource']['name'], "start_time": jsdata['resource']['start_time'], "end_time": jsdata['resource']['end_time'], "event_type": event_type, "location": jsdata['resource']['location']['type'], "tutor": tutorName}     
    return False

def getUser(uri, token):
    r = requests.get(uri, headers={"Content-Type":"application/json", "Authorization": "Bearer " + token})
    jsdata = r.json()
    if 'resource' in jsdata and 'name' in jsdata['resource']:
        return jsdata['resource']['name']
    return ""

def getEventType (uri, token):
    r = requests.get(uri, headers={"Content-Type":"application/json", "Authorization": "Bearer " + token})
    jsdata = r.json()
    if 'resource' in jsdata and 'slug' in jsdata['resource']:
        return jsdata['resource']['slug']
    return ""

def ConvertTimeZone(d):
    to_zone = tz.gettz('America/Chicago')
    central = d.astimezone(to_zone)
    return central


def main(request):
    global settings
    
    request_json = request.get_json(silent=True)

    if (
        request_json
        and 'event' in request_json
        and 'payload' in request_json
        and 'event' in request_json['payload']
    ):
        if settings is None:
            with open('settings.json') as settings_file:
                settings = json.load(settings_file)
        spreadsheetid = settings['spreadsheet']
        token = settings['calendlyToken']
        processJsonData(request_json, settings, spreadsheetid, token)
    return "Data processed"
