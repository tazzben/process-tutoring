import json
from googleapiclient import discovery
from google.oauth2 import service_account
import dateutil.parser
from dateutil import tz


credentials = None
service = None
settings = None


def processJsonData(data, settings, spreadsheetid):
    global credentials, service
    if credentials is None or service is None:
        credentials = service_account.Credentials.from_service_account_file('service-account.json').with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
        service = discovery.build('sheets', 'v4', credentials=credentials)
    elif credentials.valid is not True:
        credentials = service_account.Credentials.from_service_account_file('service-account.json').with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
        service = discovery.build('sheets', 'v4', credentials=credentials)
    if (
        'slug' in data['payload']['event_type']
        and 'uuid' in data['payload']['event']
        and 'start_time' in data['payload']['event']
        and 'end_time' in data['payload']['event']
        and 'assigned_to' in data['payload']['event']
        and 'name' in data['payload']['event_type']
        and 'duration' in data['payload']['event_type']
        and 'canceled' in data['payload']['event']
        and 'name' in data['payload']['invitee']
        and 'email' in data['payload']['invitee']
    ):
        event = data['event']
        if event == "invitee.created":
            eventtypeslug = data['payload']['event_type']['slug']
            if eventtypeslug in settings['event_types']:
                for name in data['payload']['event']['assigned_to']:
                    spreadsheetid = findSpreadsheet(settings, name)
                    eventuuid = data['payload']['event']['uuid']
                    eventtype = data['payload']['event_type']['name']
                    eventcanceled = str(data['payload']['event']['canceled'])
                    invitee = data['payload']['invitee']['name']
                    inviteeEmail = data['payload']['invitee']['email']
                    duration = data['payload']['event_type']['duration']
                    eventstart = ConvertTimeZone(dateutil.parser.parse(data['payload']['event']['start_time'])).strftime("%b %d %Y %H:%M:%S")
                    eventend = ConvertTimeZone(dateutil.parser.parse(data['payload']['event']['end_time'])).strftime("%b %d %Y %H:%M:%S")
                    row = [name, invitee, inviteeEmail, eventtype, eventstart, eventend, duration, eventcanceled, eventuuid]
                    if 'questions_and_answers' in data['payload']:
                        for question in data['payload']['questions_and_answers']:
                            if 'answer' in question:
                                row.append(question['answer'])
                    appendData(service, settings, row, spreadsheetid)
                return True
        elif event == "invitee.canceled":
            eventtypeslug = data['payload']['event_type']['slug']
            if eventtypeslug in settings['event_types']:
                eventuuid = data['payload']['event']['uuid']
                for name in data['payload']['event']['assigned_to']:
                    spreadsheetid = findSpreadsheet(settings, name)
                    getIDList(service, settings, eventuuid, data, spreadsheetid)
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


def getIDList(service, settings, uuid, data, spreadsheetid):
    sheet = service.spreadsheets()
    spreadsheet_id = spreadsheetid
    id_range = settings['idrange']
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=id_range).execute()
    values = result.get('values', [])
    flat_list = [item for sublist in values for item in sublist]
    if uuid in flat_list:
        indices = [i for i, x in enumerate(flat_list) if x == uuid]
        for i in indices:
            upDateCanceled(service, settings, i + 1, str(data['payload']['event']['canceled']), spreadsheetid)


def ConvertTimeZone(d):
    to_zone = tz.gettz('America/Chicago')
    central = d.astimezone(to_zone)
    return central


def main(request):
    global settings
    content_type = request.headers['content-type']
    if content_type == 'application/json':
        request_json = request.get_json(silent=True)
        if (
            request_json
            and 'event' in request_json
            and 'payload' in request_json
            and 'event_type' in request_json['payload']
            and 'event' in request_json['payload']
            and 'invitee' in request_json['payload']
        ):
            if settings is None:
                with open('settings.json') as settings_file:
                    settings = json.load(settings_file)
            spreadsheetid = settings['spreadsheet']
            processJsonData(request_json, settings, spreadsheetid)
    return "Data processed"
