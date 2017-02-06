from __future__ import print_function
import httplib2
from os import makedirs
from os.path import join, expanduser, exists
from operator import itemgetter

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


def get_credentials(client_secret_file, application_name, scopes):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = expanduser('~')
    credential_dir = join(home_dir, '.credentials')
    if not exists(credential_dir):
        makedirs(credential_dir)
    credential_path = join(credential_dir,
                           'sheets.googleapis.com-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_secret_file, scopes)
        flow.user_agent = application_name
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def read_sheet(spreadsheetID, rangeName, credentials):
    """
        prints top 5 slogans
    """
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    for i, row in enumerate(values):
        values[i] = row + [0] * (5 - len(row))
        try:
            values[i][4] = int(values[i][4])
        except ValueError:
            values[4] = 0
    values = sorted(values, key=itemgetter(4), reverse=True)
    if not values:
        return []
    # return columns A and E, which correspond to indices 0 and 4.
    return [(row[0], row[4]) for row in values[:5]]


if __name__ == '__main__':
    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/sheets.googleapis.com-python-quickstart.json
    scopes = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    client_secret_file = 'client_secret.json'
    application_name = 'Sheets Api'
    credentials = get_credentials(client_secret_file, application_name, scopes)

    spreadsheetId = '1xjcjBzvX4fI_baXA66VRytRFnusP60DYp0PsR_gyIOs'
    rangeName = 'Sheet1!A2:E'
    values = read_sheet(spreadsheetId, rangeName, credentials)
    for row in values:
        print('%s: %s' % (row[0], row[1]))
