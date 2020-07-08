import json
import xlrd
import requests

payloads = []

session = requests.Session()
session.headers = {'Content-Type': 'application/json'}
session.auth = 'david.milward@exclaimer.com/token', 'JgEC8hToAox1h9LTOqCs5FTYcodDzRaERc65munF'
url = 'https://exclaimersupport1592334089.zendesk.com/api/v2/users/create_many.json'

users_dict = {'users': []}
book = xlrd.open_workbook('/mnt/c/Users/korgy/Documents/Work/repos/Zendesk-api-management/files/TestUser.xlsx')
sheet = book.sheet_by_name('Sheet1')

for row in range(1, sheet.nrows):
    if sheet.row_values(row)[2]:
        users_dict['users'].append(
            {
                'name': sheet.row_values(row)[0],
                'email': sheet.row_values(row)[1],
                # 'external_id': int(sheet.row_values(row)[2]),
                'role': sheet.row_values(row)[3],
                'custom_role_id': int(sheet.row_values(row)[4]),
                'organization_id': int(sheet.row_values(row)[5]),
                'tags': (sheet.row_values(row)[6])
            }
        )

        if len(users_dict['users']) == 100:
            payloads.append(json.dumps(users_dict))
            users_dict = {'users': []}

if users_dict['users']:
    payloads.append(json.dumps(users_dict))

for user in users_dict['users']:
    print(user)

for payload in payloads:
    response = session.post(url, data=payload)
    if response.status_code != 200:
        print('Import failed with status {}'.format(response.status_code))
        exit()
    print('Successfully imported a batch of users')
    