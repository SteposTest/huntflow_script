import argparse
import os
import random
import string
from urllib.parse import urljoin

import openpyxl
import requests

parser = argparse.ArgumentParser(description='Add resume to huntflow')

parser.add_argument("token", type=str, help="huntflow token")
parser.add_argument("-d", "--base_dir", help="directory with database")
parser.add_argument("-db", "--db_name", help="database name")
parser.add_argument("-m", "--mail", help="add your mail to user-agent")

args = parser.parse_args()

BASE_HEADERS = {
    'User-Agent': 'test_script/1.0' + f'({args.mail})',
    'Authorization': f'Bearer {args.token}'
}
BASE_URL = 'https://api.huntflow.ru/'
SCRIPT_INFO = 'script_info.txt'
BASE_ROWS = {
    'position': 1,
    'name': 2,
    'money': 3,
    'comment': 4,
    'status': 5,
}


def get_row():
    try:
        with open(SCRIPT_INFO) as f_obj:
            row = int(f_obj.read())
    except:
        row = 2
        save_row(row)
    return row


def save_row(row):
    with open(SCRIPT_INFO, 'w') as f_obj:
        f_obj.write(str(row))


def get_candidate_content_name(name, position):
    for path, dirs, files in os.walk(os.path.join(args.base_dir, position)):
        for i in files:
            if _normalize_str(name) in _normalize_str(i):
                return os.path.join(path, i)


def get_random_str(length=8):
    return ''.join(random.sample(string.ascii_letters + string.digits, length))


def huntflow_request(endpoint, method='GET', headers=None, **kwargs):
    if headers is None:
        headers = {}
    headers.update(BASE_HEADERS)
    return requests.request(
        method=method,
        url=urljoin(BASE_URL, endpoint),
        headers=headers,
        **kwargs
    ).json()


def get_no_none(dict_obj, key, opt=None):
    result = dict_obj.get(key, opt)
    if result is None:
        return {}
    return result


def _normalize_str(bad_str):
    return bad_str.strip().lower().replace('й', 'и').replace('̆', '')


accounts = huntflow_request('/accounts')
account_id = accounts['items'][0]['id']

statuses_info = huntflow_request(f'/account/{account_id}/vacancy/statuses')['items']
vacancies_info = huntflow_request(f'/account/{account_id}/vacancies')['items']

filename = os.path.join(args.base_dir, args.db_name)
wb = openpyxl.load_workbook(filename)
ws = wb.active
current_row = get_row()

while True:
    candidate_info = {}
    cell_info = None

    for i, j in BASE_ROWS.items():
        cell_info = ws.cell(row=current_row, column=j).value
        if cell_info is not None:
            candidate_info[i] = str(cell_info).strip()

    if not candidate_info:
        break

    content_info = {}
    content_name = get_candidate_content_name(candidate_info['name'], candidate_info['position'])
    if content_name is not None:
        content_info = huntflow_request(
            endpoint=f'/account/{account_id}/upload',
            method='POST',
            headers={'X-File-Parse': 'true'},
            files={
                'file': (f'resume{get_random_str}.pdf', open(content_name, 'rb'), 'application/pdf')
            },
        )

    full_name = candidate_info['name'].split()
    data = {
        'last_name': get_no_none(get_no_none(content_info, 'fields', {}), 'name', {}).get('last', full_name[1]),
        'first_name': get_no_none(get_no_none(content_info, 'fields', {}), 'name', {}).get('first', full_name[1]),
        'middle_name': get_no_none(get_no_none(content_info, 'fields', {}), 'name', {}).get('middle', ''),
        'position': candidate_info['position'],
        'money': candidate_info['money'],
        'phone': '\n'.join(get_no_none(content_info, 'fields', {}).get('phones', [])),
        'company': get_no_none(get_no_none(content_info, 'fields', {}), 'experience', [{}])[0].get('company', ''),
        'birthday_day': get_no_none(get_no_none(content_info, 'fields', {}), 'birthdate', {}).get('day', ''),
        'birthday_month': get_no_none(get_no_none(content_info, 'fields', {}), 'birthdate', {}).get('month', ''),
        'birthday_year': get_no_none(get_no_none(content_info, 'fields', {}), 'birthdate', {}).get('year', ''),
        'email': get_no_none(content_info, 'fields', {}).get('email', ''),
        'photo': get_no_none(content_info, 'photo', {}).get('id', None),
        'externals': [
            {
                'data': {
                    'body': content_info.get('text', '')
                },
                'auth_type': 'NATIVE',
                'files': [
                    {
                        'id': content_info.get('id', None)
                    }
                ],
                'account_source': 73803
            }
        ]
    }

    resume_info = huntflow_request(endpoint=f'/account/{account_id}/applicants', method='POST', json=data)

    vacancy_id = sorted(
        [i for i in vacancies_info if i['position'] == candidate_info['position']],
        key = lambda k: k['created']
    )[-1]['id']
    statuses_id = [i['id'] for i in statuses_info if i['name'] == candidate_info['status']][0]

    huntflow_request(
        endpoint=f'/account/{account_id}/applicants/{resume_info["id"]}/vacancy',
        method='POST',
        json={
            'vacancy': vacancy_id,
            'status': statuses_id,
            'files':[{'id': content_info.get('id', None)}],
        },
    )

    current_row += 1
    save_row(current_row)
