import argparse
import os

import openpyxl

parser = argparse.ArgumentParser(description='Add resume to huntflow')

parser.add_argument("token", type=str, help="huntflow token")
parser.add_argument("-d", "--base_dir", help="directory with database")
parser.add_argument("-db", "--db_name", help="database name")
parser.add_argument("-m", "--mail", help="add your mail to user-agent")

args = parser.parse_args()

SCRIPT_INFO = 'script_info.txt'

base_rows = {
    'position': 1,
    'name': 2,
    'salary': 3,
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


def get_candidate_content(name, position):
    for path, dirs, files in os.walk(os.path.join(args.base_dir, position)):
        for i in files:
            if _normalize_str(name) in _normalize_str(i):
                file = open(os.path.join(path, i), 'rb')
                result = file.read()
                file.close()
                return result


def _normalize_str(bad_str):
    return bad_str.strip().lower().replace('й', 'и').replace('̆', '')


filename = os.path.join(args.base_dir, args.db_name)
wb = openpyxl.load_workbook(filename)
ws = wb.active
current_row = get_row()

while True:
    candidate_info = {}
    cell_info = None

    for i, j in base_rows.items():
        cell_info = ws.cell(row=current_row, column=j).value
        if cell_info is not None:
            candidate_info[i] = str(cell_info)

    if not candidate_info:
        break

    content = get_candidate_content(candidate_info['name'], candidate_info['position'])

    current_row += 1
    save_row(current_row)
