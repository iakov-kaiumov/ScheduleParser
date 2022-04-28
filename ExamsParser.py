import MIPTScheduleParser as sp

import xlrd
import pandas as pd
import datetime
import json

import requests
from requests.auth import AuthBase

SRC_DIR = 'Schedule/EXAMS_2021/'

TIME_COLUMN = 1
TIME_START_ROW = 7

KEY_ROW = 6
KEY_START_COLUMN = 2


class TokenAuth(AuthBase):
    """Implements a custom authentication scheme."""

    def __init__(self, token):
        self.token = 'Bearer ' + token

    def __call__(self, r):
        """Attach an API token to a custom auth header."""
        r.headers['Authorization'] = f'{self.token}'
        return r


myToken = TokenAuth(
    'MzU4Ni0yLTIyODQtY2M4OTFlNjkyNDg5ZjE3NjVkNTdmZTVlOTJlZDE3YjQ0MmFmNGY2NTU2MjA5MzFhMjllMDcyYjNjYTE5ODdjNw==')


def load_places_list():
    url = 'https://appadmin.mipt.ru/api/schedule-place?perpage=10000'
    data = requests.get(url, auth=myToken).json()
    values = []
    values_dict = dict()
    for key, value in data['data']['paginator'].items():
        values.append(value)
        values_dict[key] = value
    return values, values_dict


places, places_dict = load_places_list()


def read_auditorium_filter_dict():
    df = pd.read_excel('auditoriums_filter.xlsx', sheet_name=0)
    d = dict()
    for i in range(df.shape[0]):
        d[df.iloc[i, 0]] = df.iloc[i, 1]
    return d


auditorium_filter_dict = read_auditorium_filter_dict()


class ScheduleItem:
    def __init__(self, name, day, type, start_time, end_time, auditoriums):
        self.name = name
        self.allDay = True
        self.day = day
        self.type = type
        self.startTime = start_time
        self.endTime = end_time
        self.auditoriums = auditoriums


def find_auditorium(text):
    auditoriums = []
    values = list(map(lambda x: x.strip(), text.split(',')))
    values = list(map(lambda x: str(auditorium_filter_dict.get(x, x)), values))

    for value in values:
        if value == '':
            continue
        for place in places:
            if value in place['name']:
                auditoriums.append(place['id'])
                break
    auditoriums = list(set(auditoriums))
    auditoriums = [places_dict[key] for key in auditoriums]
    return auditoriums


def process_name(name):
    name = name.strip()
    if 'А.К.' in name:
        name = name.split(':')[-1]
    return name


def parse_file(groups, file_name):
    book = xlrd.open_workbook(file_name, formatting_info=False)
    sheet = book.sheet_by_index(0)

    dates = sheet.col_slice(colx=TIME_COLUMN, start_rowx=TIME_START_ROW)
    dates = [d.value for d in dates]

    for col in range(KEY_START_COLUMN, sheet.ncols):
        key = sheet.cell(colx=col, rowx=KEY_ROW).value.strip().replace('\n', '')

        events = []
        for row in range(KEY_ROW + 1, sheet.nrows):
            cell = sheet.cell(colx=col, rowx=row).value.strip()
            if cell == '':
                continue

            # Date
            date = dates[row - TIME_START_ROW]
            start_time = date + ' 09:00:00'
            end_time = date + ' 20:00:00'
            date = datetime.datetime.strptime(date, '%Y-%m-%d')
            day = date.weekday() + 1

            values = cell.split('/')
            if len(values) > 2:
                for i in range(0, len(values), 2):
                    name = process_name(values[i])

                    auditoriums = find_auditorium(values[i + 1])
                    events.append(
                        ScheduleItem(
                            name=name, day=day, type='exam', start_time=start_time,
                            end_time=end_time, auditoriums=auditoriums
                        )
                    )
            else:
                name = process_name(values[0])
                auditoriums = []
                if len(values) > 1:
                    auditoriums = find_auditorium(values[1])

                events.append(
                    ScheduleItem(
                        name=name, day=day, type='exam', start_time=start_time,
                        end_time=end_time, auditoriums=auditoriums
                    )
                )

        groups[key] = [item.__dict__ for item in events]


def unmerge():
    sp.delete_unmerged_files()
    for file_name in sp.get_files(SRC_DIR):
        sp.unmerge_excel_file(file_name, sp.UNMERGED_DATA + file_name.split('/')[-1], add_color=False)


def main():
    # Unmerged data
    unmerge()

    groups = dict()

    for file_name in sp.get_files(sp.UNMERGED_DATA):
        print(file_name)
        parse_file(groups, file_name)

    j = dict()
    j['timetable'] = groups

    print(json.dumps(j['timetable']['Б02-824']))

    sp.save_as_json('exams.json', j)


if __name__ == '__main__':
    main()
