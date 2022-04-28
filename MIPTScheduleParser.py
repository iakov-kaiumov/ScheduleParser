import xlrd
import os
import json
import xlwt
import time
import re

RAW_DATA = 'Schedule/raw_data_2021_2/'
UNMERGED_DATA = 'Schedule/unmerged_data/'
VERSION = '2.8'

TIME_START_ROW = 5
TIME_STOP_ROW = 15
TIME_COL = 1

KEYS_ROW = 4

LEC = [str((255, 153, 204)), str((255, 128, 128))]
LAB = [str((255, 255, 153)), str((255, 255, 153))]
SEM = [str((204, 255, 255)), str((204, 255, 255)), str((204, 204, 255))]
color_map = {"LEC": LEC, "LAB": LAB, "SEM": SEM}

BANNED_EXPRESSIONS = ['Дни', 'Часы', 'Б', 'С', 'М']
identifier = 0


class ScheduleItem:
    def __init__(self, id, name="", prof="", place="", day=0, type="SEM", startTime="9:00", endTime="10:25", notes="",
                 color="SEM"):
        self.id = id
        self.name = name
        self.prof = prof
        self.place = place
        self.day = day
        self.type = type
        self.startTime = startTime
        self.endTime = endTime
        self.color = color
        self.notes = notes
        self.synchronized = 1
        self.updated = "2021-06-30 18:39:00.000000"


def get_files(scr_dir):
    return [scr_dir + dr for dr in os.listdir(scr_dir) if not dr.startswith('.')]


def contains(expressions, s):
    for e in expressions:
        if e.lower() in s.lower():
            return True
    return False


def insert_colon(s):
    if len(s) == 3:
        return s[0] + ':' + s[1:]
    return s[:2] + ':' + s[2:]


def add_time(t1, t2):
    h1, m1 = list(map(int, t1.split(':')))
    if t1[0] == '-':
        m1 *= -1
    h2, m2 = list(map(int, t2.split(':')))
    if t2[0] == '-':
        m2 *= -1
    h = h1 + h2 + (m1 + m2) // 60
    m = (m1 + m2) % 60
    return '{0:02d}'.format(h) + ':' + '{0:02d}'.format(m)


def handle_lesson(text):
    text = ' '.join(text.replace("\"", "").split())
    split = text.split('|')
    t = split[1]
    text = split[0]
    split = text.split('/')
    name = split[0]
    prof = ""
    place = ""
    if len(split) == 3:
        prof = split[1]
        place = split[2]
    else:
        match = re.search("(?<!\d)\d{3}(?!\d)", text)
        if match:
            split = re.split("(?<!\d)\d{3}(?!\d)", text)
            name = split[0]
            place = match[0] + split[1]
    return name.strip(), prof.strip(), place.strip(), type_by_color(t)


def filtration(item):
    if item.name == '':
        return 0
    return 1


def add_schedule(groups, path):
    print(path)
    book = xlrd.open_workbook(path)
    # get the first worksheet
    sheet = book.sheet_by_index(0)

    for colx in range(sheet.ncols):
        key = sheet.cell(colx=colx, rowx=KEYS_ROW).value.split('|')[0].strip()
        # check if cell is what we need
        if key in BANNED_EXPRESSIONS or key == '':
            continue

        if 'дист' in key.lower():
            prev_key = sheet.cell(colx=colx - 1, rowx=KEYS_ROW).value.split('|')[0].split('-')[0].strip()
            key = prev_key + ' ' + key + path.split('/')[-1].split('.')[0].split(' ')[0] + ' курс'

        rowx = TIME_START_ROW - 1
        lessons = []
        times = []
        day = 0
        lesson_number = 0
        for t1 in sheet.col_slice(colx=TIME_COL, start_rowx=TIME_START_ROW, end_rowx=sheet.nrows):
            rowx += 1
            t = t1.value.split('|')[0]
            if t == '':
                if lesson_number == 7:
                    day += 1
                    lesson_number = 0
                    times = []
                continue

            text = sheet.cell(colx=colx, rowx=rowx).value
            global identifier
            identifier += 1
            item = ScheduleItem(id=identifier)
            item.name, item.prof, item.place, item.type = handle_lesson(text)
            item.color = item.type
            item.day = day + 1
            try:
                item.startTime, item.endTime = list(map(insert_colon, t.split(' - ')))
            except:
                continue

            if t in times and lesson_number != 7:
                if lessons[-1].name == '' and item.name != '':
                    lessons[-1] = item
                    item.startTime = add_time(item.endTime, '-0:40')
                if lessons[-1].name != '' and item.name == '':
                    lessons[-1].endTime = add_time(item.startTime, '0:40')

                if lessons[-1].name != '' and item.name != '' and lessons[-1].name != item.name:
                    lessons[-1].endTime = add_time(lessons[-1].endTime, '-0:45')
                    lessons.append(item)
                    lessons[-1].startTime = add_time(lessons[-2].endTime, '0:05')

            if t not in times and lesson_number != 7:
                lesson_number += 1
                times.append(t)
                if len(lessons) >= 1 and lessons[-1].name == item.name:
                    lessons[-1].endTime = item.endTime
                else:
                    lessons.append(item)
        lessons = list(filter(filtration, lessons))
        groups[key] = [item.__dict__ for item in lessons]


def create_json(data):
    return json.dumps(data, ensure_ascii=False, indent=4)


def save_as_json(path, data):
    with open(path, 'w', encoding='utf8') as outfile:
        outfile.write(create_json(data))


def type_by_color(color):
    for key, value in color_map.items():
        if color in value:
            return key
    return "SEM"


def get_color(book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]
    return pattern_colour


def delete_unmerged_files():
    for f in os.listdir(UNMERGED_DATA):
        os.remove(os.path.join(UNMERGED_DATA, f))


def unmerge_excel_file(initial_path, final_path, add_color=True):
    # read merged cells for all sheets
    book = xlrd.open_workbook(initial_path, formatting_info=True)

    # open excel file and write
    excel = xlwt.Workbook()
    for rd_sheet in book.sheets():
        # for each sheet
        wt_sheet = excel.add_sheet(rd_sheet.name)
        written_cells = []

        # over write for merged cells
        for crange in rd_sheet.merged_cells:
            # for each merged_cell
            rlo, rhi, clo, chi = crange
            cell = rd_sheet.cell(rlo, clo)
            cell_value = str(cell.value)
            if cell.ctype == xlrd.XL_CELL_DATE:
                cell = xlrd.xldate.xldate_as_datetime(cell.value, datemode=0)
                cell_value = cell.strftime("%Y-%m-%d")

            color = get_color(book, rd_sheet, rlo, clo)
            if add_color:
                cell_value += '|' + str(color)
            for rowx in range(rlo, rhi):
                for colx in range(clo, chi):
                    wt_sheet.write(rowx, colx, cell_value)
                    written_cells.append((rowx, colx))

        # write all un-merged cells
        for r in range(0, rd_sheet.nrows):
            for c in range(0, rd_sheet.ncols):
                if (r, c) in written_cells:
                    continue

                cell = rd_sheet.cell(r, c)
                cell_value = str(cell.value)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    cell = xlrd.xldate.xldate_as_datetime(cell.value, datemode=0)
                    cell_value = cell.strftime("%Y-%m-%d")
                color = get_color(book, rd_sheet, r, c)
                if add_color:
                    cell_value += '|' + str(color)
                wt_sheet.write(r, c, cell_value)

    # save the un-merged excel file
    excel.save(final_path)


def parse_schedule():
    # Delete unmerged data
    delete_unmerged_files()
    # Create unmerged data
    for file in get_files(RAW_DATA):
        unmerge_excel_file(file, UNMERGED_DATA + file.split('/')[-1])
    groups = dict()
    for file in get_files(UNMERGED_DATA):
        add_schedule(groups, file)

    j = dict()
    j['version'] = VERSION
    j['timetable'] = groups
    print(j['timetable']['Б02-824'])
    print(j['timetable'].keys())
    return j


def main():
    start_time = time.time()
    save_as_json('schedule.json', parse_schedule())
    print('Built in ' + str((time.time() - start_time)) + ' seconds')


if __name__ == '__main__':
    main()
