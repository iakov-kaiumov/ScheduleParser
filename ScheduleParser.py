import xlrd
import os
import json
import xlwt
import time


RAW_DATA = 'Schedule/raw_data_2019/'
UNMERGED_DATA = 'Schedule/unmerged_data/'
VERSION = '1.5'

TIME_START_ROW = 5
TIME_STOP_ROW = 15
TIME_COL = 1

KEYS_ROW = 4

START_ROW = 5
START_COL = 1

LEC = str((255, 153, 204))
LAB = str((255, 255, 153))
FRE = str((204, 255, 204))
SEM = str((204, 255, 255))

BANNED_EXPRESSIONS = ['Дни', 'Часы', 'Б', 'С', 'М']


class ScheduleItem:
    def __init__(self, name="", prof="", place="", day=0, type="FRE", startTime="9:00", endTime="10:25", notes=""):
        self.name = name
        self.prof = prof
        self.place = place
        self.day = day
        self.type = type
        self.startTime = startTime
        self.endTime = endTime
        self.notes = notes


def get_files(scr_dir):
    return [scr_dir + dr for dr in os.listdir(scr_dir)]


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
    return name.strip(), prof.strip(), place.strip(), type_by_color(t)


def filtration(item):
    if item.name == '':
        return 0
    return 1


def add_schedule(groups, path):
    book = xlrd.open_workbook(path)
    # get the first worksheet
    sheet = book.sheet_by_index(0)

    for colx in range(sheet.ncols):
        key = sheet.cell(colx=colx, rowx=KEYS_ROW).value.split('|')[0]
        # check if cell is what we need
        if key in BANNED_EXPRESSIONS or key == '':
            continue

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
            item = ScheduleItem()
            item.name, item.prof, item.place, item.type = handle_lesson(text)
            item.day = day + 1
            item.startTime, item.endTime = list(map(insert_colon, t.split(' - ')))

            if t in times and lesson_number != 7:
                if lessons[-1].name == '' and item.name != '':
                    lessons[-1] = item
                    item.startTime = add_time(item.endTime, '-0:40')
                if lessons[-1].name != '' and item.name == '':
                    lessons[-1].endTime = add_time(item.startTime, '0:40')
            if t not in times and lesson_number != 7:
                lesson_number += 1
                times.append(t)
                if len(lessons) > 1 and lessons[-1].name == item.name:
                    lessons[-1].endTime = item.endTime
                else:
                    lessons.append(item)
        lessons = list(filter(filtration, lessons))
        groups[key] = [item.__dict__ for item in lessons]


def create_json(data):
    return json.dumps(data, ensure_ascii=False)


def save_as_json(path, data):
    with open(path, 'w', encoding='utf8') as outfile:
        outfile.write(create_json(data))


def type_by_color(color):
    if color == SEM:
        return "SEM"
    elif color == LEC:
        return "LEC"
    elif color == LAB:
        return "LAB"
    elif color == FRE:
        return "FRE"
    return " "


def get_color(book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]
    return pattern_colour


def unmerge_excel_file(initial_path, final_path):
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
            cell_value = rd_sheet.cell(rlo, clo).value
            color = get_color(book, rd_sheet, rlo, clo)
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
                cell_value = str(rd_sheet.cell(r, c).value)
                color = get_color(book, rd_sheet, r, c)
                cell_value += '|' + str(color)
                wt_sheet.write(r, c, cell_value)

    # save the un-merged excel file
    excel.save(final_path)


def parse_schedule():
    # Create unmerged data
    # for file in get_files(RAW_DATA):
    #     unmerge_excel_file(file, UNMERGED_DATA + file.split('/')[-1])
    groups = dict()
    for file in get_files(UNMERGED_DATA):
        add_schedule(groups, file)

    print(groups['Б02-827'])
    j = dict()
    j['version'] = VERSION
    j['timetable'] = groups
    return j


def main():
    start_time = time.time()
    save_as_json('schedule.json', parse_schedule())
    print('Built in ' + str((time.time() - start_time) * 1000))


if __name__ == '__main__':
    main()
