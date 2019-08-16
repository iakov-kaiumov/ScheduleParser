import json
from ics import Calendar, Event
import time
import arrow


OUTPUT = 'Schedule/IOSOutput/'
WEEKS = 17


def save_group(key, group):
    calendar = Calendar()
    for lesson in group:
        spl = lesson['startTime'].split(':')
        s_time = "T{:02d}".format(int(spl[0])) + ':' + spl[1] + ':00+03:00'
        s_time = arrow.get('2019020' + str(3 + lesson['day']) + s_time)
        spl = lesson['endTime'].split(':')
        e_time = "T{:02d}".format(int(spl[0])) + ':' + spl[1] + ':00+03:00'
        e_time = arrow.get('2019020' + str(3 + lesson['day']) + e_time)
        for i in range(WEEKS):
            event = Event()
            event.name = lesson['name']
            event.description = lesson['prof']
            event.location = lesson['place']
            event.begin = s_time
            event.end = e_time
            calendar.events.add(event)

            s_time = s_time.replace(weeks=+1)
            e_time = e_time.replace(weeks=+1)

    with open(OUTPUT + key.replace('\n', '1') + '.ics', 'w', encoding='utf8') as file:
        file.writelines(calendar)


def main():
    print('Starting...')
    start_time = time.time()
    file = open('schedule.json', 'r', encoding='utf8')
    data = json.load(file)
    file.close()
    timetable = data['timetable']
    # save_group('Б02-827', timetable['Б02-824'])
    for key, group in timetable.items():
        # Creating CSV file
        save_group(key, group)

    print('Built in ' + str((time.time() - start_time)))


if __name__ == '__main__':
    main()
