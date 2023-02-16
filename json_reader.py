from datetime import datetime
import json

months_dict = {
    'January': 1,
    'February': 2,
    'March': 3,
    'April': 4,
    'May': 5,
    'June': 6,
    'July': 7,
    'August': 8,
    'September': 9,
    'October': 10,
    'November': 11,
    'December': 12
}


def str_to_datetime(day_of_month_str, time_str):
    time_str += ":00"
    time_str = time_str.replace(" ", "")
    month_str = day_of_month_str.split(' ')[0]
    date_str = str(datetime.now().year) + " " + day_of_month_str.replace(month_str, str(months_dict[month_str]))
    date_str = date_str.replace(' ', '-')
    datetime_str = date_str + " " + time_str
    return datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')


def format_meeting_details():
    meeting_list = []
    with open('meetings.json', 'r') as f:
        meetings = json.load(f)
        for _ in meetings:
            print(_['title'])
            # print(_['full title'])
            title = _['title'].split('from')[0].strip()
            print("formatted title: " + title)
            time_start = ' '.join(_['title'].split('from')[1].split('to')[0].split(' ')[4:])
            time_end = _['title'].split('to')[1].split(' ')
            if len(time_end) == 3:
                time_end = ' '.join(time_end[1:])
            else:
                time_end = ' '.join(time_end[-2:])

            day_of_week = _['title'].split('from')[1].split('to')[0].split(',')[0].strip()
            day_of_month_start = _['title'].split('from')[1].split('to')[0].split(',')[0].split(" ")[1].strip() + " " + \
                                 _['title'].split('from')[1].split('to')[0].split(',')[0].split(" ")[2].strip()

            if "AM" in time_start:
                time_start = time_start.replace("AM", "")
                time_start.strip()
            if "PM" in time_start:
                time_start = time_start.replace("PM", "")
                time_start.strip()
                hour = int(time_start.split(":")[0])
                if hour != 12:
                    hour += 12
                else:
                    hour = 0
                time_start = str(hour) + ":" + time_start.split(":")[1]
            if "AM" in time_end:
                time_end = time_end.replace("AM", "")
                time_end.strip()
            if "PM" in time_end:
                time_end = time_end.replace("PM", "")
                time_end.strip()
                hour = int(time_end.split(":")[0])
                if hour != 12:
                    hour += 12
                else:
                    hour = 0
                time_end = str(hour) + ":" + time_end.split(":")[1]

            time_start = str_to_datetime(day_of_month_start, time_start)
            time_end = str_to_datetime(day_of_month_start, time_end)
            overdue = datetime.now() > time_start
            # if overdue:
            #     delay_offset = 0
            # else:
            #     delay_offset =  time_start - datetime.now()
            #     # datetime to seconds
            #     delay_offset = int(delay_offset.total_seconds())

            # meeting_length = time_end - time_start

            # print(_['id'])
            # print("Title: " + title)
            # print("Time Start: " + str(time_start))
            # print("Time End: " + str(time_end))
            # print("Day of Week: " + day_of_week)
            # print("Day of Month: " + day_of_month)
            # print("Overdue: " + str(overdue))
            # # print("Delay Offset: " + str(delay_offset))
            # print("Meeting Length: " + str(meeting_length))
            # print("#############################################")
            meeting = {'id': _['id'], 'title': title, 'time_start': str(time_start), 'time_end': str(time_end)}
            meeting_list.append(meeting)
    # write to json file
    with open('meetings_formatted.json', 'w') as f:
        json.dump(meeting_list, f)


def get_list_from_json():
    with open('meetings_formatted.json', 'r') as f:
        meetings = json.load(f)
    for _ in meetings:
        _['time_start'] = datetime.strptime(_['time_start'], '%Y-%m-%d %H:%M:%S')
        _['time_end'] = datetime.strptime(_['time_end'], '%Y-%m-%d %H:%M:%S')
        # _['attended'] = False
    return meetings

if __name__ == '__main__':
    format_meeting_details()