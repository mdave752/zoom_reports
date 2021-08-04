import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side

# Which file to parse
filename = "filename.csv"

f = open(filename, 'r', encoding='utf-8')
file = f.read().split('\n')
f.close()

new_list = list()
user_data = [0, 0, 0]
user_name, email = list(), list()
user = list()
for x in file[1:]:
    answ = x.split(',')
    if answ[0] == '':
        test_list = [meet_id, duration, topic, start_t, end_t, user]
        user_name = []
        email = []
        new_list.append(test_list)
        test_list = []
        user = []
        continue
    meet_id = answ[1]
    duration = answ[10]
    topic = answ[0]
    start_time = answ[8].split()
    right_time = datetime.datetime.strptime(
        start_time[0], "%m/%d/%Y").strftime("%d.%m.%Y")
    start_t = f"{right_time} {start_time[1]}"
    end_time = answ[9].split()
    eright_time = datetime.datetime.strptime(
        end_time[0], "%m/%d/%Y").strftime("%d.%m.%Y")
    end_t = f"{eright_time} {end_time[1]}"
    user_data = [answ[12], answ[13], answ[16]]
    user.append(user_data)


workbook = Workbook()
sheet = workbook.active

sheet.column_dimensions["A"].width = 23
sheet.column_dimensions["B"].width = 20
sheet.column_dimensions["C"].width = 28
sheet.column_dimensions["D"].width = 25


thin = Side(border_style="thin", color="303030")
medium = Side(border_style="medium", color="303030")
i = 0
for x in new_list:

    sheet[f"A{1+i}"] = 'Meeting ID : '
    sheet[f"A{1+i}"].border = Border(top=medium,
                                     left=medium, right=thin, bottom=thin)

    sheet[f"B{1+i}"] = x[0]  # Meet_id
    sheet[f'B{1+i}'].border = Border(top=medium,
                                     left=thin, right=thin, bottom=thin)

    sheet[f"A{2+i}"] = 'Time (minutes) : '
    sheet[f"A{2+i}"].border = Border(top=thin,
                                     left=medium, right=thin, bottom=thin)

    sheet[f"B{2+i}"] = x[1]  # duration
    sheet[f"B{2+i}"].border = Border(top=thin,
                                     left=thin, right=thin, bottom=thin)

    sheet[f"A{4+i}"] = 'Name'
    sheet[f"A{4+i}"].border = Border(top=thin, left=medium, bottom=thin)
    # for loop with  all name
    s = 5
    for nam in x[5]:
        sheet[f"C{s+i}"] = nam[1]
        sheet[f"A{s+i}"] = nam[0]
        sheet[f"D{s+i}"] = nam[2]
        s += 1

    sheet[f"C{1+i}"] = 'Topic : '
    sheet[f"C{1+i}"].border = Border(top=medium,
                                     left=thin, right=thin, bottom=thin)

    sheet[f"D{1+i}"] = x[2]  # Topic
    sheet[f"D{1+i}"].border = Border(top=medium,
                                     left=thin, right=medium, bottom=thin)

    sheet[f"C{2+i}"] = 'Start time : '
    sheet[f"C{2+i}"].border = Border(top=thin,
                                     left=thin, right=thin, bottom=thin)

    sheet[f"D{2+i}"] = x[3]  # start_t
    sheet[f"D{2+i}"].border = Border(top=thin,
                                     left=thin, right=medium, bottom=thin)

    sheet[f"C{4+i}"] = 'E-mail'
    sheet[f"C{4+i}"].border = Border(top=thin, bottom=thin)
    # For loop for user emails

    sheet[f"C{3+i}"] = 'End time : '
    sheet[f"C{3+i}"].border = Border(top=thin,
                                     left=thin, right=thin, bottom=medium)

    sheet[f"D{3+i}"] = x[4]  # End_t
    sheet[f"D{3+i}"].border = Border(top=thin,
                                     left=thin, right=medium, bottom=medium)

    # Boarders
    sheet[f'D{4 + i}'].border = Border(top=thin, right=medium, bottom=thin)
    sheet[f'A{3 + i}'].border = Border(top=thin,
                                       left=medium, right=thin, bottom=medium)
    sheet[f'B{3 + i}'].border = Border(top=thin,
                                       left=thin, right=thin, bottom=medium)
    sheet[f'B{4 + i}'].border = Border(top=thin,  bottom=thin)

    for m in range(5, len(x[5]) + 6):
        sheet[f'A{m + i}'].border = Border(left=medium)
        sheet[f'D{m + i}'].border = Border(right=medium)

    sheet[f'A{len(x[5]) + 5 + i}'].border = Border(left=medium,
                                                   bottom=medium)  # Bottom-Left
    # Bottom - mid
    sheet[f'B{len(x[5]) + 5 + i}'].border = Border(bottom=medium)
    # Bottom - mid
    sheet[f'C{len(x[5]) + 5 + i}'].border = Border(bottom=medium)
    sheet[f'D{len(x[5]) + 5 + i}'].border = Border(right=medium,
                                                   bottom=medium)  # Bottom-Right

    i += len(x[5]) + 6


workbook.save(filename='.\Zoom_test\report.xlsx')
