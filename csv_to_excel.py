from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.styles import Font
import datetime
import csv


class ZoomToXlsx:
    REGION_PREFERENCES = {
        # REGIONAL AND LANGAUGE SETTINGS, CHANGE VALUES TO DIFFERENT LANAGUE IF NEEDED
        "meetingId": "Meeting ID",
        "topic": "Topic" + " :",
        "meetingDuration": "Meeting Duration (minutes)" + " :",
        "startTime": "Start time",
        "endTime": "End time",
        "name": "Name",
        "email": "Email ",
        "timeOnline": "Duration Online",
        # Takes datetime date value, check strftime documentation for valid dates
        "dateToLocal": "%d.%m.%Y",
    }

    def __init__(self, input, output):
        self.input = input
        self.output = output
        self.workbook = Workbook()
        self.sheet = self.workbook.active

    def convertToDict(self):
        try:
            return csv.DictReader(open(self.input))
        except:
            raise Exception("File is not avilable or is of a wrong type")

    def filterDataByDate(self):
        filterByDate = dict()
        lastDate = None
        for record in self.convertToDict():
            date = record["Start Time"].split(" ")[0]
            if lastDate and date == lastDate:

                filterByDate[date].append(record)
            else:
                filterByDate[date] = [record]
            lastDate = date
        return filterByDate

    def filterUniqueUsers(self, type, values):
        filteredDict = dict()
        for user in values:
            name = user["Name (Original Name)"]
            email = user["User Email"]
            online = int(user["Duration (Minutes)"])
            user = str()
            if type == "email":
                user = email
            else:
                user = name
            if user in filteredDict.keys():
                filteredDict[user]["online"] += online
            else:
                filteredDict[user] = {
                    "name": name,
                    "email": email,
                    "online": online
                }
        return filteredDict

    def addUsersToTable(self, row, values, medium):
        sheet = self.sheet
        for _, value in values.items():
            sheet[f"A{row}"] = value["name"]
            sheet[f"B{row}"] = value["email"]
            sheet[f"C{row}"] = value["online"]
            sheet[f'A{row}'].border = Border(left=medium)
            sheet[f'D{row}'].border = Border(right=medium)
            row += 1
        return row

    def csvToXlsx(self, byNameOrEmail="name"):
        sheet = self.sheet

        # Table Width
        sheet.column_dimensions["A"].width = 25
        sheet.column_dimensions["B"].width = 30
        sheet.column_dimensions["C"].width = 20
        sheet.column_dimensions["D"].width = 35

        medium = Side(border_style="medium", color="303030")
        row = 1
        for _, meeting in self.filterDataByDate().items():

            #  HEADER
            sheet[f"A{row}"] = self.REGION_PREFERENCES["meetingId"]
            sheet[f"A{row}"].font = Font(bold=True)
            sheet[f"A{row}"].border = Border(top=medium, left=medium)
            sheet[f"B{row}"] = meeting[0]["Meeting ID"]  # Meet_id
            sheet[f'B{row}'].border = Border(top=medium)

            sheet[f"C{row}"] = self.REGION_PREFERENCES["topic"]
            sheet[f"C{row}"].font = Font(bold=True)
            sheet[f"C{row}"].border = Border(top=medium)
            sheet[f"D{row}"] = meeting[0]["\ufeffTopic"]  # Topic
            sheet[f"D{row}"].border = Border(top=medium, right=medium)

            row += 1
            sheet[f"A{row}"] = self.REGION_PREFERENCES["meetingDuration"]
            sheet[f"A{row}"].font = Font(bold=True)
            sheet[f"A{row}"].border = Border(left=medium)
            sheet[f"B{row}"] = meeting[0]["Duration (Minutes)"]
            sheet[f"C{row}"] = self.REGION_PREFERENCES["startTime"]
            sheet[f"C{row}"].font = Font(bold=True)
            sheet[f"D{row}"] = self.changeDateFormat(
                meeting[0]["Start Time"])
            sheet[f"D{row}"].border = Border(right=medium)

            row += 1
            sheet[f"C{row}"] = self.REGION_PREFERENCES["endTime"]
            sheet[f"C{row}"].font = Font(bold=True)
            sheet[f"D{row}"] = self.changeDateFormat(meeting[0]["End Time"])
            sheet[f"D{row}"].border = Border(right=medium)
            sheet[f'A{row}'].border = Border(left=medium)

            # Space before user list
            row += 1
            sheet[f'A{row}'].border = Border(left=medium)
            sheet[f'D{row}'].border = Border(right=medium)
            row += 1

            # User header
            sheet[f"A{row}"] = self.REGION_PREFERENCES["name"]
            sheet[f"A{row}"].font = Font(bold=True)
            sheet[f"B{row}"] = self.REGION_PREFERENCES["email"]
            sheet[f"B{row}"].font = Font(bold=True)
            sheet[f"C{row}"] = self.REGION_PREFERENCES["timeOnline"]
            sheet[f"C{row}"].font = Font(bold=True)
            sheet[f"A{row}"].border = Border(left=medium)
            sheet[f'D{row}'].border = Border(right=medium)

            # User Table
            row += 1
            filteredDict = self.filterUniqueUsers(byNameOrEmail, meeting)
            row = self.addUsersToTable(row, filteredDict, medium)
            sheet[f'A{row}'].border = Border(top=medium)
            sheet[f'B{row}'].border = Border(top=medium)
            sheet[f'C{row}'].border = Border(top=medium)
            sheet[f'D{row}'].border = Border(top=medium)
            row += 1

        self.workbook.save(filename=self.output)
        print(f"File written @ {self.output}")

    def changeDateFormat(self, dateString):
        time = dateString.split()
        rightDate = datetime.datetime.strptime(
            time[0], "%m/%d/%Y").strftime(self.REGION_PREFERENCES["dateToLocal"])
        return f"{rightDate} {time[1]}"


# RUN PROGRAM
# input = "./input_files/meetinglistdetails_00000000_00000000.csv"
# output = "./output_files/report.xlsx"
# test = ZoomToXlsx(input, output)
# test.csvToXlsx()
