from openpyxl import Workbook
from src.CRUDcontroller import *
from openpyxl import load_workbook
import os


def init() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "MappingTraineeId"
    sheet["A1"] = "CourseId"
    sheet["B1"] = "TraineeId"
    workbook.create_sheet("Trainees")
    sheet = workbook["Trainees"]
    sheet["A1"], sheet["B1"] = "id", "Name"
    sheet["C1"], sheet["D1"], sheet["E1"] = "Course", "Background/degree", "WorkExperience"
    workbook.create_sheet("CourseDetails")
    sheet = workbook["CourseDetails"]
    sheet["A1"], sheet["B1"] = "CourseId", "Description"
    workbook.create_sheet("Trainers")
    sheet = workbook["Trainers"]
    sheet["A1"], sheet["B1"] = "id", "name"
    sheet["C1"], sheet["D1"] = "email", "phone number"
    workbook.create_sheet("MappingCourseTrainer")
    sheet = workbook["MappingCourseTrainer"]
    sheet["A1"], sheet["B1"] = "Trainerid" , "CourseId"
    workbook.create_sheet("Managers")
    sheet = workbook["Managers"]
    sheet["A1"], sheet["B1"] = "id", "name"
    sheet["C1"], sheet["D1"] = "email", "phone number"
    workbook.save(filename="course_manager.xlsx")

def setAttendance(sheet: Worksheet, absent: list):
    print(sheet["A1"])
    for row in sheet.rows:
        if row[0].value in absent:
            row[2].value = "A"
if not os.path.isfile("course_manager.xlsx"):
    init()
workbook = load_workbook(filename="course_manager.xlsx")
sheet = workbook["Trainees"]
print(sheet["A1"].row)
choice = None
controller = CRUDcontroller(sheet)
while choice != "q":
    choice = input("""
    1 Show sheets
    2 Set active sheet
    3 Create entry
    4 Update entry
    5 Delete entry
    6 Add session
    q Quit
    """)
    match(choice):
        case "1":
            print(", ".join(workbook.sheetnames))
        case "2":
            try:
                sheet = input("Enter name of sheet ")
                sheet = workbook[sheet]
                controller.data = sheet
            except KeyError:
                print("No such sheet")
                continue
        case "3":
            newId = input("Enter the id ")
            if controller.find(newId) != 0:
                print("id must be unique")
                continue
            controller.add(newId, [input(f"Enter the new value for {cell.value} ") for cell in sheet[1][1:]])
        case "4":
            row = controller.find(input("Enter the id "))
            if not row:
                print("id not found")
                continue
            controller.update(row, [input(f"Enter the new value for {cell.value} ") for cell in sheet[1][1:]])
        case "5":
            controller.delete(input("Enter the id "))
        case "6":
            date = input("Enter date ")
            workbook.create_sheet(date)
            attendance = workbook[date]
            attendance["A1"], attendance["B1"], attendance["C1"] = "id", "name", "present/absent"
            trainees = workbook["Trainees"]
            rows = trainees.rows
            next(rows)
            for row in rows:
                attendance.append([row[0].value, row[1].value, "P"])
            setAttendance(attendance, [input("Enter id of absent trainee ") for i in range(int(input("Enter number of absent students ")))])

            
        case "q":
            break
workbook.save(filename="course_manager.xlsx")