from openpyxl import Workbook
workbook=Workbook()
sheet=workbook.active
def validateTime(startTime, endTime):
    if startTime.isdigit()==False or endTime.isdigit()==False:
        print("Please enter valid numeric values for time.")
        return False
    elif int(startTime)<0 or int(startTime)>23 or int(endTime)<0 or int(endTime)>23 or int(startTime)>=int(endTime):
        print("Please enter time values between 0 and 23.")
        return False
    return True

def validateDay(day):
    days=["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    if day.lower() not in days:
        print("Please enter a valid day of the week.")
        return False
    return True

while (True):
    for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=6), start=1):
        for j, cell in enumerate(row, start=1):
            sheet.cell(row=cell.row, column=cell.column).value = f"R{str(i)}C{str(j)}"
    course=input("Write the course name: ")
    startTime=input("Write the start time (hour in 24h format): ")
    endTime=input("Write the end time (hour in 24h format): ")
    if validateTime(startTime, endTime)==False:
        continue
    day=input("Write the day of the week: ")
    if validateDay(day)==False:
        continue
    nameFile=input("Write the file name: ")

    choose=input("Do you want to add another course? (yes/no): ")
    workbook.save(f"{nameFile}.xlsx")
    if choose.lower() == "no":
        break
    