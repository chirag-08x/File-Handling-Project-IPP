import pandas as pd

data = pd.read_excel("Results.xlsx")
values = data.values;

def filterData():
    for i in range(3,7):
        currentRow = values[i]
        datePlate = currentRow[3].date()

        name, regNo, program, total, result = currentRow[0], currentRow[1], currentRow[2], currentRow[-2], currentRow[-1]
        day, month, year = datePlate.strftime('%d'), datePlate.strftime('%b'), datePlate.strftime('%y')

        rowData = f"Mr. {name} bearing registration number {regNo}, enrolled in program+ {program} has appeared in final exams on {day}-{month}-{year}. Student has successfully qualified the following courses: Microsoft Excel, VBA, SQL, Power BI, and have scored a grand total of {total} marks."

        if result == "F":
            rowData += f"But he failed the {program} course. He may have to reappear in this exam." 
        
        file = open(f"{name}.txt", "w")
        file.write(rowData)
        file.close()

filterData()
