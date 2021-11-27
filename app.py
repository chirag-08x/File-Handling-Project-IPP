import pandas as pd

data = pd.read_excel("Results.xlsx")
values = data.values;

def filterData():
    for i in range(3,7):
        currentRow = values[i]
        datePlate = currentRow[3].date()

        rowData = f"Mr. {currentRow[0]} bearing registration number {currentRow[1]}, enrolled in program+ {currentRow[2]} has appeared in final exams on {datePlate.strftime('%d')}-{datePlate.strftime('%b')}-{datePlate.strftime('%y')}. Student has successfully qualified the following courses: Microsoft Excel, VBA, SQL, Power BI, and have scored a grand total of {currentRow[-2]} marks."

        if currentRow[-1] == "F":
            rowData += f"But he failed the {currentRow[2]} course. He may have to reappear in this exam." 
        
        name = currentRow[0]
        file = open(f"{name}.txt", "w")
        file.write(rowData)
        file.close()

filterData()
