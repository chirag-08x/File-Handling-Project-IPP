import pandas as pd

data = pd.read_excel("Results.xlsx")
values = data.values;

def filterData():
    for i in range(3,7):
        currentRow = values[i]
        datePlate = currentRow[3].date()

        stdinf = {
            "name":currentRow[0],
            "regNo":currentRow[1],
            "program":currentRow[2],
            "total" :currentRow[-2],
            "result":currentRow[-1],
            "day":datePlate.strftime('%d'),
            "month":datePlate.strftime('%b'),
            "year":datePlate.strftime('%y'),
            "courses": list(filter(lambda x:type(x)!=float, values[1])),
            "hasFailed":[]
        }
        
        rowData = f"""Mr. {stdinf["name"]} bearing registration number {stdinf["regNo"]}, enrolled in program+ {stdinf["program"]} has appeared in final exams on {stdinf["day"]}-{stdinf["month"]}-{stdinf["year"]}. Student has successfully qualified the following courses: {", ".join(stdinf["courses"])}, and have scored a grand total of {stdinf["total"]} marks."""

        c = 0
        if stdinf["result"] == "F":
            for i in range(5, len(currentRow)-1, 2):
                if currentRow[i] == "Fail":
                    stdinf["hasFailed"].append(stdinf["courses"][c])
                c += 1 
            rowData += f"""But he failed the {", ".join(stdinf["hasFailed"])} course. He may have to reappear in this exam."""

        file = open(f"{stdinf['name']}.txt", "w")
        file.write(rowData)
        file.close()

filterData()