#!bin/python3.6

import sys
import os
import datetime
import xlrd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


#An array containing the first saturday of the resort of each year.
SATURDAYS = [
    datetime.datetime.strptime("05/26/2018", "%m/%d/%Y"),
    datetime.datetime.strptime("05/25/2019", "%m/%d/%Y"),
    datetime.datetime.strptime("05/23/2020", "%m/%d/%Y"),
    datetime.datetime.strptime("05/29/2021", "%m/%d/%Y"),
    datetime.datetime.strptime("05/28/2022", "%m/%d/%Y"),
]
WEEKS = []

#Mappings
AGE_GROUPS = [
    {"Group": "IN", "rbName": "Infant", "cmName": ""},
    {"Group": "JT", "rbName": "Jr. Toddler", "cmName": "Junior Toddlers"},
    {"Group": "ST", "rbName": "Sr. Toddler", "cmName": "Senior Toddlers"},
    {"Group": "JM", "rbName": "Jr. Moppet", "cmName": "Junior Moppets"},
    {"Group": "SM", "rbName": "Sr. Moppet", "cmName": "Senior Moppets"},
    {"Group": "JR", "rbName": "Junior", "cmName": "Juniors"},
    {"Group": "SR", "rbName": "Senior", "cmName": "Seniors"},
    {"Group": "PT", "rbName": "Preteen", "cmName": "Preteens"},
    {"Group": "JT", "rbName": "Jr. Teen", "cmName": "Junior Teens"},
    {"Group": "ST", "rbName": "Sr. Teen", "cmName": "Senior Teens"},
    {"Group": "AD", "rbName": "Adult", "cmName": "Post-Group Teens"}
]


#Sets resort week numbers
def weeks():
    for i in SATURDAYS:
        for j in range(18):
            weekNo = j + 1
            fDOW = i + datetime.timedelta(days=7 * j)
            lDOW = i + datetime.timedelta(days=7 * j + 7)
            week = {
                "weekNo":weekNo,
                "fDOW":fDOW,
                "lDOW":lDOW,
            }
            WEEKS.append(week)
    return WEEKS


# Rectifies common date DQ issues seen within the exports
def cleanDate(date):
    if date == '':
        return date
    if str(date).find("/") == -1:
        result = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(date) - 2)
    else:
        if str(date) == "01/01/001":
            result = datetime.datetime.strptime("01/01/0001", "%m/%d/%Y")
        else:
            result = datetime.datetime.strptime(str(date), "%m/%d/%Y")
    result = result.strftime("%m/%d/%Y")
    return result


# Interprets a campminder export excel file and processes the data within it
def processCampMinder(fileName):
    data = []

    wb = xlrd.open_workbook(fileName)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    sheet.cell_value(0, 0)

    #Loop over all rows of the sheet
    for i in range(1, sheet.nrows):
        rowCells = []

        #Loop over all columns of the row
        for j in range(sheet.ncols):
            #Additional logic to handle dates
            if j == 3:
                rowCells.append(cleanDate(sheet.cell_value(i, j)))
            else:
                rowCells.append(sheet.cell_value(i, j))

        gender = rowCells[2]
        gender = gender[:1]
        row = {
            "firstName":                rowCells[0],
            "lastName":                 rowCells[1],
            "gender":                   gender,
            "birthDate":                rowCells[3],
            "schoolGrade":              rowCells[4],
            "kidsGroup":                rowCells[5],
            "arrival":                  " - ",
            "enrolledChildSessions":    rowCells[6],
            "guestAccommodation":    rowCells[7],
            "changes":                  "",
        }
        data.append(row.copy())
    return data

# Interprets a ResBill export excel file and processes the data within it
def processResBill(fileName):
    data = []
    colTitles = []
    rowChildren = []

    #Open resbill workbook path provided by main
    wb = xlrd.open_workbook(fileName)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    sheet.cell_value(0, 0)

    #Loop over all rows in the sheet
    for i in range(0, sheet.nrows):
        childId = 1
        rowCommonFields = {}
        childDict = {}

        #Loop over all columns within the row
        for j in range(sheet.ncols):
            #On first row set column titles array
            if i == 0:
                colTitles.append(sheet.cell_value(i, j))
            else:
                title = colTitles[j]

                # if j is onto next child eith find common family data or append it to child dict
                if not title.endswith(str(childId)):
                    if j < 2:
                        # if first child for the family set common fields for a family
                        if title == "arrival":
                            rowCommonFields.update({colTitles[j]: cleanDate(sheet.cell_value(i, j))})
                        else:
                            rowCommonFields.update({colTitles[j]: sheet.cell_value(i, j)})
                    else:
                        # If not first child for family, populate child info with common data
                        childDict.update({colTitles[0]: rowCommonFields["arrival"]})
                        childDict.update({colTitles[1]: rowCommonFields["accom"]})
                        if childDict["childlast"] != "":
                            rowChildren.append(childDict)
                        childDict = {}
                        childId += 1

                #If j is into child info columns set their data, else set the common fields across all children from the family
                if j >= 2 and title.endswith(str(childId)):
                    # Cut child numbers off column names eg. childlast2
                    title = title[:-1]
                    #Fix formatting of date if it is incorrect
                    if title == "dob":
                        childDict.update({title: cleanDate(sheet.cell_value(i, j))})
                    elif title == "agegrp" and sheet.cell_value(i, j) != '':
                        for k in AGE_GROUPS:
                            if sheet.cell_value(i, j) == k["rbName"]:
                                childDict.update({title: k["cmName"]})
                                break
                            else:
                                childDict.update({title: sheet.cell_value(i, j)})
                    else:
                        childDict.update({title: sheet.cell_value(i, j)})

    #Process pivoted child records
    for r in rowChildren:
        arrival = datetime.datetime.strptime(str(r["arrival"]), "%m/%d/%Y")
        # Map arrival dates to week numbers for comparison
        for i in WEEKS:
            if  i["fDOW"] <= arrival < i["lDOW"]:
                r["enrolledChildSessions"] = 'Week %d' % i["weekNo"]
                break
            elif i["weekNo"] == 1 and (i["fDOW"] - datetime.timedelta(days=7)) <= arrival < i["fDOW"]:
                r["enrolledChildSessions"] = "Week 1"
                break
            else:
                r["enrolledChildSessions"] = 'Unable to find match. Check date arrival date'
        row = {
            "firstName":                r["childfirst"],
            "lastName":                 r["childlast"],
            "gender":                   r["sex"],
            "birthDate":                r["dob"],
            "schoolGrade":              "-",
            "kidsGroup":                r["agegrp"],
            "arrival":                  r["arrival"],
            "enrolledChildSessions":    r["enrolledChildSessions"],
            "guestAccommodation":       r["accom"],
            "changes":                  "",
        }
        data.append(row.copy())
    return data


def compare(campMinder, resBill):
    i = 1
    matches = []
    cmOnly = []
    rbOnly = []
    data = []
    # for each camp minder record check all resBill records for a match
    for cmChild in campMinder:
        for rbChild in resBill:
            # Completes match on first name and last name fields
            if rbChild["firstName"].lower() == cmChild["firstName"].lower() \
                    and rbChild["lastName"].lower() == cmChild["lastName"].lower() \
                    and (rbChild["birthDate"] == cmChild["birthDate"] or rbChild["birthDate"] == '' or cmChild["birthDate"] == ''):
                rbChild.update({"matchKey": i, "src": "ResBill"})
                cmChild.update({"matchKey": i, "src": "CampMinder"})
                matches.append(rbChild)
                resBill.remove(rbChild)
        if "matchKey" in cmChild.keys() and cmChild["matchKey"] != -999:
            matches.append(cmChild)
            i += 1
        else:
            cmChild.update({"matchKey": -999, "src": "CampMinder"})
            cmOnly.append(cmChild)

    for rbChild in resBill:
        rbChild.update({"matchKey": -999, "src": "ResBill"})
        rbOnly.append(rbChild)

    matchKey = -1
    rowComp = {}
    for index, row in enumerate(matches):
        if row["matchKey"] == matchKey:
            if row["enrolledChildSessions"] != rowComp["enrolledChildSessions"]:
                row["enrolledChildSessions"] = "?(!)?" + row["enrolledChildSessions"]
                matches[index].update({"changes": matches[index]["changes"] + "Enrolled child session, "})
            if row["guestAccommodation"] != rowComp["guestAccommodation"]:
                row["guestAccommodation"] = "?(!)?" + row["guestAccommodation"]
                matches[index].update({"changes": matches[index]["changes"] + "Accommodation, "})
            if row["kidsGroup"] != rowComp["kidsGroup"]:
                row["kidsGroup"] = "?(!)?" + row["kidsGroup"]
                matches[index].update({"changes": matches[index]["changes"] + "Kids Group, "})
        else:
            rowComp = row
            matchKey = row["matchKey"]

    data.append(cmOnly)
    data.append(rbOnly)
    data.append(matches)
    return data

# Takes a list of visuals and writes them to a csv file
def writeToXls(data, sourceDirectory):
    cmOnly = data[0]
    rbOnly = data[1]
    matches = data[2]

    # Excel file output
    workbook = xlsxwriter.Workbook(sourceDirectory + "CM-RB-Compare-" + datetime.datetime.now().strftime("%Y%m%d-%I%M") + ".xlsx")
    sheets = {
        "worksheet2": workbook.add_worksheet("Changes"),
        "worksheet0": workbook.add_worksheet("CampMinder Only"),
        "worksheet1": workbook.add_worksheet("ResBill Only"),
    }
    bold = workbook.add_format({'bold': True})
    change = workbook.add_format({'bold': True, 'font_color': 'red'})
    topBorder = workbook.add_format({'top': 1})


    prevMatchNo = 0

    for i in range(3):
        for rowNum, rowData in enumerate(data[i]):
            for colNum, (colKey, colData) in enumerate(rowData.items()):
                if rowNum == 0:
                    sheets["worksheet" + str(i)].write(rowNum, colNum, colKey, bold)
                if str(colData).startswith("?(!)?"):
                    sheets["worksheet" + str(i)].write(rowNum + 1, colNum, colData[5:], change)
                else:
                    sheets["worksheet" + str(i)].write(rowNum + 1, colNum, colData)

                if colKey == "matchKey":
                    if colData != prevMatchNo:
                        sheets["worksheet" + str(i)].set_row(rowNum + 1, None, topBorder)
                        prevMatchNo = colData
                    else:
                        prevMatchNo = colData
            if rowNum >= len(data[i])-1:
                i += 1

    workbook.close()


def main(sourceDirectory):
    campMinder = []
    resBill = []
    # Checks and alters the formatting of the passed path string if required
    if sourceDirectory.endswith('"') and sourceDirectory.startsWith('"'):
        sourceDirectory = sourceDirectory[1:-1]
    if not sourceDirectory.endswith("/"):
        sourceDirectory = sourceDirectory + "/"

    # Iterate through the excel files within the folder path specified
    for fileName in os.listdir(sourceDirectory):
        # Check that the file is an excel file
        if not fileName.startswith("~") and not fileName.startswith("CM-RB")\
                and fileName.endswith(".xlsx") or fileName.endswith(".xls") or fileName.endswith(".xlsm") or fileName.endswith(".xlsb"):
            # Parse the entire excel file pathname to a variable
            pathName = os.path.join(sourceDirectory, fileName)

            print("Scanning: " + fileName)
            if "campminder" in fileName.lower():
                campMinder.extend(processCampMinder(pathName))
            elif "resbill" in fileName.lower():
                resBill.extend(processResBill(pathName))
            else:
                print("Info: File name must contain 'resbill' or 'campminder' to be fully scanned. File skipped")
            print("Finished: " + fileName + "\n")
        else:
            print("Info: " + fileName + " Is not a valid excel file. Ignored.\n")

    data = compare(campMinder, resBill)
    #writeToCsv(campMinder, sourceDirectory)
    writeToXls(data, sourceDirectory)
    messagebox.showinfo("CM-RB Compare", "Success! the file: CM-RB-Compare-" + datetime.datetime.now().strftime("%Y%m%d-%I%M") + ".xlsx has been generated and saved in the selected directory!")
    sys.exit("************** Process Complete **************")


############# Start ##############
# Check current python version and exit if not using python 3
if sys.version_info[0] < 3:
    print(sys.version_info)
    print(sys.executable)
    input("\n **You must be using python version >3 to use this script. Press Enter to exit...")
    messagebox.showerror("CM-RB Compare", "There has been an Error. Error: 1")
    sys.exit("************** Process Complete **************")

root = tk.Tk()
root.withdraw()
if messagebox.askokcancel("CM-RB Compare","Click OK to select location of files to compare or click cancel to exit!"):
    weeks()
    # Prompt for folder path and execute main function
    print("Please wait for folder select prompt to select the folder containing the excel files to compare...")
    directory = filedialog.askdirectory()
    if directory != '':
        main(directory)
    else:
        sys.exit("You have chosen to cancel")
else:
    sys.exit("You have chosen to cancel")