#!/usr/local/bin/python3

import tkinter as tk
import openpyxl
import os

# important places in the sheet
CHECKPOINT_NUMBER_LOCATION = 'D4'
CHECKPOINT_NAME_LOCATION = 'F4'
CHECKPOINT_TIME_TARGET_ROW = 7

def writeSpreadSheet(filename,checkpoint_number, checkpoint_name, checkpoint_time_target, ):

    max_points = 20

    path = 'Templates'
    save_path = 'Out'
    sheetName = os.path.join(path, "Template.xlsx")
    wb = openpyxl.load_workbook(sheetName)
    sheet = wb['Sheet1']
    sheet[CHECKPOINT_NUMBER_LOCATION] = checkpoint_number
    sheet[CHECKPOINT_NAME_LOCATION] = checkpoint_name
    # for loop over 20 integers
    row = CHECKPOINT_TIME_TARGET_ROW
    for i in range(0, 20):
        row = row + 1
        cell = 'B' + str(row)
        score_cell = 'C' + str(row)
        team_cell = 'D' + str(row)
        sheet[cell] = '00:' + str(int(checkpoint_time_target) + i)
        sheet[score_cell] = max_points - i
        sheet[team_cell] = i + 1

    resultsName = os.path.join(save_path, filename + ".xlsx")
    wb.save(resultsName)

window = tk.Tk()
window.title("Downsman Results Generator")

# Create a frame for the text entry box
content = tk.Frame(window)

def newLabel(text):
    label = tk.Label(content, text=text)
    name = tk.Entry(content)
    return [label, name]

def setGrid(row, label):
    label[0].grid(column=0, row=row, columnspan=1)
    label[1].grid(column=1, row=row, columnspan=2)



fileName = newLabel("Enter File Name:")
checkPointName = newLabel("Enter Checkpoint Name:")
checkPointNumber = newLabel("Enter Checkpoint Number:")
timeTarget = newLabel("Enter Time Target:")

def okPressed():
    
    f = fileName[1].get()
    fileName[1].delete(0, tk.END)
    cpName = checkPointName[1].get()
    checkPointName[1].delete(0, tk.END)
    cpNumber = checkPointNumber[1].get()
    checkPointNumber[1].delete(0, tk.END)
    time = timeTarget[1].get()
    timeTarget[1].delete(0, tk.END)
    writeSpreadSheet(f, cpNumber, cpName, time)
    print("OK")

ok = tk.Button(content, text="Okay", command=okPressed)
cancel = tk.Button(content, text="Cancel", command=window.destroy)

content.grid(column=0, row=0)

row = 0
setGrid(row, fileName)
setGrid(row + 1, checkPointName)
setGrid(row + 2, checkPointNumber)
setGrid(row + 3, timeTarget)


ok.grid(column=1, row=row + 4)
cancel.grid(column=2, row=row + 4)

window.mainloop()



