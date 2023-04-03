#!/usr/local/bin/python3

import openpyxl
import os
import sys
import time
import tkinter as tk

# important places in the sheet
CHECKPOINT_NUMBER_LOCATION = 'D4'
CHECKPOINT_NAME_LOCATION = 'F4'
CHECKPOINT_TIME_TARGET_ROW = 7






    
checkpoint_number ='1'
checkpoint_name = 'scotts checkpoint'
checkpoint_time_target = 31
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
    sheet[cell] = '00:' + str(checkpoint_time_target + i)
    sheet[score_cell] = max_points - i
    sheet[team_cell] = i + 1

resultsName = os.path.join(save_path, "Results.xlsx")
wb.save(resultsName)




