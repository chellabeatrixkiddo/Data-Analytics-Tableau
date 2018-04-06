#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr  5 16:03:24 2018

@author: Admin
"""

import csv
import xlsxwriter
import os

csv_folder = '/Users/Admin/Downloads/NIRF/Engineering/required'

workbook = xlsxwriter.Workbook('Data.xlsx')
sheet1 = workbook.add_worksheet("Sheet 1")
sheet2 = workbook.add_worksheet("Sheet 2")
sheet3 = workbook.add_worksheet("Sheet 3")
sheet4 = workbook.add_worksheet("Sheet 4")
sheet5 = workbook.add_worksheet("Sheet 5")
sheet6 = workbook.add_worksheet("Sheet 6")
sheet7 = workbook.add_worksheet("Sheet 7")
sheet8 = workbook.add_worksheet("Sheet 8")
sheet9 = workbook.add_worksheet("Sheet 9")
sheet10 = workbook.add_worksheet("Sheet 10")
sheet11 = workbook.add_worksheet("Sheet 11")
sheet12 = workbook.add_worksheet("Sheet 12")
sheet13 = workbook.add_worksheet("Sheet 13")

rowtracker = [0,0,0,0,0,0,0,0,0,0,0,0,0]
coltracker = [0,0,0,0,0,0,0,0,0,0,0,0,0]
# Write the Header values
csv_path = os.path.join(csv_folder, "IR-1-E-E-U-0014.csv")

f = open(csv_path, 'r')
reader = csv.reader(f)
for rowi, row in enumerate(reader):
    for coli, value in enumerate(row):
        #[1,3,6,9,12,15,18,21,25,28,32,36,38]:
        if rowi == 0:
            sheet1.write(rowtracker[0], coltracker[0], value)
            coltracker[0] += 1
        elif rowi == 2:
            sheet2.write(rowtracker[1], coltracker[1], value)
            coltracker[1] += 1
        elif rowi == 5:
            sheet3.write(rowtracker[2], coltracker[2], value)
            coltracker[2] += 1
        elif rowi == 8:
            sheet4.write(rowtracker[3], coltracker[3], value)
            coltracker[3] += 1
        elif rowi == 11:
            sheet5.write(rowtracker[4], coltracker[4], value)
            coltracker[4] += 1
        elif rowi == 14:
            sheet6.write(rowtracker[5], coltracker[5], value)
            coltracker[5] += 1
        elif rowi == 17:
            sheet7.write(rowtracker[6], coltracker[6], value)
            coltracker[6] += 1
        elif rowi == 20:
            sheet8.write(rowtracker[7], coltracker[7], value)
            coltracker[7] += 1
        elif rowi == 24:
            sheet9.write(rowtracker[8], coltracker[8], value)
            coltracker[8] += 1
        elif rowi == 27:
            sheet10.write(rowtracker[9], coltracker[9], value)
            coltracker[9] += 1
        elif rowi == 31:
            sheet11.write(rowtracker[10], coltracker[10], value)
            coltracker[10] += 1
        elif rowi == 35:
            sheet12.write(rowtracker[11], coltracker[11], value)
            coltracker[11] += 1
        elif rowi == 37:
            sheet13.write(rowtracker[12], coltracker[12], value)
            coltracker[12] += 1

# Add the institute id as the last column in each sheet
sheet1.write(rowtracker[0], coltracker[0], "Institute ID")
sheet2.write(rowtracker[1], coltracker[1], "Institute ID")
sheet3.write(rowtracker[2], coltracker[2], "Institute ID")
sheet4.write(rowtracker[3], coltracker[3], "Institute ID")
sheet5.write(rowtracker[4], coltracker[4], "Institute ID")
sheet6.write(rowtracker[5], coltracker[5], "Institute ID")
sheet7.write(rowtracker[6], coltracker[6], "Institute ID")
sheet8.write(rowtracker[7], coltracker[7], "Institute ID")
sheet9.write(rowtracker[8], coltracker[8], "Institute ID")
sheet10.write(rowtracker[9], coltracker[9], "Institute ID")
sheet11.write(rowtracker[10], coltracker[10], "Institute ID")
sheet12.write(rowtracker[11], coltracker[11], "Institute ID")
sheet13.write(rowtracker[12], coltracker[12], "Institute ID")

for i in range(13):
    rowtracker[i] = 1

f.close()
'''
# Write the values
for filename in os.listdir(csv_folder):
    if ".csv" in filename:
        csv_path = os.path.join(csv_folder, filename)
    else:
        continue
    f = open(csv_path, 'r')
    reader = csv.reader(f)
    print(csv_path)
    instid = filename[:-4]
    for rowi, row in enumerate(reader):
        for coli, value in enumerate(row):
            if rowi > 0 and rowi < 2:
                sheet1.write(rowtracker[0], coli, value)
                if (coltracker[0] - coli) == 1:
                    sheet1.write(rowtracker[0], coltracker[0], instid)
                    rowtracker[0] += 1
            elif rowi > 2 and rowi < 5:
                sheet2.write(rowtracker[1], coli, value)
                if (coltracker[1] - coli) == 1:
                    sheet2.write(rowtracker[1], coltracker[1], instid)
                    rowtracker[1] += 1
            elif rowi > 5 and rowi < 8:
                sheet3.write(rowtracker[2], coli, value)
                if (coltracker[2] - coli) == 1:
                    sheet3.write(rowtracker[2], coltracker[2], instid)
                    rowtracker[2] += 1
            elif rowi > 8 and rowi < 11:
                sheet4.write(rowtracker[3], coli, value)
                if (coltracker[3] - coli) == 1:
                    sheet4.write(rowtracker[3], coltracker[3], instid)
                    rowtracker[3] += 1
            elif rowi > 11 and rowi < 14:
                sheet5.write(rowtracker[4], coli, value)
                if (coltracker[4] - coli) == 1:
                    sheet5.write(rowtracker[4], coltracker[4], instid)
                    rowtracker[4] += 1
            elif rowi > 14 and rowi < 17:
                sheet6.write(rowtracker[5], coli, value)
                if (coltracker[5] - coli) == 1:
                    sheet6.write(rowtracker[5], coltracker[5], instid)
                    rowtracker[5] += 1
            elif rowi > 17 and rowi < 20:
                sheet7.write(rowtracker[6], coli, value)
                if (coltracker[6] - coli) == 1:
                    sheet7.write(rowtracker[6], coltracker[6], instid)
                    rowtracker[6] += 1
            elif rowi > 20 and rowi < 24:
                sheet8.write(rowtracker[7], coli, value)
                if (coltracker[7] - coli) == 1:
                    sheet8.write(rowtracker[7], coltracker[7], instid)
                    rowtracker[7] += 1
            elif rowi > 24 and rowi < 27:
                sheet9.write(rowtracker[8], coli, value)
                if (coltracker[8] - coli) == 1:
                    sheet9.write(rowtracker[8], coltracker[8], instid)
                    rowtracker[8] += 1
            elif rowi > 27 and rowi < 31:
                sheet10.write(rowtracker[9], coli, value)
                if (coltracker[9] - coli) == 1:
                    sheet10.write(rowtracker[9], coltracker[9], instid)
                    rowtracker[9] += 1
            elif rowi > 31 and rowi < 35:
                sheet11.write(rowtracker[10], coli, value)
                if (coltracker[10] - coli) == 1:
                    sheet11.write(rowtracker[10], coltracker[10], instid)
                    rowtracker[10] += 1
            elif rowi > 35 and rowi < 37:
                sheet12.write(rowtracker[11], coli, value)
                if (coltracker[11] - coli) == 1:
                    sheet12.write(rowtracker[11], coltracker[11], instid)
                    rowtracker[11] += 1
            elif rowi > 37:
                sheet13.write(rowtracker[12], coli, value)
                if (coltracker[12] - coli) == 1:
                    sheet13.write(rowtracker[12], coltracker[12], instid)
                    rowtracker[12] += 1
'''
workbook.close()