import os
from openpyxl import load_workbook
from Libraries.Styles import *

class Format():

    def __init__(self,file):
        self.file = file

    def formatTable(self):
        filePath = os.path.join(os.getcwd(), self.file)
        UserPass = load_workbook(filePath)

        Sheet1 = UserPass["login"]
        Sheet3 = UserPass["test_login_results"]

        number = 1
        for value in Sheet3.iter_rows(min_row=1, min_col=1, max_col=4, values_only=True):
            Sheet3["A" + str(number)] = ""
            Sheet3["B" + str(number)] = ""
            Sheet3["C" + str(number)] = ""
            Sheet3["D" + str(number)] = ""
            number += 1

        Sheet3.column_dimensions['A'].width = 30.0
        Sheet3.column_dimensions['B'].width = 30.0
        Sheet3.column_dimensions['C'].width = 20.0
        Sheet3.column_dimensions['D'].width = 20.0

        Sheet3["A1"].alignment = position_left()
        Sheet3["A1"].font = styleBold()
        Sheet3["A1"] = "username"
        Sheet3["B1"].alignment = position_left()
        Sheet3["B1"].font = styleBold()
        Sheet3["B1"] = "password"
        Sheet3["C1"].alignment = position_center()
        Sheet3["C1"].font = styleBold()
        Sheet3["C1"] = "correctness"
        Sheet3["D1"].font = styleBold()
        Sheet3["D1"].alignment = position_center()
        Sheet3["D1"] = "test results"

        number = 2

        for value in Sheet1.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
            username = value[0]
            password = value[1]
            correctness = value[2]

            Sheet3["A" + str(number)].alignment = position_left()
            Sheet3["A" + str(number)] = username
            Sheet3["B" + str(number)].alignment = position_left()
            Sheet3["B" + str(number)] = password
            Sheet3["C" + str(number)].alignment = position_center()
            Sheet3["C" + str(number)] = correctness

            number+=1

        UserPass.save(self.file)

