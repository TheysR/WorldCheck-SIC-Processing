#######################################################################
# common functions/ classes for WordlCheck Triages
#
#######################################################################
from mysqlx import ColumnType
from openpyxl import Workbook, load_workbook
import sys

class ExcelHeader:
    def __init__(self, ws):
        self.col = {
        'AdditionalInfo' : 0,
        'Categories' : 0,
        'OfficialLists' : 0,
        'Reports' : 0,
        'Status' : 0,
        'Type' : 0
        }
    # worksheet is pointer to open ws
        if not ws:
            # raise error
            print('read_header() : No open or valid workheet')
            sys.exit()
        # read fist line
        c = 0
        for c_col in ws.columns:
            c += 1
            if (ws.cell(row=1, column=c).value) == "Categories":
                self.col['Categories'] = c
            if (ws.cell(row=1, column=c).value) == "AdditionalInformation":
                self.col['AdditionalInfo'] = c
            if (ws.cell(row=1, column=c).value) == "OfficialLists":
                self.col['OfficialLists'] = c
            if (ws.cell(row=1, column=c).value) == "REPORTS":
                self.col['Reports'] = c
            if (ws.cell(row=1, column=c).value) == "TYPE":
                self.col['Type'] = c
            if (ws.cell(row=1, column=c).value) == "STATUS" or (ws.cell(row=1, column=c).value) == "Status":
                self.col['Status'] = c
        if self.col['Status'] == 0:
            self.col['Status']= c + 1

    