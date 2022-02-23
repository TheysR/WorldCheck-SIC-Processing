#######################################################################
# common functions/ classes for WordlCheck Triages
# ExcelHeader: readdsadrs header and assigns column numbers
#######################################################################
from sre_compile import isstring
from openpyxl import Workbook
import sys, re
##########################################################################
class ExcelHeader:
    """reads header row from excel worksheet and assigns column numbers."""
# called with ws (worksheet class)
# Init: assigns column numbers to the headers of a SIC file, using
# label as index. Adds Remarks label in first row at the end.
# properties:
# col (type dict): col['Header_Name'] containes column number
# last_col: contains last valid column
# values are assigned at initialisation
# ver 1.0 : read headers and add status and remarks if not found
# ver 1.1 : added last_column property
# ver 1.2 : added Addcolumn method (not tested)
# ver 1.3 : added ProfilType to TYPE match in headers (found in disqualified)
########################################################################## 
    def __init__(self, ws):
        self.col = {
            'AdditionalInfo' : 0,
            'Bio' : 0,
            'Categories' : 0,
            'OfficialLists' : 0,
            'Remarks' : 0,
            'Reports' : 0,
            'Status' : 0,
            'Type' : 0
            }
        # worksheet is pointer to open ws
        if not ws:
            #  error
            print('ExcelHeader() : ws: No open or valid workheet')
            sys.exit()
        # read first line
        c = 0
        for c_col in ws.columns:
            c += 1
            match (ws.cell(row=1, column=c).value):
                case "Categories":
                    self.col['Categories'] = c
                case "BIO":
                    self.col['Bio'] = c
                case "AdditionalInformation":
                    self.col['AdditionalInfo'] = c
                case "OfficialLists":
                    self.col['OfficialLists'] = c
                case "REPORTS":
                    self.col['Reports'] = c
                case "TYPE" | "ProfileType":
                    self.col['Type'] = c
                case "STATUS" | "Status":
                    self.col['Status'] = c
        # report and additional info column must be present to process file
        if self.col['Reports'] == 0:
            print('ERROR: Reports column missing, check file')
            sys.exit()
        if self.col['AdditionalInfo'] == 0:
            print('ERROR: AdditionalInformation column missing, check file')
            sys.exit()
        if self.col['Categories'] == 0:
            print('ERROR: Caterories column missing, check file')
            sys.exit()
        # if Status column was not found, add it
        if self.col['Status'] == 0:
            c += 1
            self.col['Status'] = c
            ws.cell(row=1, column=self.col['Status'], value='STATUS') # write missing status column
        # next free column, add Remarks column
        c += 1
        self.col['Remarks'] = c
        c += 1
        ws.cell(row=1, column=self.col['Remarks'], value = 'REMARKS')
        # if 'Type' column does not exist, we create element as next column so col references do not raise errors.
        # 'Type' is not mandatory for processing
        if self.col['Type'] == 0:
            c +=1
            print('WARNING: Type column missing')
            self.missing_col = 'Type'
            self.col['Type'] = c
        else:
            self.missing_col = ''
        self.last_column = c
    # end _init
##################################################################################################
    def AddColumn(self, ws, label):
        ''' Adds a user defined column to the open worksheet '''
        # not tested
##################################################################################################
        if not isstring(label):
            print ('AddColumn(): wrong type label :', label, 'argument must be a string')
            sys. exit() # fatal error
        if not ws:
            #  error
            print('AddColumn(): No open or valid workheet')
            sys.exit()            
        self.last_column +=1  
        dct = { label : self.last_column }
        self.col.extend(dct)
        ws.cell(row=1, column=self.last_column, value = label)
        return self.last_column
        
# end class ExcelHeader
############################################################################################
def RegexSearch(regex, String, r):
    """"Search for string in text using regex."""
# helper function to search for regular expression to catch error
# and use case ignore flag
#############################################################################################
    try:
        p = re.compile(regex, re.IGNORECASE)
    except:
        print(r, 'Wrong regex:', regex)
        sys.exit()
    mtch = p.search(String)
    return mtch
# end function        