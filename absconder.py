#######################################################################
# parse Excel Worksheet for correct SIC flag
# LOGIC FOR ABSCONDER OR FUGITIVE
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 1.0 2022-02-01
# run version 2.0 with pre conv only (meant for crt)
#######################################################################
# modules/libararies needed
from openpyxl import load_workbook, Workbook
import re  # regex
import sys
import argparse
from common import ExcelHeader
# definition of crime categories
# order of some crimes in list is important for logic and efficiency
# first crime found and convicted for, ends check for further crimes, that's why
# put most frequent ones first 
ver = '1.0'
Triage = [
    r"wanted by",
    r"abscond(ed)?",
    r"abscond(er|ing)",
    r"at large",
    r"escape(d)?"
    r"fled",
    r"flee(s)?"
    r"fugitive(s)?",
    r"in absentia",
    r"most wanted",
    r"uamvs",
    r"unkonwn whereabouts",
    r"uzmvd",
    r"vnmps-mw",
    r"whereabouts unknown",
    r"placed on( international)? wanted list",
    r"wanted for"

    ]
ReverseTriage = [
    r"imprisoned",
    r"busted",
    r"captured",
    r"custody",
    r"detained",
    r"electronic monitoring",
    r"electronic surveillance measures",
    r"incarceration",
    r"no longer (listed|wanted)", 
    r"surrendered to (authorities|police)", 
    r"under house arrest",
    r"under( special)? surveillance"
]
words_apart = 20 # maximum distance of words apart from crime and conviction when matching cirme frst and conviction second
pre_conv = False
DebugFlg = False
# functions
#####################################################################
def check_issue(issues, str_Triage, r):
# checks if crime was found and convicted
# # returns True (crime found and written in record), False, and None (to review) 
# writes record if correct
#####################################################################
    global pre_conv, DebugFlg, ListCheck, ReverseTriage, ReverseTag
    sic_crime = False
    ReverseTag = False
    for x_crime in issues:
        try:
            p = re.compile(x_crime, re.I)
        except:
            print(r, 'Regex error:', x_crime)
            print(p.error)
            sys.exit()
        s_crime = p.search(str_Triage)
        if s_crime:
            sic_crime = True
            # issue found, check for reverse triage kword if it follows
            # there is normally only one issue per record
            for kword in ReverseTriage:
                s_str = x_crime + '.* ' + kword # or use x_crime instead of s_crime.group()
                try:
                    q = re.compile(s_str, re.I) # to ignore case
                except:
                    print(r, 'Regex error:', s_str)
                    sys.exit()
                s_counter = q.search(str_Triage)
                if s_counter:
                    # counter found, exit incorrect
                    ReverseTag = True
                    return False
            # end for reverse
        # end if
    # end for issue    
    if sic_crime:    
        if ListCheck:
            ws.cell(row=r, column=head.col['Status'], value="SIC CORRECT (LIST)")
        else:
            ws.cell(row=r, column=head.col['Status'], value="SIC CORRECT")
        return True
    # end if (s_crime true)    
    return sic_crime
# end functions
###################################################

# start program
DebugFlg = False
preconv_option = False
CheckTag = False
ReverseTag = False
parser = argparse.ArgumentParser(description='Process Narcitics` SIC', prog='nacrotics')
parser.add_argument("--version",help="Displays version only", action='version', version='%(prog)s ' + ver)
parser.add_argument("--debug", help="Debug mode", action='store_true')
parser.add_argument('filename', help="filename to read")
args = parser.parse_args()
if args.debug:
    DebugFlg = True
org_file = args.filename
if ".xlsx" not in org_file:
    dest_file = org_file + ' Passed.xlsx'
    WorkSheet = org_file 
    org_file = org_file + '.xlsx'
else:
    file_parts = org_file.split('.')
    dest_file = file_parts[0] + ' Passed.xlxs'
    WorkSheet = file_parts[0]
if DebugFlg:
    print('Org: ', org_file)
    print('Dest: ', dest_file)
    input('enter > ')
# open workbook

print( 'Loading spreadsheet', org_file)
# check if filename exists
#
try:     
    wb = load_workbook(filename=org_file)
except:
    print("cannot open file", org_file)
    sys.exit()
ws = wb[WorkSheet]
r = 0
print("Processing worksheet")
head = ExcelHeader(ws)
for row in ws.rows:
    pre_conv = False
    r += 1
    if r == 1:
        # ws.cell(row=1, column=23, value='Status')
        continue
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=head.col['Categories']).value       
    c_AdditionalInfo = ws.cell(row=r,column=head.col['AdditionalInfo']).value 
    c_OfficialLists = ws.cell(row=r, column=head.col['OfficialLists']).value   
    c_Reports = ws.cell(row=r,column=head.col['Reports']).value         
    c_Type = ws.cell(row=r, column=head.col['Type']).value           
    c_status= ws.cell(row=r, column=head.col['Status']).value          
    c_Triage = c_Reports
    TagStr = [] # resets list of OifficialLists
    Extra = False
    LongReport = False
    ListCheck = False
    # if "CRIME" not in c_categories:
    #    continue

    # Note: generalisation: one could filter only for review manaully records (c_status == "REVIEW MANUALLY"). e.g.:
    # if "REVIEW" not in c_status:
    #       continue
    # Entities are flagged for manual review
    if c_Type == "E":
        # entity, flag for manual review
        print(r, "Entity: Review manually", end='\r')
        ws.cell(row=r, column=head.col['Status'], value="ENTITY: REVIEW MANUALLY")
        print(r, "Entity.                                 ", end='\r')
        continue
    if DebugFlg:
        print(r, c_Reports)
        input('Enter > ')
    if not c_Reports:
        ws.cell(row=r, column=head.col['Status'], value='NO REPORT')
        print(r, "No report found.                         ", end='\r')
        continue
    if len(c_Reports) > 800:
        if not CheckTag:
            ws.cell(row=r, column=head.col['Status'], value="LONG REPORT: REVIEW MANUALLY")
            print(r, "Report too long.                         ", end='\r')
            continue
    # check if in additional lists
    if c_OfficialLists:
        # extract lists from string
        # split string
        if DebugFlg:
            print(r, "List found")
        l_list = c_OfficialLists.split(';')
        i=0
        for tag in l_list:
            # look for tag in c_AdditionalInfo and extract string
            # for fraud, if HSS is found, tag a CORRECT HSS and no further processing
            regex = '\['+tag+'\].*?\['
            #_DEBUG print (regex)
            p = re.compile(regex)
            x = p.search(c_AdditionalInfo)
            if x:
                TagStr.append(x.group())
                 # we do not need to strip the brackets
                print(r, "List match ", tag, "found            ", end="\r")
                i += 1
                Extra = True
            # end if
    #   # end for
        if DebugFlg:
            print(TagStr)
    # end if (lists)
    # we now have TagInfo populated
    # len(TagInfo) = mumber of elements (>0) or i , Ectra = True as flag, and i as the 
    # for str in TagInfo:
    #  we check crimes and convictions there as well at the end ofr the following loop
    # check for convvicted crimes in Report
    sic_crime = check_issue(Triage, c_Triage, r)
    # now we check for Addidional List Tags
    if sic_crime == True:
        continue # go to next record
    if Extra:
        print("Checking additional in Lists", end="\r")
        LongReport = False
        ListCheck = True
        for x_Triage in TagStr:
            if len(x_Triage) > 720 and not CheckTag:
                LongReport = True
                ws.cell(row=r, column=head.col['Status'], value="LONG LIST ENTRY: REVIEW")
                print(r, "Long list enry.                         ", end='\r')
                continue
            sic_crime = check_issue(Triage, x_Triage, r)
            if sic_crime == True:
                break
        # end for
    # end if
    # if sic_crime was true, it was already written
    if sic_crime == True:
        continue # to next record
    if sic_crime == False:
        print(r, 'SIC incorrect                                   ', end='\r')
        if ReverseTag:
            ws.cell(row=r, column=head.col['Status'], value='SIC INCORRECT (REV KWORD)')
        else:
            ws.cell(row=r, column=head.col['Status'], value="SIC INCORRECT")
        continue
    if sic_crime == None:
        print(r, "Review manually                            ", end='\r')
        ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
    
    # end loop through rows
# write to new workbook
# delete the other sheet
print('\nWriting and saving results in file', dest_file, '...' )
try:
    wb.save(dest_file)
except:
    input("\nCannot write to file. Try to close it first and press enter > ")
    print("Saving...")
    wb.save(dest_file)
print('Done')
# end program ######################################################################################################
 