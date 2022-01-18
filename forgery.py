#######################################################################
# parse Excel Worksheet for correct SIC flag
# LOGIC FOR FORGARY & UTTER
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 1.0 commited 20222-01-18
# Notes: TODO refine logic with client
#######################################################################
# modules/libararies needed
from openpyxl import load_workbook, Workbook
import re  # regex
import sys
# definition of crime categories
# order of some crimes in list is important for logic and efficiency
# first crime found and convicted for, ends check for further crimes, that's why
# put most frequent ones first 
ver = '1.0'
crimes = [
    r"(fake|false) (report|invoice)",
    r"(false|counterfeit) \w*passport",
    r"false (public|private) document",
    r"false travel document",
    r"(false|fake) document",
    r"false information",
    r"false statement",
    r"false and misleading information",
    r"false inaccurate or incomplete documents and statements",
    r"false income document",
    r"false or misleading (information|statment)",
    r"false (record|report)",
    r"false social security number",
    r"false tax return",
    r"false testimony",
    r"falsehood in public document",
    r"falsely reporting expense",
    r"falsification (?!.*?(currency|note))",
    r"falsified",
    r"falsifying",
    r"fictitious (statement|invoice)",
    r"forg[ei]",
    r"(inaccurate)|(incomplete) documents and statements",
    r"misleading information",
    r"(providing|submitting) false information",
    r"(providing|submitting) false or misleading information",
    r"uttering",
    r"counterfeit(ing)* document",
    r"counterfeit of a private document",
    r"counterfeit of official document",
    r"counterfeiting official document",
    r"document falsehood"
    ]
words_apart = 20 # maximum distance of words apart from crime and conviction when matching cirme frst and conviction second
pre_conv = False
# functions
############################################################
def check_conviction(type, str_report, n):
# returning True, False, or None
# checks if there was a convitcion for the crime type
# type : crime (string)
# str_report : record (report column) (string)
# r: row begin processed, for informational purposes only (debugging)
############################################################
    global words_apart, DebugFlg
    post_conv = 0
    phrase = [
        r"convicted",
        r"sentence[d]*",
        r"pleaded guilty",
        r"pleaded no contest",
        r"found guilty",
        r"imprisoned",
        r"fined",
        r"arrested .+ serve",
        r"for conviction",
        r"ordered .*\s*to pay",
        r"incarcerated",
        r"amditted guilt",
        r"served probation",
        r"to serve .* imprisonment"
    ]
    # keywords must be near crime type if before conv
    # build search string with crime type
   
    for str in phrase:
        # search keyword after conviction. 
        #  if specified after.
        s_str = str + ' .*?' + type  # RegEx  non-greedy
        p = re.compile(s_str, re.IGNORECASE)
        x = p.search(str_report)
        if x:
            words = re.split("\s", x.group())
            if len(words) > words_apart: # too many words in between, but there could be further mention of conviction
                # look for conviction further ahead
                y = p.search(str_report, x.start()+1)
                if y:
                    words = re.split("\s", y.group())
                    if len(words) > words_apart:
                        print (n, "Too many words between issue and conv", end="\r")
                        post_conv = -1 # to flag for review
                else:
                    return 1 
            else:
                return 1
        # if not found, check the other way around. problem is if there was a conviction for somthing different, in which
        # case we should not check for preious mentions of issues
        # this is difficult. Here some tries just to catch these common ones
        if "sentenced for" in str_report:
            continue
        if "pleaded guilty to" in str_report:
            continue
        if "found guilty for" in str_report:
            continue
        if "pleaded no contest to" in str_report:
            continue
        s_str = "sentence[d]* .*?for "
        p = re.compile(s_str, re.I)
        x = p.search(str_report)
        if x:
            continue
        # now the other way around
        s_str = type + r'.*? (' + str + r')'
        p = re.compile(s_str, re.IGNORECASE)
        x = p.search(str_report)
        if x:
            if DebugFlg:
                print(n, x.group())
            # found. if there are too many words between the type and the conviction phrase, assume review
            # there may be a later repetiion of the word wich decreases the word count, do we check twice
            words = re.split("\s", x.group()) # split into words

            if len(words) > words_apart: # too many words in between, but there could be further mention of issue
                # look for issue further ahead
                y = p.search(str_report, x.start()+1)
                if y:
                    words = re.split("\s", y.group())
                    if len(words) > words_apart:
                        print (n, "Too many words between issue and conv", end="\r")
                        post_conv = -1 # to flag for review
                else:
                    return 2
            else:
                return 2
            # end if
        else:
            pass
        # end if
    # end for
    return post_conv
# end function ######################################################
#####################################################################
def check_issue(issues, str_Triage, r):
# checks if crime was found and convicted
# # returns True (crime found and written in record), False, and None (to review) 
# writes record if correct
#####################################################################
    global DebugFlg
    sic_crime = False
    for x_crime in issues:
        p = re.compile(x_crime, re.I) # to ignore case
        s_crime = p.search(str_Triage)
        if s_crime:
            # put exlusions here
            # end excluisions ##########
            # check conviction for crime
            pre_conv = True # crime found, no conviction (yet)
            chk = check_conviction(x_crime, str_Triage, r)
            if chk == -1:
                # too far away, flag for review (for now)
                print(r, "SIC Review                     ", end='\r')
                sic_crime = None
                # we do not break as we could find a valid record for another kewword
            if chk == 1:
                # write correct to sheet
                print(r, "SIC Correct                             ", end='\r')
                ws.cell(row=r, column=15, value="CORRECT CONV")
                return True
            if chk == 2:
                # write correct to sheet
                print(r, "SIC Correct                             ", end='\r')
                ws.cell(row=r, column=15, value="CORRECT INF")
                return True
                    
        # end if
    # end for
    return sic_crime
# end funcfions
###################################################

# start program
DebugFlg = False
NumArgs = len(sys.argv)
if NumArgs > 1:
    print(NumArgs)
    if sys.argv[1] == '--version':
        print('crime version ', ver)
        sys.exit()
    elif sys.argv[1] == '--debug':
        DebugFlg = True
        print("Running in debug mode")
    else:
        print('Usage: forgery [--version|debug]')
        sys.exit()
# open workbook
print( 'Loading spreadsheet F&U.xlsx...')    
wb = load_workbook(filename="F&U.xlsx")
ws = wb['Post Conv Forgery & Utter']
r = 0
print("Processing sheet")
for row in ws.rows:
    r += 1
    pre_conv = False
    if r == 1:
        continue # skip header (first row)
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=5).value       # column G
    c_OfficialLists = ws.cell(row=r, column=6).value    # column H
    c_AdditionalInfo = ws.cell(row=r,column=7).value    # cloumn I
    c_Reports = ws.cell(row=r,column=8).value          # column L
    c_Type = ws.cell(row=r, column=10).value            # column Y
    c_status= ws.cell(row=r, column=15).value           # column AH
    c_Triage =c_Reports
    TagStr = [] # resets list of OifficialLists
    Extra = False
    # skip non-crime records
    if "CRIME" not in c_categories:
        continue

    # Note: generalisation: one could filter only for review manaully records (c_status == "REVIEW MANUALLY"). e.g.:
    # if "REVIEW" not in c_status:
    #       continue
    # Entities are flagged for manual review
    if c_Type == "E":
        # entity, flag for manual review
        print(r, "Entity: Review manually", end='\r')
        ws.cell(row=r, column=15, value="ENTITY: REVIEW MANNUALLY")
        continue
    if len(c_Reports) > 700:
        ws.cell(row=r, column=15, value="TOO LONG REPORT: REVIEW MANUALLY")
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
            regex = '\['+tag+'\].*?\['
            if DebugFlg: 
                print (regex)
            p = re.compile(regex)
            x = p.search(c_AdditionalInfo)
            if x:
                TagStr.append(x.group())
                 # we do not need to strip the brackets
                print(r, "LIst match ", tag, "found", end="\r")
                i += 1
                Extra = True
            # end if
    #   # end for
        if DebugFlg:
            print(TagStr)
    # end if
    # we now have TagInfo populated
    # len(TagInfo) = mumber of elements (>0) or i , Ectra = True as flag, and i as the 
    # for str in TagInfo:
    #  we check crimes and convictions there as well at the end ofr the following loop
    
    # check for convvicted crimes in Report
    sic_crime = check_issue(crimes, c_Triage, r)
    # now we check for Addidional List Tags
    if sic_crime == True:
        continue # go to next record
    if Extra:
        print("Checking additional in Lists", end="\r")
        for x_Triage in TagStr:
            sic_crime = check_issue(crimes, x_Triage, r)
            if sic_crime == True:
                break
        # end for
    # end if
    # if sic_crime was true, it was already written
    if sic_crime == True:
        continue
    if sic_crime == False:
        print(r, "SIC incorrect                              ", end='\r')
        if pre_conv == False:
            ws.cell(row=r, column=15, value="INCORRECT")
        else:
            ws.cell(row=r, column=15, value="INCORRECT NO CONV (REVIEW)")
    if sic_crime == None:
        print(r, "Review manually                            ", end='\r')
        ws.cell(row=r, column=15, value="REVIEW MANUALLY")
    
    # end loop through rows
# write to new workbook

print('\nWriting and saving results spreadsheet F&U Passed.xlsx ...')
wb.save('F&U Passed.xlsx')
print('Done')
# end program ######################################################################################################
 