#######################################################################
# parse Excel Worksheet for correct SIC flag
# LOGIC FOR NARCOTICS
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 3.0, new trtiage for complete database
# run version 2.0 with pre conv only (meant for crt)
#######################################################################
# modules/libararies needed
from weakref import WeakSet
from mysqlx import Row
from openpyxl import load_workbook, Workbook
import re  # regex
import sys
import argparse
from common import ExcelHeader, RegexSearch

# definition of crime categories
# order of some crimes in list is important for logic and efficiency
# first crime found and convicted for, ends check for further crimes, that's why
# put most frequent ones first 
ver = '3.0'
crimes = [
    r"racketeering and narcotics",
    r"involved in narcotics business",
    r"(traffick|distribut|import|export|transport|smuggl|production).*?((narcotics?)|(drugs?))",
    r"((narcotics?)|drug) ((traffick(ing)?)|distribution|import|transport|smuggling|violation|sale)",
    r"(drug|(narcotics?)) (delivery|distribution|(cultivat(e|ation))|(manufactur(e|ing))|supply)",
    r"(traffick|distribut|import|export|transport|smuggl|production).*?((drugs?)|(narcotics?))",
    r"dealing in (drugs|nacrotics)",
    r"(sale|robbery|theft|transfer) of (drugs|narcotics)",
    r"(selling|(steal(ing)?)|transfer) (drugs|narcotics)",
    r"posess and distribute (drugs|narcotics)",
    r"theft in nacrotics",
    r"violating federal narcotics law",
    r"violating( the)? law(s)? of narcotics",
    r"(sale|robbery|theft|transfer) of (drugs|narcotics)",
    r"(selling|(steal(ing)?)|transfer) (drugs|Narcotics)",
    r"posession for the purpose of trafficking",
    r"narcotics( .+?)? (charges|(crimes(s)?)|factory|felony|(for sale)|(offence(s)?))",
    r"(drugs?|nartotics?)( .+?){0,3}? ((manufactur(e|ing))|conspiracy|dealing|supply|trade)",
    r"((drugs?)|(narcotics?))(( .+){0,2}?[ -]related)? ((offence[s]?)|(crime[s]?|)(charge[s]?)|activities)",
    r"produc(e|ing)( .+?){0,2}? ((drugs?)|(narcotics?))",
    r"production of (drugs|narcotics)",
    r"manufactur(e|ing) of( a)? ((drugs?)|(narcotics?))",
    r"(narcotic|drug) posession.*?((deal(ing)?)|(sell(ing)?)|sale|(distribut(ion|e|ing))|(traffick(ing?))|(cultivat(ion|e))|(supply(ing)?))",
    r"posession of (drugs|natotics).*?((deal(ing)?)|sale|(sell(ing)?)|(distribut(ion|e|ing))|(supply(ing?))|traffick(ing)|(cultivat(e|ion)))",
    r"(distribut|sell).*?(drugs|narcotics)",
    r"smuggl(e|ing|ed)( .+?){0,3}? ((drugs?)|(narcotics?))",
    r"link between narcotic cartels",
    r"narcotics seized",
    r"seizure of narcotics",
    r"unlawful sale and promotion of prescription drugs",
    r"(smuggl|distribut|sell).*?(ketamine)",
    r"aggravated trafficking",
    r"dispensing of control(led)? substance(s)?",
    r"distribution of a listed chemical",
    r"intent on dealing",
    r"intent to distribute",
    r"narcotics[ -]trac[k]?fficking",
    r"((narcotics?)|(drugs?)) ((charge[s]?)|racket|activities)",
    r"narcotics( .+?)? (charges|(crimes(s)?)|factory|felony|(for sale)|(offence(s)?))",
    r"((narcotics?)|(drugs?)) (production|precursors|cultivation)",
    r"relat(ed|ing) to (drugs|Narcotics)",
    r"smuggl.* ((drugs?)|(narcotics?))",
    r"smuggl(e|ing|ed)( .+?){0,3}? ((drugs?)|(narcotics?))",
    r"((drugs?)|(Nartocis?)) ((charge[s]?)|racket|activities)",
    r"$$ ((charge[s]?)|racket|activities)",
    r"(traffick|distribut|import|export|transport|smuggl|production).*?$$",
    r"$$ (traffick|distribution|import|transport|smuggling|violation|sale)",
    r"(deliver|distribut|cultivat|manufactur|dealing).*?$$",
    r"$$ (delivery|distribution|cultivation|manufacture|supply)",
    r"(sell|suppl).*?$$",
    r"dealing in $$",
    r"posess and distribute $$",
    r"(sale|robbery|theft|transfer) of $$",
    r"(selling|(steal(ing)?)|transfer) $$",
    r"posess and distribute $$",
    r"posess(ing)? with( the)? intent to (distribute|supply) $$",
    r"$$ for sale or supply",
    r"racketeering involving $$",
    r"ditribut\w+(. +?)? $$",
    r"$$ (production|precursors|cultivation)",
    r"production of $$",
    r"$$(. +?){0,3}? ((manufactur(e|ing))|conspiracy|dealing|supply|trade)",
    r"$$(( .+){0,2}?[ -]related)? ((offence[s]?)|(crime[s]?|)(charge[s]?)|activities)",
    r"produc(e|ing)( .+?){0,2}? $$",
    r"manufactur(e|ing) of( a)? $$",
    r"$$ posession.*?((deal(ing)?)|(sell(ing)?)|sale|(distribut(ion|e|ing))|(traffick(ing?))|(cultivat(ion|e))|(supply(ing)?))",
    r"posession of $$.*?((deal(ing)?)|sale|(sell(ing)?)|(distribut(ion|e|ing))|(supply(ing?))|traffick(ing)|(cultivat(e|ion)))",
    r"(distribut|sell).*?$$",
    r"$$ ((charge[s]?)|racket|activities)",
    r"relat(ed|ing) to $$",
    r"smuggl.* $$",
    r"smuggl(e|ing|ed)( .+?){0,3}? $$",
    ]
# most common drugs

drugs = [
    r"cocaine( .+?)?",
    r"controlled substance[s]?",
    r"crack",
    r"ecstasy",
    r"mari[hj]uana",
    r"hydrocodone( .+?)?",
    r"isomethadone( .+?)?",
    r"narcotic[s]?",
    r"drug[s]?"
]
acquittals = [
    r"a[c]?quitt(al|ed)",
    r"pardon(ed)?"
]
dismissals = [
    r"dismiss(ed|al)",
    r"dropped",
    r"case filed"
]

words_apart = 20 # maximum distance of words apart from crime and conviction when matching cirme frst and conviction second
pre_conv = False
DebugFlg = False
# functions
####################################################################
def check_list_sic(list_tag, r):
#  returns true or false for trapping positive sic lists
####################################################################
# lists that trigger positive sic tag. could be read from a file
    TrueLists =[
        r"BDDNC",
        r"INNCB",
        r"NPNCB",
        r"NDLEA",
        r"PKANF",
        r"PHPDEA",
        r"RUFDCS",
        r"BPI-SDNT",
        r"BPI-SDNTK",
        r"SDNT",
        r"SDNTK",
        r"USINL",
        r"USSS-FRAA"
    ]
    list_status = [ False, "Null"]
    for str_list in TrueLists:
        if str_list in list_tag:
            list_status = [ True , str_list]
            return list_status
    return list_status 
# end check_sic_list()
############################################################
def check_conviction(c_crime, str_report, n):
# returning True, False, or None
# checks if there was a convitcion for the crime type
# c_crime : crime (string)
# str_report : record (report column) (string)
# r: row begin processed, for informational purposes only (debugging)
# returns 1 if found and issue follows conviction
# returns 2 if found and issue is followed by conviction
# returns -1 if issue is folloed by conviction but too far apart
# returns 0 is no conviction was found at all
############################################################
    post_conv = 0
    long_flag = False
    global words_apart, DebugFlg, preconv_option
    phrase = [
        r"found guilty",
        r"convicted",
        r"sentence[d]*",
        r"pleaded guilty",
        r"pleaded no contest",
        r"imprisoned",
        r"fined",
        r"arrested .+ serve",
        r"for conviction",
        r"ordered .*\s*to (pay|serve)",
        r"incarcerated",
        r"admitted guilt",
        r"served probation",
        r"to serve .* imprisonment",
        r'previous conviction[s]* .*?'
    ]
    # keywords must be near crime type if before conv
    # build search string with crime type
   
    for conviction in phrase:
        long_flag = False
        # search keyword after conviction. Distance of words are not checked as crime usually follows conviction after a few words
        #  if specified after.
        s_str = conviction + ' .*?' + c_crime  # conviction before crime
        if DebugFlg:
            print(n, s_str)
            input("Press return ")
        x = RegexSearch(s_str, str_report, n)
        if x:
            post_conv = 2
            if DebugFlg:
                print(n, x.group())
            words = re.split('\s', x.group())
            if len(words) > words_apart:
                # crime too far from sentence, look for another sentence further ahead, in case there are two
                y = RegexSearch(str_report, x.start()+1, n)
                if y:
                    words = re.split("\s", y.group())
                    if len(words) > words_apart:
                        print (n, "Too many words between offense and conv", end="\r")
                        post_conv = -1 # to flag for review
                        # issue is too far for a conclusive conviction
                        # let's ignore this., but flag for review in case we do not find further evidence (return -1)
                        long_flag = True
                # issue is too far for a conclusive conviction
                # let's ignore this., but flag for review in case we do not find further evidence (return -1)
            # end if (len)
            # we have foud conviction, check for acquittals
            if long_flag == False:
                for tag in acquittals:
                    s_str = c_crime + '.*' + tag
                    s_acquitt = RegexSearch(s_str, str_report, n)
                    if s_acquitt:
                        print("Acquittal found                                ", end='\r')   
                        return -2 # although not coreect, is behaves like correct as no further offences are affeced if there is a dismissal
                        # this may be revised
                # end for (acquitals)
            # end if (long_flag)
            return post_conv
        # end if (x, conviction found)       
        # if not found, check the other way around. problem is if there was a conviction for somthing different, in which
        # case we should not check for previous mentions of issues and we should assume no conviction for this loop item
        # this is difficult. Here some tries just to catch these common ones
        if "pleaded guilty to" in str_report:
            continue
        if "found guilty for" in str_report:
            continue
        if "pleaded no contest to" in str_report:
            continue
        s_str = 'sentenced for \d+ years'
        x = RegexSearch(s_str, str_report, n)
        
        if not x:
            s_str = "sentence[d]* .*? *for "
            x = RegexSearch(s_str, str_report, n)
            if x:
                continue
        else:
            if DebugFlg:
                print(c_crime, "\n", str)
                print(n, "sentenced for x years found")
                input("enter")
        s_str = "sentence[d]* .*? *on charges of"
        x=  RegexSearch(s_str, str_report, n)
        if x:
            continue
        s_str = "found guilty .*? *on charges of"
        x= RegexSearch(s_str, str_report, n)
        if x:
            continue
        s_str = 'pleaded guilty .*? *to'
        x=  RegexSearch(s_str, str_report, n)
        if x:
            continue
        
        # now as previous were cleared, we check the other way around, conviction after crime
        s_str = c_crime + r'.*? ' + conviction
        x = RegexSearch(s_str, str_report, n)
        if x:
            post_conv = 1
            if DebugFlg:
                print(n, x.group())
            # found. if there are too many words between the type and the conviction phrase, assume review
            # there may be a later repetiion of the word wich decreases the word count, do we check twice
            words = re.split("\s", x.group()) # split into words

            if len(words) > words_apart: # too many words in between, but there could be further mention of issue
                # look for issue further ahead
                y = RegexSearch(s_str, x.start()+1, n)
                if y:
                    words = re.split("\s", y.group())
                    if len(words) > words_apart:
                        print (n, "Too many words between issue and conv", end="\r")
                        post_conv = -1 # to flag for review
                        long_flag = True
                else:
                    long_flag = True
                    post_conv = -1
                # end if (y)
            # end if (Len)
            if long_flag == False:
                for tag in acquittals:
                    s_str = c_crime + '.*' + tag
                    s_acquitt = RegexSearch(s_str, str_report, n)
                    if s_acquitt:
                        print(n, "Acquittal found                                ", end='\r')   
                        return -2 
                    # this may be revised
                # end for (aquitals)
            # end if (long_flag)
            return post_conv
        # end if (x)
    # end for (str)
    return post_conv
#####################################################################
def check_item(item, str_Triage, r):
# treturns True, False or None
#####################################################################
    global pre_conv, preconv_option, DebugFlg, ListCheck, ws, dismissals
    sic_crime = False
    s_crime = RegexSearch(item, str_Triage, r)
    if s_crime:
        pre_conv = True
        if len(s_crime.group()) > 100:
            # should mark as review, as we found a remote connection
            sic_crime = None
            return sic_crime
        chk = check_conviction(item, str_Triage, r)
        if chk == -1:
            # too far away, flag for review (for now)
            sic_crime = None
        if chk == 1:
            # write correct to sheet
            print(r, "SIC Correct                             ", end='\r')
            ws.cell(row=r, column=head.col['Status'], value="SIC TAG CORRECT")
            if ListCheck:
                ws.cell(row=r, column=head.col['Remarks'], value="Conviction after Triage. From List")
            else:
                ws.cell(row=r, column=head.col['Remarks'], value="Conviction after Tag")
            return True
        if chk == 2:
            # write correct to sheet
            print(r, "SIC Correct                             ", end='\r')
            ws.cell(row=r, column=head.col['Status'], value="SIC TAG CORRECT")
            if ListCheck:
                ws.cell(row=r, column=head.col['Remarks'], value="Conviction before Triage. From List")
            else:
                ws.cell(row=r, column=head.col['Remarks'], value="Conviction before Tag")
            return True
        if chk == -2:
            print(r, 'Review manually, acquittal found                            ', end='\r')
            ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
            if ListCheck:
                ws.cell(row=r, column=head.col['Remarks'], value="Offence tag with conviction but with acquittal. From List")
            else:
                ws.cell(row=r, column=head.col['Remarks'], value="Offence tag with conviction but with acquittal")
            return True
        if chk == 0:
            # no conviction found, look for dismissals
            for tag in dismissals:
                s_str = item + '.*' + tag # we may ommit item in search string
                s_diss = RegexSearch(s_str, str_Triage, r)
                if s_diss:
                    print(r, "Review, dismissal found                                ", end='\r')   
                    ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
                    ws.cell(row=r, column=head.col['Remarks'], value="Dismissal found (no conviction).")
                    return True # behaves like correct as no further offences are affeced if there is a dismissal
                # end if
            # end for
            print (r, 'Tag correct - pre conv', end='\r')
            ws.cell(row=r, column=head.col['Status'], value="SIC TAG CORRECT")
            if ListCheck:
                ws.cell(row=r, column=head.col['Remarks'], value="Offence tag found (no conviction). From List")
            else:
                ws.cell(row=r, column=head.col['Remarks'], value="Offence tag found (no conviction).")
            return True
    # end if (crime found)
    return sic_crime
    
#####################################################################
def check_issues(issues, str_Triage, r):
# checks if crime was found and convicted
# # returns True (crime found and written in record), False, and None (to review) 
# writes record if correct
#####################################################################
    global pre_conv, DebugFlg, pws, NoChemCheck
    sic_crime = False

    for x_crime in issues:
        if "$$" not in x_crime:
            sic_crime = check_item(x_crime, str_Triage, r)
            if sic_crime:
                break
        else:
            # loop for each drug in common list
            for drug in drugs:
                # replace $$ with proper drug
                s_str = x_crime.replace( "$$", drug)
                sic_crime = check_item(s_str, str_Triage, r)
                if sic_crime:
                    break
            # end for (drug loop
            if sic_crime:
                break
            # not found yet, let's check in chemmicals list then
            #enforce option not to check for chemicals
            if NoChemCheck:
                continue
            n = 0
            for ph_row in pws.rows:
                n += 1
                chemical = pws.cell(row=n, column=1).value
                s_str = x_crime.replace( "$$", chemical)
                if DebugFlg:
                    print(s_str)
                    print(r, 'chemical in row', n, chemical)
                    input('continue >')
                sic_crime = check_item(s_str, str_Triage, r)
                if sic_crime:
                    break
            # end loop trough drugs in excel list
            if sic_crime:
                break
    # end for (issues loop)
    return sic_crime
# end functions
###################################################

# start program
Testflag = False
DebugFlg = False
preconv_option = False
TrueCondition = False
offence_found = False
RowLimit = 0
NoChemCheck = False
parser = argparse.ArgumentParser(description='Process Narcotics SIC', prog='nacrotics.py')
parser.add_argument("--pc", help="Chcek pre-convition only)", action='store_true')
parser.add_argument("--version",help="Displays version only", action='version', version='%(prog)s ' + ver)
parser.add_argument("--debug", help="Debug mode", action='store_true')
parser.add_argument('filename', help="filename to read")
parser.add_argument('-t', '--test', help='run for a limited number of rows', type=int)
parser.add_argument("--nolist", help="No chemicals list is checked", action='store_true')
args = parser.parse_args()
if args.debug:
    DebugFlg = True
if args.pc:
    preconv_option = True
    print("Pre conviction option")
if args.nolist:
    NoChemCheck = True
    print("Chemical file is not checked.")
if args.test:
    Testflag = True
    RowLimit = args.test   
    print ('Test: processing only', RowLimit, ' rows') 
org_file = args.filename
if ".xlsx" not in org_file:
    if preconv_option:
        dest_file = org_file + ' Preconv Passed.xlsx'
    else:
        dest_file = org_file + ' Passed.xlsx'
    WSheet = org_file
    org_file = org_file + '.xlsx'
else:
    file_parts = org_file.split('.')
    if preconv_option:
        dest_file = file_parts[0] + ' Preconv Passed.xlxs'
    else:
        dest_file = file_parts[0] + ' Passed.xlxs'
    WSheet = file_parts[0]
if DebugFlg:
    print('Org: ', org_file)
    print('Dest: ', dest_file)
    input('enter > ')

print( 'Loading spreadsheet', org_file)
# check if filename exists
#
try:     
    wb = load_workbook(filename=org_file)
except:
    print("cannot open file", org_file)
    sys.exit()
ws = wb[WSheet]
head = ExcelHeader(ws)
if DebugFlg:
 print(head.col)
 input('Enter > ')
 # loaf chemicals list
if not NoChemCheck:
    print('Loading chemicals list from file.')
    try:
        pwb = load_workbook('NarcList.xlsx')
    except:
        print('Coud not open file NarcList.xlsx. Is it open?')
        input('Enter > ')
        pwb = load_workbook('NarcList.xlsx')
    # load workheet for pharma
    pws = pwb['Sheet1']
r = 0
print("Processing worksheet")
for row in ws.rows:
  
    pre_conv = False
    r += 1
    if r == 1:
        continue
    if Testflag:
        if r > RowLimit:
            break
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=head.col['Categories']).value       
    c_OfficialLists = ws.cell(row=r, column=head.col['OfficialLists']).value    
    c_AdditionalInfo = ws.cell(row=r,column=head.col['AdditionalInfo']).value    
    c_Reports = ws.cell(row=r,column=head.col['Reports']).value     
    c_Type = ws.cell(row=r, column=head.col['Type']).value
    c_status= ws.cell(row=r, column=head.col['Status']).value
    c_Triage = c_Reports
    TagStr = [] # resets list of OifficialLists
    Extra = False
    LongReport = False
    ListCheck = False
    

    # Entities are flagged for manual review
    if c_Type == "E":
        # entity, flag for manual review
        print(r, "Entity: Review manually", end='\r')
        ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
        ws.cell(row=r, column=head.col['Remarks'], value='Entity')
        print(r, "Entity.                                 ", end='\r')
        continue
    if DebugFlg:
        print(r, c_Reports)
        input('Enter > ')
    if not c_Reports:
        ws.cell(row=r, column=head.col['Status'], value='SIC TAG NOT FOUND')
        ws.cell(row=r, column=head.col['Remarks'], value='No report column')
        print(r, "No report found.                         ", end='\r')
        continue
    if len(c_Reports) > 800:
        ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
        ws.cell(row=r, column=head.col['Remarks'], value='Report or list content too long')
        print(r, "Report too long.                         ", end='\r')
        continue
    # check if in additional lists
    if c_OfficialLists:
        sic_list = check_list_sic(c_OfficialLists, r)
        if sic_list[0] == True:
            print(r, "Tagged list", end='\r')
            ws.cell(row=r, column=head.col['Status'], value="SIC TAG CORRECT")
            ws.cell(row=r, column=head.col['Remarks'], value='Triggered by official List presence' + sic_list[1])
            continue
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
            x = RegexSearch(regex, c_AdditionalInfo, r)
            if x:
                TagStr.append(x.group())
                 # we do not need to strip the brackets
                print(r, "List match ", tag, "found            ", end="\r")
                i += 1
                Extra = True
            # end if
        # end for
    # end if (OfficialLists)
    # we now have TagInfo populated
    # len(TagInfo) = mumber of elements (>0) or i , Extra = True as flag, and i as the 
    # for str in TagInfo:
    #  we check crimes and convictions there as well at the end of the following loop
    # check for convvicted crimes in Report
    print(r, '                                                  ', end='\r')
    sic_crime = check_issues(crimes, c_Triage, r)
    # now we check for Addidional List Tags
    if sic_crime == True:
        continue # go to next record
    # if not found yet, continue checking in lists if present
    if Extra:
        print(r, "Checking additional in Lists", end="\r")
        LongReport = False
        ListCheck = True
        for x_Triage in TagStr:
            if len(x_Triage) > 720:
                LongReport = True
                ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
                ws.cell(row=r, column=head.col['Remarks'], value="List entry too long.")
                print(r, "Long list enry.                         ", end='\r')
                continue # check next list tag, if found successfull, the record will be overwritten
            sic_crime = check_issues(crimes, x_Triage, r)
            if sic_crime == True:
                break
        # end for
        # if long report was detected and no crime found
        if LongReport:
            continue
    # end if
    # if sic_crime was true, it was already written
    if sic_crime == True:
        continue
    if sic_crime == False:
        if preconv_option:
            print(r, 'SIC incorrect                                   ', end='\r')
            ws.cell(row=r, column=head.col['Status'], value="TAG SHOULD BE REMOVED")
            ws.cell(row=r, column=head.col['Remarks'], value="No offence tag found")
            continue
        print(r, "SIC incorrect                              ", end='\r')
        if pre_conv:
            ws.cell(row=r, column=head.col['Status'], value="TAG SHOULD BE REMOVED")
            ws.cell(row=r, column=head.col['Remarks'], value="No offence tag found")
        else:
            ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
            ws.cell(row=r, column=head.col['Remarks'], value="Offence found with possible aquittals")
    if sic_crime == None:
        print(r, "Review manually. Long text.                            ", end='\r')
        ws.cell(row=r, column=head.col['Status'], value="REVIEW MANUALLY")
        ws.cell(row=r, column=head.col['Remarks'], value="Long distance between offence and conviction")
    
    # end loop through rows
# write to new workbook
print('\nWriting and saving results in file', dest_file, '...' )
try:
    wb.save(dest_file)
except:
    input("\nCannot write to file. Try to close it first and press enter > ")
    print("Saving...")
    wb.save(dest_file)
print('Done')
# end program ######################################################################################################
 