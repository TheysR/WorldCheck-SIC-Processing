#######################################################################
# parse Excel Worksheet for correct SIC flag
# lOGIC FOR CRIMES
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 1.4 commited 2022-01-16
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
ver = '1.'
crimes = [
    r"stolen",
    r"steal",
    r"stole public",
    r"misappropriat",
    r"embezzl",
    r"peculat",
    r"larceny",
    r"to rob",
    r"robbing",
    r"robber",
    r"grand theft",
    r"theft",
    r"theft of", # must be after theft
    # r"organized crime",
    r"burglar",
    r"pilfering",
    r"heist",
    r"shoplift",
    r"siphon",
    r"diverted|diversion", 
    # r"illicit enrichment",
    r"malversati",
    # r"public deposit",
    # r"illegally receiving public",
    r"absorbing public deposits|funds",
    r"illegal absorption of public",
    r"illegally receiving|acquiring .*deposits",
    r"illegally receiving|acquiring .*public",
    r"illegally accepting|absorbing .*public",
    r"illegally accepting|absorbing .*deposits",
    r"illicit enrichment",
    r"illegally obtain",
    r"illegally receiving pension",
    r"illegally receiving survivor benefits",
    # r"illegal gain",
    r"funds .* illegally",
    # r"financial management irregularities",
    r"larceny",
    r"ditribuit.* state .* asssets"
    # "Misappropriation of funds", included in misappropation
    # "Trafficking in Stolen Goods",
    # "Selling stolen goods",
    # "Receipt of stolen goods",
    # "Intellectual Property",
    ]
words_apart = 11 # maximum distance of words apart from crime and conviction when matching cirme frst and conviction second
# functions
############################################################
def check_conviction(type, str_report, n):
# returning True, False, or None
# checks if there was a convitcion for the crime type
# type : crime (string)
# str_report : record (report column) (string)
# r: row begin processed, for informational purposes only (debugging)
############################################################
    post_conv = False
    phrase = [
        r"convicted",
        r"sentence[d]*",
        r"pleaded guilty",
        r"found guilty",
        r"imprisoned",
        r"fined",
        r"arrested .+ serve"
    ]
    # keywords must be near crime type if before conv
    # build search string with crime type
   
    for str in phrase:
        # search keyword after conviction. Distance of words are not checked as crime usually follows conviction after a few words
        #  if specified after.
        s_str = str + ' .* ' + type # RegEx word followed by space and anythnig in between and the second word
        #_DEBUG print(n, s_str)
        #_DEBUG input("Press return")
        p = re.compile(s_str, re.IGNORECASE)
        x = p.search(str_report)
        if x:
            #_DEBUG print(n, x.group())
            return True # exit fucntion for any crime found. these are most cases.
        # if not found, check the other way around
        s_str = type + ' .* ' + str
        p = re.compile(s_str, re.IGNORECASE)
        x = p.search(str_report)
        if x:
            #_DEBUG print(n, x.group())
            # found. if there are too many words between the type and the conviction phrase, assume review
            i_str = re.split("\s", x.group()) # split into words

            if len(i_str) > words_apart:
                print (n, " Too many words between crime and conv", end="\r")
                post_conv = None # to flag for review
            else:
                return True
            # end if
        else:
            #_DEBUG print(n, str_report)
            pass
        # end if
    # end for
    return post_conv
# end function ######################################################
#####################################################################
def check_issue(issues, str_Triage):
# checks if crime was found and convicted
# # returns True (crime found and written in record), False, and None (to review) 
#
#####################################################################
    sic_crime = False
    for x_crime in issues:
        s_crime = re.search(x_crime, str_Triage)
        if s_crime:
            # found issue in string
            # first, exclude identity theft and theft by deception #########
            if x_crime == "theft":
                x = re.search("[Ii]dentity theft", str_Triage)
                if x:
                    # idenfity theft found. look for another theft not preceded by Identity
                    y = re.search(r"(?<![Ii]dentity) theft", str_Triage)
                    if y == None:
                        # no theft not preceeded by identity was found, i.e. only identity theft was found
                        continue
                # check for theft by deception
                x = re.search(r"[Tt]heft by deception", str_Triage)
                if x:
                    # see if ther are no other thefts apart form deception
                    y = re.search(r"[Tt]heft (?!by deception)", str_Triage)
                    if y == None:
                        # no theft not followed by 'by deecption' was found, i.e. only theft by deception was found
                        continue
                # catch exxlusive 'theft of', otherwise this iteration could be flagged as correct
                x = re.search(r"[Tt]heft of", str_Triage)
                if x:
                    y = re.search(r"[Th]heft (?!of)", str_Triage)
                    if y == None:
                        continue # if found, will be processed in next iteration
            if x_crime == "theft of":
                # exlude theft of identity
                x = re.search("theft of identi", str_Triage)
                if x:
                    # look for negative
                    y = re.search(r"theft of (?!identi)", str_Triage)
                    if y == None:
                        continue # only theft of identity was found
            if x_crime == "stolen":
                # ecxclude specific stolen stuff
                x = re.search(r"stolen identi" , str_Triage)
                if x:
                    # look for negative
                    y = re.search(r"stolen (?!identi)", str_Triage)
                    if y == None:
                        continue # only stolen identity ewas found, no more soltne, copntien for next phrase
            if x_crime == "stolen":
                x = re.search("stolen credit card scheme" , str_Triage)
                if x:
                    # look for negative
                    y = re.search(r"stolen (?!credit card scheme)", str_Triage)
                    if y == None:
                        continue
            if x_crime == "stolen":
                x = re.search(r"stolen personal identification" , str_Triage)
                if x:
                    # look for negative
                    y = re.search(r"stolen (?!personal identification)", str_Triage)
                    if y == None:
                        continue
            if x_crime == "steal":
                x = re.search("steal trade secrets" , str_Triage)
                if x:
                    # look for negative
                    y = re.search(r"steal (?!trade secrets)", str_Triage)
                    if y == None:
                        continue
            
            # end exclude identity theft, theft by deception and stolen identity, etc ##########
            # check conviction for crime
            chk = check_conviction(x_crime, str_Triage, r)
            if chk == None:
                # too far away, flag for review (for now)
                print(r, "SIC Review", end='\r')
                sic_crime = None
                # we do not break as we could find a valid record for another kewword
            if chk == True:
                # write correct to sheet
                print(r, "   SIC Correct", end='\r')
                ws.cell(row=r, column=15, value="SIC CORRECT")
                return True
                
        # end if
    # end for
    return sic_crime
# end funcfions
###################################################

# start program
NumArgs = len(sys.argv)
if NumArgs > 1:
    print(NumArgs)
    if sys.argv[1] == '--version':
        print('crime version ', ver)
        sys.exit()
    else:
        print('Usage: crime [--version]')
        sys.exit()
# open workbook
print( 'Loading spreadsheet TE2 RECORDS done.xlsx...')    
wb = load_workbook(filename="TE2 RECORDS done.xlsx")
ws = wb['Theft & Embezz Post Conv']
r = 0
print("Processing sheet")
for row in ws.rows:
    r += 1
    post_conv = False
    if r == 1:
        continue # skip header (first row)
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=5).value
    c_OfficialLists = ws.cell(row=r, column=6).value
    #_DEBUG print(r, c_OfficialLists)
    c_AdditionalInfo = ws.cell(row=r,column=7).value
    c_Reports = ws.cell(row=r,column=8).value
    c_Type = ws.cell(row=r, column=9).value
    c_status= ws.cell(row=r, column=15).value
    c_Triage = c_Reports
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
        ws.cell(row=r, column=15, value="REVIEW MANUALLY")
        continue
    # check if in additional lists
    if c_OfficialLists:
        # extract lists from string
        # split string
        #_DEBUG print(r, "List found")
        l_list = c_OfficialLists.split(';')
        i=0
        for tag in l_list:
            # look for tag in c_AdditionalInfo and extract string
            regex = '\['+tag+'\].*?\['
            #_DEBUG print (regex)
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
        #_DEBUG print(TagStr)
    # end if
    # we now have TagInfo populated
    # len(TagInfo) = mumber of elements (>0) or i , Ectra = True as flag, and i as the 
    # for str in TagInfo:
    #  we check crimes and convictions there as well at the end ofr the following loop
    
    # check for convvicted crimes in Report
    sic_crime = check_issue(crimes, c_Triage)
    # now we check for Addidional List Tags
    if sic_crime == True:
        continue # go to next record
    if Extra:
        print("Checking additional in Lists", end="\r")
        for x_Triage in TagStr:
            sic_crime = check_issue(crimes, x_Triage)
        # end for
    # end if
    # if sic_crima was true, it was already written
    if sic_crime == True:
        continue
    if sic_crime == False:
        # SIC is not correct
        #_DEBUG print(r, "No match: ", c_Reports)
        print(r, "  SIC incorrect", end='\r')
        ws.cell(row=r, column=15, value="SIC INCORRECT")
    if sic_crime == None:
        print(r, "Review manually", end='\r')
        ws.cell(row=r, column=15, value="REVIEW MANUALLY")
    # end loop thtough keywords 
# end loop through rows
# write to new workbook

print('\nWriting and saving results spreadsheet vacas4.xlsx ...')
wb.save('vacas4.xlsx')
print('Done')
# end python program ######################################################################################################
 