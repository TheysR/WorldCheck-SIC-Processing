#######################################################################
# parse Excel Worksheet for correct SIC flag
# lOGIC FOR CRIMES
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 1.2 commited 2022-01-12
# Notes: TODO refine logic with client
#######################################################################
# modules/libararies needed
from openpyxl import load_workbook, Workbook
import re  # regex
# definition of crime categories
# order of some crimes in list is important for logic and efficiency
# fisrt crime found and convicted for ends check for further crimes, that's why 
crimes = [
    "stolen",
    "steal",
    "stole public funds",
    "misappropriat",
    "Embezzl",
    "embezzl", 
    "peculation",
    "larceny",
    "to rob",
    "robbing",
    "robber",
    "Robber",
    "grand theft",
    "theft",
    "theft of",
    # "organized crime",
    "burglar",
    "pilfering",
    "heist",
    "shoplift",
    "siphon",
    "divert", 
    "diversion",
    # "illicit enrichment",
    "malversati",
    # "public deposit",
    # "illegally receiving public",
    "absorbing public deposits illegally",
    "illegally absorbing of public deposits",
    "illegally absorbing public deposits",
    "illegal absorption of public deposits",
    "illegally accepting public deposits",
    "illegally receiving deposits",
    "illicit enrichment",
    # "illegal gain",
    "funds were obtained illegally",
    "funds obtained illegally"
    # "financial management irregularities",
    "Larceny",
    "larceny"
    # "Misappropriation of funds", included in misappropation
    # "Trafficking in Stolen Goods",
    # "Selling stolen goods",
    # "Receipt of stolen goods",
    # "Intellectual Property",
    ]
words_apart = 12 # maximum distance of words apart from crime and conviction
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
        r"[Cc]onvicted",
        r"[Ss]entence[d]*",
        r"[Pp]leaded guilty",
        r"[Ff]ound guilty",
        r"[Ii]mprisoned"
    ]
    # the following checks are Logic A.
    # However, they are the same as in phrase plus punctuation marks
    
    #if ". Pleaded guilty to charges." in str_report:
    #    post_conv = True
    #if ". Pleaded guilty." in str_report:
    #    post_conv = True
    #if ". Found guilty of charges." in str_report:
    #    post_conv = True
    #if ". Found guilty." in str_report:
    #    post_conv = True
    #if ". Imprisoned." in str_report:
    #    post_conv = True
    #if ". Fined." in str_report:
    #    post_conv = True
    #if ". Sentenced." in str_report:
    #    post_conv = True
    #if ". Convicted." in str_report:
    #    post_conv = True
    
    # if post_conv == True:
    #     return post_conv

    # logic B/C
    # keywords must be near crime type
    # build search string with crime type
   
    for p in phrase:
        # search keyword after conviction. Distance of words are not checked as crime usually follows conviction after a few words
        #  if specified after.
        s_str = p + ' .+ ' + type # RegEx word followed by space and anythnig in between and the second word
        x = re.search(s_str, str_report)
        if x:
            #_DEBUG print(n, x.group())
            return True
        # if not found, check the other way around
        s_str = type + '.+' + p
        x = re.search(s_str, str_report)
        if x:
            #_DEBUG print(n, x.group())
            # found. if there are too many words between the type and the conviction phrase, assume review
            i_str = re.split("\s", x.group()) # split into words
            if len(i_str) > words_apart:
                #_DEBUG print (n, " Too many words between crime and conv")
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

# end functions

# start program
# open workbook
print( 'Loading spreadsheet TE.xlsx...')    
wb = load_workbook(filename="TE.xlsx")
ws = wb['Sheet1']
r = 0
for row in ws.rows:
    r += 1
    post_conv = False
    if r == 1:
        continue # skip header (first row)
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=5).value
    c_OfficialLists = ws.cell(row=r, column=6).value
    c_AdditionalInfo = ws.cell(row=r,column=7).value
    c_Reports = ws.cell(row=r,column=8).value
    c_status= ws.cell(row=r, column=15).value
    c_Triage = c_Reports
    # skip non-crime records
    if "CRIME" not in c_categories:
        continue

    # Note: generalisation: one could filter only for review manaully records (c_status == "REVIEW MANUALLY"). e.g.:
    # if "REVIEW" in c_status:
    #       continue

    # find words that indicate crime, loop through crimes
    sic_crime = False
    for x_crime in crimes:
        if x_crime in c_AdditionalInfo:  # see if it is better to use regex instead (in which case we can use them in rcimes list)
            # found crime in rpoert
            # first, exclude identity theft and theft by deception #########
            if x_crime == "theft":
                x = re.search("[Ii]dentity theft", c_Triage)
                if x:
                    # idenfity theft found. look for another theft not preceded by Identity
                    y = re.search(r"(?<![Ii]dentity) theft", c_Triage)
                    if y == None:
                        # no theft not preceeded by identity was found, i.e. only identity theft was found
                        continue
                # check for theft by deception
                x = re.search(r"[Tt]heft by deception", c_Triage)
                if x:
                    # see if ther are no other thefts apart form deception
                    y = re.search(r"[Tt]heft (?!by deception)", c_Triage)
                    if y == None:
                        # no theft not followed by 'by deecption' was found, i.e. only theft by deception was found
                        continue
            if x_crime == "theft of":
                # exlude theft of identity
                x = re.search("theft of identity", c_Triage)
                if x:
                    # look for negative
                    y = re.search(r"theft of (?!identity)", c_Triage)
                    if y == None:
                        continue # only theft of identity was found 
            # end exclude identity theft and theft by deception ##########
            # check conviction for crime
            chk = check_conviction(x_crime, c_Triage, r)
            if chk == None:
                # too far away, flag for review (for now)
                print(r, "SIC Review")
                sic_crime = None
                # we do not break as we could find a valid record for another kewword
            if chk == True:
                # write correct to sheet
                print(r, " SIC Correct")
                ws.cell(row=r, column=15, value="SIC CORRECT")
                sic_crime = True
                break # exit crime type loop as one crime conviction was found
        # end if
    # end for
    if sic_crime == False:
        # SIC is not correct
        #_DEBUG print(r, "No match: ", c_Reports)
        print(r, "  SIC incorrect")
        ws.cell(row=r, column=15, value="SIC INCORRECT")
    if sic_crime == None:
        ws.cell(row=r, column=15, value="REVIEW MANUALLY")
    # end loop thtough keywords 
# end loop through rows
# write to new workbook
print('Writing and saving results spreadsheet vacas4.xlsx ...')
wb.save('vacas4.xlsx')
print('Done')
# end python program ######################################################################################################
 