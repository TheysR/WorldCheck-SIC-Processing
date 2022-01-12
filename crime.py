#######################################################################
# parse Excel Worksheet for correct SIC flag
# lOGIC FOR CRIMES
# crime must match and must be convicted for it
# Parsing Reports Column (H) and writing results in column 15 ()
# (c) 2022 Theys Radmann
# ver 1.1 commited 2022-01-10
# Notes: TODO refine logic with client
#######################################################################
# modules/libararies needed
from openpyxl import load_workbook, Workbook
import re  # regex
# definition of crime categories
crimes = [
    "misappropriation",
    "embezzlement",
    "speculation",
    "larceny",
    "robbing",
    "robbed",
    "robbery",
    "robberies",
    "theft",
    "of theft",
    "theft of",
    "theft by",
    "to theft,",
    "for theft",
    "grand theft",
    "organized crime",
    "burglary",
    "steal",
    "stealing",
    "pilfering",
    "heist",
    "shoplifting",
    "siphoning",
    "siphoned",
    "diverting",
    "diverted",
    "diversion",
    "stolen",
    "illicit enrichment",
    "malversation",
    "malversating",
    "public deposit",
    "illegally receiving public",
    "absorbing public deposits illegally",
    "illegally absorbing of public deposits",
    "illegal absorption of public deposits",
    "illegally accepting public deposits",
    "to rob",
    "illegally receiving deposits",
    "stole public funds",
    "illicit enrichment",
    "illegal gain",
    "funds were obtained illegally",
    "funds obtained illegally",
    "financial management irregularities",
    "Larceny",
    "Misappropriation of funds",
    "Speculation",
    "Robbery",
    "Trafficking in Stolen Goods",
    "Selling stolen goods",
    "Receipt of stolen goods",
    "Intellectual Property",
    "Embezzlement"
    ]
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
        "[Cc]onvicted",
        "[Ss]entenced",
        "[Pp]leaded guilty",
        "[Ff]ound guilty",
        "[Ii]mprisoned"
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
            if len(i_str) > 20:
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
            # first, exclude identity theft #########
            if x_crime == "theft":
                x = re.search("[Ii]dentity theft", c_AdditionalInfo)
                if x:
                    # idenfity theft found. look for another theft not preceded by Identity
                    y = re.search(r"(?<![Ii]dentity) theft", c_AdditionalInfo)
                    if y == None:
                        # no theft not preceeded by identity was found, i.e. only identity theft was found
                        continue
            if x_crime == "theft of":
                # exlude theft of identity
                x = re.search("theft of identity", c_AdditionalInfo)
                if x:
                    # look for negative
                    y = re.search(r"theft of (?!identity)", c_AdditionalInfo)
                    if y == None:
                        continue # for now
            if x_crime == "to theft":
               # exlude theft of identity
                x = re.search("to theft of identity", c_AdditionalInfo)
                if x:
                    # negative
                    y = re.search(r"to theft (?!of identity)", c_AdditionalInfo)
                    if y == None:
                        continue # for now
            # end exclude identity theft ##########
            # check conviction for crime
            chk = check_conviction(x_crime, c_AdditionalInfo, r)
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
wb.save('vacas4.xlsx')
# end python program ######################################################################################################
 