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
from openpyxl import Workbook, load_workbook
import re  # regex
import sys
import argparse
# definition of crime categories
# order of some crimes in list is important for logic and efficiency
# first crime found and convicted for, ends check for further crimes, that's why
# put most frequent ones first 
ver = '1.3'
crimes = [
    r"(fake|altered|altering|false|counterfeit|ficticious)( private| public| identity)? document[s]?",
    r"(fake|altered|altering|false|counterfeit) report[s]?",
    r"counterfeit of( a)? (private|public|offical) document[s]?",
    r"fake( .+?){0,3} ((certificate[s]?)|(document[s]?))",
    r"fake papers",
    r"(fake|false|(counterfeit(ing)?))(( credit[s]?)| debit| prepaid)? card[s]?",
    r"(counterfeit|fake)( .+?)? licen[sc]e[s]?",
    r"altering a",
    r"altering(  +?){0,15} ((reimbursement[s]?)|rececipts)",
    r"fake commercial private instrument",
    r"altering an insurance ralated document",
    r"counterfeiting (certificates|documents)",
    r"(counterfeit|fake|false)( .+?){0,3} ticket[s]?",
    r"counterfeit(ing)? ATM card[s]?",
    r"countrfeit bills",
    r"counterfeit credit and debit cards",
    r"(counterfeit|fake) health insurance",
    r"fake( .+?){0,2} health insurance",
    r"(false|ficticious|fake) invoice[s]?",
    r"fake electricity bills",
    r"false ((record[s]?)|subscription)",
    r"false ID",
    r"false issuance of special VAT invoices",
    r"(fake|false) travel document[s]?",
    r"(false|fake|counterfeit)( .+?){0,2} identit(y|ies|ification)",
    r"fake( .+?)? visa",
    r"(fake|false|counterfeit|counterfeiting|tampered)( .+?){0,2} passport",
    r"(fake|counterfeit|counterfeiting) ATM",
    r"false expense",
    r"fake medical (job|(report[s]?))",
    r"(false|fake) (vehicle|car)",
    r"fake bank (guarantee|accounts|cards)",
    r"fake bearer cheques",
    r"fake( payments)? receipt[s]?",
    r"(fake|falsely)( GST)? billing",
    r"(fake|false) credentials",
    r"fake government (identity|identification|papers|(document[s]?))",
    r"fake gst billing racket",
    r"fake insurance",
    r"counterfeiting document[s]?",
    r"false( .+?){0,4} ((document[s]?)|(report[s]?))",
    r"false (invoicing|(claim[s]?)|logbook|names)",
    r"false( print)? receipts",
    r"counterfeit printing",
    r"tampered( .+?){0,2} passport[s]?",
    r"tampering with a public record",
    r"counterfeit(ing)( .+?){0,2} passport[s]?",
    r"(fake|counterfeit|counterfeiting)( .+?){0,4} stamp[s]?",
    r"(counterfeit|false|fake)( .+?){0,2} (identification|identity)",
    r"((counterfeit(ing)*)|fake)( .+?){0,3} residence card[s]?",
    r"fraudulent( identification)? document[s]?",
    r"(false|fake) documentation",
    r"false (information|(instrument[s]?))",
    r"false( account)? statement[s]?",
    r"misleading information",
    r"false( and| or) misleading ((document[s]?)|information|(statement[s]*))",
    r"false( and| or) misleading promotional material*",
    r"false( and| or) incomplete document[s]?",
    r"false inaccurate or incomplete documents",
    r"false income document[s]*",
    r"false or misleading (information|(statement[s]*))",
    r"false (record|(report[s]?))",
    r"false social security number",
    r"(false|fake) tax (refund|return)",
    r"false federal income tax",
    r"false( and fictitious)? securities",
    r"false signature[s]?",
    r"false account statements",
    r"false social security number",
    r"falsifying( .+?){0,4}(signatures|form)",
    r"falsifying ((document[s]?)|books)"
    r"fake (accounts|degree|demand|joining letters|lottery|letterheads|titles)",
    r"(fake|false) land document[s]*",
    r"(fake|false) (((para)?medical)|affidavit)",
    r"fake navy bills",
    r"fake receipts scam",
    r"fake admission[s]?",
    r"fake agreement",
    r"fake job ((certificate[s]?)|racket)",
    r"false annual compliance certification",
    r"answering firm's annual comliance questionnary falsely",
    r"applications that had been falsified",
    r"fake education(al)? ((certificate[s]?)|(documment[s]*))",
    r"false (testimony|writing)",
    r"false value added tax",
    r"fake e-pass",
    r"fake examination[- ]certificate racket",
    r"fake army appointment letter",
    r"(false|fake)( .+?){0,3}? licen[cs]e[s]*",
    r"fictitious auto loan",
    r"ficticious prime bank securities",
    r"counterfeit(ing)?( .+?){0,2}? licen[cs]e[s]*",
    r"counterfeit(ing)? credit card[s]*",
    r"counterfeit (certificates|stamps)",
    r"counterfeiting of subpoenas",
    r"false entr(y|ies)",
    r"falsehood in( public)? document",
    r"(false|fake)( \w)? document[s]?",
    r"falsely (billing|(seeking (reimbursement[s]?))|signing|submitting)",
    r"falsely certif(ied|ying)",
    r"(falsely|fraudulently) inflat(e|ed|ing)",
    r"falsely making out invoices",
    r"falsely reflect conversations with customers",
    r"falsely claiming (assistance|survivor benefits|reimbursement)",
    r"falsely (obtaining|receiving|recording|(report(ing)?)|(represent(ing|ed)))",
    r"falsely reflect conversations",
    r"false claims for expenses",
    r"false firm's documents",
    r"fictitious( .+?){0,3}? ((statement[s]?)|(document[s]?)|(report[s]?)|(reimbursement[s]?))",
    r"forged ((instrument[s]?)|(document[s]?))",
    r"false and ((incomplete)|(misleading)) document[s]?",
    r"(inaccurate|incomplete) documents and statements",
    r"falsely seeking reimbursement",
    r"falsely signing( .+?){0,2}? customer names",
    r"falsely submitting personal expenses",
    r"falsely named as beneficiary",
    r"misleading information",
    r"false( or misleading)? information",
    r"counterfeit of a private document",
    r"counterfeit of official document[s]?",
    r"counterfeiting official document[s]?",
    r"counterfeit of government lottery",
    r"(counterfeit|fake) mark[ -]?sheet",
    r"(document|record) fals\w+",
    r"(document|record) forg\w+",
    r"unsworn falsfication",
    r"fabricating a graft report",
    r"fabricating an account document",
    r"fabricating arbirtation documents",
    r"inaccurate documents and statements",
    r"improper alterations and other falsifications",
    r"(fabricated|fabricating)(. +?){0,4}? ((document[s]?)|((proposer )?information)|(report[s]?)|expense|reimbursement)",
    r"fabricating( total)? profit",
    r"counterfeit(ing)?( .+?){0,4} document[s]?",
    r"fake land (documents|papers|letters|title)",
    r"falsif((y(ing)?)|ied)( .+?){0,1}? expense reimbursement",
    r"falsif((y(ing)?)|ied)( .+?){0,4}? ((passport[s]?)|(document[s]?)|(record[s]?)|information|(form[s]?)|(letter[s]?))",
    r"falsif((y(ing)?)|ied)( .+?){0,5}? (travel|letters|(signature[s]?)|(affidavit[s]?)|insurance|statements|claims)",
    r"falsif((y(ing)?)|ied)( .+?){0,4}? (books|(report[s]?)|(authori[zs]ation)|(account[s]?)|funds|(certificate[s]?))",
    r"falsif((y(ing)?)|ied)( .+?){0,4}? (promissory notes|variable annuity|paperwork|federal tax income|resident cards)",
    r"falsif((y(ing)?)|ied)( .+?){0,2}? (wire transfer|data|results|annuity withdrawal|((customer(s)?)? application[s]?))",
    r"falsif((y(ing)?)|ied)( .+?){0,4}? (commision|account|cards|electoral|identity|loan applications)",
    r"falsif((y(ing)?)|ied)( .+?)? (((payroll|loan|mortgage) application(s)?)|((new )?account)|((sales|personal|income) tax returns))",
    r"falsif((y(ing)?)|ied) (timesheets|paperwork|data submitted|evidence to prove own regitration|tax receipts)",
    r"falsif((y(ing)?)|ied) meeting minutes for procurement of monies",
    r"falsif((y(ing)?)|ied) (receipts for false claims|return-to-work data)",
    r"falsification of( .+?){0,2}? (an offcial seal|certificates|documents|(evidence[s]?)|(work(er permit applications)?))",
    r"falsification of( .+?){0,2}? (expense|claims|reimbursement)",
    r"falsification of (funeral contracts|official statements in court|registered capital)",
    r"falsification of (disater relief vouchers|value added tax(vat) invoices)",
    r"falsification of( .+?){0,2}? (identity|(public and private documents))",
    r"forg(ing|ed)( .+?){0,3}? (account|(invoice[s]?))",
    r"credit card counterfeiting",
    r"(document|ideological) falsehood",
    r"applications that had been falsified",
    r"answering firm's annual compliance questionnaire falsely",
    r"fake Aadhaar card[s]?",
    r"fake pattadar passbooks",
    r"manipulation of (public procurement|land records|of the shares)"
    r"(false|inflated|inflating)( .+?){0,5}? ((reimbursement[s]?)|(billing{s]?)|(expense[s]?))",
    r"forging",
    r"uttering",
    r"forgery",
    r"falsification[.,]",
    r"forge"
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
    global words_apart, DebugFlg, preconv_option
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
        r"ordered .* *to pay",
        r"incarcerated",
        r"amditted guilt",
        r"served probation",
        r"to serve .* imprisonment",
        r'previous conviction[s]* .*?'
    ]
    # keywords must be near crime type if before conv
    # build search string with crime type
   
    for str in phrase:
        # search keyword after conviction. 
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
        # if not found, check the other way around. problem is if there was a conviction for something different, in which
        # case we should not check for preious mentions of issues
        # this is difficult. Here some tries just to catch these common ones
        # ther eare a few cases in which 
        s_str = 'sentenced for \d+ years'
        p = re.compile(s_str, re.I)
        x = p.search(str_report)
        if not x:
            s_str = "sentence[d]* .*? *for "
            p = re.compile(s_str, re.I)
            x = p.search(str_report)
            if x:
                continue
        else:
            if DebugFlg:
                print(type, "\n", str)
                print(n, "sentenced for x years found")
                input("enter")
            pass
        if "pleaded guilty to" in str_report:
            continue
        if "found guilty of" in str_report:
            continue
        if "pleaded no contest to" in str_report:
            continue
        if "convicted for" in str_report:
            continue
        s_str = "sentence[d]* .*? *on charges of"
        p= re.compile(s_str, re.I)
        x= p.search(str_report)
        if x:
            continue
        s_str = "found guilty .*? *on charges of"
        p= re.compile(s_str, re.I)
        x= p.search(str_report)
        if x:
            continue
        s_str = 'pleaded guilty .*? *to'
        p= re.compile(s_str, re.I)
        x= p.search(str_report)
        if x:
            continue
        # no convitcion for crime found and no specific conviction noticed. We check the other way around
        s_str = type + r'.*? ' + str
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
    global DebugFlg, preconv_option, pre_conv, ListCheck
    sic_crime = False
    i = 0
    for x_crime in issues:
        i += 1
        print(r, 'Checking issue:', i , "                                             ", end='\r')
        try:
            p = re.compile(x_crime, re.I) # to ignore case
        except:
            print('wrong regx string: ', x_crime)
            sys.exit()
        s_crime = p.search(str_Triage)
        if s_crime:
            # put exclusions here
            if x_crime == "uttering":
                x = re.search(r"uttering[.].*( currency|( cheque[s]*)|( bank[ ]?note)| money)", str_Triage)
                if x:
                    # crime is to be evaluated manually, may be forgery alone w/o currency.
                    return -1
                x = re.search(r"uttering.*( currency|( cheque[s]*)|( bank[ ]?note)| money)", str_Triage)
                if x:
                    continue
            if x_crime == "forge":
                x = re.search(r"forge.*( currency|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    # invalid triage
                    continue
            if x_crime == "forgery":
                x = re.search(r"forgery[.,].*( currency|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    # crime is to be evaluated manually, may be forgery alone w/o currency.
                    return -1
                x = re.search(r"forgery.*(( currency)|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    continue
            if x_crime == "forging":
                x = re.search(r"forging[.,].*( currency|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    # crime is to be evaluated manually, may be forgery alone w/o currency.
                    return -1
                x = re.search(r"forging.*(( currency)|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    continue
            if x_crime == r"falsification[.,]":
                x = re.search(r"falsification[.,].*(( currency)|( cheque[s]*)|( bank[ ]?note)|( money))", str_Triage)
                if x:
                    return -1
                # end exclusions ##########
            # check conviction for crime
            #? check if match string is too long?
            pre_conv = True # crime found, no conviction (yet)
            if preconv_option:
                print(r, 'SIC Correct                              ', end='\r')
                if ListCheck:
                    ws.cell(row=r, column=15, value="CORRECT (LIST)")
                else:
                    ws.cell(row=r, column=15, value="CORRECT")
                return True
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
print('Forgery & Uttering Processing\n' 'Version', ver)
DebugFlg = False
preconv_option = False
parser = argparse.ArgumentParser(description='Process Forgery & Uttering SIC', prog='forgery.py')
parser.add_argument("--pc", help="Chcek pre-conviction only", action='store_true')
parser.add_argument("--version",help="Displays version only", action='version', version='%(prog)s ' + ver)
parser.add_argument("--debug", help="Debug mode (verbose)", action='store_true')
parser.add_argument('filename', help="filename to read")
args = parser.parse_args()
if args.debug:
    print("Debug mode")
    DebugFlg = True
if args.pc:
    print("Pre Conv mode")
    preconv_option = True
org_file = args.filename
if ".xlsx" not in org_file:
    if preconv_option:
        dest_file = org_file + ' Preconv Passed.xlsx'
    else:
        dest_file = org_file + ' Passed.xlsx'
    WorkSheet = org_file
    org_file = org_file + '.xlsx'
else:
    file_parts = org_file.split('.')
    if DebugFlg:
        print(file_parts)
    WorkSheet = file_parts[0]
    if preconv_option:
        dest_file = file_parts[0] + ' Preconv Passed.xlxs'
    else:
        dest_file = file_parts[0] + ' Passed.xlxs'
# open workbook

print( 'Loading spreadsheet ', org_file)
# check if filename exists
#
try:     
    wb = load_workbook(filename=org_file)
except:
    print("cannot open file ", org_file)
    sys.exit()
if preconv_option:
    sheet = 'Pre Conv Forgery & Utter'
else:
    sheet = 'Post Conv Forgery & Utter'
ws = wb[sheet]
r = 0
print("Processing sheet")
for row in ws.rows:
    r += 1
    pre_conv = False
    if r == 1:
        continue # skip header (first row)
    # read row into variables (only useful ones)
    c_categories = ws.cell(row=r, column=5).value       # column E
    c_OfficialLists = ws.cell(row=r, column=6).value    # column F
    c_AdditionalInfo = ws.cell(row=r,column=7).value    # cloumn G
    c_Reports = ws.cell(row=r,column=8).value           # column H
    c_Type = ws.cell(row=r, column=10).value            # column J
    c_status= ws.cell(row=r, column=15).value           # column O
    c_Triage =c_Reports
    TagStr = [] # resets list of OifficialLists
    Extra = False
    ListCheck =False
    # skip non-crime records
    # if "CRIME" not in c_categories:
    #    continue

    # Note: generalisation: one could filter only for review manaully records (c_status == "REVIEW MANUALLY"). e.g.:
    # if "REVIEW" not in c_status:
    #       continue
    # Entities are flagged for manual review
    if c_Type == "E":
        # entity, flag for manual review
        print(r, "Entity: Review manually", end='\r')
        ws.cell(row=r, column=15, value="ENTITY: REVIEW MANNUALLY")
        continue
    if not c_Reports:
        ws.cell(row=r, column=15, value='NO REPORT')
        print(r, "No report found.                         ", end='\r')
        continue
    if len(c_Reports) > 750:
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
                print(r, "List match ", tag, "found", end="\r")
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
    #  we check crimes and convictions there as well at the end of the following loop
    
    # check for convvicted crimes in Report
    sic_crime = check_issue(crimes, c_Triage, r)
    # now we check for Addidional List Tags
    if sic_crime == True:
        continue # go to next record
    if Extra:
        print(r, "Checking additional in Lists", end="\r")
        ListCheck = True
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
        if preconv_option:
            print(r, "SIC incorrect                         ", end="\r")
            ws.cell(row=r, column=15, value="INCORRECT")
            continue
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
if preconv_option:
    dws = wb['Post Conv Forgery & Utter']
else:
    dws = wb['Pre Conv Forgery & Utter']
# delete the other sheet
wb.remove(dws)
print('\nWriting and saving results spreadsheet ', dest_file)
try:
    wb.save(dest_file)
except:
    input('Cannot write to '+ dest_file + ', Try to close an press enter> ')
    wb.save(dest_file)
print('Done')
# end program ######################################################################################################
 