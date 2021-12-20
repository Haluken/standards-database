import json
import xlsxwriter
from datetime import datetime


###CREATE WORKBOOK
workbook = xlsxwriter.Workbook('Prototype Standards Database.xlsx')

###FORMATS
bold = workbook.add_format({'bold': True}) #bold font
date_format = workbook.add_format({'num_format': 'mm/d/yyyy', 'bottom':1, 'right':1}) #date format with bottom and right border
orangeheader = workbook.add_format({'font_color':'#000000','bottom':6,'right':1,'bg_color':'#F8CBAD'}) #Light orange background, black font, double border bottom, single border right
blueheader = workbook.add_format({'font_color':'#000000','bottom':6,'right':1,'bg_color':'#C5D9F1'}) #Black font, double border on bottom, single border on right, light blue background
greenheader = workbook.add_format({'font_color':'#000000','bottom':6,'right':1,'bg_color':'#c4d79b'}) #Black font, double border on bottom, single border on right, light green background
body_format = workbook.add_format({'bottom':1, 'right':1}) #border line on bottom and right
next_month = workbook.add_format({'bg_color':'#FFFF01'}) #Yellow#FFFF01
this_month = workbook.add_format({'bg_color':'#ffbe01'}) #Orange#FF8C01
next_week = workbook.add_format({'bg_color':'#ff6f01'}) #Little more red#FF5001
this_week = workbook.add_format({'bg_color':'#FF0101'}) #Red#FF0101
#tomorrow = workbook.add_format({'bg_color':'#FF0101'}) #Red
#today = workbook.add_format({'bg_color':'#FF0101'}) #Red


##############################ASTM F47###########################################
###LOAD INFO FROM SOURCE
with open('ASTMF47temp.txt') as f:
    ASTMF47_info = json.load(f)

###CREATE SHEET
F47worksheet = workbook.add_worksheet('ASTM - F47')

###SET WIDTHS
F47worksheet.set_column('A:A', 50)
F47worksheet.set_column('B:B', 14)
F47worksheet.set_column('C:C', 15)
F47worksheet.set_column('D:D', 12)
F47worksheet.set_column('E:E', 27)
F47worksheet.set_column('F:F', 15)
F47worksheet.set_column('G:G', 27)
F47worksheet.set_column('H:H', 10.5)


###CREATE TABLE
F47worksheet.add_table(0,0,len(ASTMF47_info),7, {'style':'Table Style Medium 21',
                                          'columns':[{'header':'Title', 'header_format':greenheader},
                                                     {'header':'ID Number', 'header_format':greenheader},
                                                     {'header':'Subcommittee', 'header_format':greenheader},
                                                     {'header':'Start Date', 'header_format':greenheader},
                                                     {'header':'Status', 'header_format':greenheader},
                                                     {'header':'Last Updated', 'header_format':greenheader},
                                                     {'header':'Project Lead', 'header_format':greenheader},
                                                     {'header':'Link', 'header_format':greenheader}]})


###WRITE INFO
count = 1
for item in ASTMF47_info:
    F47worksheet.write(count, 0, ASTMF47_info[item]["TITLE"], body_format)
    F47worksheet.write(count, 1, ASTMF47_info[item]["ID_NUMBER"], body_format)
    F47worksheet.write(count, 2, ASTMF47_info[item]["SUBCOMMITTEE"], body_format)


    try:
        startdate = datetime.strptime(ASTMF47_info[item]["START_DATE"], "%m-%d-%Y")
        F47worksheet.write(count, 3, startdate, date_format)
    except:
        F47worksheet.write(count, 3, "", body_format)

    F47worksheet.write(count, 4, ASTMF47_info[item]["STATUS"], body_format)

    try:
        lastupdated = datetime.strptime(ASTMF47_info[item]["LAST_UPDATED"], "%b %d, %Y")
        F47worksheet.write(count, 5, lastupdated, date_format)
    except:
        F47worksheet.write(count, 5, "", body_format)

    F47worksheet.write(count, 6, ASTMF47_info[item]["PROJECT_LEAD"], body_format)
    F47worksheet.write(count, 7, ASTMF47_info[item]["LINK"], body_format)

    F47worksheet.write(count, 8, " ")
    count += 1


##############################ISO SC 13##########################################
###LOAD INFO FROM SOURCE
with open('ISOSC13temp.txt') as f:
    ISOSC13_info = json.load(f)

###CREATE SHEETS
SC13worksheet = workbook.add_worksheet('ISO - SC13')

###SET WIDTHS
SC13worksheet.set_column('A:A', 75)
SC13worksheet.set_column('B:B', 20)
SC13worksheet.set_column('C:C', 27)
SC13worksheet.set_column('D:D', 22.5)
SC13worksheet.set_column('E:E', 27)
SC13worksheet.set_column('F:F', 17.5)

###CREATE TABLE
SC13worksheet.add_table(0,0,len(ISOSC13_info),5, {'style':'Table Style Medium 21',
                                          'columns':[{'header':'Title', 'header_format':blueheader},
                                                     {'header':'ID Number', 'header_format':blueheader},
                                                     {'header':'Completion Status', 'header_format':blueheader},
                                                     {'header':'Completion Status Date', 'header_format':blueheader},
                                                     {'header':'In Progress/Published/Withdrawn', 'header_format':blueheader},
                                                     {'header':'ISO Link', 'header_format':blueheader}]})
###WRITE INFO
count = 1
for item in ISOSC13_info:
    SC13worksheet.write(count, 0, ISOSC13_info[item]["TITLE"], body_format)
    SC13worksheet.write(count, 1, ISOSC13_info[item]["ID_NUMBER"], body_format)
    SC13worksheet.write(count, 2, ISOSC13_info[item]["COMPLETION_CODE_TEXT"], body_format)

    stagedate = datetime.strptime(ISOSC13_info[item]["STAGE_DATE"], "%Y-%m-%d")
    SC13worksheet.write(count, 3, stagedate, date_format)

    SC13worksheet.write(count, 4, ISOSC13_info[item]["PROGRESS_CLASS"], body_format)
    SC13worksheet.write(count, 5, ISOSC13_info[item]["LINK"], body_format)

    SC13worksheet.write(count, 6, " ")
    count += 1





##############################ISO SC 14##########################################
###LOAD INFO FROM SOURCE
with open('ISOSC14temp.txt') as f:
    ISOSC14_info = json.load(f)

###CREATE SHEETS
SC14worksheet = workbook.add_worksheet('ISO - SC14')

###Conditionally formats ballot close dates to highlight those closing soon
SC14worksheet.conditional_format('G2:G'+str(len(ISOSC14_info)+1), {'type':'formula',
                                                           'criteria':'=AND($G2-TODAY()>=0, $G2-TODAY()<=7)',
                                                           'format':this_week})
SC14worksheet.conditional_format('G2:G'+str(len(ISOSC14_info)+1), {'type':'formula',
                                                           'criteria':'=AND($G2-TODAY()>=0, $G2-TODAY()<=21)',
                                                           'format':next_week})
SC14worksheet.conditional_format('G2:G'+str(len(ISOSC14_info)+1), {'type':'formula',
                                                           'criteria':'=AND($G2-TODAY()>=0, $G2-TODAY()<=35)',
                                                           'format':this_month})
SC14worksheet.conditional_format('G2:G'+str(len(ISOSC14_info)+1), {'type':'formula',
                                                           'criteria':'=AND($G2-TODAY()>=0, $G2-TODAY()<=49)',
                                                           'format':next_month})

###SET WIDTHS
SC14worksheet.set_column('A:A', 35)
SC14worksheet.set_column('B:B', 20)
SC14worksheet.set_column('C:C', 27)
SC14worksheet.set_column('D:D', 22.5)
SC14worksheet.set_column('E:E', 27)
SC14worksheet.set_column('F:F', 17.5)
SC14worksheet.set_column('G:G', 17.5)
SC14worksheet.set_column('H:H', 20.5)
SC14worksheet.set_column('I:I', 13.5)
SC14worksheet.set_column('J:J', 10.5)
SC14worksheet.set_column('K:K', 10.5)


###CREATE TABLE
SC14worksheet.add_table(0,0,len(ISOSC14_info),10, {'style':'Table Style Medium 21',
                                          'columns':[{'header':'Title', 'header_format':orangeheader},
                                                     {'header':'ID Number', 'header_format':orangeheader},
                                                     {'header':'Completion Status', 'header_format':orangeheader},
                                                     {'header':'Completion Status Date', 'header_format':orangeheader},
                                                     {'header':'In Progress/Published/Withdrawn', 'header_format':orangeheader},
                                                     {'header':'AIAA Ballot Type', 'header_format':orangeheader},
                                                     {'header':'Ballot Close Date', 'header_format':orangeheader},
                                                     {'header':'Proposed US Position', 'header_format':orangeheader},
                                                     {'header':'Project Lead', 'header_format':orangeheader},
                                                     {'header':'ISO Link', 'header_format':orangeheader},
                                                     {'header':'AIAA Link', 'header_format':orangeheader}]})


###WRITE INFO
count = 1
for item in ISOSC14_info:
    SC14worksheet.write(count, 0, ISOSC14_info[item]["TITLE"], body_format)
    SC14worksheet.write(count, 1, ISOSC14_info[item]["ID_NUMBER"], body_format)
    SC14worksheet.write(count, 2, ISOSC14_info[item]["COMPLETION_CODE_TEXT"], body_format)

    stagedate = datetime.strptime(ISOSC14_info[item]["STAGE_DATE"], "%Y-%m-%d")
    SC14worksheet.write(count, 3, stagedate, date_format)

    SC14worksheet.write(count, 4, ISOSC14_info[item]["PROGRESS_CLASS"], body_format)
    SC14worksheet.write(count, 5, ISOSC14_info[item]["BALLOT_TYPE"], body_format)

    try:
        closedate = datetime.strptime(ISOSC14_info[item]["CLOSE_DATE"], "%Y-%m-%d")
        SC14worksheet.write(count, 6, closedate, date_format)
    except:
        SC14worksheet.write(count, 6, "", body_format)
    
    SC14worksheet.write(count, 7, ISOSC14_info[item]["PROPOSED_POSITION"], body_format)
    SC14worksheet.write(count, 8, ISOSC14_info[item]["PROJECT_LEAD"], body_format)
    SC14worksheet.write(count, 9, ISOSC14_info[item]["LINK"], body_format)
    if ISOSC14_info[item]["AIAA_LINK"] == "":
        SC14worksheet.write(count, 10, " ", body_format)
    else:
        SC14worksheet.write(count, 10, ISOSC14_info[item]["AIAA_LINK"], body_format)
    SC14worksheet.write(count, 11, " ")
    count += 1










workbook.close()
