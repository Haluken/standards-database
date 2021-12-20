from selenium import webdriver
from selenium.webdriver.common.by import By
import re
import json

###Finds the intersection of two lists
def intersection(lst1, lst2):
    return list(set(lst1) & set(lst2))

###Makes Words Look Like This
def to_title_case(word):
    return word[0].upper() + word[1:len(word)].lower()

    
#########################################ISO SC 13#######################################################
###Converts ISO completion codes to their text format
completion_codes = {'00.00':'Proposal for new project received',
                    '00.20':'Proposal for new project under review',
                    '00.60':'Close of review',
                    '00.98':'Proposal for new project abandoned',
                    '00.99':'Approval to ballot proposal for new project',
                    '10.00':'Proposal for new project registered',
                    '10.20':'New project ballot initiated',
                    '10.60':'Close of voting',
                    '10.92':'Proposal returned to submitter for further definition',
                    '10.98':'New project rejected',
                    '10.99':'New project approved',
                    '20.00':'New project registered in TC/SC work programme',
                    '20.20':'Working draft study initiated',
                    '20.60':'Close of comment period',
                    '20.98':'Project deleted',
                    '20.99':'Working draft approved for registration as committee draft',
                    '30.00':'Committee draft registered',
                    '30.20':'Committee draft study/ballot initiated',
                    '30.60':'Close of voting/comment period',
                    '30.92':'Committee draft referred back to working group',
                    '30.98':'Project deleted',
                    '30.99':'Committee draft approved for registration as draft international standard',
                    '40.00':'Draft international standard registered',
                    '40.20':'Draft international standard ballot initiated: 12 weeks',
                    '40.60':'Close of voting',
                    '40.92':'Full report circulated: draft international standard referred back to TC or SC',
                    '40.93':'Full report circulated: decision for new draft international standard ballot',
                    '40.98':'Project deleted',
                    '40.99':'Full report circulated: draft international standard approved for registration as final draft international standard',
                    '50.00':'Final text received or final draft international standard registered for formal approval',
                    '50.20':'Proof sent to secretariat or final draft international standard ballot initiated: 8 weeks',
                    '50.60':'Close of voting. Proof returned by secretariat',
                    '50.92':'Final draft international standard or proof referred back to TC or SC',
                    '50.98':'Project deleted',
                    '50.99':'Final draft international standard or proof approved for publication',
                    '60.00':'International standard under publication',
                    '60.60':'International standard published',
                    '90.20':'International standard under systematic review',
                    '90.60':'Close of review',
                    '90.92':'International standard to be revised',
                    '90.93':'International standard confirmed',
                    '90.99':'Withdrawal of international standard proposed by TC or SC',
                    '95.20':'Withdrawal ballot initiated',
                    '95.60':'Close of voting',
                    '95.92':'Decision not to withdraw international standard',
                    '95.99':'Withdrawal of international standard'}


###Open main iso page
driver = webdriver.Chrome()
driver.get("https://www.iso.org/committee/46612/x/catalogue/p/0/u/1/w/0/d/0")

###Gets links to individual ISO standard pages
standardpages = driver.find_elements(By.XPATH, "//a[contains(@href, '/standard/')]")
ISO_links = []
for standardpage in standardpages:
    ISO_links.append(standardpage.get_attribute("href"))



ISO_info = {}
count = 0
###Iterate thru each standard webpage
for standardpage_num in range(len(standardpages)):
    count += 1 #debug

    #Goes to ISO standard information page
    ISO_link = ISO_links[standardpage_num]
    driver.get(ISO_link)

    #Title of standard
    title = driver.find_element(By.CLASS_NAME, "no-uppercase").text

    #Filters out ID number
    ID_number = driver.find_element(By.TAG_NAME, "h1").text
    ID_number_trimmed = re.search("( [0-9]{4,5})(-[0-9]+)?((-|:)([0-9]{4}))?",ID_number).group()

    #Finds stage codes and pairs them with their respective stage date
    stage_code_objlist = driver.find_elements(By.CSS_SELECTOR, "span[class='stage-code'")
    stage_date_objlist = driver.find_elements(By.CSS_SELECTOR, "span[class='stage-date'")
    stage_code_list = []
    stage_date_list = []

    for i in range(len(stage_code_objlist)):
        #Filter out bad stage codes, match up with stage dates, pick last one
        if stage_code_objlist[i].get_attribute('innerHTML')[0] != '<':
            stage_code_list.append(stage_code_objlist[i].get_attribute('innerHTML'))

    for i in range(len(stage_date_objlist)): #add date strings to list
        stage_date_list.append(stage_date_objlist[i].get_attribute('innerHTML'))

    #Starts from the bottom, finds the first (last) one with a stage date, takes that as the most recent progress
    for i in range(len(stage_date_list)-1,0,-1):
        if stage_date_list[i] != '':
            stage_date = stage_date_list[i]
            stage_code = stage_code_list[i]
            break

    #Looks for a date in the reaffirmation header and handles exception if does not exist
    try: 
        reaffirmation_class = driver.find_element(By.CLASS_NAME, "no-margin").text
        reaffirmation_date = re.search("[0-9]{4}", reaffirmation_class).group()
    except: #an unlimited try except clause *cringe*
        reaffirmation_date = 'current'

    #Looks for a publication date, default Not Published
    try:
        publication_date = driver.find_element(By.CSS_SELECTOR, "span[itemprop='releaseDate']").get_attribute('innerHTML')
    except:
        publication_date = "Not published"        

    #Determine if standard is published, in progress, or withdrawn
    stage_code_int = float(stage_code)
    if stage_code_int < 51: #cutoff for in progress is 50.99
        progress_class = "In Progress"
    if stage_code_int >= 60 and stage_code_int < 91: #cutoff for published is 60.00 to 90.99
        progress_class = "Published"
    if stage_code_int > 95: #cutoff for withdrawal is 95.20
        progress_class = "Withdrawn"

    ##Assemble info data    
    ISO_info[ID_number_trimmed] = {'TITLE':title,
            'ID_NUMBER':ID_number,
            'STAGE_DATE':stage_date,
            'COMPLETION_CODE':stage_code,
            'COMPLETION_CODE_TEXT':completion_codes[stage_code],
            'REAFFIRMATION_DATE':reaffirmation_date,
            'PUBLICATION_DATE':publication_date,
            'LINK':ISO_link,
            "BALLOT_TYPE":"",
            "PROJECT_LEAD":"",
            "PROPOSED_POSITION":"",
            "CLOSE_DATE":"",
            "AIAA_LINK":"",
            "PROGRESS_CLASS":progress_class}

    #if count > 0:
    #    break




###Write to file
with open('ISOSC13temp.txt', 'w') as file:
    file.write(json.dumps(ISO_info))



driver.close()
