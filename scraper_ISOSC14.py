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


#########################################ISO SC 14#######################################################
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
driver.get("https://www.iso.org/committee/46614/x/catalogue/")

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

    #Assemble info data    
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

    #if count > 40:
    #    break





################################################AIAA################################################
###Gets links to AIAA standard pages
driver.get("https://aiaa.kavi.com/apps/org/workgroup/portal/my_ballots.php#")


###Finds username, password fields
username = driver.find_element(By.ID, "__ac_name")
password = driver.find_element(By.ID, "__ac_password")
remember = driver.find_element(By.CSS_SELECTOR, "label[for='__ac_remember_1'")
login = driver.find_element(By.CSS_SELECTOR, "input[name='submit'")

###Logs in using tara halt's login
username.send_keys("Tara.R-CTR.Halt")
password.send_keys("ISOstandards2021")
remember.click()
login.click()


###Get data from AIAA pages
AIAA_ID_numbers = []
ballot_types = []
project_leads = []
proposed_positions = []
close_dates = []
AIAA_links = []
breakout = False
while not breakout:  
    ballot_titles = driver.find_elements(By.CLASS_NAME, "ballottitle")
    for ballot_title in ballot_titles:
        text = ballot_title.get_attribute("innerHTML")
        #Determines ID number of standard
        try:
            AIAA_ID_number = re.search("( [0-9]{4,5})(-[0-9]+)?((-|:)([0-9]{4}))?", text).group()
        except:
            AIAA_ID_number = ""

        #Determines if standard is up for systematic review or approval - All systematic review ballots have systematic review in title
        try:
            ballot_type = re.search("Systematic Review", text).group()
        except:
            ballot_type = "Approval"
        AIAA_ID_numbers.append(AIAA_ID_number)
        ballot_types.append(ballot_type)


    #Finds project leads and proposed positions. The info is contained in three or four lines and may or may not exist
    #adding the results from a regex scan of each line to a buffer, which the actual value is taken from
    project_lead_objlist = driver.find_elements(By.CSS_SELECTOR, 'span[class = "itemdetails"]>*')
        
    PL_buffer = []
    PP_buffer = []
    for i in range(len(project_lead_objlist)):
        obj = project_lead_objlist[i]
        nextobj = project_lead_objlist[min(i+1,len(project_lead_objlist)-1)]
        
        text = obj.get_attribute("innerHTML")
        try:
            project_lead = re.search("\((\w*\W*)?([A-Z][a-z]* [A-Z][a-z]*)([^\)]\W*.*)?\)", text).group(2)
        except:
            project_lead = "null"

        try:
            proposed_position = re.search('"([\w ]*)"', text).group(1)
        except:
            proposed_position = "null"

        PL_buffer.append(project_lead)
        PP_buffer.append(proposed_position)
        #Each field starts with a class == ballottitle
        if nextobj.get_attribute("class") == "ballottitle" or i>=(len(project_lead_objlist)-1):
            #Scan thru both buffers, take anything thats not null as the value. reset buffers
            for name in PL_buffer:
                if name != "null":
                    project_lead = name
            for position in PP_buffer:
                if position != "null":
                    proposed_position = to_title_case(position)
                    break
            PL_buffer = []
            PP_buffer = []
            project_leads.append(project_lead)
            proposed_positions.append(proposed_position)

    #Finds close dates and links. this is easy
    close_dates.extend([x.get_attribute("innerHTML") for x in driver.find_elements(By.CSS_SELECTOR, "td[class = 'close_date']")])
    AIAA_links.extend([x.get_attribute("href") for x in driver.find_elements(By.CLASS_NAME, 'ballot-status')])

    #Goes to the next page, breaks if on last page already
    next_page_button = driver.find_element(By.CLASS_NAME, "next")
    if not next_page_button.get_attribute("href"):
        breakout = True
    next_page_button.click()


    
###Build AIAA_info
AIAA_info = {}
for i in range(len(AIAA_ID_numbers)):
    AIAA_info[AIAA_ID_numbers[i]] = {"BALLOT_TYPE":ballot_types[i],
                                     "PROJECT_LEAD":project_leads[i],
                                     "PROPOSED_POSITION":proposed_positions[i],
                                     "CLOSE_DATE":close_dates[i],
                                     "AIAA_LINK":AIAA_links[i]}
###Scan and merge ISO_info and AIAA_info
for id_number in AIAA_info:
    if id_number in ISO_info:
        ISO_info[id_number].update(AIAA_info[id_number])

"""
#print(ISO_info)
print(AIAA_ID_numbers)
print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
print([x for x in ISO_info.keys()])
print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
print(intersection(AIAA_ID_numbers,[x for x in ISO_info.keys()]))
"""

###Write to file
with open('ISOSC14temp.txt', 'w') as file:
    file.write(json.dumps(ISO_info))




driver.close()
