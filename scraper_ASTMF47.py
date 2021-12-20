from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import json

def to_title_case(word):
    return word[0].upper() + word[1:len(word)].lower()


###Open Main ASTMF47 page
driver = webdriver.Chrome()
driver.get("https://member.astm.org/MyASTM/MyCommittees/StandardsTracking")

###Wait for page to load
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[id="userName"')))
except:
    print("timeout login")

###Login
username = driver.find_element(By.CSS_SELECTOR, 'input[id="userName"')
password = driver.find_element(By.ID, "password")
login = driver.find_element(By.CSS_SELECTOR, 'button[onclick="passEnc()"')

username.send_keys("james.davis02@faa.gov")
password.send_keys("Starlord69")
login.click()

###Navigate to standards list page
##Click standards tracking
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-testid="standardsTrackingLink')))
except:
    print("timeout standards tracking link")

standards_tracking = driver.find_element(By.CSS_SELECTOR, 'a[data-testid="standardsTrackingLink"')
standards_tracking.click()

##Click All F47
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-testid="standardsDetailsLink')))
except:
    print("timeout all f47 link")

standards_tracking = driver.find_element(By.CSS_SELECTOR, 'a[data-testid="standardsDetailsLink"')
standards_tracking.click()


###Get links to standards
#Wait for page to load
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "astm-link")))
except:
    print("timeout get links")
    
#gets active standards
standard_link_objs = driver.find_elements(By.XPATH, "//a[contains(@href, 'astm.org/F')]")
active_links = []
for obj in standard_link_objs:
    active_links.append(obj.get_attribute('href'))
    
#gets proposed new standards
new_links = []
standard_link_objs = driver.find_elements(By.XPATH, "//a[contains(@href, '/WorkItemDetails/')]")
for obj in standard_link_objs:
    new_links.append(obj.get_attribute('href'))


data = {}


###Get data from each ACTIVE standard page
count = 0
for standardpage_num in range(len(active_links)):
    count += 1 #debug

    #Goes to ASTM F47 standard information page
    LINK = active_links[standardpage_num]
    driver.get(LINK)

    
    #Might need to select FAA account
    try:
        element = WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((By.CLASS_NAME, "account-select-button")))
        driver.find_element(By.CLASS_NAME, "account-select-button").click()
    except:
        print("timeout active standard")
        
    

    last_updated = driver.find_element(By.CLASS_NAME, "last-updated").get_attribute("innerHTML").replace("Last Updated: ","")
    ID_number = driver.find_element(By.CLASS_NAME, "sku").get_attribute("innerHTML").replace("&nbsp;"," ")
    title = driver.find_element(By.CLASS_NAME, "name").get_attribute("innerHTML")

    print(last_updated)
    print(ID_number)
    print(title)
    print(LINK)
    print("-----------------------------------")

    data[ID_number] = {'TITLE':title,
                       'ID_NUMBER':ID_number,
                       'SUBCOMMITTEE':'',
                       'START_DATE':'',
                       'PROJECT_LEAD':'',
                       'STATUS':'Published',
                       'LAST_UPDATED':last_updated,
                       'LINK':LINK}
    
    #if count > 0:
    #    break


###Get data from each NEW PROPOSED standard page
count = 0
for standardpage_num in range(len(new_links)):
    count += 1 #debug

    #Goes to ASTM F47 standard information page
    LINK = new_links[standardpage_num]
    driver.get(LINK)

    
    #Might need to select FAA account
    try:
        element = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.CLASS_NAME, "account-select-button")))
        driver.find_element(By.CLASS_NAME, "account-select-button").click()
    except:
        print("timeout new proposed")
    
    

    ID_number = driver.find_element(By.CLASS_NAME, "astm-type-heading--h3").get_attribute("innerHTML")
    title = driver.find_element(By.CLASS_NAME, "astm-type-heading--h2").text
    subcommittee = driver.find_element(By.XPATH, "//a[contains(@href, 'SUBCOMMIT/F')]").get_attribute("innerHTML")

    info2s = driver.find_elements(By.CLASS_NAME, "info2")
    date_initiated = info2s[0].get_attribute("innerHTML")
    project_lead = info2s[1].get_attribute("innerHTML")
    status = info2s[4].get_attribute("innerHTML")
        
    
    print(ID_number)
    print(title)
    print(subcommittee)
    print(date_initiated)
    print(project_lead)
    print(status)
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    data[ID_number] = {'TITLE':title,
                       'ID_NUMBER':ID_number,
                       'SUBCOMMITTEE':subcommittee,
                       'START_DATE':date_initiated,
                       'PROJECT_LEAD':project_lead,
                       'STATUS':status,
                       'LAST_UPDATED':'',
                       'LINK':'https://www.astm.org/get-involved/technical-committees/committee-f47/subcommittee-f47#'}
    
    #if count > 0:
    #    break


###Write to file
with open('ASTMF47temp.txt', 'w') as file:
    file.write(json.dumps(data))


driver.close()
