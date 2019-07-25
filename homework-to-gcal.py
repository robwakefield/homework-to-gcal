from __future__ import print_function
from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options  

from time import sleep

import os

cwd = os.getcwd() + "\\"
#print(cwd)
calendarID = "ID_OF_YOUR_PERSONAL_CALENDAR"

def addToCalendar(title, description, date, color):

    # build scopes using creds
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    store = file.Storage(cwd + 'storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(cwd + 'client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    GCAL = discovery.build('calendar', 'v3', http=creds.authorize(Http()))

    EVENT = {
        # title
        'summary': title,
        # desc
        'description': description,
        # color of event
        'colorId': color,
        # all day event
        'start':  {'date': date}, # '2019-03-21'
        'end':    {'date': date},
        # notify at 5pm day before
        'reminders': {
        'useDefault': False,
        'overrides': [
          {'method': 'popup', 'minutes': (24 - 17) * 60},
        ],
        },
        # set user as not busy
        'transparency': "transparent",

    }

    e = GCAL.events().insert(calendarId=calendarID,
            sendUpdates='all', body=EVENT).execute()

    # give success output
    print('Added to calendar: ' + title)

def searchCalendar(classCode, description):

    SCOPES = 'https://www.googleapis.com/auth/calendar'
    store = file.Storage(cwd + 'storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(cwd + 'client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    GCAL = discovery.build('calendar', 'v3', http=creds.authorize(Http()))

    page_token = None
    while True:
      events = GCAL.events().list(calendarId=calendarID,q=classCode, pageToken=page_token).execute()
      for event in events['items']:
        try:
            if event['description'] == description:
                #print(event['description'])
                return True
        except:
            print("no description")
      page_token = events.get('nextPageToken')
      if not page_token:
        break
    return False

def getData(username, password):
    output = []
    # create url with credentials (this is using my school's homework page")
    url = "https://" + username + ":" + password + "@slg.bournemouth-school.org/SLG/Students/SLGPages/ViewHomework.aspx"
    # start webdriver
    chrome_options = Options()
    # run headless ass no need to see what is happening (also runs faster)
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    # open webpage
    driver.get(url)
    # get homework elements
    try:
        homeworks = driver.find_elements_by_xpath("//div[contains(@class,'ganttview-block-text')]")
    except:
        ###ERROR: unable to load elements
        print("unable to load homework elements")
        driver.quit()
        return []
    # verify that homeworks were found
    if len(homeworks) <= 0:
        ###ERROR: no homeworks found
        print("no homework elements found")
        driver.quit()
        return []

    # loop through each homework
    for hw in homeworks:
        # click on homework bar
        actions = ActionChains(driver)
        actions.move_to_element_with_offset(hw, 0, 0).click().perform()

        # wait for popup to appear
        # FINDS ALL DIFFERENT PARTS OF HOMEWORK INFO THAT IS NEEDED
        try:
            parent = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "homeworkDetailToolTip")))
            try:
                title = parent.find_element_by_tag_name("h3").text
            except:
                title = "n/a"
            try:
                description = parent.find_element_by_tag_name("blockquote").text
            except:
                description = "n/a"
            try:
                setBy = parent.find_element_by_class_name("popSetByData").text
            except:
                setBy = "n/a"
            try:
                dueOn = parent.find_element_by_class_name("popDueOnData").text
            except:
                dueOn = "n/a"
            try:
                daysLeft = parent.find_element_by_class_name("popDaysLeftData").text
            except:
                daysLeft = "n/a"
        except:
            ###ERROR: popup did not appear
            print("popup did not appear")
            driver.quit()
            return []
            
        # wait
        sleep(1)
        # try to click on close button
        try:
            close_button = driver.find_element_by_class_name("closeButtonContainer")
            actions = ActionChains(driver)
            actions.move_to_element(close_button).click().perform()
        except:
            ###ERROR: could not find close button
            print("could not find close button")
            driver.quit()
            return []
        # store all information gathered in a dictionary
        current = {
        "subject": hw.text,
        "title": title,
        "description": description,
        "setBy": setBy,
        "dueOn": dueOn,
        "daysLeft": daysLeft
        }
        output.append(current)
        # wait
        sleep(1)

    # close webdriver
    driver.quit()

    return output


def sanitize(inData):
    # clean up some data sp that fields not found are accounted for
    temp = []
    for element in inData:
        try:
            if element['title'] == "n/a":
                continue
            long = element['setBy']
            setBy = long[:long.find(" on ")]
            setOn = long[long.find(" on ") + 4:long.find(" for ")]
            classCode = long[long.find("-") + 1:long.find("-") + long[long.find("-"):].find(" ")]
            daysLeft = int(element['daysLeft'][:element['daysLeft'].find(" ")])

            ## changes for personal choice
            if element['subject'].capitalize() == "Mathematic":
                subject = "Maths"
            elif element['subject'].capitalize() == "F. maths":
                subject = "F. Maths"
            else:
                subject = element['subject'].capitalize()
                
            current = {
            "subject": subject,
            "title": element['title'].capitalize(),
            "description": element['description'],
            "setOn": setOn,
            "dueOn": element['dueOn'],
            "daysLeft": daysLeft,
            "setBy": setBy,
            "classCode": classCode
            }
            temp.append(current)
        except:
            print("error on sanitize")

    output = []
    while len(temp) > 0:
        lowest = 100
        for element in temp:
            if element['daysLeft'] < lowest:
                lowest = element['daysLeft']
                lowest_element = element
        output.append(lowest_element)
        temp.remove(lowest_element)

    return output

def check(inData):
    # checks for duplicates which occur sometimes
    output = []
    for i in range(len(inData)):
        match = False
        for j in range(len(inData)):
            if not inData[i] == inData[j]:
                if inData[i]['description'] == inData[j]['description'] and inData[i]['description'] != "n/a":
                    match = True
        if not match:
            output.append(inData[i])
        else:
            print("Removed: " + inData[i]['title'])
            
    return output

def getDate(raw):
    # works out date in format that gcal accepts
    months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
    raw = raw[raw.find(",") + 2:].replace(",", "")
    raw = raw.split(" ")
    month = str(months.index(raw[0]) + 1)
    if len(month) < 2:
        month = "0" + month
    day = raw[1]
    year = raw[2]
    return year + "-" + month + "-" + day

def main():

    # used to set the colour of gcal event
    colorIDs = ["","F. Maths","","Physics","","Computing","Maths"]
    
    # login information
    username = "USERNAME"
    password = "PASSWORD"

    try:
        data = getData(username, password)
        data = sanitize(data)
        data = check(data)
        #print(data)
        print(len(data), "homeworks found")
        for d in data:
            try:
                title = "[" + d['subject'] + "] - " + d['title']
                maxLength = 100
                # splice string if too long
                if len(title) > maxLength:
                    title = title[:maxLength] + "..."
                    description = "..." + title[maxLength:]
                else:
                    description = ""
                description += d['description']
                if not description == "n/a":
                    description += "\n\nSet By: " + d['setBy'] + " [" + d['classCode'] + "]\nSet On: " + d['setOn']
                else:
                    description = "Set By: " + d['setBy'] + " [" + d['classCode'] + "]\nSet On: " + d['setOn']
                date = getDate(d['dueOn'])
                colorID = str(colorIDs.index(d['subject']))
                try:
                    if not searchCalendar(d['classCode'], description):
                        addToCalendar(title, description, date, colorID)
                    else:
                        print("already in calendar")
                except Exception as e:
                    print(e)
                    print("unable to add to calendar");
            except Exception as e:
                print(e)
                print("error in main loop")
    except Exception as e:
        print(e)
        print("could not get data")

try:
    main()
except Exception as e:
    print(e)
    print("MAIN ERROR")
    sleep(10)
