#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!pip install --upgrade gspread
#!pip install gspread_formatting
#!pip install selenium
#!pip install numpy
#!pip install pandas


# In[2]:


import pandas as pd
import gspread
import numpy as np
from selenium import webdriver
import selenium.common.exceptions
import time
import datetime
from tkinter import *
from tkinter import scrolledtext
from tkinter import ttk
import os
import json

setting = json.load(open("settings.json"))

weeksList = [datetime.date(2021, 11, 15)]
for programLength in range(1, 24):
    weeksList.append(weeksList[programLength - 1] + datetime.timedelta(days = 7))
for item in range(len(weeksList)):
    splited = weeksList[item].strftime('%Y-%m-%d').split('-')
    weeksList[item] = 'Week ' + str(item + 1) + ' ' + splited[1].lstrip('0') + '/' + splited[2].lstrip('0')

reviewPayCheckWeeks = [datetime.date(2021, 11, 19)]
for programLength in range(1, 13):
    reviewPayCheckWeeks.append(reviewPayCheckWeeks[programLength - 1] + datetime.timedelta(days = 14))
for item2 in range(len(reviewPayCheckWeeks)):
    splited = reviewPayCheckWeeks[item2].strftime('%Y-%m-%d').split('-')
    reviewPayCheckWeeks[item2] = 'Review Paycheck ' + splited[1].lstrip('0') + '/' + splited[2].lstrip('0')


# In[3]:


def initialSetup():
    global USERNAME
    global PASSWORD
    global ALL_FLAG
    global gc
    global ssheet
    global gsheets
    global _target
    ALL_FLAG = False

    employeesOnTeam = teamCheck()
    setting["Username"] = username.get()
    setting["Week"] = weekdrop.get()
    setting["Review Pay Check Week"] = review.get()
    setting["Team"] = teamDrop.get()
    setting["Employee"] = targetDropdown.get()
    with open("settings.json", "w") as f:
        f.seek(0)
        f.write(json.dumps(setting, indent=4))
        f.truncate()

    USERNAME = setting["Username"]
    PASSWORD = setting["LOGINS"][USERNAME]
    # setup
    gc = gspread.service_account(filename=os.path.join(os.getcwd(), "service_account.json"))

    ssheet = teamDrop.get()
    gsheets = gc.open_by_key(setting["TEAMSHEETS"][ssheet])

    if targetDropdown.get() == 'All' and not reviewFlag.get():
        _target = employeesOnTeam[1]
        targetSetup(_target)
        ALL_FLAG = True
    else:
        _target = targetDropdown.get()
        targetSetup(targetDropdown.get())
        ALL_FLG = False


# In[4]:


def targetSetup(target):
    global employees
    global times
    global week
    global paddedWeek
    global finalReviewWeek
    global reviewcolum
    global timecolum
    global cws
    global __target
    # read data and put it in a dataframe

    try:
        target + ''
        __target = target
    except Exception:
        __target = target[0]

    cws = gsheets.worksheet(str(__target) + "'s Sites")
    sheets = gsheets.worksheet(str(__target) + "'s Sites").get_all_values()
    week = weekdrop.get().split(' ')[2]
    
    
    if reviewFlag.get():
        reviewcolum = gsheets.worksheet(str(__target) + "'s Sites").find(str(review.get())).col
        reviewWeek = review.get()
        paddedReviewWeek = reviewWeek.split(' ')
        finalReviewWeek = paddedReviewWeek[2]
    else:
        timecolum = gsheets.worksheet(str(__target) + "'s Sites").find(str(week)).col
        paddedWeek = week.split('/')
        paddedWeek[0] = paddedWeek[0].zfill(2)
        paddedWeek[1] = paddedWeek[1].zfill(2)

    # first arg is range of rows, second is the row used as header
    if rangeFlag.get():
        df = pd.DataFrame(sheets[int(rangemin.get()) - 3:int(rangemax.get())], columns=sheets[0])
    else:
        df = pd.DataFrame(sheets[3:], columns=sheets[0])
    # print(df.head())

    filtered_df = df[['Employee ID', 'Youth Name', week]].copy()

    for column_name in filtered_df.columns[2:]:
        filtered_df[column_name] = pd.to_numeric(filtered_df[column_name], errors='coerce')

    #filtered_df['Employee ID'] = pd.to_numeric(filtered_df['Employee ID'], errors='coerce')
    filtered_df = filtered_df.fillna(0)
    #filtered_df['Employee ID'] = filtered_df['Employee ID'].astype(np.int64)

    #filtered_df.dtypes
    # for column_name in filtered_df.columns[2:]:
    # filtered_df[column_name] = filtered_df[column_name].astype(np.int64)
    currentEmployeeLast = ''
    currentEmployeeFirst = ''
    #print(final_df.head(30))

    final_df = filtered_df[['Employee ID', 'Youth Name', week]]

    employees = final_df['Employee ID'].to_list()
    times = final_df[week].to_list()

    auditLogStart(__target)


# In[5]:


def auditLogStart(name):
    logLabel.config(state=NORMAL)
    logLabel.delete('1.0', END)
    logLabel.insert(END,'======================================================================\n')
    logLabel.insert(END,str("Beginning New Audit and Time Entry Log For " + name + ":\n"))
    logLabel.insert(END,'======================================================================\n')
    logLabel.config(state=DISABLED)
    logLabel.update()


# In[6]:


class Driver:
    def __init__(self):
        """self.options = webdriver.ChromeOptions()
        self.options.add_argument(r"--user-data-dir=" + setting["Chrome Profile Path"])
        self.options.add_argument(r'--profile-directory=' + setting["Profile"])
        self.driver = webdriver.Chrome(options=self.options)"""

        self.driver = webdriver.Chrome()

    def _test(self):
        self.driver.get('http://access.boston.gov/')

    def login(self, USERNAME, PASSWORD):
        self._test()
        self.driver.find_element_by_xpath('//*[@id="username"]').clear()
        self.driver.find_element_by_xpath('//*[@id="password"]').clear()
        self.driver.find_element_by_xpath('//*[@id="username"]').send_keys(USERNAME)
        self.driver.find_element_by_xpath('//*[@id="password"]').send_keys(PASSWORD)
        try:
            self.driver.find_element_by_xpath('/html/body/main/div/form/div[3]/div[1]/div[3]/button').click()
        except Exception:
            time.sleep(0)
        """try:
            self.driver.find_element_by_xpath('/html/body/main/div/form/div[3]/div[1]/div[3]/button').click()
        except Exception:
            try:
                self.driver.find_element_by_xpath('//*[@id="app-icon-BAIS HCM"]').click()
            except Exception:
                time.sleep(0)
            time.sleep(0)
        """
        while True:
            try:
                self.driver.find_element_by_xpath('//*[@id="app-icon-BAIS HCM"]').click()
                break
            except Exception:
                time.sleep(0.2)
        self.driver.find_element_by_xpath('//*[@id="app-icon-BAIS HCM"]').click()
        handles = self.driver.window_handles
        for i in range(len(handles) - 1):
            self.driver.switch_to.window(handles[i])
            self.driver.close()
        self.driver.switch_to.window(handles[len(handles) - 1])

        try:
            self.driver.find_element_by_xpath('/html/body/main/div/form/div[3]/div[1]/div[3]/button').click()
        except Exception:
            time.sleep(0)


# In[7]:


def timeSheets():
    chrome.driver.switch_to.default_content()
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="pthnavbca_MYFAVORITES"]').click()
            time.sleep(0.15)
            chrome.driver.find_element_by_xpath('//*[@id="crefli_fav_HC_TL_MSS_EE_PRD_GBL3"]/a').click()
            break
        except Exception:
            time.sleep(0.25)
    chrome.driver.switch_to.frame('ptifrmtgtframe')
    time.sleep(0.5)


# In[8]:


def reviewPayPage():
    chrome.driver.switch_to.default_content()
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="pthnavbca_MYFAVORITES"]').click()
            time.sleep(0.15)
            chrome.driver.find_element_by_xpath('//*[@id="crefli_fav_HC_PAY_CHECK_USA5"]/a').click()
            break
        except Exception:
            time.sleep(0.25)
    chrome.driver.switch_to.frame('ptifrmtgtframe')
    time.sleep(0.5)


# In[9]:


def loginPS():
    global chrome
    chrome = Driver()
    chrome.login(USERNAME, PASSWORD)


# In[10]:


def checkTime(employeeID, timesheet):
    global employee
    if timesheet == 0.0 or employeeID == '0' or employeeID == '':
        return 0
    try:
        chrome.driver.switch_to.frame('ptifrmtgtframe')
    except Exception:
        time.sleep(0)
    while chrome.driver.find_element_by_xpath('//*[@id="VALUE$1"]').get_attribute('value') != str(employeeID):
        try:
            time.sleep(0.05)
            chrome.driver.find_element_by_xpath('//*[@id="VALUE$1"]').clear()
            chrome.driver.find_element_by_xpath('//*[@id="VALUE$1"]').send_keys(employeeID)
            chrome.driver.find_element_by_xpath('//*[@id="TL_MSS_SRCH_WRK_GET_EMPLOYEES"]').click()
            time.sleep(0.2)
        except Exception:
            time.sleep(0)
    try:
        employee = chrome.driver.find_element_by_xpath('//*[@id="EMPLID$0"]').text
    except selenium.common.exceptions.NoSuchElementException:
        logLabel.config(state=NORMAL)
        logLabel.insert(END,'----------------------------------------------------------------------\n')
        logLabel.insert(END, str(employeeID + ' may be inactive\n'))
        logLabel.insert(END,'----------------------------------------------------------------------\n')
        logLabel.config(state=DISABLED)
        timeSheets()
        return 0
    while employee != employeeID:
        chrome.driver.find_element_by_xpath('//*[@id="TL_MSS_SRCH_WRK_GET_EMPLOYEES"]').click()
        try:
            time.sleep(0.15)
            employee = chrome.driver.find_element_by_xpath('//*[@id="EMPLID$0"]').text
        except selenium.common.exceptions.StaleElementReferenceException:
            time.sleep(0.3)
    while chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY12"]').get_attribute('value') != paddedWeek[0] + '/'+ paddedWeek[1]+'/2021':
        time.sleep(0.1)
        chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY12"]').clear()
        time.sleep(0.1)
        chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY12"]').send_keys(week + '/2021')
        time.sleep(0.1)
        chrome.driver.find_element_by_xpath('//*[@id="TL_RTA_WRK_REFRESH_ICN"]/img').click()
        time.sleep(0.3)
        break
    while True:
        try:
            TIME_ENTERED = format(float(chrome.driver.find_element_by_xpath('//*[@id="TOTAL_RPTD_HRS1$0"]').text), '.2f')
            break
        except Exception:
            time.sleep(0.1)
    while True:
        try:
            FIRSTNAME = chrome.driver.find_element_by_xpath('//*[@id="FIRST_NAME$0"]').text
            LASTNAME = chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').text
            break
        except Exception:
            time.sleep(0.1)
    if TIME_ENTERED == format(timesheet, '.2f'):
        logLabel.config(state=NORMAL)
        logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + TIME_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **CORRECT**\n'))
        logLabel.config(state=DISABLED)
        return 1
    else:
        return -1
    """if auditFlag.get() == 0:
        logLabel.config(state=NORMAL)
        logLabel.insert(END,'**********************************************************************\n')
        logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + TIME_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **INCORRECT**\n'))
        logLabel.insert(END,'**********************************************************************\n')
        logLabel.config(state=DISABLED)"""


# In[11]:


def enterTime(employeeID, timesheet, flag = True):
    FIRSTNAME = chrome.driver.find_element_by_xpath('//*[@id="FIRST_NAME$0"]').text
    LASTNAME = chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').text
    HOURS_ENTERED = format(float(chrome.driver.find_element_by_xpath('//*[@id="TOTAL_RPTD_HRS1$0"]').text), '.2f')
    if not flag:
        if HOURS_ENTERED != timesheet:
            logLabel.config(state=NORMAL)
            logLabel.insert(END,'**********************************************************************\n')
            logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + HOURS_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **INCORRECT**\n'))
            logLabel.insert(END,'**********************************************************************\n')
            logLabel.config(state=DISABLED)
            return -1
        else:
            logLabel.config(state=NORMAL)
            logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + HOURS_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **CORRECT**\n'))
            logLabel.config(state=DISABLED)
            return 1

    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').click()
            break
        except Exception:
            time.sleep(0.2)
            chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').click()
            break
        except Exception:
            time.sleep(0.15)

    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY1"]').get_attribute('value')
            break
        except Exception:
            time.sleep(0.1)
        
    while chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY1"]').get_attribute('value') != week + '/2021':
        try:
            time.sleep(0.1)
            chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY1"]').clear()
            time.sleep(0.1)
            chrome.driver.find_element_by_xpath('//*[@id="DATE_DAY1"]').send_keys(week + '/2021')
            time.sleep(0.1)
            chrome.driver.find_element_by_xpath('//*[@id="TL_LINK_WRK_REFRESH_ICN"]/img').click()
            time.sleep(0.4)
            break
        except Exception:
            try:
                chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').click()
            except Exception:
                time.sleep(0.15)
    
    try:
        if chrome.driver.find_element_by_xpath('//*[@id="DERIVED_TL_WEEK_TL_COMMENTS"]').text.split(' ')[3] == 'inactive':
            time.sleep(0.5)
            if chrome.driver.find_element_by_xpath('//*[@id="DERIVED_TL_WEEK_TL_COMMENTS"]').text.split(' ')[3] == 'inactive':
                logLabel.config(state=NORMAL)
                logLabel.insert(END,'**********************************************************************\n')
                logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + HOURS_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **INACTIVE**\n'))
                logLabel.insert(END,'**********************************************************************\n')
                logLabel.config(state=DISABLED)
                timeSheets()
                return -1
    except Exception:
        time.sleep(0)
    time.sleep(1)
    for x in range(5):
        for weekend in range(6, 8):
            while True:
                try:
                    if chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(weekend) + '$0"]').get_attribute('value') != '0.000000':
                        chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(weekend) + '$0"]').clear()
                        chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(weekend) + '$0"]').send_keys(str('0.000000'))
                        #chrome.driver.find_element_by_xpath('//*[@id="YTRC$0"]').send_keys('')
                        time.sleep(0.2)
                    break
                except Exception:
                    time.sleep(0.2)
        for t in range(1, 6):
            while True:
                try:
                    if t < 5:
                        try:
                            if chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(t) + '$0"]').get_attribute('value') != str(format(timesheet // 5, '.2f') + '0000'):
                                chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(t) + '$0"]').clear()
                                chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY' + str(t) + '$0"]').send_keys(str(format(timesheet // 5, '.2f') + '0000'))
                                #chrome.driver.find_element_by_xpath('//*[@id="YTRC$0"]').send_keys('')
                                time.sleep(0.2)
                            break
                        except Exception:
                            time.sleep(0.2)
                    else:
                        if chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY5$0"]').text != str(format(timesheet % 5 + timesheet // 5, '.2f') + '0000'):
                            chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY5$0"]').clear()
                            chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY5$0"]').send_keys((str(format(timesheet % 5 + timesheet // 5, '.2f') + '0000')))
                            #chrome.driver.find_element_by_xpath('//*[@id="YTRC$0"]').send_keys('')
                        break
                except Exception:
                    chrome.driver.find_element_by_xpath('//*[@id="LAST_NAME$0"]').click()
                    continue
    """time.sleep(0.25)
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY4$0"]').clear()
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY4$0"]').send_keys(str(format(timesheet // 5, '.2f')))
    time.sleep(0.25)
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY5$0"]').clear()
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY5$0"]').send_keys(str(format(timesheet // 5, '.2f')))
    time.sleep(0.25)
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY6$0"]').clear()
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY6$0"]').send_keys(str(format(timesheet // 5, '.2f')))
    time.sleep(0.25)
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY7$0"]').clear()
    chrome.driver.find_element_by_xpath('//*[@id="QTY_DAY7$0"]').send_keys((str(format(timesheet % 5 + timesheet // 5, '.2f'))))"""
    time.sleep(0.3)
    if chrome.driver.find_element_by_xpath('//*[@id="YTRC$0"]').get_attribute('value') != 'SREG':
        chrome.driver.find_element_by_xpath('//*[@id="YTRC$0"]').send_keys(str("SREG"))
    time.sleep(0.3)
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="TL_LINK_WRK_SUBMIT_PB"]').click()
            break
        except Exception:
            time.sleep(0)
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="DERIVED_ETEO_SAVE_PB"]').click()
            break
        except Exception:
            try:
                chrome.driver.find_element_by_xpath('//*[@id="TL_LINK_WRK_SUBMIT_PB"]').click()
            except Exception:
                time.sleep(0.5)
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="TL_LINK_WRK_TL_SELECT_TEXT1"]').click()
            break
        except Exception:
            time.sleep(1)
    while True:
        try:
            FINAL_HOURS_ENTERED = format(float(chrome.driver.find_element_by_xpath('//*[@id="TOTAL_RPTD_HRS1$0"]').text), '.2f')
            break
        except Exception:
            time.sleep(0.15)
    if float(FINAL_HOURS_ENTERED) != timesheet:
        logLabel.config(state=NORMAL)
        logLabel.insert(END,'**********************************************************************\n')
        logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + FINAL_HOURS_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **INCORRECT**\n'))
        logLabel.insert(END,'**********************************************************************\n')
        logLabel.config(state=DISABLED)
        return -1
    else:
        logLabel.config(state=NORMAL)
        logLabel.insert(END, str(employeeID.ljust(7, ' ') + (FIRSTNAME + ' ' + LASTNAME + ' ').ljust(38, ' ') + FINAL_HOURS_ENTERED.ljust(5, ' ') + ' ' + format(timesheet, '.2f').ljust(5, ' ') +' **CORRECT**\n'))
        logLabel.config(state=DISABLED)
        return 1
        #print(FIRSTNAME + ' ' + LASTNAME + ' has ' + FINAL_HOURS_ENTERED + ' hours entered and checked.')


# In[12]:


def reviewPaycheck(employee):
    try:
        chrome.driver.switch_to.frame('ptifrmtgtframe')
    except Exception:
        time.sleep(0.1)
    while True:
        try:
            chrome.driver.find_element_by_xpath('//*[@id="#ICClear"]').click()
            break
        except Exception:
            time.sleep(0.1)
    while True:
        try:
            time.sleep(0.2)
            chrome.driver.find_element_by_xpath('//*[@id="Y_ZZ_PAY_CHK_VW_PAY_END_DT"]').send_keys(str(finalReviewWeek) + '/2021')
            chrome.driver.find_element_by_xpath('//*[@id="Y_ZZ_PAY_CHK_VW_EMPLID"]').send_keys(employee)
            chrome.driver.find_element_by_xpath('//*[@id="#ICSearch"]').click()
            break
        except Exception:
            chrome.driver.find_element_by_xpath('//*[@id="#ICClear"]').click()

    while True:
        try:
            CHECKED_TIME = float(chrome.driver.find_element_by_xpath('//*[@id="PAY_SPCL_EARNS_SPCL_HRS$0"]').text)
            break
        except Exception:
            try:
                if chrome.driver.find_element_by_xpath('//*[@id="win0divSEARCHRESULT"]/h2').text == 'No matching values were found.':
                    logLabel.config(state=NORMAL)
                    logLabel.insert(END, str('No Check Cut For: ' + employee + '\n'))
                    logLabel.config(state=DISABLED)
                    return 0
                else:
                    continue
            except Exception:
                time.sleep(0.2)
            time.sleep(0.2)
    """for x in range(10):
        if x == 9:
            logLabel.config(state=NORMAL)
            logLabel.insert(END, str('No Check Cut For: ' + employee + '\n'))
            logLabel.config(state=DISABLED)
            return 0
        try:
            CHECKED_TIME = float(chrome.driver.find_element_by_xpath('//*[@id="PAY_SPCL_EARNS_SPCL_HRS$0"]').text)
            break
        except Exception:
            time.sleep(1.2)"""
    while True:
        try:
            time.sleep(0.2)
            chrome.driver.find_element_by_xpath('//*[@id="#ICList"]').click()
            break
        except Exception:
            time.sleep(1)
    if CHECKED_TIME % 1.00 == 0.0:
        return int(CHECKED_TIME)
    return CHECKED_TIME


# In[13]:


def auditTime():
    offset = int(rangemin.get()) - 2 if rangeFlag.get() else 4
    timeSheets()
    time.sleep(0.3)
    entered = 0
    zeros = 0
    failCount = 0
    timer = time.perf_counter()
    logLabel.config(state=NORMAL)
    logLabel.insert(END, str(('Employee Info').ljust(45, ' ') + "PPSF".ljust(5, ' ') + " " + "SPSH".ljust(5, ' ') + '\n'))
    logLabel.insert(END,'----------------------------------------------------------------------\n')
    logLabel.config(state=DISABLED)
    for i in range(len(employees)):
        while True:
            try:
                auditResult = checkTime(employees[i], times[i])
                timeFlag = 0
                logLabel.see('end')
                logLabel.update()
                if (auditResult == 0):
                    zeros = zeros + 1
                elif (auditResult == -1 and auditFlag.get() == 0):
                    timeFlag = enterTime(employees[i], times[i])
                elif (auditResult == -1 and auditFlag.get() == 1):
                    timeFlag = enterTime(employees[i], times[i], False)
                if (timeFlag == 1 or auditResult == 1):
                        entered = entered + 1
                        cws.format(chr(ord('@')+timecolum) + str(offset+i),{
                            "backgroundColor": {
                                "red": 0.788235294,
                                "green": 0.854901961,
                                "blue": 0.97254902
                            }
                        })
                elif (timeFlag == -1):
                    failCount = failCount + 1
                progress['value'] = 100*(i+1)/len(employees)
                progressText.config(text = 'Processed: ' + str(i+1) + '/' + str(len(employees)))
                break
            except Exception:
                timeSheets()
                time.sleep(0.3)
    timestop = time.perf_counter()
    logLabel.config(state=NORMAL)
    logLabel.insert(END, str('=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-\n'))
    logLabel.insert(END, str('Total of ' + str(len(employees)) + ' were processed\n'))
    logLabel.insert(END, str('Time Entered For ' + str(entered) + ' Youth\n' ))
    if failCount > 0:
        logLabel.insert(END, str(str(failCount) + ' had time entry error, please review audit log\n'))
    logLabel.insert(END, str(str(zeros) + ' Youths had no time entered\n'))
    logLabel.insert(END, str('Operation Took ' +  str(int((timestop - timer) // 60)) + 'm' + str(int((timestop - timer) % 60)) + 's\n'))
    logLabel.insert(END, str('For an average of ' + str(str(format((timestop - timer) / entered, '.2f'))) + ' seconds per youth\n'))
    logLabel.see('end')
    logLabel.update()
    logLabel.config(state=DISABLED)
    try:
        os.mkdir('auditLogs')
    except Exception:
        time.sleep(0)
    system_time = datetime.datetime.now()
    timestamp = system_time.strftime("%m-%d-%Y_%H%M")
    weekfixed = week.split('/')
    auditLogtxt = open('auditLogs/' + __target + '_auditLog_' + weekfixed[0] + '-' + weekfixed[1] + '_' + timestamp + '.txt', 'w')
    auditLogtxt.writelines(logLabel.get('1.0', END).strip(' \t\n\r'))


# In[14]:


def reviewPay():
    if rangeFlag.get():
        offset = int(rangemin.get()) + 4
    else:
        offset = 4
    for j in range(len(employees)):
        if employees[j] == '':
            continue
        while True:
            try:
                #cws.update_cell(int(j+4), int(reviewcolum), reviewPaycheck(employees[j]))
                cws.update(chr(ord('@')+reviewcolum) + str(j+offset), reviewPaycheck(employees[j]))
                cws.format(chr(ord('@')+reviewcolum) + str(j+offset),{
                    "backgroundColor": {
                        "red": 0.917647059,
                        "green": 0.6,
                        "blue": 0.6
                    }
                })
                logLabel.update()
                break
            except Exception:
                reviewPayPage()


# In[15]:


def runBot():
    if reviewFlag.get():
        reviewPayPage()
        reviewPay()
    else:
        auditTime()
        if ALL_FLAG:
            team = teamCheck()
            for i in range(2, len(team)):
                targetSetup([team[i]])
                auditTime()


# In[16]:


def teamSelect(event):
    if teamDrop.get() == 'Team 1':
        targetDropdown.config(value=team1)
    if teamDrop.get() == 'Team 2':
        targetDropdown.config(value=team2)
    if teamDrop.get() == 'Team 3':
        targetDropdown.config(value=team3)
    targetDropdown.current(0)
    return


# In[17]:


def teamCheck():
    if teamDrop.get() == 'Team 1':
        return team1
    if teamDrop.get() == 'Team 2':
        return team2
    if teamDrop.get() == 'Team 3':
        return team3


# In[18]:


def passwordCheck(event):
    if password.get() != 'Password':
        return
    password.delete('0', 'end')
    return

def minrangeCheck(event):
    if rangemin.get() != 'Start':
        return
    rangemin.delete('0', 'end')
    return

def maxrangeCheck(event):
    if rangemax.get() != 'Finish':
        return
    rangemax.delete('0', 'end')
    return


# In[19]:


root = Tk()
root.title("Time Entry: Electric Boogaloo")
root.columnconfigure(0, weight=1)

settings = Frame(root)
settings.grid(row=0, column=0, pady= 10, sticky=N+S)

username = Entry(settings, width=35, borderwidth=3)
username.insert(0, setting["Username"])
username.grid(row=0, column=0, padx=1, pady=1, columnspan=1, sticky=N)
password = Entry(settings, width=35, borderwidth=3)
password.insert(0, 'Password')
password.bind("<FocusIn>", passwordCheck)
password.grid(row=1, column=0, padx=1, pady=1, columnspan=1, sticky=N)

weeks = weeksList
weekdrop = StringVar()
weekdrop.set(setting["Week"])
weekDropdown = OptionMenu(settings, weekdrop, *weeks)
weekDropdown.grid(row=2, column=0, pady=1, columnspan=1, sticky=N)
weekDropdown.config(width=30, relief=RIDGE)

auditFlag = IntVar()
auditWithEntry = Radiobutton(settings, text="Enter Time", variable=auditFlag, value= 0)
auditWithEntry.grid(row=3, column= 0, pady=1,padx = 15, columnspan=1, sticky=N+E)
auditselection = Radiobutton(settings, text="Only Audit", variable=auditFlag, value= 1)
auditselection.grid(row=3, column= 0, pady=1, padx = 15, columnspan=1, sticky=N+W)

reviewFlag = BooleanVar()
reviewselection = Checkbutton(settings, text="Review Paycheck?", variable=reviewFlag)
reviewselection.grid(row=5, column=0, pady=1, padx= 1, columnspan=1, sticky=N)
reviews = reviewPayCheckWeeks
review = StringVar()
review.set(setting["Review Pay Check Week"])
reviewDrop = OptionMenu(settings, review, *reviews)
reviewDrop.grid(row=4, column=0, pady=1, columnspan=2, sticky=N)
reviewDrop.config(width=30, relief=RIDGE)


team1 = ['All', 'My Ho', 'William', 'Lien']
team2 = ['All', 'Haniyah', 'Jaquille', 'Justin']
team3 = []


teams = ['Team 1', 'Team 2']
teamDrop = ttk.Combobox(settings, value=teams)
teamDrop.set(setting["Team"])
teamDrop.grid(row=6, column=0, pady=1, columnspan=1, sticky=N)
teamDrop.config(width=30)
teamDrop.bind("<<ComboboxSelected>>", teamSelect)
teamDrop.config(state='readonly')

targets = team1
targetDropdown = ttk.Combobox(settings)
if setting['Employee'] in team1:
    targetDropdown.config(value=team1)
elif setting['Employee'] in team2:
    targetDropdown.config(value=team2)
else:
    targetDropdown.config(value=team3)
targetDropdown.set(setting["Employee"])
targetDropdown.grid(row=7, column=0, pady=1, columnspan=1, sticky=N)
targetDropdown.config(width=30)
targetDropdown.config(state='readonly')

rangemin = Entry(settings, width=10, borderwidth=3)
rangemin.insert(0, 'Start')
rangemin.bind("<FocusIn>", minrangeCheck)
rangemin.grid(row=8, column=0, pady=1, padx= 42, columnspan= 1, stick=N+W)
rangemax = Entry(settings, width=10, borderwidth=3)
rangemax.insert(0, 'Finish')
rangemax.bind("<FocusIn>", maxrangeCheck)
rangemax.grid(row=8, column=0, pady=1, padx=40, columnspan=1, stick=N+E)
rangelabel = Label(settings, text='to')
rangelabel.grid(row=8, column=0, pady=1, padx=1, columnspan=1, stick=N)
rangeFlag = BooleanVar()
rangecheck = Checkbutton(settings, text="Specify (Spreadsheet) Range?", variable=rangeFlag)
rangecheck.grid(row=9, column=0, pady=1, padx=1, columnspan=1, sticky=N)

finalize = Button(settings, text='Apply Settings', command=initialSetup, width=15)
finalize.grid(row=10, column=0, pady=1, columnspan=1, sticky=S)

login = Button(settings, text='Login', command=loginPS, width=35)
login.grid(row=11, column=0, pady=1, columnspan=1, sticky=S)

enteringTime = Button(settings, text='Run Bot', command=runBot, width=35)
enteringTime.grid(row=12, column=0, pady=1, columnspan=1, sticky=S)

auditLog = LabelFrame(root, text='Audit Log, there are many like it but this is mine')
auditLog.grid(row=0, column=2, columnspan=5, rowspan=9, padx=5, pady=2, sticky=N+S+W+E)
logLabel = scrolledtext.ScrolledText(auditLog, width=70, height=35, state=DISABLED)
font_tuple = ('Hack', 10)
logLabel.config(font = font_tuple)
auditLog.columnconfigure(0, weight=1)
auditLog.rowconfigure(0, weight=1)
logLabel.pack()

progress = ttk.Progressbar(root, orient=HORIZONTAL, mode='determinate' , length=70)
progress.grid(row=14, column=1, columnspan=6, rowspan=1,padx=5,pady=2, stick=N+S+W+E)
progressText = Label(root, text='')
progressText.grid(row=14,column=0, pady=2, stick=S)

root.update()
root.minsize(root.winfo_width(), root.winfo_height())
root.maxsize(root.winfo_width(), root.winfo_height())

root.update_idletasks()
root.mainloop()

