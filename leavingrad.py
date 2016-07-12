from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time, sys, os, calendar, openpyxl as oxl


## Getting Data

# get rads IDs (from excel)
refer = oxl.load_workbook('rad_reference.xlsx')
sheet = refer.get_sheet_by_name('Sheet1')
ID_nums = []
leave_date = []
j = 2
while True:
    if sheet['B{0}'.format(j)].value == None:
        break
    else:
        ID_nums.append(str(sheet['B{0}'.format(j)].value))
        leave_date.append(sheet['A{0}'.format(j)].value)
    j += 1

# navigate to the rad priv report
fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting",False)
fp.set_preference("browser.download.dir",os.getcwd())
fp.set_preference("browser.helperApps.neverAsk.saveToDisk",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
fp.set_preference("browser.download.manager.showAlertOnComplete",False)
driver = webdriver.Firefox(fp)
driver.get('http://jacob.scherber:IneedOne2@vradhome/Privileging/Reports')
drop_down = driver.find_element_by_id('reportSelect')
drop_down.click()
for i in range(4):
    drop_down.send_keys(Keys.DOWN)
drop_down.send_keys(Keys.ENTER)

# change types searched
close_stat = driver.find_element_by_class_name('select2-search-choice-close')
open_stat = driver.find_element_by_id('s2id_autogen10')
ActionChains(driver).move_to_element(close_stat).click().perform()
ActionChains(driver).move_to_element(open_stat).click(
    ).send_keys('init').send_keys(Keys.RETURN).perform()

# enter rad names (need to build try-except for rads removed)
rads_in = driver.find_element_by_id('s2id_autogen1')

for i in range(len(ID_nums)):
    try:
        actions = ActionChains(driver)
        actions.move_to_element(rads_in)
        actions.click()
        actions.send_keys(ID_nums[i]).perform()
        time.sleep(0.5)
        actions = ActionChains(driver)
        actions.send_keys(Keys.RETURN).perform()
    except:
        for i in range(2,len(ID_nums[i])):
            actions = ActionChains(driver)
            actions.send_keys(Keys.BACK_SPACE).perform()
    continue
# run and export
driver.find_element_by_class_name('pull-right').click()
time.sleep(7)
driver.find_element_by_class_name('headerButton').click()
time.sleep(1)
## Comparing Data
test_wb = oxl.load_workbook('output.xlsx')
test = test_wb.get_sheet_by_name('Worklist')
t_length = 1

# get workbook length
while True:
    if test['A{0}'.format(t_length)].value == None:
        break
    else:
        t_length += 1
t_length -= 1
positives = []
# find matching records
for row in range(2, t_length):
    current_record = []
    if test['S{0}'.format(row)].value == "":
        test['S{0}'.format(row)].value = datetime.strptime(
            '01/01/2099', '%m/%d/%Y')
    else:
        test['S{0}'.format(row)].value = datetime.strptime(
                                    test['S{0}'.format(row)].value, '%m/%d/%Y')
    for rad in range(len(ID_nums)):
        if (ID_nums[rad] == str(test['A{0}'.format(row)].value) and
        leave_date[rad] < (test['S{0}'.format(row)].value +
                                relativedelta(months=+1)) and
        test['J{0}'.format(row)].value in ['REAP-RC','REAP-ELEC','REAP-ORIG',
                                           'INP', 'INP-QA', 'INP-QA',
                                           'ORIG-SIG','ELEC-SIG', 'REAP-INP',
                                           'REAP-QA'] and
            'eaving' not in test['M{0}'.format(row)].value):
            date_reform = leave_date[rad].strftime('%m/%d/%Y')
            positives.append([date_reform,
                              test['B{0}'.format(row)].value,
                              test['E{0}'.format(row)].value])
print(positives)
# ID = A
# Name = B
# Facility = E
# Status = J
# Next Action Note = M
# Expiration Date = S



os.remove('output.xlsx')
