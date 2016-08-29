from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime
from dateutil.relativedelta import relativedelta
import time, sys, os, calendar, openpyxl as oxl
import selenium

## Getting Data
def strip_rad_reference():
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
            leave_date.append(sheet['A{0}'.format(j)].value)
            ID_nums.append(str(sheet['B{0}'.format(j)].value))
        j += 1
    return ID_nums, leave_date

def remove_and_restart(bad_ID, driver):
    driver.quit()
    refer = oxl.load_workbook('rad_reference.xlsx')
    sheet = refer.get_sheet_by_name('Sheet1')
    ID_nums = []
    leave_date = []
    names = []
    j = 2
    while True:
        if sheet['B{0}'.format(j)].value == None:
            break
        else:
            leave_date.append(sheet['A{0}'.format(j)].value)
            sheet['A{0}'.format(j)].value = None
            ID_nums.append(str(sheet['B{0}'.format(j)].value))
            sheet['B{0}'.format(j)].value = None        
            names.append(sheet['C{0}'.format(j)].value)
            sheet['C{0}'.format(j)].value = None
        j += 1
    for i in reversed(range(len(ID_nums))):
        if ID_nums[i] == bad_ID:
            del(ID_nums[i])
            del(leave_date[i])
            del(names[i])
    for i in range(2,len(ID_nums)+2):
        sheet['A{0}'.format(i)].value = leave_date[i-2]
        sheet['B{0}'.format(i)].value = int(ID_nums[i-2])
        sheet['C{0}'.format(i)].value = names[i-2]
    refer.save('rad_reference.xlsx')
    # need to restart script
    execfile('leavingrad.py')
    
def get_output(ID_nums):
    # navigate to the rad priv report
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList",2)
    fp.set_preference("browser.download.manager.showWhenStarting",False)
    fp.set_preference("browser.download.dir",os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    fp.set_preference("browser.download.manager.showAlertOnComplete",False)
    driver = webdriver.Firefox(fp)
    driver.get('http://jacob.scherber:IneedOne3@vradhome/Privileging/Reports')
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

    # enter rad names (try-except for rads removed still buggy...)
    rads_in = driver.find_element_by_id('s2id_autogen1')

    for i in range(len(ID_nums)):
        try:
            actions = ActionChains(driver)
            actions.move_to_element(rads_in)
            actions.click()
            actions.send_keys(ID_nums[i]).perform()
            time.sleep(1)
            actions = ActionChains(driver)
            actions.send_keys(Keys.RETURN).perform()
        except:
            # need to add code to remove the record automatically and then restart
            print('ID: {0} is no longer in the system'
                  ', remove and restart'.format(ID_nums[i-1]))
            remove_and_restart(ID_nums[i-1], driver)
            time.sleep(2)
            sys.exit()
        continue
    # run and export
    driver.find_element_by_class_name('pull-right').click()
    time.sleep(11)
    driver.find_element_by_class_name('headerButton').click()
    time.sleep(7)
    driver.quit()
    
def get_positives(ID_nums, leave_date, remove=True):
    ## Comparing Data
    output_wb = oxl.load_workbook('output.xlsx')
    if remove == True:
        os.remove('output.xlsx')
    test = output_wb.get_sheet_by_name('Worklist')
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
            test['S{0}'.format(row)].value = datetime.datetime.strptime(
                '01/01/2099', '%m/%d/%Y')
        else:
            test['S{0}'.format(row)].value = datetime.datetime.strptime(
                                        test['S{0}'.format(row)].value, '%m/%d/%Y')
        for rad in range(len(ID_nums)):
            if (ID_nums[rad] == str(test['A{0}'.format(row)].value) and
            leave_date[rad] < (
                test['S{0}'.format(row)].value + relativedelta(months=+1)) and
            test['J{0}'.format(row)].value in ['REAP-RC','REAP-ELEC','REAP-ORIG',
                                               'INP', 'INP-QA', 'INP-QA',
                                               'ORIG-SIG','ELEC-SIG', 'REAP-INP',
                                               'REAP-QA'] and
                'eaving' not in test['M{0}'.format(row)].value and
            'TERMING' not in test['M{0}'.format(row)].value and
                'erming' not in test['M{0}'.format(row)].value and
                'LEAVING' not in test['M{0}'.format(row)].value):
                date_reform = leave_date[rad].strftime('%m/%d/%Y')
                positives.append([date_reform,
                                  test['B{0}'.format(row)].value,
                                  (test['E{0}'.format(row)].value + ' ' + test['F{0}'.format(row)].value)])
    return positives
# how to deal with C2 sites

def input_information(positives):
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get('http://vradhome/Privileging/Facility?facility')
    time.sleep(2)
    for k in range(len(positives)):
        fac_search = driver.find_element_by_id('s2id_autogen1')
        actions = ActionChains(driver)
        actions.move_to_element(fac_search)
        actions.click()
        actions.send_keys(positives[k][2]).perform()
        time.sleep(1.5)
        actions = ActionChains(driver)
        actions.send_keys(Keys.RETURN).perform()
        time.sleep(10)
        ex_test = driver.find_elements_by_class_name("expand")
        # maybe avoide C2
        if len(ex_test) < 3:
            continue
        else:
            try:
                name_spot = driver.find_element_by_xpath(
                    "//*[contains(text(), '{0}')]".format(positives[k][1]))
                actions = ActionChains(driver)
                actions.move_to_element(name_spot)
                actions.click()
                actions.perform()
                expand = driver.find_elements_by_class_name("expand")[2].click()
                tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).strftime('%m/%d/%y')
                next_act = driver.find_elements_by_class_name("form-control")[0]
                comment = driver.find_elements_by_class_name("form-control")[9]
                actions = ActionChains(driver)
                actions.move_to_element(next_act).click()
                actions.send_keys('{0}'.format(tomorrow)).send_keys(Keys.TAB)
                actions.send_keys("See Note (Rad Leaving)").perform()
                actions.move_to_element(comment).click().send_keys(
                    'Rad last read date on: {0}. Please stop app'
                    ' process and update status to WD. Please contact'
                    ' Roberta with questions.'.format(positives[k][0])
                    ).perform()
                driver.find_element_by_xpath('//*[@title="Submit Mass Update"]').click()
                print('{0} : {1} \nRad last read date on: {2}.'.format(
                      positives[k][1],positives[k][2],positives[k][0]))
                
            except selenium.common.exceptions.NoSuchElementException:
                    print("** Unable to find {0} at {1} **\n".format(positives[k][1],positives[k][2]))
            continue
          
        
        time.sleep(3)
    driver.quit()

ID_nums, leave_date = strip_rad_reference()    
get_output(ID_nums)
positives = get_positives(ID_nums, leave_date)
input_information(positives)

