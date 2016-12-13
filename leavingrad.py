from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from dateutil.relativedelta import relativedelta
from tkinter import *
import time, sys, os, calendar, datetime, selenium, os, openpyxl as oxl

# maybe rewrite using a class to store physician and site info
class Physician(object):
    def __init__(self, name, ID, end_date):
        self.name = name
        self.ID = ID
        self.end_date = end_date
        
class Site(object):
    def __init__(self, name, state, priv_end):
        self.name = name
        self.state = state
        self.priv_end = priv_end


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
        if sheet['B{0}'.format(j)].value == None: # at end of information stop loop
            break
        else: # otherwise take all names into a list
            leave_date.append(sheet['A{0}'.format(j)].value)
            sheet['A{0}'.format(j)].value = None
            ID_nums.append(str(sheet['B{0}'.format(j)].value))
            sheet['B{0}'.format(j)].value = None        
            names.append(sheet['C{0}'.format(j)].value)
            sheet['C{0}'.format(j)].value = None
        j += 1
    for i in reversed(range(len(ID_nums))): # find the bad entry and remove it
        if ID_nums[i] == bad_ID:
            del(ID_nums[i])
            del(leave_date[i])
            del(names[i])
    for i in range(2,len(ID_nums)+2): # recreate spreadsheet
        sheet['A{0}'.format(i)].value = leave_date[i-2]
        sheet['B{0}'.format(i)].value = int(ID_nums[i-2])
        sheet['C{0}'.format(i)].value = names[i-2]
    refer.save('rad_reference.xlsx') # resave
    execfile('leavingrad.py') # restart script
    
def get_output(ID_nums,password,username):
    # navigate to the rad priv report
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList",2)
    fp.set_preference("browser.download.manager.showWhenStarting",False)
    fp.set_preference("browser.download.dir",os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    fp.set_preference("browser.download.manager.showAlertOnComplete",False)
    driver = webdriver.Firefox(fp)
    driver.get('http://'+username+':'+password+'@vradhome/Privileging/Reports')
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
            # doesn't catch if the last record is missing
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

def generate_excel_output(positives):
    wb = oxl.Workbook()
    sheet = wb.active
    # headers
    sheet['A1'] = "Physician"
    sheet['B1'] = "Site"
    sheet['C1'] = "Note"
    sheet['D1'] = "Comment"
    current_r = 2
    for entry in positives: # adds a line for each antry
        sheet['A'+str(current_r)] = entry[1]
        sheet['B'+str(current_r)] = entry[2]
        sheet['C'+str(current_r)] = "See Note (Rad Leaving)"
        sheet['D'+str(current_r)] = 'Rad last read date on: {0}. Please stop app'\
                    ' process and update status to WD. Please contact'\
                    ' Roberta with questions.'.format(entry[0])
        current_r+=1
    fileExists = True
    
    duplicate = ""
    counter = 1
    while fileExists:
        increment = False
        for filename in os.listdir(os.path.join(os.getcwd(), \
                        "Old")):
            # maybe add a 3 month filter to this loop, based on date modified: then delete
            if filename == "Rad exiting Audit_"+datetime.date.today().strftime("%Y%m%d")+duplicate+".xlsx":
                duplicate = " ({0})".format(counter)
                counter +=1
                increment = True
        if not increment:
            fileExists = False
            
    # saves the workbook in same directory
    wb.save("./Old/Rad exiting Audit_"+datetime.date.today().strftime("%Y%m%d")+duplicate+".xlsx")
    # opens the excel workbook once finished
    path = os.path.join(os.getcwd(), \
                        "Old\\Rad exiting Audit_"+datetime.date.today().strftime("%Y%m%d")+duplicate+".xlsx")
    os.startfile(path) # doesn't work with relative paths

def runAll(event=None):
    password = p_ent.get()
    username = u_ent.get()
    root.destroy() # quit the window after the information is gotten
    print("Getting information from rad_reference...")
    ID_nums, leave_date = strip_rad_reference()
    print("Getting information from MRPA...")
    get_output(ID_nums,password,username)
    print("Checking affiliations against end dates...")
    positives = get_positives(ID_nums, leave_date)
    print("Generating excel output...")
    generate_excel_output(positives)
    
    
def main():

    global p_ent # need to be global to be accessed by the callback
    global u_ent # need to be global to be accessed by the callback
    global root # need to be global to be accessed by the callback
    root = Tk()
    root.geometry("250x80+300+300") # width x height + location on screen
    p_ent = Entry(root, bg = 'white')
    u_ent = Entry(root, bg = 'white')
    u_ent.insert(0,'jacob.scherber') # use my username as default for now
    button = Button(root,text = 'Execute', command = runAll)
    p_lab = Label(root, text = 'Password')
    u_lab = Label(root, text = "Username")
    # organizing elements
    u_lab.grid(row=0,column=0)
    u_ent.grid(row=0,column=1)
    p_lab.grid(row=1,column=0)
    p_ent.grid(row=1,column=1)
    p_ent.focus()
    button.grid(row=3,column=3)
    root.bind('<Return>', runAll) # passes an event object to function
    root.mainloop()

main()
    

##    u_lab.pack(anchor=E, padx=10, pady=10)
##    u_ent.pack(anchor=E, padx=10, pady=10)
##    p_lab.pack(anchor=E, padx=10, pady=10)
##    p_ent.pack(anchor=E, padx=10, pady=10)
##    button.pack(anchor=W,padx=10, pady=10)

#print("Getting information from rad_reference...")
#ID_nums, leave_date = strip_rad_reference()
#print("Getting information from MRPA...")
#get_output(ID_nums)
#print("Checking affiliations against end dates...")
#positives = get_positives(ID_nums, leave_date)
#print("Generating excel output...")
#generate_excel_output(positives)
#input_information(positives)

# not currently used, quit working
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
        time.sleep(2)
        actions = ActionChains(driver)
        actions.send_keys(Keys.RETURN).perform()
        driver.implicitly_wait(10)
        ex_test = driver.find_elements_by_class_name("expand")
        # maybe avoide C2
        if len(ex_test) < 3:
            continue
        else:
            try:
                driver.implicitly_wait(20)
                name_spot = driver.find_element_by_xpath(
                    "//*[contains(text(), '{0}')]".format(positives[k][1]))
                driver.execute_script("return arguments[0].scrollIntoView();",name_spot)
                actions = ActionChains(driver)
                actions.move_to_element(name_spot).click().perform()
                expand = driver.find_elements_by_class_name("expand")[2]
                driver.execute_script("return arguments[0].scrollIntoView();",expand)
                expand.click()
                tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).strftime('%m/%d/%y')
                next_act = driver.find_elements_by_class_name("form-control")[0]
                comment = driver.find_elements_by_class_name("form-control")[9]
                actions = ActionChains(driver)
                actions.move_to_element(next_act).click()
                actions.send_keys('{0}'.format(tomorrow))
                actions.send_keys(Keys.TAB).perform()
                actions = ActionChains(driver)
                actions.send_keys("See Note (Rad Leaving)").perform()
                actions = ActionChains(driver)
                actions.move_to_element(comment).click().send_keys(
                    'Rad last read date on: {0}. Please stop app'
                    ' process and update status to WD. Please contact'
                    ' Roberta with questions.'.format(positives[k][0])
                    ).perform()
                
                driver.find_element_by_xpath('//*[@title="Submit Mass Update"]').click()
                time.sleep(7)
                print('{0} : {1} \nRad last read date on: {2}.\n'.format(
                      positives[k][1],positives[k][2],positives[k][0]))
                
            except selenium.common.exceptions.NoSuchElementException:
                    print("** Unable to find {0} at {1} **\n".format(positives[k][1].encode('ascii','ignore')
                                                                     ,positives[k][2].encode('ascii','ignore')))
            continue
          
        
        time.sleep(3)
    driver.quit()
