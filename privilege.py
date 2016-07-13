#['12/18/2016', 'Davis, Mark', 'Roger Williams Medical Center']

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time, datetime



driver = webdriver.Chrome()
driver.get('http://vradhome/Privileging/Facility?facility')
positives = '12/20/2016'
# new code
fac_search = driver.find_element_by_id('s2id_autogen1')
actions = ActionChains(driver)
actions.move_to_element(fac_search)
actions.click()
actions.send_keys('Roger Williams Medical Center').perform()
time.sleep(0.5)
actions = ActionChains(driver)
actions.send_keys(Keys.RETURN).perform()
time.sleep(2.8)

driver.find_element_by_xpath("//*[contains(text(), 'Davis, Mark')]").click()
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
    'process and update status to WD. Please contact'
    'Roberta with questions.'.format(positives)
    ).perform()
