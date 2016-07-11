from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

# navigate to the rad priv report
driver = webdriver.Firefox()
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

# get rads IDs (from excel)
ID_nums = ['5471','4816']

# enter rad names (need to build try-except for rads removed)
rads_in = driver.find_element_by_id('s2id_autogen1')

for i in range(len(ID_nums)):
    actions = ActionChains(driver)
    actions.move_to_element(rads_in)
    actions.click()
    actions.send_keys(ID_nums[i]).perform()
    time.sleep(0.5)
    actions = ActionChains(driver)
    actions.send_keys(Keys.RETURN).perform()

# run and export
driver.find_element_by_class_name('pull-right').click()
time.sleep(3)
driver.find_element_by_class_name('headerButton').click()
