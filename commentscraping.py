import time
import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from bs4 import BeautifulSoup
import threading
import requests
import pandas as pd

#detecting whether an element has changed/being refreshed
class element_changes(object):

  def __init__(self, length):
    self.length = length

  def __call__(self, driver):
    elements = browser.find_elements(By.CLASS_NAME, 'replyContent')
    if len(elements) > self.length:
      return len(elements)
    else:
      return False

#Defining a set of functions for time controlling
#reference: https://stackoverflow.com/questions/57060448/cannot-take-screenshot-with-0-width
def wait_selector_present(dv, xp: str, timeout: int = 5):
    cond = EC.presence_of_element_located((By.XPATH, xp))
    try:
        WebDriverWait(dv, timeout).until(cond)
    except TimeoutException as e:
        raise ValueError(f'Cannot find {xp} after {timeout}s') from e


def wait_selector_update(pre:int, timeout:int=5):
    cond = element_changes(pre)
    WebDriverWait(browser, timeout).until(cond)

def wait_selector_visible(xp: str, timeout: int = 5):
    cond = EC.visibility_of_any_elements_located((By.XPATH, xp))
    try:
        WebDriverWait(browser, timeout).until(cond)
    except TimeoutException as e:
        raise ValueError(f'Cannot find any visible {xp} after {timeout}s') from e

def wait_selector_clickable(xp: str, timeout: int = 5):
    cond = EC.element_to_be_clickable((By.XPATH, xp))
    try:
        WebDriverWait(browser, timeout).until(cond)
    except TimeoutException as e:
        raise ValueError(f'Cannot find any visible {xp} after {timeout}s') from e

#fetch the main textual content from a comment given its id
def fetch_website(id):
    global hash_map
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    thread_browser = webdriver.Chrome(driver_root_file, options=options)
    thread_browser.set_page_load_timeout(30)
    #intercept timeout exceptions to make sure that the page gets loaded
    while True:
        try:
            thread_browser.get(base_link+id)
            wait_selector_present(thread_browser, '//*[@id="replyContentMain"]',timeout=10)
            txt_element = thread_browser.find_element(By.ID, 'replyContentMain')
            if not txt_element.get_attribute('innerText'):
                continue
            else:
                hash_map[id] = txt_element.get_attribute('innerText')
                thread_browser.quit()
                break
        except TimeoutException:
            continue
driver_root_file = ChromeDriverManager().install()
browser = webdriver.Chrome(driver_root_file)
browser.get("http://liuyan.people.com.cn/threads/list?fid=5063&position=1&utm_source=pocket_mylist")
XPATH_IDS = {'dropdown':'/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul',
             'load':'/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul/div[1]'}
dropdown_len = 0
id_lst = []
title_lst = []
time_lst = []
hash_map = dict()
#click on the "加载更多" button until all data after June has been loaded
while True:
    wait_selector_present(browser, XPATH_IDS['dropdown'])
    dropdown = browser.find_element(By.XPATH,XPATH_IDS['dropdown'])
    wait_selector_update(dropdown_len, timeout=10)
    dropdown_len = len(dropdown.find_elements(By.CLASS_NAME, 'replyContent'))
    post_time = browser.find_element(By.XPATH,'/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul/li[{}]/div[2]/div/p'.format(dropdown_len))
    #if the time of the post is in May, exit the loop
    if int(post_time.text.split('-')[1]) == 5:
        break
    wait_selector_clickable(XPATH_IDS['load'])
    load_button = browser.find_element(By.XPATH, XPATH_IDS['load'])
    load_button.click()
#fetch the ids, the titles, and the posting time of the comments
for i in range(dropdown_len):
    id_element = browser.find_element(By.XPATH,'/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul/li[{}]/div[2]/div/h2/span[2]'.format(i+1))
    title_element = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul/li[{}]/div[1]/h1'.format(i+1))
    time_element = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/main/div/div/div[2]/div/ul/li[{}]/div[2]/div/p'.format(i+1))
    id_lst.append(id_element.text.split(':')[1])
    title_lst.append(title_element.text)
    time_lst.append(time_element.text)

base_link = 'http://liuyan.people.com.cn/threads/content?tid='
# Use the ids and the base link to access the full content of the comments
#Multi-threading is applied here to speed up the process of fetching the webpages
for i in range(0, len(id_lst), 8):
    try:
        t1 = threading.Thread(target=fetch_website,args=(id_lst[i],))
        t2 = threading.Thread(target=fetch_website, args=(id_lst[i+1],))
        t3 = threading.Thread(target=fetch_website,args=(id_lst[i+2],))
        t4 = threading.Thread(target=fetch_website, args=(id_lst[i+3],))
        t5 = threading.Thread(target=fetch_website,args=(id_lst[i+4],))
        t6 = threading.Thread(target=fetch_website, args=(id_lst[i+5],))
        t7 = threading.Thread(target=fetch_website,args=(id_lst[i+6],))
        t8 = threading.Thread(target=fetch_website, args=(id_lst[i+7],))
        t1.start()
        t2.start()
        t3.start()
        t4.start()
        t5.start()
        t6.start()
        t7.start()
        t8.start()

        t1.join()
        t2.join()
        t3.join()
        t4.join()
        t5.join()
        t6.join()
        t7.join()
        t8.join()
    except IndexError:
        # if less than 8 ids are left, goes here
        t_lst = []
        for j in range(i, len(id_lst), 1):
            t_lst.append(threading.Thread(target=fetch_website,args=(id_lst[j],)))
        for t in t_lst:
            t.start()
        for t in t_lst:
            t.join()
assert len(hash_map) == len(id_lst)
#output the information as an excel file
out_dict = dict()
out_dict['id'] = id_lst
out_dict['title'] = title_lst
out_dict['time'] = time_lst
text_lst = []
for id in id_lst:
    text_lst.append(hash_map[id])
out_dict['main text'] = text_lst
pd.DataFrame(out_dict).to_excel('comments.xlsx', index=False)



