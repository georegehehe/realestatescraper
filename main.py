import selenium.common.exceptions
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from PIL import Image
import numpy as np
import pandas as pd
import time
import re
import sys

#headless mode, defaulted to True, False only for debugging purposes
HEADLESS = True

#find the amount of offset for the slider in order to solve the captcha
def find_offset():
    #Looking for a white rectangle within the image, which indicates the location of the missing block
    #defining the parameter of the white rectangle block
    PATTERN_WIDTH = 20
    PATTERN_HEIGHT = 34
    img = Image.open('Captcha.png')
    img_matrix = np.asarray(img)
    #take rgb
    img_matrix = img_matrix[:, :, :3]
    max_color_diff = 0
    max_offset = 0
    #max_pattern_matrix = []
    # Pool the original image matrix into a smaller matrix that summarizes the average color of each rectangle
    # of pattern block size, with top left corner at location (i,j)
    # and find the pattern block with the largest average color--indicating closeness to color white
    for i in range(img_matrix.shape[0]-PATTERN_HEIGHT):
        for j in range(img_matrix.shape[1]-PATTERN_WIDTH):
            pattern_matrix = img_matrix[i:i+PATTERN_HEIGHT, j:j+PATTERN_WIDTH]
            curr_color_diff = pattern_matrix.mean(axis=2).mean(axis=None)
            if curr_color_diff >= max_color_diff:
                max_color_diff = curr_color_diff
                max_offset = j
                #max_pattern_matrix = pattern_matrix
    #return the offset of the slider.
    return max_offset-PATTERN_WIDTH#, max_pattern_matrix

#simulate the dragging of a slider
#https://blog.csdn.net/Mingyueyixi/article/details/104345623
def simpleSimulateDragX(source, targetOffsetX):
    """
    简单拖拽模仿人的拖拽：快速沿着X轴拖动，直接一步到达正确位置，再暂停一会儿，然后释放拖拽动作
    B站是依据是否有暂停时间来分辨人机的，这个方法适用。
    :param source:
    :param targetOffsetX:
    :return: None
    """
    #参考`drag_and_drop_by_offset(eleDrag,offsetX-10,0)`的实现，使用move方法
    action_chains = ActionChains(browser)
    # 点击，准备拖拽
    action_chains.click_and_hold(source)
    action_chains.pause(0.2)
    action_chains.move_by_offset(targetOffsetX,0)
    action_chains.pause(0.6)
    action_chains.release()
    action_chains.perform()

#detecting whether an element has changed/being refreshed
class element_changes(object):

  def __init__(self, locator, previous):
    self.locator = locator
    self.previous = previous

  def __call__(self, driver):
    element = driver.find_element(By.XPATH, self.locator)
    element_text = element.get_attribute("innerText")
    if element_text != self.previous:
      return element_text
    else:
      return False

#Defining a set of functions for time controlling
#reference: https://stackoverflow.com/questions/57060448/cannot-take-screenshot-with-0-width
def wait_selector_present(xp: str, timeout: int = 5):
    cond = EC.presence_of_element_located((By.XPATH, xp))
    try:
        WebDriverWait(browser, timeout).until(cond)
    except TimeoutException as e:
        raise ValueError(f'Cannot find {xp} after {timeout}s') from e

def wait_selector_invisible(xp: str, timeout: int = 5):
    cond = EC.invisibility_of_element_located((By.XPATH, xp))
    WebDriverWait(browser, timeout).until(cond)

def wait_table_refreshed(tb: str,pre:str, timeout: int = 5):
    cond = element_changes(tb, pre)
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

#create the list of companies to search for
def create_search_lst():
    xls = pd.ExcelFile('违约房地产公司2021-2022.xlsx')
    company_lst = pd.read_excel(xls, '2022')['公司全称'].tolist()
    return company_lst

#assemble the strings taken from the raw html file into a dataframe
def process_table_text(txt, company):
    txt_lst = list(filter(lambda x: x, re.split('\n\t\n', txt)))
    company_name = []
    institute_names = []
    acceptance_gross = []
    acceptance_balance = []
    overdue_gross = []
    overdue_balance = []
    for i in range(len(txt_lst)):
        res = (i + 1) % 10
        if res == 1:
            company_name.append(company)

        if res == 2:
            institute_names.append(txt_lst[i])
        elif res == 3:
            acceptance_gross.append(txt_lst[i])
        elif res == 4:
            acceptance_balance.append(txt_lst[i])
        elif res == 5:
            overdue_gross.append(txt_lst[i])
        elif res == 6:
            overdue_balance.append(txt_lst[i])
    assert len(company_name) == len(institute_names) == len(acceptance_balance) == len(acceptance_gross)\
           == len(overdue_balance) == len(overdue_gross)
    if len(company_name) == 10:
        print('The date for {} may be incomplete, please verify', company)
    new_df = pd.DataFrame(zip(company_name, institute_names, acceptance_gross, acceptance_balance, overdue_gross, overdue_balance),
                          columns=['公司名称', '承兑人开户机构名称', '累计承诺发生额', '承兑余额', '累计逾期发生额','逾期余额'])
    return new_df

#assemble all the steps needed to solve the captcha
def solve_captcha():
    wait_selector_present(XPATH_IDS['image'])
    wait_selector_visible(XPATH_IDS['image'])
    captcha_img = browser.find_element(By.XPATH,XPATH_IDS['image'])
    captcha_img.screenshot('Captcha.png')
    offset = find_offset()
    slider = browser.find_element(By.XPATH, XPATH_IDS['slider'])
    simpleSimulateDragX(slider, offset)

# initialize driver
if HEADLESS:
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    browser = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
else:
    print('Script running, Please reframe from touching the keyboard/mouse until task completion')
    browser = webdriver.Chrome(ChromeDriverManager().install())
browser.get("https://disclosure.shcpe.com.cn/#/infoQuery/ticketStateQuery")
XPATH_IDS = {'first choice':'/html/body/div[1]/section/section/main/section/div[1]/section/nav/div/form/div[2]/div/div/div[2]/div[1]/div[1]/ul/li[1]',
             'month':'/html/body/div[2]/div[1]/div/div[2]/table[3]/tbody/tr[2]/td[2]',
             'image':'//*[@id="app"]/section/section/main/section/div[1]/section/nav/div/div/div/div[2]/div/div/div/div[2]/div[1]/img',
             'slider':'//*[@id="app"]/section/section/main/section/div[1]/section/nav/div/div/div/div[2]/div/div/div/div[3]/div/div',
             'table':'//*[@id="app"]/section/section/main/section/div[2]/div[1]/section[2]/div[2]/div/div[3]/table',
             'search button':'//*[@id="search-btn"]'}

#set month to June
input_elements = browser.find_elements(By.CLASS_NAME, 'el-input__inner')
date_box = input_elements[0]
text_box = input_elements[1]
date_box.click()
month_element = browser.find_element(By.XPATH, XPATH_IDS['month'])
browser.implicitly_wait(10)
ActionChains(browser).move_to_element(month_element).click(month_element).perform()
# modify this list to continue from the previous checkpoint
companies = create_search_lst()
pre = "N/A"
#perform the following steps on each company in the company list
"""
the following format of code would safeguard any unexpected issues, including connection issues and so on

while True:
    try:
    
    except:


"""
for i, company in enumerate(companies):
    while True:
        try:
            #progrss bar
            output_str = '\r{} out of {} companies'.format(i+1, len(companies))
            sys.stdout.write(output_str)
            sys.stdout.flush()
            #input company name
            text_box.clear()
            text_box.send_keys(company)
            #select the first item from the pop-up menu
            try:
                first_choice = browser.find_element(By.XPATH, XPATH_IDS['first choice'])
                wait_selector_present(XPATH_IDS['first choice'])
                wait_selector_visible(XPATH_IDS['first choice'])
                wait_selector_clickable(XPATH_IDS['first choice'])
                browser.implicitly_wait(10)
                ActionChains(browser).move_to_element(first_choice).click(first_choice).perform()
            except StaleElementReferenceException:
                try:
                    first_choice = browser.find_element(By.XPATH, XPATH_IDS['first choice'])
                except NoSuchElementException:
                    #unless company doesn't exist, in which case continue to the next company
                    print('company {} not found, continuing...'.format(company))
                    break
                wait_selector_present(XPATH_IDS['first choice'])
                wait_selector_visible(XPATH_IDS['first choice'])
                wait_selector_clickable(XPATH_IDS['first choice'])
                browser.implicitly_wait(10)
                ActionChains(browser).move_to_element(first_choice).click(first_choice).perform()
            #click the search button
            search_button = browser.find_element(By.XPATH, XPATH_IDS['search button'])
            wait_selector_clickable(XPATH_IDS['search button'])
            browser.implicitly_wait(10)
            ActionChains(browser).move_to_element(search_button).click(search_button).perform()
            #loop until captcha gets solved correctly
            while True:
                solve_captcha()
                try:
                    wait_selector_invisible(XPATH_IDS['image'])
                    break
                except TimeoutException:
                    print('Solving Captcha unsuccessful, reattempting...')
            timeout = False
            try:
                #make sure that the data in the table is refreshed before fetching
                #condition being that any text in the table changes
                wait_table_refreshed(XPATH_IDS['table'], pre)
            except TimeoutException:
                #unless timeout, need to check the integrity of the data later then
                print('The data for {} needs to be verifed'.format(company))
                break
            table_element = browser.find_element(By.XPATH, XPATH_IDS['table'])
            #for debugging
            assert table_element.get_attribute("innerText") != pre
            pre = table_element.get_attribute("innerText")
            concat_df = process_table_text(pre, company)
            try:
                #if excel file exists, append to the end
                orig_df = pd.read_excel('result.xlsx')
                final_df = pd.concat([orig_df, concat_df])
                final_df.to_excel('result.xlsx', index=False)
            except FileNotFoundError:
                #else, write a new excel file
                concat_df.to_excel('result.xlsx', index=False)
            break
        except:
            continue




