import datetime
from re import U
from bottle import redirect
from eel import init, expose, sleep, start
import eel
import tkinter
from tkinter import filedialog
from numpy.lib.histograms import histogram
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException as NSEE, NoSuchWindowException
from selenium.common.exceptions import TimeoutException as TE
from selenium.common.exceptions import NoSuchWindowException as NSWE
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from os import path
from win10toast import ToastNotifier

home = path.expanduser('~')
location = path.join(home, 'Downloads')

file = open(path.join(location, 'log_file.txt'), "a")

PATH = "geckodriver.exe"
options = webdriver.ChromeOptions()
options.add_argument('--log-level=3')
options.set_capability("silent", True)
# options.add_argument("--headless")  # Runs Chrome in headless mode.
options.add_argument('--no-sandbox')  # Bypass OS security model
options.add_argument('--disable-gpu')  # applicable to windows os only
options.add_argument('start-maximized')
options.add_argument('disable-infobars')
options.add_argument("--disable-extensions")

init('web')


midway_btn = '//*[@id="enterprise-access-plugin-btn"]'

upload_url = "https://selection.amazon.com/home"
midway_url = 'https://midway-auth.amazon.com/login'

file_path = None
upload_driver = None
midway_auth = False

start_no = 88
all_xpath = {
    "pollaris_modal":  "//*[@id='polaris_root']/awsui-modal/div[2]/div/div/div[1]/awsui-button/button",
    "product-type":  "//*[@id='awsui-multiselect-0-dropdown-option-1']/div/awsui-checkbox/label/input",
    "UK": "//*[@id='awsui-multiselect-1-dropdown-option-1']/div/awsui-checkbox/label/input",
    "DE": "//*[@id='awsui-multiselect-1-dropdown-option-2']/div/awsui-checkbox/label/input",
    "FR": "//*[@id='awsui-multiselect-1-dropdown-option-3']/div/awsui-checkbox/label/input",
    "IT": "//*[@id='awsui-multiselect-1-dropdown-option-12']/div/awsui-checkbox/label/input",
    "ES": "//*[@id='awsui-multiselect-1-dropdown-option-14']/div/awsui-checkbox/label/input",
    "input_asin": "//*[@class='asin']/awsui-form-field/div/div/div/div/span/awsui-textarea/textarea",
    "add-attribute": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/section/div[2]/div/div[1]/div/div[1]/div/section[1]/div/div[2]/awsui-button/button",
    "search-btn": "/html/body/div/div/div/awsui-app-layout/div/div[1]/nav/div/span/div/awsui-form/div/div[4]/span/div[1]/awsui-button[1]/button",
    "submit_new_contribution": "/ html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/awsui-tabs/div/ul/li[2]/a/span/span/awsui-tooltip/span",
    "data_augmenter_id": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/awsui-tabs/div/div[2]/span/section/awsui-column-layout/div/span/div/div/awsui-form-field/div/div[2]/div/div[1]/span/awsui-input/div/input",
    "search_icon": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/awsui-tabs/div/div[2]/span/section/awsui-column-layout/div/span/div/div/awsui-form-field/div/div[2]/div/div[2]/span/awsui-button/button",
    "next_btn": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[4]/span/div/awsui-button[2]",
    "attribute_dropdown_btn": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div[1]/button",
    "attribute_input": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/input",
    "attribute_search_list": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/ul",
    "attribute_search_li": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/ul/li[1]",
    "index_suppressed_dropdown_btn": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[2]/awsui-column-layout/div/span/div/div[1]/div/div/div/div/div/div/div/div/div/div/div/awsui-form-field/div/div/div/div/span/awsui-select/div/div/awsui-select-trigger/div/div",
    "index_suppressed_input": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[2]/awsui-column-layout/div/span/div/div[1]/div/div/div/div/div/div/div/div/div/div/div/awsui-form-field/div/div/div/div/span/awsui-select/div/div/awsui-select-dropdown/div/div[1]/span/awsui-select-filter/div/awsui-input/div/input",
    "index_suppressed_false_li": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[2]/awsui-column-layout/div/span/div/div[1]/div/div/div/div/div/div/div/div/div/div/div/awsui-form-field/div/div/div/div/span/awsui-select/div/div/awsui-select-dropdown/div/div[2]/ul/li",
    "index_suppressed_next_btn": "/html/body/div/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[4]/span/div/awsui-button[3]"
}
mp_id = {
    "UK": 3,
    "DE": 4,
    'FR': 3
}

cookie = ''
new_url = 'https://selection.amazon.com/item/3/B07MMWLRN7/'


def for_single_Asin(asin, mp, id, brand):
    try:
        print("starting one")
        mp = mp_id[mp]
        upload_url = 'https://selection.amazon.com/item/'+str(mp)+'/'+str(asin)
        global upload_driver
        upload_driver.get(upload_url)
        sleep(10)
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, all_xpath['pollaris_modal']))
        )
        # click close button on modal
        upload_driver.find_element_by_xpath(
            all_xpath['pollaris_modal']).click()
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['add-attribute']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['add-attribute']).click()
        upload_driver.find_element_by_xpath(
            all_xpath['submit_new_contribution']).click()
        upload_driver.find_element_by_xpath(
            all_xpath['data_augmenter_id']).send_keys(id)
        upload_driver.find_element_by_xpath(all_xpath['search_icon']).click()
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_dropdown_btn']).click()
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['attribute_input']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_input']).send_keys('index_suppressed')
        sleep(2)
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['attribute_search_li']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_search_li']).click()
        upload_driver.find_element_by_xpath(all_xpath['next_btn']).click()
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['index_suppressed_dropdown_btn']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['index_suppressed_dropdown_btn']).click()
        upload_driver.find_element_by_xpath(
            all_xpath['index_suppressed_input']).send_keys('false')
        sleep(2)
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['index_suppressed_false_li']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['index_suppressed_false_li']).click()
        upload_driver.find_element_by_xpath(
            all_xpath['index_suppressed_next_btn']).click()
        return True
    except NSEE as e:
        print('error', e)
        return False


def start_task(data):
    # print(data)
    length = len(data['asin'])
    length = 1
    for i in range(0, length):
        asin = data['asin'][i]
        mp = data['mp'][i]
        id = data['id'][i]
        brand = data['brand'][i]
        response = for_single_Asin(asin, mp, id, brand)
        if(response):
            print('done')


def get_mid_auth():
    try:
        global upload_driver
        global midway_auth
        global cookie
        print("Starting midway")
        upload_driver = webdriver.Firefox()
        upload_driver.get(new_url)
        upload_driver.maximize_window()
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, all_xpath['pollaris_modal']))
        )
        midway_auth = True
        return True

    except (NSEE):
        midway_auth = True
        return 'Proceed'
    except (TE, NSWE):
        print("Timeout miday")
        return 'Mideway timeout'


# {'asin': ['B00P2RNJ54', 'B07QJCTTN9', 'B07MMWLRN7', 'B00OKV2USU', 'B08S7SZWH8'], 'mp': ['FR', 'DE', 'UK', 'DE', 'ES'], 'id': [
#    119898864412, 118720037412, 119882024212, 118720037412, 119805860512], 'brand': ['Puma', 'Puma', 'Puma', 'Puma', 'Puma']}

def separate_filename(f_path):
    excel_data = pd.read_excel(f_path)
    all_data = {
        'asin': excel_data['ASIN'].to_list(),
        'mp': excel_data['MP'].to_list(),
        'id': excel_data['ID'].to_list(),
        'brand': excel_data['Brand'].to_list()
    }
    return all_data


def reset():
    global file_path
    global upload_driver
    global midway_auth
    upload_driver = None
    file_path = None
    midway_auth = None


def showNotification(msg):
    toaster = ToastNotifier()
    toaster.show_toast("Picture ", msg,
                       duration=20, icon_path='icon.ico', threaded=True)


@ expose
def start_driver_upload(f_path):
    global file_path
    global upload_driver
    global midway_auth
    file_path = f_path
    try:
        all_data = separate_filename(f_path)
        get_mid_auth()
        if(midway_auth):
            start_task(all_data)
            print('start')
        else:
            # upload_driver.quit()
            # reset()
            msg = 'Midway Authentication Failed, Try Again'
            print(msg)
            showNotification(msg)
            return [msg, 'red']
    except Exception as e:
        # if(upload_driver is not None):
        # upload_driver.quit()
        # reset()
        msg = f'Something went wrong, Restart the Program and try again \n Error occured: {e}'
        print(msg)
        showNotification(msg)
        return [msg, 'red']


start_driver_upload(
    'C:/Users/vayushi/Desktop/Listing and Troubleshooting.xlsx')


@ expose
def get_file_path():
    root = tkinter.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(
        title='Please select your excel sheet', initialdir='/', filetypes=[('Excel Files', '*.xlsx')])
    return file_path


# if __name__ == "__main__":
#    start('index.html', size=(900, 710))
