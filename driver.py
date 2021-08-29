import datetime
from eel import init, expose, start
import eel
import tkinter
from tkinter import filedialog
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException as NSEE
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

PATH = "chromedriver.exe"
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
ids = {
    "uk": "//*[@id='awsui-checkbox-"+start_no+"']",
    "de": "//*[@id='awsui-checkbox-"+start_no+1+"']",
    "fr": "//*[@id='awsui-checkbox-"+start_no+2+"']",
    "it": "//*[@id='awsui-checkbox-"+start_no+11+"']",
    "es": "//*[@id='awsui-checkbox-"+start_no+13+"']"
}


def start_task():
    try:
        global upload_driver
        upload_driver.get(upload_url)
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH,  "//*[@id='polaris_root']/awsui-modal/div[2]/div/div/div[1]/awsui-button/button"))
        )
        # click close button on modal
        upload_driver.find_element_by_xpath(
            "//*[@id='polaris_root']/awsui-modal/div[2]/div/div/div[1]/awsui-button/button").click()
        # select items radio button
        upload_driver.find_element_by_css_selector(
            "input[type='radio'][value='items-search']").click()
        # select only asin in filter
        upload_driver.find_element_by_xpath(
            "//*[@id='awsui-checkbox-13']").click()
        # insert asin
        upload_driver.find_element_by_xpath(
            "//*[@id='awsui-textarea-0']").send_keys('ABCDEF')
    except NSEE as e:
        print('error', e)
        return 'Error on page'


def get_mid_auth():
    try:
        global upload_driver
        global midway_auth
        upload_driver = webdriver.Chrome(options=options, executable_path=PATH)
        upload_driver.get(upload_url)
        upload_driver.maximize_window()
        midway_auth = True
        return True
    except (NSEE):
        midway_auth = True
        return 'Proceed'
    except (TE, NSWE):
        return 'Mideway timeout'


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


@expose
def start_driver_upload(f_path):
    global file_path
    global upload_driver
    global midway_auth
    file_path = f_path
    try:
        all_data = separate_filename(f_path)
        print(all_data)
        get_mid_auth()
        if(midway_auth):
            start_task()
            print('midway done')
        else:
            upload_driver.quit()
            reset()
            msg = 'Midway Authentication Failed, Try Again'
            showNotification(msg)
            return [msg, 'red']
    except Exception as e:
        print(e)
        if(upload_driver is not None):
            upload_driver.quit()
        reset()
        msg = f'Something went wrong, Restart the Program and try again \n Error occured: {e}'
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
