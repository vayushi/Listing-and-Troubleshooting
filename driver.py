import eel
from tkinter import filedialog
from tkinter import Tk
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException as NSEE, NoSuchWindowException
from selenium.common.exceptions import TimeoutException as TE
from selenium.common.exceptions import NoSuchWindowException as NSWE
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pandas import DataFrame
from pandas import read_excel
from os import path
from win10toast import ToastNotifier
from random import choice
from string import ascii_lowercase
from datetime import date, datetime

home = path.expanduser('~')
location = path.join(home, 'Downloads')

PATH = "geckodriver.exe"

eel.init('web')

upload_url = "https://selection.amazon.com/home"
midway_url = 'https://midway-auth.amazon.com/login'

file_path = None
upload_driver = None
midway_auth = False
response_list = {'result': [], 'header': [], 'desc': []}
log_file = open(path.join(location, 'log_file.txt'), 'a+')
is_log = True

all_xpath = {
    "midway_btn": '//*[@id="enterprise-access-plugin-btn"]',
    "pollaris_modal":  "//*[@id='polaris_root']/awsui-modal/div[2]/div/div/div[1]/awsui-button/button",
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
    "brand_value_input": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[2]/awsui-column-layout/div/span/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div[2]/awsui-form-field/div/div/div/div/span/div/awsui-input/div/input",
    "suppressed_value_drop_down": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[3]/div/div",
    "suppress_current_value_btn": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[3]/div/div[2]/awsui-tooltip/span/span/button",
    "suppress_remove_btn": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[3]/div/div[2]/div[1]/div/div[2]/div/div/button",
    "suppress_reason": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[2]/span/span/div/div/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/awsui-form-field/div/div/div/div/span/div/awsui-input/div/input",
    "suppressed_next_btn": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[4]/span/div/awsui-button[3]/button",
    "submit_contribution_btn": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[2]/div/div/span/section/div/awsui-wizard/div/div/awsui-form/div/div[4]/span/div/awsui-button[3]/button",
    "message_header_span": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[1]/div/span/awsui-flashbar/div/awsui-flash/div/div[2]/div/div[1]",
    "message_desc_span": "/html/body/div[2]/div/div/awsui-app-layout/div/main/div[1]/div/span/awsui-flashbar/div/awsui-flash/div/div[2]/div/div[2]"
}

mp_id = {
    "UK": 3,
    "DE": 4,
    'FR': 5,
    "IT": 35691,
    "ES": 44551
}


def log(data):
    if is_log == True:
        log_file.write(">> " + str(datetime.now()) + " >> " + str(data) + "\n")


def for_single_Asin(asin, mp, id, brand, task_id, i):
    try:
        mp = mp_id[mp]
        upload_url = 'https://selection.amazon.com/item/'+str(mp)+'/'+str(asin)
        log("Open URL" + upload_url)
        global upload_driver
        upload_driver.get(upload_url)
        eel.sleep(10)
        if(i == 0):
            eel.sleep(5)
            log("Finding Modal")
            WebDriverWait(upload_driver, 60).until(
                EC.presence_of_element_located(
                    (By.XPATH, all_xpath['pollaris_modal']))
            )
            # click close button on modal
            upload_driver.find_element_by_xpath(
                all_xpath['pollaris_modal']).click()
            log("Closed Modal")
        log("Finding Add Attribute")
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['add-attribute']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['add-attribute']).click()
        log("Clicked Add Attribute")
        upload_driver.find_element_by_xpath(
            all_xpath['submit_new_contribution']).click()
        log("Submit new Contribution button clicked")
        upload_driver.find_element_by_xpath(
            all_xpath['data_augmenter_id']).send_keys(id)
        log("Added Augmented id")
        eel.sleep(2)
        upload_driver.find_element_by_xpath(all_xpath['search_icon']).click()
        log("Search icon")
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_dropdown_btn']).click()
        eel.sleep(2)
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['attribute_input']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_input']).send_keys('brand')
        log("Choosing brand value")
        eel.sleep(2)
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['attribute_search_li']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['attribute_search_li']).click()
        eel.sleep(2)
        upload_driver.find_element_by_xpath(all_xpath['next_btn']).click()
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['brand_value_input']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['brand_value_input']).clear()
        upload_driver.find_element_by_xpath(
            all_xpath['brand_value_input']).send_keys(brand)
        log("Entering brand value")
        eel.sleep(2)
        upload_driver.find_element_by_xpath(
            all_xpath['suppressed_value_drop_down']).click()
        eel.sleep(2)
        try:
            upload_driver.find_element_by_xpath(
                all_xpath['suppress_remove_btn'])
            log("Suprress already present")
        except NSEE:
            upload_driver.find_element_by_xpath(
                all_xpath['suppress_current_value_btn']).click()
            log("New supress value")
        finally:
            upload_driver.find_element_by_xpath(
                all_xpath['suppress_reason']).clear()
            upload_driver.find_element_by_xpath(
                all_xpath['suppress_reason']).send_keys(task_id)
            log("Entering task id")
            eel.sleep(2)
        upload_driver.find_element_by_xpath(
            all_xpath['suppressed_next_btn']).click()
        log("Supress next button clicked")
        eel.sleep(2)
        WebDriverWait(upload_driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,  all_xpath['submit_contribution_btn']))
        )
        upload_driver.find_element_by_xpath(
            all_xpath['submit_contribution_btn']).click(
        )
        log("Submit contribution button clicked")
        eel.sleep(2)
        msg_header = upload_driver.find_element_by_xpath(
            all_xpath['message_header_span']).text
        msg_desc = upload_driver.find_element_by_xpath(
            all_xpath['message_desc_span']).text
        log("Final Response \n" + "Done \n" + msg_header + msg_desc)
        return ['Done', msg_header, msg_desc]
    except NSEE as e:
        log("NSEE as" + e)
        return ['Error Occured', '', '']


def start_task(data):
    global response_list
    length = len(data['asin'])
    eel.get_curr_status(f'Process started')
    log("Started process")
    header_list = []
    result_list = []
    desc_list = []
    for i in range(0, length):
        a = f'Progress Status: {i+1} out of {length}'
        eel.get_curr_status(a)
        log(a)
        asin = data['asin'][i]
        mp = data['mp'][i]
        id = data['id'][i]
        brand = data['brand'][i]
        task_id = data['task_id'][i]
        [result, header, desc] = for_single_Asin(
            asin, mp, id, brand, task_id, i)
        result_list.append(result)
        header_list.append(header)
        desc_list.append(desc)
    response_list['result'] = result_list
    response_list['header'] = header_list
    response_list['desc'] = desc_list
    log("All Asin Completed, Please Wait Generating Excel File")
    eel.get_curr_status(
        f'All Asin Completed, Please Wait Generating Excel File')
    return


def get_mid_auth():
    try:
        global upload_driver
        global midway_auth
        authStatus = True
        msg = f'Completed Midway authentication'
        eel.get_curr_status(f'Starting Firefox Driver')
        log("Starting driver")
        upload_driver = webdriver.Firefox()
        log("Opening url" + upload_url)
        upload_driver.get(upload_url)
        eel.get_curr_status(f'Starting Midway authentication, Check Firefox')
        upload_driver.maximize_window()
        WebDriverWait(upload_driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, all_xpath['pollaris_modal']))
        )
    except NSEE as e:
        log("NSEE as" + e)
        authStatus = False
        msg = f'Midway authentication Failed'
    except TE as e:
        log("TE as" + e)
        authStatus = False
        msg = f'Midway authentication Failed, Timeout miday'
    except NSWE as e:
        log("NSWE as" + e)
        authStatus = False
        msg = f'Midway authentication Failed, Timeout miday'
    finally:
        midway_auth = authStatus
        eel.get_curr_status(msg)
        log("Midway Status" + str(midway_auth))
        return authStatus


def separate_filename(f_path):
    eel.get_curr_status(f'Reading Excel file...')
    log("Reading File")
    excel_data = read_excel(f_path)
    all_data = {
        'asin': excel_data['ASIN'].to_list(),
        'mp': excel_data['MP'].to_list(),
        'id': excel_data['ID'].to_list(),
        'brand': excel_data['Brand'].to_list(),
        "task_id": excel_data['Task_ID'].to_list()}
    log("File readed")
    eel.get_curr_status(f'Reading Excel file completed')
    return all_data


def get_random_string():
    length = 10
    letters = ascii_lowercase
    result_str = ''.join(choice(letters) for _ in range(length))
    return result_str


def export_excel_file(data):
    try:
        global response_list
        if(len(response_list['result']) == 0):
            return
        df = DataFrame({
            'ASIN': data['asin'],
            'MP': data['mp'],
            'ID': data['id'],
            "Brand": data['brand'],
            "Task_ID": data['task_id'],
            "Result": response_list['result'],
            "Message": response_list['header'],
            "Description": response_list['desc']
        })
        file_name = 'Listing and TroubleShooting'
        output_file = file_name + ' - ' + get_random_string() + '.xlsx'
        df.to_excel(location+'/'+output_file, index=False)
        log("Generating Excel file")
    except Exception as e:
        eel.get_curr_status(f'Something went wrong with excel')
        log("Error while generating excel as "+e)


def reset():
    global file_path
    global upload_driver
    global midway_auth
    if(upload_driver is not None):
        upload_driver.quit()
    upload_driver = None
    file_path = None
    midway_auth = None


def showNotification(msg):
    toaster = ToastNotifier()
    toaster.show_toast("Picture ", msg,
                       duration=20, icon_path='icon.ico', threaded=True)


@ eel.expose
def start_driver_upload(f_path):
    global file_path
    global upload_driver
    global midway_auth
    file_path = f_path
    msg = 'Task Completed'
    color = 'green'
    try:
        all_data = separate_filename(f_path)
        get_mid_auth()
        if(midway_auth):
            start_task(all_data)
        else:
            msg = 'Midway Authentication Failed, Try Again'
            color = 'red'
    except Exception as e:
        msg = f'Something went wrong, Restart the Program and try again \n Error occured: {e}'
        color = 'red'
    finally:
        log(msg)
        export_excel_file(all_data)
        reset()
        showNotification(msg)
        return [msg, color]


@ eel.expose
def get_file_path():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(
        title='Please select your excel sheet', initialdir='/', filetypes=[('Excel Files', '*.xlsx')])
    log("File path = " + file_path)
    return file_path


if __name__ == "__main__":
    if log_file == True:
        log_file.write("\n \n \n \n>>>>>>>>>>> Starting Program at " +
                       str(datetime.now()) + " >>>>>>>> \n \n")
    eel.start('index.html', size=(600, 600))
