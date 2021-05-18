import selenium
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

from datetime import date, timedelta
from email.message import EmailMessage
import pandas as pd
import numpy as np
import glob
import time
import os

import smtplib


# get username
import getpass

delay = 10

global browser_handle_track
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.page_load_strategy = 'normal'
path_driver = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
driver = path_driver
browser = webdriver.Chrome(executable_path=path_driver, options=chrome_options)
delay = 15


def login_cdms(username, pass1, pass2):
    try:

        # chrome_options =  Options()
        # chrome_options.add_experimental_option("detach", True)
        # chrome_options.page_load_strategy = 'normal'
        url = "https://cdms.exxonmobil.com"
        # path_driver = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
        # browser = webdriver.Chrome(executable_path=path_driver, options=chrome_options)
        browser.get(url)

        user = username
        password = pass1
        password2 = pass2

        signing_button = '//*[@id="loginForm"]/div[7]/a'

        box1 = WebDriverWait(browser, delay).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "pf.username")))

        box1.clear()

        box1.send_keys(user)

        box2 = WebDriverWait(browser, delay).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "pf.pass"))
        )

        box2.clear()

        box2.send_keys(password)

        browser.find_element_by_xpath(signing_button).click()

        time.sleep(1)

        guid = browser.window_handles

        browser_handle_track = guid[-1]

        browser.switch_to.window(guid[-1])

        dd_data_source = browser.find_element_by_id("ddlDataSource_Input")

        dd_data_source.click()
        time.sleep(1)
        dd_options = browser.find_element_by_id(
            "ddlDataSource_DropDown"
        ).find_element_by_css_selector("ul")

        for nice_option in dd_options.find_elements_by_css_selector("li"):

            if nice_option.get_attribute("textContent") == "Production-SCL":
                nice_option.click()
                break
        time.sleep(2)

        # user send keys
        box4 = WebDriverWait(browser, delay).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "T1"))
        )

        box4.send_keys(user)

        # password send keys
        box5 = WebDriverWait(browser, delay).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "T2"))
        )

        box5.send_keys(password2)

        browser.find_element_by_id("Submit2").click()
        print("Login Successful")
        time.sleep(2)
    except Exception as e:
        print(e)
        browser.quit()
        login_cdms(username="SWPabourg", pass1="31Forestavenue*", pass2="05forestavenue")


def get_report(report_name, start_date, end_date):
    repo_input = browser.find_element_by_id(
        "ctl00_ContentPlaceHolder1_ddlReportName_Input"
    )
    repo_input.click()
    time.sleep(2)

    repo_input.send_keys(report_name)
    repo_input.send_keys(Keys.RETURN)
    time.sleep(2)

    # Backspace Text Box
    for key in range(15):
        browser.find_element_by_id(
            "ctl00_ContentPlaceHolder1_txtFromDate_dateInput"
        ).send_keys(Keys.BACKSPACE)

    time.sleep(2)
    browser.find_element_by_id(
        "ctl00_ContentPlaceHolder1_txtFromDate_dateInput"
    ).send_keys(start_date)
    time.sleep(2)

    # Backspace Text Box
    for key in range(15):
        browser.find_element_by_id(
            "ctl00_ContentPlaceHolder1_txtToDate_dateInput"
        ).send_keys(Keys.BACKSPACE)
    time.sleep(2)
    browser.find_element_by_id(
        "ctl00_ContentPlaceHolder1_txtToDate_dateInput"
    ).send_keys(end_date)
    time.sleep(2)

    # Select Excel
    browser.find_element_by_xpath(
        '//*[@id="ctl00_ContentPlaceHolder1_rblistOptions_1"]').click()
    # Download Report Copy
    # browser = webdriver.Chrome(chrome_options=options)
    browser.find_element_by_id(
        "ctl00_ContentPlaceHolder1_btnDisplayReports"
    ).click()
    time.sleep(30)


def email_auto_script(to_email, subject, body, file_path):
    print('Sending Email!!!')

    gmail_user = "graham.scripting@gmail.com"
    gmail_password = "directorchris"

    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = gmail_user
        msg["To"] = to_email
        msg.set_content(body)
        # Raw text for path needs to be added here
        for file in file_path:
            print(file)
            with open(file, "rb") as f:
                file_data = f.read()
                file_name = os.path.split(file)
                file_name = file_name[-1]

                msg.add_attachment(
                    file_data,
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=file_name,
                )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_password)
            smtp.send_message(msg)
            print("Email sent!")
    except:
        print("Something went wrong...Email not Sent.")


def get_downloaded_file(file_extension: str):
    ##############################################################
    # parameter is a type string eg: 'pdf', 'xlsx', 'xls', 'doc'
    # returns: File path of latest downloaded file.
    ##############################################################

    username = getpass.getuser()
    list_of_files = glob.glob(f"C:/Users/{username}/Downloads/*.{file_extension}")
    downloaded_file_path = max(list_of_files, key=os.path.getctime)
    return downloaded_file_path


def module(row):
    if row["Module"] == "Material":
        return row["OT_Hours"]
    else:
        return row["ST_Rate"] * row["ST_Hours"] + row["OT_Rate"] * row["OT_Hours"] + row["DT_Hours"] * row["DT_Rate"]


def auth(row):
    if row["timesheet_reference"] > 0:
        return "Authorized"
    else:
        return "Unauthorized"


def status(row):
    if row["Paid"] == "OutStanding":
        return "OutStanding"
    else:
        return "Paid"


def week_ending(row):
    ini_date = pd.to_datetime(row["TS_Int"])
    new_date = timedelta((12 - ini_date.weekday()) % 7)
    week_end_date = ini_date + new_date
    return week_end_date


def emp_unit(row):
    if row["Module"] == "Labour":
        return f'{row["Name1"]} {row["Name2"]}'

    elif row["Module"] == "Equipment":
        return f'{row["employee_or_equipment_id"]} {row["Name1"]}'

    elif row["Module"] == "Material":
        return row["Name1"]

    else:
        return "Check"


def rec_pay(rec_pay_path, remit_file, lookup_file):

    df = pd.read_excel(rec_pay_path, engine='openpyxl')
    df_remit = pd.read_excel(remit_file, engine='openpyxl')
    df_lookup = pd.read_excel(lookup_file, engine='openpyxl', sheet_name='PO Approvers')

    df_lookup_new = df_lookup.iloc[:, :2]

    df_billable = df.loc[df['Skil_Eqip_Mat'] != 'Nonbillable'].copy()

    df_billable['Line Cost'] = df_billable.apply(lambda row: module (row), axis=1)

    df_billable['Authorized'] = df_billable.apply(lambda row: auth (row), axis=1)

    df_billable['Week Ending'] = df_billable.apply(lambda row: week_ending (row), axis=1)

    new_header = df_remit.iloc[2] #grab the first row for the header
    df_remit_new = df_remit[3:] #take the data less the header row
    df_remit_new.columns = new_header #set the header row as the df header

    df_temp = pd.DataFrame(df_remit_new[['Amount', 'Date Paid']]).copy()

    df_temp.rename(columns={'Amount': 'timesheet_reference'}, inplace=True)

    df_temp.dropna(inplace=True)

    df_billable['timesheet_reference'].round(2)

    df_billable = pd.merge(df_billable,
                         df_temp,
                         on ='timesheet_reference',
                         how ='left')

    df_billable['Date Paid'].replace(np.nan, 'OutStanding', inplace=True)

    df_billable.rename(columns={'Date Paid': 'Paid'}, inplace=True)

    df_billable['Status'] = df_billable.apply(lambda row: status(row), axis=1)

    df_billable['Employee/Unit'] = df_billable.apply(lambda row: emp_unit(row), axis=1)

    df_lookup_new.rename(columns={'Area ID': 'area_id', 'Approver Name': 'Approver'}, inplace=True)

    df_billable = pd.merge(df_billable,
                         df_lookup_new,
                         on ='area_id',
                         how ='left')

    auth_column = df_billable.pop('Authorized')
    df_billable.insert(0, 'Authorized', auth_column)

    status_column = df_billable.pop('Status')
    df_billable.insert(0, 'Status', status_column)

    paid_column = df_billable.pop('Paid')
    df_billable.insert(0, 'Paid', paid_column)

    week_ending_column = df_billable.pop('Week Ending')
    df_billable.insert(0, 'Week Ending', week_ending_column)

    emp_unit_column = df_billable.pop('Employee/Unit')
    df_billable.insert(0, 'Employee/Unit', emp_unit_column)

    approver_column = df_billable.pop('Approver')
    df_billable.insert(0, 'Approver', approver_column)

    line_cost_column = df_billable.pop('Line Cost')
    df_billable.insert(0, 'Line Cost', line_cost_column)

    df_billable["Week Ending"] = df_billable["Week Ending"].astype(str)
    df_billable["TS_Int"] = df_billable["TS_Int"].astype(str)
    df_billable['Agreement'] = df_billable['Agreement'].astype(str)

    df_billable.to_excel("test.xlsx", index_label=False, index=False)