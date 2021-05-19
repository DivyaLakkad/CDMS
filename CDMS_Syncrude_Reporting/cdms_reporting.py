from functions import login_cdms
from functions import get_report
from functions import browser
from functions import rec_pay
import glob
import os, sys
import shutil
import datetime


if __name__ == '__main__':

    login_cdms('SWPbjaved', 'D1rectorchris', 'D1rectorchris!03')
    
    # Go to report download page
    repo_link: str = "https://cdms.exxonmobil.com/reports.aspx"
    # Report 1 (Regular Swipe Report)
    browser.get(repo_link)
    #
    report_name = 'Reconcile and Payment Status Detailed Listing (Excel)'
    from_date = '06/28/2020'
    x = datetime.datetime.now().date()
    to_date = x.strftime('%m/%d/%Y')
    #
    get_report(report_name, from_date, to_date)
    
    browser.quit()
    
    temp_from_date = from_date.replace('/', '.')
    temp_to_date = to_date.replace('/', '.')
    
    file_name = f'Rec_Pay_Status_{temp_from_date}_to_{temp_to_date}.xlsx'
    
    list_of_files = glob.glob(r"C:/Users/divyal/Downloads/*.xlsx")
    
    excel_file_path = max(list_of_files, key=os.path.getctime)
    
    if os.path.exists(file_name):
        os.remove(file_name)
    
    os.rename(excel_file_path, file_name)
    file_name = 'Rec_Pay_Reports/Rec_Pay_Status_06.28.2020_to_05.19.2021.xlsx'
    remit_file = 'Extra/Remittances.xlsx'
    lookup_file = 'Extra/CDMS Look ups.xlsx'

    rec_pay(file_name, remit_file, lookup_file)