from functions import login_cdms
from functions import get_report
from functions import browser
from functions import rec_pay
import glob
import os


if __name__ == '__main__':

    # # login_cdms('SWPbjaved', 'D1rectorchris', 'D1rectorchris!03')
    # #
    # # # Go to report download page
    # # repo_link: str = "https://cdms.exxonmobil.com/reports.aspx"
    # # # Report 1 (Regular Swipe Report)
    # # browser.get(repo_link)
    # #
    # report_name = 'Reconcile and Payment Status Detailed Listing (Excel)'
    # from_date = '06/28/2020'
    # to_date = '05/15/2021'
    # # get_report(report_name, from_date, to_date)
    # #
    # # browser.quit()
    #
    # save_name = f'Rec_Pay_Status_{from_date}_to_{to_date}.xlsx'
    # list_of_files = glob.glob("C:\\Users\\" + "divyal" + "\\Downloads/*.xlsx")
    # excel_file_path = max(list_of_files, key=os.path.getctime)
    # if os.path.exists(save_name):
    #     os.remove(save_name)
    # # file_name_arr = excel_file_path.split('\\')
    # # file_name = file_name_arr[-1]
    # # full_file_path = os.path.join(excel_file_path, file_name)
    # save_name_path = os.path.join(r'C:\Users\divyal\Downloads', save_name)
    # os.rename(excel_file_path, save_name_path)

    path = 'reconsile_payment.xlsx'
    remit_file = 'Remittances.xlsx'
    lookup_file = 'CDMS Lookups.xlsx'

    rec_pay(path, remit_file, lookup_file)