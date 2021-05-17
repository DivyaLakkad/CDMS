from functions import login_cdms
from functions import get_report
from functions import browser


if __name__ == '__main__':

    login_cdms('SWPbjaved', 'D1rectorchris', 'D1rectorchris!03')

    # Go to report download page
    repo_link: str = "https://cdms.exxonmobil.com/reports.aspx"
    # Report 1 (Regular Swipe Report)
    browser.get(repo_link)

    get_report('Reconcile and Payment Status Detailed Listing (Excel)', '2021/04/01', '2021/05/01')