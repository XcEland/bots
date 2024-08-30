from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

@task
def robot_spare_bin_python():
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    # fill_and_submit_sales_form()
    fill_form_with_excel_data()

def open_the_intranet_website():
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    page = browser.page()
    page.fill("#username","maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sales_form(sales_rep):
    page = browser.page()
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname",sales_rep["Last Name"])
    page.fill("#salesresult",str(sales_rep["Sales Target"]))
    page.select_option("#salestarget", str(sales_rep["Sales"]))
    page.click("text=Submit")

def download_excel_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)
