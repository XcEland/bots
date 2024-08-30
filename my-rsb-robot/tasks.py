from robocorp.tasks import task
from robocorp import browser

@task
def robot_spare_bin_python():
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    log_in()
    fill_and_submit_sales_form()

def open_the_intranet_website():
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    page = browser.page()
    page.fill("#username","maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sales_form():
    page = browser.page()
    page.fill("#firstname", "John")
    page.fill("#lastname","Smith")
    page.fill("#salesresult","123")
    page.select_option("#salestarget", "10000")
    page.click("text=Submit")
