from os import name
from robocorp.tasks import task
from robocorp import browser

from RPA.Robocorp.Vault import Vault

_secret = Vault().get_secret("credentials")
gmail_account = _secret["email"]
gmail_password = _secret["password"]


@task
def minimal_task():
    login_to_gmail()

def login_to_gmail():
    browser.goto("https://mail.google.com/mail")
    page = browser.page()
    page.fill("#identifierId", gmail_account)
    page.click("button:text('Next')")

    page.fill("#password", gmail_password)
    page.click("button:text('Next')")

if name == 'main':
    minimal_task()