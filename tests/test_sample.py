#!/usr/bin/env python

from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from gmail import Gmail

gmail = Gmail()
toaddr = "nhan.vo@menlosecurity.com"
  
opts = Options()
opts.headless = True
driver = Chrome(options=opts, executable_path='chromedriver.exe')

try:
    driver.get('http://webcode.me')

    assert 'My html page1' == driver.title

    messages = "Ok"

except:
    messages = "Error"

finally:
    gmail.send_email(toaddr, messages)
    driver.quit()

