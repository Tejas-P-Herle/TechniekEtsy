#!/usr/bin/env python3

import openpyxl
import time

from tkinter import Tk

import requests
import grequests
import json
import sys

from lxml.html import fromstring
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


cookies = {}
csrf_token = ""
page_guid = ""
# filepath = "sellers_names.xlsx"
filepath = input("Filepath(sellers_names.xlsx): ")
if not filepath.strip():
    filepath = "sellers_names.xlsx"

message_file = input("Message File(Default Message - 'Hi'): ")
if not message_file.strip():
    message = "HI"
else:
    with open(message_file) as file:
        message = file.read()

message = message.replace('"', '\\"').replace("'", "\\'")
print("Message", message)

num_msg = 20
batch_size = 30
TIMEOUT=20

sleep_time = 0


def set_cookies(driver):
    global cookies

    all_cookies = driver.get_cookies()
    cookies = {}
    for cookie in all_cookies:
        cookies[cookie['name']] = cookie['value']
    print("COOKIES", cookies)


def set_codes(driver):
    global csrf_token, page_guid

    csrf_token = driver.execute_script(
        "return document.querySelector('meta[name=\"csrf_nonce\"]').content")
    page_guid = driver.execute_script("guid_start = document.body.innerHTML.indexOf('page_guid\":') + 12; return document.body.innerHTML.slice(guid_start, document.body.innerHTML.indexOf('\"', guid_start+1))")


def get_captcha(driver):
    return driver.execute_async_script("""
        async function getParams() {
          let resp = await fetch("https://www.etsy.com/api/v3/ajax/member/conversations/recaptcha-data");
          let respJson = await resp.json();

          let captcha_code = respJson["id"];
            
          gresp = grecaptcha.enterprise.getResponse();
          
          return [captcha_code, gresp]
        }
        params = getParams()
        params.then(resp => arguments[0](resp))
    """)

    
def sendMsgCore(user, message, captcha_code, recaptchaResp, timeout=TIMEOUT):

    reqJson = {
        "subject": "hi",
        "message": message,
        "attachments": "{}",
        "recipient_loginname": user,
        "recipient_id": None,
        "captcha_code": captcha_code,
        "g-recaptcha-response": recaptchaResp,
        "api_context": "buyer_conversations"
    }

    resp = requests.post(
        "https://www.etsy.com/api/v3/ajax/member/conversations",
        json=reqJson, headers={
          "x-csrf-token": csrf_token,
          "x-detected-locale": "INR|en-GB|IN",
          "page_guid": page_guid
        }, cookies=cookies, timeout=timeout
    )

    return resp


def sendMsgProxyPool(user, message, captcha_code, recaptchaResp, proxyBatch, timeout=TIMEOUT):

    reqJson = {
        "subject": "hi",
        "message": message,
        "attachments": "{}",
        "recipient_loginname": user,
        "recipient_id": None,
        "captcha_code": captcha_code,
        "g-recaptcha-response": recaptchaResp,
        "api_context": "buyer_conversations"
    }
    
    rs = []
    for proxy in proxyBatch:
        proxies = {"http": proxy, "https": proxy}
        rs.append(grequests.post(
            "https://www.etsy.com/api/v3/ajax/member/conversations",
            json=reqJson, headers={
              "x-csrf-token": csrf_token,
              "x-detected-locale": "INR|en-GB|IN",
              "page_guid": page_guid
            }, cookies=cookies, proxies=proxies, timeout=timeout
        ))

    return grequests.map(rs)

def reset_captcha(driver, captcha_code):
    
    return driver.execute_script(
        f"document.querySelector('#{captcha_code}'); grecaptcha.enterprise.reset();")


def sendMsg(user, message, proxyPool=False, timeout=TIMEOUT):
    captcha_code, gresp = get_captcha(driver)

    if proxyPool is False:
        msg_resp = sendMsgCore(user, message, captcha_code, gresp, timeout=timeout)
        # reset_captcha(driver, captcha_code)
        return msg_resp, captcha_code

    src = 'https://api64.ipify.org'
    response = requests.get('https://free-proxy-list.net/')
    parser = fromstring(response.text)
    proxies = []

    for i in parser.xpath('//tbody/tr'):
        if i.xpath('.//td[7][contains(text(),"yes")]'):
            proxy = ":".join([i.xpath('.//td[1]/text()')[0],
                              i.xpath('.//td[2]/text()')[0]])
            proxies.append(proxy)

    print("Got Proxies")
    batches = [proxies[s:s+batch_size] for s in range(0, len(proxies), batch_size)]
    for ix, batch in enumerate(batches):
        print("Testing Batch", ix)
        resps = sendMsgProxyPool(user, message, captcha_code, gresp, batch, timeout)
        print("Batch Done", ix, resps)
        # print("RESPS", resps)
        for resp in resps:
            if resp and resp.status_code == 201:
                # reset_captcha(driver, captcha_code)
                return resp, captcha_code
    return None, ""


def get_working_proxies(req_no=5):
    src = 'https://api64.ipify.org'
    response = requests.get('https://free-proxy-list.net/')
    parser = fromstring(response.text)
    proxies = []

    for i in parser.xpath('//tbody/tr'):
        if i.xpath('.//td[7][contains(text(),"yes")]'):
            proxy = ":".join([i.xpath('.//td[1]/text()')[0],
                              i.xpath('.//td[2]/text()')[0]])
            proxies.append(proxy)

    working_proxies = []
    batches = [proxies[s:s+batch_size] for s in range(0, len(proxies), batch_size)]
    for batch in batches:
        print("Testing Batch")
        rs = (grequests.get(src, proxies={'http': proxy, 'https': proxy}, timeout=20) for u in batch)
        for i, resp in enumerate(grequests.map(rs)):
            if resp and resp.ok:
                working_proxies.append(batch[i])
        print("Batch Complete, Working:", len(working_proxies))
        if len(working_proxies) >= req_no:
            break
    print("Working Proxies:", working_proxies)

    return working_proxies


def is_captcha_solved(driver):
    try:
        return driver.execute_script(
            'return (typeof grecaptcha !== "undefined") && grecaptcha && grecaptcha.enterprise?.getResponse().length > 0')
    except:
        return False
    

def is_browser_open(driver):
   try:
       driver.current_url
       return EC.alert_is_present()
   except:
       return False


def get_sellers(sheet_obj):
    sellers = []
    i = 1 + 1
    max_i = i + num_msg
    seller_col = 1
    print("Seller COL", seller_col)

    for k in range(1, 100):
        col_val = sheet_obj.cell(row=1, column=k).value
        if "Seller" in col_val or "Name" in col_val:
            seller_col = k
            break

    while sheet_obj.cell(row=i, column=seller_col).value:
        seller = sheet_obj.cell(row=i, column=seller_col).value
        sellers.append(seller)
        i += 1
    return sellers
  


def main():
    global driver

    # proxies = get_working_proxies()
    # root = Tk()
    # self.master.title("Techneik Etsy Messenger")

    # root.mainloop()

    wb_obj = openpyxl.load_workbook(filepath)
    sheet_obj = wb_obj.active

    sellers = get_sellers(sheet_obj)

    print("Sellers", sellers)
    
    chrome_options = Options()
    if sys.platform == "linux":
        chrome_options.add_argument("user-data-dir=selenium")
    else:
        script_dir = pathlib.Path().absolute()
        chrome_options.add_argument(f"user-data-dir={script_dir}\\selenium")

    driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://www.etsy.com/messages/sent")
    wait = WebDriverWait(driver, 1800)
    print("Loaded, Searching for button")
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[data-test-id=convo-compose-button]")))

    driver.execute_script("document.querySelector('[data-test-id=convo-compose-button]').click()")

    set_cookies(driver)
    set_codes(driver)
    
    for seller in sellers:
        while not is_captcha_solved(driver):
            if not is_browser_open(driver):
                print("Browser Closed")
                return

        resp, captcha_code = sendMsg(seller, message)
        if resp and resp.status_code == 201:
            print("Message sent successfully to", seller)
        else:
            print("Resp", resp)
            if resp is not None:
                if resp.status_code == 200:
                    print("Rate Limit HIT")
                print("Content", resp.content)
        # captcha_code, gresp = get_captcha(driver)
        time.sleep(sleep_time)
        if captcha_code and captcha_code.strip():
            reset_captcha(driver, captcha_code)


if __name__ == "__main__":
    main()

