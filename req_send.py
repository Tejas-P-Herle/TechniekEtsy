#!/usr/bin/env python3

import requests
import grequests
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
batch_size = 100


def set_cookies(driver):
    global cookies

    # cookies = driver.execute_script("return document.cookie")
    # print("COOKIES", cookies)
    # cookies = driver.execute_script("""
    #         let output = {};
    #         document.cookie.split(/\s*;\s*/).forEach(function(pair) {
    #           pair = pair.split(/\s*=\s*/);
    #         output[pair[0]] = pair.splice(1).join('=');
    #         });
    #         return output;
    #     """)
    all_cookies = driver.get_cookies()
    cookies = {}
    for cookie in all_cookies:
        cookies[cookie['name']] = cookie['value']
    print("COOKIES", cookies)


def set_codes(driver):
    global csrf_token, page_guid

    csrf_token = driver.execute_script("return document.querySelector('meta[name=\"csrf_nonce\"]').content")
    page_guid = driver.execute_script("guid_start = document.body.innerHTML.indexOf('page_guid\":') + 12; return document.body.innerHTML.slice(guid_start, document.body.innerHTML.indexOf('\"', guid_start+1))")


def get_captcha(driver):
    return driver.execute_async_script("""
        async function getParams() {
          let resp = await fetch("https://www.etsy.com/api/v3/ajax/member/conversations/recaptcha-data");
          let respJson = await resp.json();

          let captcha_code = respJson["id"];
            
          gresp = grecaptcha.enterprise.getResponse();
          
          document.querySelector("#" + captcha_code);
          grecaptcha.enterprise.reset();
          
          return [captcha_code, gresp]
        }
        params = getParams()
        params.then(resp => arguments[0](resp))
    """)

    
def sendMsgCore(user, message, captcha_code, recaptchaResp, proxy=None, timeout=30):

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

    
    proxies = None
    if proxy:
        proxies = {"http": proxy, "https": proxy}

    resp = requests.post(
        "https://www.etsy.com/api/v3/ajax/member/conversations",
		json=reqJson, headers={
          "x-csrf-token": csrf_token,
          "x-detected-locale": "INR|en-GB|IN",
          "page_guid": page_guid
        }, cookies=cookies, proxies=proxies, timeout=timeout
    )

    return resp


def sendMsg(user, message):
    captcha_code, gresp = get_captcha(driver)

    return sendMsgCore(user, message, captcha_code, gresp)


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



def main():
    global driver

    # proxies = get_working_proxies()
    
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


if __name__ == "__main__":
    main()

