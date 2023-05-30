
import openpyxl
import time
import sys
import pathlib

from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class MsgSender:

    def __init__(self, filepath, username, password, msg_template,
            start_num, num_msg):

        self.username = username
        self.password = password
        # self.url_col = url_col
        self.msg_template = msg_template
        self.start_num = start_num
        self.num_msg = num_msg

        wb_obj = openpyxl.load_workbook(filepath)
        self.sheet_obj = wb_obj.active

        chrome_options = Options()
        if sys.platform == "linux":
            chrome_options.add_argument("user-data-dir=selenium")
        else:
            script_dir = pathlib.Path().absolute()
            chrome_options.add_argument(f"user-data-dir={script_dir}\\selenium")

        self.driver = webdriver.Chrome(options=chrome_options)

    def get_sellers(self):
        i = self.start_num + 1
        max_i = i + self.num_msg
        urls = []
        messages = []
        url_col = 2

        for k in range(2, 100):
            col_val = self.sheet_obj.cell(row=2, column=k).value
            print("COL VAL", col_val)
            if col_val and "etsy.com" in col_val:
                if (col_val.startswith("http") or col_val.startwith("www")
                        or col_val.startswith("etsy")):
                    url_col = k
                    break
        print("URL COL", url_col)
                
        while self.sheet_obj.cell(row=i, column=url_col).value:
            url = self.sheet_obj.cell(row=i, column=url_col).value
            params = []
            j = 1
            while self.sheet_obj.cell(row=i, column=url_col + j).value:
                params.append(self.sheet_obj.cell(
                    row=i, column=url_col + j).value)
                j += 1
                
            # message = self.sheet_obj.cell(row=i, column=msg_col).value
            print("Params", params)
            params += ['' * 1000]
            message = self.msg_template.format(*params)
            urls.append(url)
            messages.append(message)
            i += 1
            if i >= max_i:
                break
        return urls, messages
    
    def safe_evaluate(self, script, tries=1):
        while tries > 0:
            try:
                return self.driver.execute_script(script)
            except Exception as err:
                print("Evaluation Error:", err)
            tries -= 1

    def is_captcha_solved(self):
        return self.safe_evaluate('return (typeof grecaptcha !== "undefined") && grecaptcha && grecaptcha.enterprise?.getResponse().length > 0')


    def set_textarea(self, selector, value):
        value = value.replace('"', '\\"').replace("'", "\\'")
        self.safe_evaluate(
            "var valSetter=Object.getOwnPropertyDescriptor("
            "window.HTMLTextAreaElement.prototype, 'value').set;"
            f"var input = document.querySelector('{selector}');"
            f"valSetter.call(input,'{value}');"
            "var ev = new Event('input',{bubbles:true});"
            "input.dispatchEvent(ev);", 3)

    def is_browser_open(self):
        try:
            self.driver.current_url
            return True
        except:
            return False

    def start_sending(self):
        seller_pages, messages = self.get_sellers()

        driver = self.driver
        driver.get("https://www.etsy.com")

        wait = WebDriverWait(driver, 180)

        if not self.safe_evaluate('return document.querySelector(".welcome-message-text")', 3):
            
            print("NO Login History")
            driver.execute_script('document.querySelector("button.select-signin").click()')
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#join_neu_email_field')))

            driver.execute_script(f'document.querySelector("#join_neu_email_field").value = "{self.username}"')
            driver.execute_script(f'document.querySelector("#join_neu_password_field").value = "{self.password}"')
            driver.execute_script('document.querySelector("button[name=\'submit_attempt\']").click()')

            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.welcome-message-text')))
        print("Logged IN")
        for i, page in enumerate(seller_pages):
            print("GET Page -", "'" + page + "'")
            driver.get(page)
            # driver.execute_script('document.querySelector("#desktop_shop_owners_parent").querySelector("a.wt-btn.wt-btn--outline.wt-width-full.contact-action.convo-overlay-trigger.inline-overlay-trigger").click()')
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'a.contact-action')))
            # driver.execute_script('document.querySelector("a.contact-action").querySelector("a.wt-btn.wt-btn--outline.wt-width-full.contact-action.convo-overlay-trigger.inline-overlay-trigger").click()')

            driver.execute_script("document.querySelector('div.shop-owner-small-screens').querySelector('a.contact-action').click()")
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,
                "textarea[placeholder='Write a message']")))
            
            mod_message = messages[i]
            value = mod_message
            self.set_textarea('textarea[placeholder="Write a message"]', value)

            while not self.is_captcha_solved() and self.is_browser_open():
                self.set_textarea('textarea[placeholder="Write a message"]', value)

            if not self.is_browser_open():
                return 1
                
            print(f"Send Message: '{mod_message}' to {page}")
            driver.execute_script(f'document.querySelector("button.cheact-arrow-container").click()')
            time.sleep(2)
            # input("Click ENTER to goto next page")
        
        driver.close()
        return 0

