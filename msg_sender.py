
import openpyxl

from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class MsgSender:

    def __init__(self, filepath, username, password, url_col, msg_template,
            start_num, num_msg):

        self.username = username
        self.password = password
        self.url_col = url_col
        self.msg_template = msg_template
        self.start_num = start_num
        self.num_msg = num_msg

        wb_obj = openpyxl.load_workbook(filepath)
        self.sheet_obj = wb_obj.active

        chrome_options = Options()
        chrome_options.add_argument("user-data-dir=selenium")
        self.driver = webdriver.Chrome(options=chrome_options)

    def get_sellers(self):
        i = 2
        urls = []
        messages = []
        while self.sheet_obj.cell(row=i, column=self.url_col).value:
            url = self.sheet_obj.cell(row=i, column=self.url_col).value
            params = []
            j = 1
            while self.sheet_obj.cell(row=i, column=self.url_col + j).value:
                params.append(self.sheet_obj.cell(
                    row=i, column=self.url_col + j).value)
                j += 1
                
            # message = self.sheet_obj.cell(row=i, column=msg_col).value
            print("Params", params)
            params += ['' * 1000]
            message = self.msg_template.format(*params)
            urls.append(url)
            messages.append(message)
            i += 1
        return urls, messages


    def is_captcha_solved(self):
        return self.driver.execute_script('return (typeof grecaptcha !== "undefined") && grecaptcha && grecaptcha.enterprise?.getResponse().length > 0')


    def set_textarea(self, selector, value):
        self.driver.execute_script(
            "var valSetter=Object.getOwnPropertyDescriptor("
            "window.HTMLTextAreaElement.prototype, 'value').set;"
            f"var input = document.querySelector('{selector}');"
            f"valSetter.call(input,'{value}');"
            "var ev = new Event('input',{bubbles:true});"
            "input.dispatchEvent(ev);")

    def start_sending(self):
        seller_pages, messages = self.get_sellers()

        driver = self.driver
        driver.get("https://www.etsy.com")

        wait = WebDriverWait(driver, 180)

        if not driver.execute_script('return document.querySelector(".welcome-message-text")'):
            
            print("NO Login History")
            driver.execute_script('document.querySelector("button.select-signin").click()')
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#join_neu_email_field')))

            driver.execute_script(f'document.querySelector("#join_neu_email_field").value = "{email}"')
            driver.execute_script(f'document.querySelector("#join_neu_password_field").value = "{password}"')
            driver.execute_script('document.querySelector("button[name=\'submit_attempt\']").click()')

            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.welcome-message-text')))
        print("Logged IN")
        for i, page in enumerate(seller_pages):
            print("GET Page -", "'" + page + "'")
            driver.get(page)
            driver.execute_script('document.querySelector("#desktop_shop_owners_parent").querySelector("a.wt-btn.wt-btn--outline.wt-width-full.contact-action.convo-overlay-trigger.inline-overlay-trigger").click()')
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "textarea[placeholder='Write a message']")))
            mod_message = messages[i]
            # driver.execute_script(f'document.querySelector(").value = "{mod_message}"')
            
            value = mod_message
            self.set_textarea('textarea[placeholder="Write a message"]', value)

            # while not is_captcha_solved(driver):
                # driver.execute_script(f'document.querySelector("textarea[placeholder=\'Write a message\']").value = "{mod_message}"')
            while not self.is_captcha_solved():
                self.set_textarea('textarea[placeholder="Write a message"]', value)
                # pass
                
            print(f"Send Message: '{mod_message}' to {page}")
            # driver.execute_script(f'document.querySelector("button.cheact-arrow-container").click()')
            # input("Click ENTER to goto next page")
        
        driver.close()

