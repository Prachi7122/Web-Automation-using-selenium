
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
import xlu
from multiprocessing import Process
from selenium.common.exceptions import WebDriverException

import os
load_dotenv()
# for holding website elements
ask_question = ""
start_writing = ""
title_txt = ""
question_txt = ""
ans_txt= ""
tags_txt = ""
# contetnt to be post on websites
title = ""
question = ""
answer =""
tag = ""
review = ""
# additional varible for diffrent websites process
value =""
general =""
lg_button =""
row_Q =2
# login element credentials
id =""
pw =""
lg =""
# user id and password holder
email =""
password = ""

# Reading sheets from .env file
login_sheet = os.getenv("LOGIN_SHEET")
ans_sheet = os.getenv("ANS_SHEET")
web_email = os.getenv("EMAIL_EXCEL")  # 1
web_password = os.getenv("PASSWORD_EXCEL")  # 2
web_url = os.getenv("URLS")  # 3
web_IdField = os.getenv("ID_TEXTFIELD")  # 4
web_PwField = os.getenv("PW_TEXTFIELD")  # 5
web_LoginButton = os.getenv("LOGIN_BUTTON")  # 6
web_AskQuestion = os.getenv("ASK_QUESTION")  # 7
web_TitleText = os.getenv("TITLE_TEXT")  # 8
web_QuestionText = os.getenv("QUESTION_TEXT")  # 9
web_AnsText = os.getenv("ANS_TEXT")  # 10
web_TagsText = os.getenv("TAGS_TEXT")  # 11
web_Review = os.getenv("REVIEW")  # 12
web_StartWriting = os.getenv("START_WRITING")  # 13
web_Keys = os.getenv("KEYS")  # 14
web_Post = os.getenv("POST")  # 15
web_GeneralQ = os.getenv("GENERAL_Q")  # 16
web_LoginButtonValue = os.getenv("LOGIN_BUTTON_VALUE")  # 17
web_WithEmail = os.getenv("WITH_EMAIL")  # 18
web_CreateQue = os.getenv("CREATE_QUE")  # 19
web_Values = os.getenv("VALUES")  # 20
web_ValueStack = os.getenv("VALUE_STACK")  # 21
# sheet2
web_Title = os.getenv("TITLE")  # 1
web_Questions = os.getenv("QUESTION")  # 2
web_Answer = os.getenv("ANSWER")  # 3
web_Tags = os.getenv("TAGS")  # 4


class WebAutomation:
    def __init__(self, path):
        self.path = path
        # self.driver = webdriver.Chrome()
        self.driver = None

    def open_new_tab(self, url):
        self.driver.execute_script("window.open('', '_blank');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.get(url)
        time.sleep(3)

    def login(self, row):
        web = xlu.readData(self.path, login_sheet, row, web_url)
        self.driver.execute_script("window.open('', '_blank');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.get(web)

        time.sleep(2)
        global id, pw, lg, email, password
        email = xlu.readData(self.path, login_sheet, row, web_email)
        password = xlu.readData(self.path, login_sheet, row, web_password)
        id = xlu.readData(self.path, login_sheet, row, web_IdField)
        pw = xlu.readData(self.path, login_sheet, row, web_PwField)
        lg = xlu.readData(self.path, login_sheet, row, web_LoginButton)

        self.driver.find_element(By.ID, id).send_keys(email)
        self.driver.find_element(By.ID, pw).send_keys(password)
        time.sleep(2)

        self.driver.find_element(By.XPATH, lg).click()
        time.sleep(3)

    def type2_post(self, row):  # stack exchange

        value = xlu.readData(self.path, login_sheet, row, web_ValueStack)
        if (value == 'ask'):
            web = xlu.readData(self.path, login_sheet, row, web_url)
            self.driver.execute_script("window.open('', '_blank');")
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.get(web)

        global ask_question, start_writing, title_txt, question_txt, tags_txt, title, question, tag, review, row_Q

        ask_question = xlu.readData(self.path, login_sheet, row, web_AskQuestion)
        start_writing = xlu.readData(self.path, login_sheet, row, web_StartWriting)
        title_txt = xlu.readData(self.path, login_sheet, row, web_TitleText)
        question_txt = xlu.readData(self.path, login_sheet, row, web_QuestionText)
        tags_txt = xlu.readData(self.path, login_sheet, row, web_TagsText)
        title = xlu.readData(self.path, ans_sheet, row_Q, web_Title)
        question = xlu.readData(self.path, ans_sheet, row_Q, web_Questions)
        tag = xlu.readData(self.path, ans_sheet, row_Q, web_Tags)
        review = xlu.readData(self.path, login_sheet, row, web_Review)
        value = xlu.readData(self.path, login_sheet, row, web_Values)

        self.driver.find_element(By.XPATH, ask_question).click()
        time.sleep(10)
        if value != "fault":
            self.driver.find_element(By.XPATH, start_writing).click()
            time.sleep(5)

        self.driver.find_element(By.XPATH, title_txt).send_keys(title)
        time.sleep(3)
        self.driver.find_element(By.XPATH, question_txt).send_keys(question)
        time.sleep(3)
        self.driver.find_element(By.XPATH, tags_txt).send_keys(tag)
        time.sleep(5)
        self.driver.find_element(By.XPATH, review).click()
        time.sleep(3)

    def type3_post(self, row):  # proprofs
        global ask_question, general, title_txt, question_txt, title, question, row_Q

        ask_question = xlu.readData(self.path, login_sheet, row, web_AskQuestion)
        general = xlu.readData(self.path, login_sheet, row, web_GeneralQ)
        title_txt = xlu.readData(self.path, login_sheet, row, web_TitleText)
        question_txt = xlu.readData(self.path, login_sheet, row, web_QuestionText)
        title = xlu.readData(self.path, ans_sheet, row_Q, web_Title)
        question = xlu.readData(self.path, ans_sheet, row_Q, web_Questions)
        time.sleep(15)
        # self.driver.refresh()

        self.driver.find_element(By.XPATH, general).click()

        time.sleep(3)
        self.driver.find_element(By.XPATH, ask_question).click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, title_txt).send_keys(title)
        time.sleep(3)
        self.driver.find_element(By.XPATH, question_txt).send_keys(question)
        time.sleep(3)

    def type4_post(self, row):
        web = xlu.readData(self.path, login_sheet, row, web_url)
        # self.driver.execute_script("window.open('', '_blank');")
        # self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver = webdriver.Chrome()
        self.driver.get(web)
        self.driver.maximize_window()
        time.sleep(3)
        global id, pw, lg, row_Q

        email = xlu.readData(self.path, login_sheet, row, web_email)
        lg_button = xlu.readData(self.path, login_sheet, row, web_LoginButtonValue)
        with_email = xlu.readData(self.path, login_sheet, row, web_WithEmail)
        password = xlu.readData(self.path, login_sheet, row, web_password)
        id = xlu.readData(self.path, login_sheet, row, web_IdField)
        pw = xlu.readData(self.path, login_sheet, row, web_PwField)
        lg = xlu.readData(self.path, login_sheet, row, web_LoginButton)

        self.driver.find_element(By.XPATH, lg_button).click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, with_email).click()
        time.sleep(3)

        self.driver.find_element(By.ID, id).send_keys(email)
        self.driver.find_element(By.ID, pw).send_keys(password)
        time.sleep(2)

        self.driver.find_element(By.XPATH, lg).click()
        time.sleep(3)

        create_Q = xlu.readData(self.path, login_sheet, row, web_CreateQue)

        ask_question = xlu.readData(self.path, login_sheet, row, web_AskQuestion)

        title_txt = xlu.readData(self.path, login_sheet, row, web_TitleText)
        time.sleep(3)

        title = xlu.readData(self.path, ans_sheet, row_Q, web_Title)
        print("prachi")

        time.sleep(10)
        self.driver.find_element(By.XPATH, create_Q).click()
        time.sleep(3)

        self.driver.find_element(By.XPATH, ask_question).click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, title_txt).send_keys(title)
        time.sleep(3)

    def type5_post(self, row):
        web = xlu.readData(self.path, login_sheet, row, web_url)
        self.driver.refresh()
        self.driver.execute_script("window.open('', '_blank');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.get(web)
        time.sleep(10)
        global id, pw, lg, ask_question, general, title_txt, question_txt, title, question, lg_button, row_Q, ans_txt, answer

        lg_button = xlu.readData(self.path, login_sheet, row, web_LoginButtonValue)
        email = xlu.readData(self.path, login_sheet, row, web_email)
        password = xlu.readData(self.path, login_sheet, row, web_password)
        id = xlu.readData(self.path, login_sheet, row, web_IdField)
        pw = xlu.readData(self.path, login_sheet, row, web_PwField)
        lg = xlu.readData(self.path, login_sheet, row, web_LoginButton)
        self.driver.find_element(By.LINK_TEXT, lg_button).click()
        time.sleep(3)
        self.driver.find_element(By.ID, id).send_keys(email)
        self.driver.find_element(By.ID, pw).send_keys(password)
        time.sleep(2)

        self.driver.find_element(By.CLASS_NAME, lg).click()
        time.sleep(3)

        ask_question = xlu.readData(self.path, login_sheet, row, web_AskQuestion)
        # general = xlu.readData(self.path, "Sheet1", row, 16)
        title_txt = xlu.readData(self.path, login_sheet, row, web_TitleText)
        question_txt = xlu.readData(self.path, login_sheet, row, web_QuestionText)
        tags_txt = xlu.readData(self.path, login_sheet, row, web_TagsText)
        ans_txt = xlu.readData(self.path, login_sheet, row, web_AnsText)
        title = xlu.readData(self.path, ans_sheet, row_Q, web_Title)
        question = xlu.readData(self.path, ans_sheet, row_Q, web_Questions)
        answer = xlu.readData(self.path, ans_sheet, row_Q, web_Answer)

        tag = xlu.readData(self.path, ans_sheet, row_Q, web_Tags)
        time.sleep(15)

        # self.driver.find_element(By.XPATH, general).click()
        # time.sleep(3)
        self.driver.find_element(By.XPATH, ask_question).click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, title_txt).send_keys(title)
        time.sleep(3)
        self.driver.find_element(By.XPATH, question_txt).send_keys(question)
        time.sleep(3)
        self.driver.find_element(By.XPATH, ans_txt).send_keys(answer)
        time.sleep(3)
        self.driver.find_element(By.XPATH, tags_txt).send_keys(tag)
        time.sleep(3)

    def close_driver(self):
        if self.driver:
            self.driver.quit()


# code with try-catch
def process_websites(start_row, end_row, path, failed_websites):
    web_automation = WebAutomation(path)
    for row in range(start_row, end_row + 1):
        print(row)
        key = xlu.readData(path, login_sheet, row, web_Keys)
        try:
            if (key == "ans_com"):
                web_automation.type4_post(row)
            elif (key == 'procode'):
                web_automation.type5_post(row)
            else:
                web_automation.login(row)
                # webites with similar login process
                if (key == 'stack'):
                    web_automation.type2_post(row)  # stack exchange
                else:
                    web_automation.type3_post(row)

        except WebDriverException as e:
            print(f"Exception on row {row} with key value {key}: {e}")
            failed_websites.append(row)
            print("appended")

    web_automation.close_driver()


def main():
    path = "C:/Users/prachi soni/OneDrive/Desktop/Zing/sheet2.xlsx"
    rows = xlu.getRowCount(path, login_sheet)
    web_automation = WebAutomation(path)
    failed_website = []

    for row in range(2, rows + 1):
        print(row)
        key = xlu.readData(path, login_sheet, row, web_Keys)
        value = xlu.readData(path, login_sheet, row, web_ValueStack)

        try:
            if (key == "ans_com"):
                web_automation.type4_post(row)
            elif (key == 'procode'):
                web_automation.type5_post(row)
            else:
                if (value != 'ask'):
                    web_automation.login(row)
                if (key == 'stack'):
                    web_automation.type2_post(row)  # stack exchange
                else:
                    web_automation.type3_post(row)  # proprofs
        except WebDriverException as e:
            print(f"Exception on row {row} with key value {key}: {e}")
            failed_website.append(row)
    web_automation.close_driver()
    print("this website has passed exception")
    print(failed_website)

    #ans phase
    rows = xlu.getRowCount(path, login_sheet)
    web_automation = WebAutomation(path)
    failed_website = []
    #
    # for row in range(2, rows + 1):
    #     print(row)
    #     key = xlu.readData(path, login_sheet, row, web_Keys)
    #     value = xlu.readData(path, login_sheet, row, web_ValueStack)
    #
    #     try:
    #         if (key == "ans_com"):
    #             web_automation.type4_post(row)
    #         elif (key == 'procode'):
    #             web_automation.type5_post(row)
    #         else:
    #             if (value != 'ask'):
    #                 web_automation.login(row)
    #             if (key == 'stack'):
    #                 web_automation.type2_post(row)  # stack exchange
    #             else:
    #                 web_automation.type3_post(row)  # proprofs
    #     except WebDriverException as e:
    #         print(f"Exception on row {row} with key value {key}: {e}")
    #         failed_website.append(row)
    # web_automation.close_driver()
    # print("this website has passed exception")
    # print(failed_website)


if __name__ == "__main__":
    main()