from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import pandas as pd
import logging

GECKO_DRIVER_PATH = './geckodriver'

EMAIL = "elliottban@sjtu.edu.cn"
SCHOOL = "TELECOM"
START_DATE = "2017-08-28"
END_DATE = "2019-02-20"
SOURCE_FILE = "./source.xlsx"

def set_env():
    if EMAIL == "":
        EMAIL = input("邮箱：")
    if SCHOOL == "":
        SCHOOL = input("精准外校名称：")
    if START_DATE == "":
        START_DATE = input("开始交换时间 yyyy-mm-dd：")
    if END_DATE == "":
        END_DATE = input("结束交换时间 yyyy-mm-dd：")
    if SOURCE_FILE == "":
        SOURCE_FILE = input("学分转换表路径：")
    logging.info("基本参数变量获取成功")

    return


def read_application_form(source_file_path):
    data_source = pd.read_excel(source_file_path, sheet_name=0, header=None)
    start_line = (data_source.loc[:, 0].str.contains(
        "Course Code") == True).nonzero()[0][0]
    end_line = (data_source.loc[:, 0].isna()).nonzero()[0][0]
    data_source.columns = data_source.loc[start_line]
    data_source = data_source.loc[start_line+1:end_line-1]

    description = pd.read_excel("./source.xlsx", sheet_name=1, header=0)

    return pd.merge(data_source, description, how="left", left_on=data_source.columns[0], right_on="Course code")


def new_application(form_row):
    driver.switch_to_window(driver.window_handles[1])

    input_dict = {
        "input#V1_CTRL4": EMAIL,
        "#V1_CTRL79": START_DATE,
        "#V1_CTRL15": END_DATE,
        "#V1_CTRL16": form_row['Course code'],
        "#V1_CTRL17": list(form_row['学分       \n(Credit Point)'])[0],
        "#V1_CTRL18": form_row['Course Name in school for exchange'],
        "#V1_CTRL19": form_row['学时                   \n(Credit Hour)'],
        "#V1_CTRL20": form_row['成绩（Grade）'],
        "#V1_CTRL21": form_row['Brief Description: '],
    }

    blank_element = driver.find_element_by_css_selector("h1:nth-child(2)")

    for selector in input_dict:
        element = driver.find_element_by_css_selector(selector)
        try:
            element.send_keys(str(input_dict[selector]))
        except Exception as e:
            print(e)
        blank_element.click()

    time.sleep(0.5)
    selector = ".select2-container"
    elements = driver.find_elements_by_css_selector(selector)
    element = elements[0]
    element.click()
    selector = ".select2-search__field"
    element = driver.find_element_by_css_selector(selector)
    element.click()
    time.sleep(0.5)
    content = "Fra"
    element.send_keys(content)
    time.sleep(1)

    selector = "li.select2-results__option:nth-child(1)"
    element = driver.find_element_by_css_selector(selector)
    element.click()

    element = elements[1]
    element.click()
    time.sleep(0.5)
    selector = ".select2-search__field"
    element = driver.find_element_by_css_selector(selector)
    element.click()
    content = SCHOOL
    element.send_keys(content)
    time.sleep(2)

    selector = "li.select2-results__option:nth-child(1)"
    element = driver.find_element_by_css_selector(selector)
    element.click()

    selector = "#V1_CTRL7"
    element = driver.find_element_by_css_selector(selector)
    element = Select(element)
    element.select_by_visible_text(form_row['课程名称      \n(Course Name)'])


if __name__ == "__main__":
    set_env()

    logging.info("脚本正在启动……您将看到一个新的浏览器窗口")
    driver = webdriver.Firefox(executable_path=GECKO_DRIVER_PATH)
    driver.get("http://my.sjtu.edu.cn/")
    input("现在，请在新打开的浏览器中前往成绩转换申请页面，完成后敲击回车")

    application_form = read_application_form(SOURCE_FILE)
    print(application_form)
    index = input("从哪条记录开始（含）？：")

    try:
        index = int(index)
    except Exception as e:
        index = 0
        logging.warning("Error with input:" + e)
    
    while index < len(application_form):
        index += 1
        form_row = application_form.loc[index]
        print(form_row)
        response = input("确认填写？y/N")
        if response == 'y':
            new_application(form_row)
