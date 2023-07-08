from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pytest
import xlutls
import time


class TestCompanyInfoLI:
    @pytest.fixture()
    def setup_teardown(self):
        global driver
        global path
        global rows
        ser = Service('F:/Selenium/chromedriver_win32_108/chromedriver.exe')
        driver = webdriver.Chrome(service=ser)
        driver.implicitly_wait(3)
        driver.get("https://www.linkedin.com/home")
        time.sleep(2)
        driver.maximize_window()
        time.sleep(2)
        driver.execute_script("window.scrollTo(0,100)")
        time.sleep(2)
        path = r"C:\Users\DSG\Desktop\login_testdata.xlsx"
        rows = xlutls.getRowCount(path, "Sheet1")
        time.sleep(3)
        yield
        driver.close()
        print('test completed')


    def test_company_details(self,setup_teardown):
        for r in range(2, rows + 1):
            username = xlutls.readData(path, "Sheet1", r, 1)
            password = xlutls.readData(path, "Sheet1", r, 2)
            time.sleep(3)
            driver.find_element(By.ID, "session_key").send_keys(username)
            time.sleep(2)
            driver.find_element(By.ID, "session_password").send_keys(password)
            time.sleep(2)
            driver.find_element(By.XPATH, "//*[@id='main-content']/section[1]/div/div/form[1]/div[2]/button").click()
            time.sleep(2)
            driver.find_element(By.CLASS_NAME, "search-global-typeahead__input").send_keys("tata consultancy services")
            time.sleep(2)
            driver.find_element(By.XPATH,"(// div // span // span[@class ='search-global-typeahead-hit__text t-16 t-black t-normal'])[1]").click()
            time.sleep(2)
            driver.find_element(By.LINK_TEXT, "Tata Consultancy Services").click()
            time.sleep(2)
            driver.find_element(By.LINK_TEXT, "see more").click()
            time.sleep(2)
            company_info = driver.find_element(By.CLASS_NAME, "org-about-module__description").text
            print(company_info)
            xlutls.writeData(path, "Sheet1", r, 3, company_info)
            time.sleep(5)













