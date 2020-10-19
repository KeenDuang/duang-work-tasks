import time
import selenium
from selenium import webdriver

#参数定义
loginname='#####'      #账户名
password='######'    #密码
date='2020-10-08'       #处理数据的日期

def main():
    opt=webdriver.ChromeOptions()
    driver=webdriver.Chrome(options=opt)
    driver.get('http://10.18.30.14:7001/aml_v2/login')
    driver.maximize_window()
    driver.find_element_by_id('loginname').send_keys(loginname)
    driver.find_element_by_id('password').send_keys(password)
    driver.find_element_by_id('loginbutton').click()
    time.sleep(2)
    driver.switch_to.frame(0)
    driver.find_element_by_xpath("//*[@id='mp02_organ']/div/div[5]/div/div[1]/div/ul/li[1]").click()

    #查询某个日期的数据
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
    driver.switch_to.frame("100504_IScolor_f")
    driver.find_element_by_xpath('//*[@id="create_date_q"]').send_keys(date)
    driver.find_element_by_xpath('//*[@id="create_date_z"]').send_keys(date)
    driver.find_element_by_xpath('//*[@id="submitForm"]').click()

    #每次从上到下点击100条数据
    for i in range(100):
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
        driver.switch_to.frame("100504_IScolor_f")
        driver.find_element_by_xpath('//*[@id="bb14_survey"]/div/div/div[2]/div/div/table/tbody/tr[1]/td[1]/a/i').click()
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
        driver.switch_to.frame("100504_IScolor_f")
        driver.find_element_by_xpath('//*[@id="myForm"]/div[1]/div/div/a[1]').click()
    driver.quit()


if __name__=="__main__":
    try:
        for i in range(100):
            main()
            time.sleep(10)
    except selenium.common.exceptions.NoSuchElementException:
        print('Done!')

