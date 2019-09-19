import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
import xlrd
import json
wb = xlrd.open_workbook('3rd.xls')
sheet = wb.sheet_by_index(0)
atds = {}
i = 2
while True:
    try:
        usn = sheet.cell_value(i, 1)
        dob = sheet.cell_value(i, 3)
        i = i+1
        print(usn+"  "+dob)
        d = dob[0:2]
        mm = dob[3:5]
        yyyy = dob[6:]
    except:
        break
    sa = {}
    driver = webdriver.Chrome('/home/akash/Downloads/chromedriver')
    driver.get('http://119.161.98.138/feedback/index.php')
    search_box = driver.find_element_by_id('username')
    try:
        search_box.send_keys(usn)
        dd = Select(driver.find_element_by_id('dd'))
        dd.select_by_visible_text(d)
        dd = Select(driver.find_element_by_id('mm'))
        dd.select_by_index(int(mm))
        dd = Select(driver.find_element_by_id('yyyy'))
        dd.select_by_visible_text(yyyy)
    except:
        print("Waste DOB")
        continue
    submit = driver.find_element_by_name('submit')
    submit.click()
    try:
        validate = driver.find_element_by_xpath('//*[@id="sims-container"]/table[1]/tbody/tr[2]/td')
        if "Pending Feedbacks" in validate.text:
            driver.get('http://119.161.98.138/sims')
            text1 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[1]/table/tbody/tr/td[2]')
            sub1 = text1[0].text
            p1 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[1]/div[4]/div[2]/table/tbody/tr[3]/td')
            a1 = p1[0].text
            text2 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[3]/table/tbody/tr/td[2]')
            sub2 = text2[0].text
            p2 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[3]/div[4]/div[2]/table/tbody/tr[3]/td')
            a2 = p2[0].text
            text3 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[5]/table/tbody/tr/td[2]')
            sub3 = text3[0].text
            p3 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[5]/div[4]/div[2]/table/tbody/tr[3]/td')
            a3 = p3[0].text
            text4 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[7]/table/tbody/tr/td[2]')
            sub4 = text4[0].text
            p4 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[7]/div[4]/div[2]/table/tbody/tr[3]/td')
            a4 = p4[0].text
            text5 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[9]/table/tbody/tr/td[2]')
            sub5 = text5[0].text
            p5 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[9]/div[4]/div[2]/table/tbody/tr[3]/td')
            a5 = p5[0].text
            text6 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[11]/table/tbody/tr/td[2]')
            sub6 = text6[0].text
            p6 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[11]/div[4]/div[2]/table/tbody/tr[3]/td')
            a6 = p6[0].text
            text7 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[13]/table/tbody/tr/td[2]')
            sub7 = text7[0].text
            p7 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[13]/div[4]/div[2]/table/tbody/tr[3]/td')
            a7 = p7[0].text
            text8 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[15]/table/tbody/tr/td[2]')
            sub8 = text8[0].text
            p8 = driver.find_elements_by_xpath('//*[@id="left-column"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/div[15]/div[4]/div[2]/table/tbody/tr[3]/td')
            a8 = p8[0].text
            sa[sub1] = a1
            sa[sub2] = a2
            sa[sub3] = a3
            sa[sub4] = a4
            sa[sub5] = a5
            sa[sub6] = a6
            sa[sub7] = a7
            sa[sub8] = a8
            atds[usn] = sa
        driver.close()
    except:
        print("Wrong DOB")
        driver.close()

with open('output.json','w') as outfile:
    json.dump(atds,outfile, indent=4)