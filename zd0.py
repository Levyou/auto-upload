#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2018/11/16 0016 9:50
# @Author  : y
from selenium import webdriver
import time
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import pandas as pd
import win32com.client

data = pd.read_csv(r'C:\Users\Mr.You\Desktop\rocca\22.csv')

driver = webdriver.Chrome()
driver.get('http://pppluopan.rocappp.com:8080/pppcompass/page/login/login')

time.sleep(30)


for i in range(78, len(data.index)):
    try:
        driver.switch_to.frame('iframe_11')
        driver.find_element_by_xpath("//div[@class='btn-group']/button").click()
        time.sleep(1)

        driver.find_element_by_id('infoTitle').send_keys(data['政策名称'][i])

        driver.find_element_by_id('infoAuthor').send_keys(data['发文单位'][i])

        driver.find_element_by_xpath("//div[@class='clearfix']/button").click()
        time.sleep(1)
        driver.find_element_by_xpath("//div[@class='treeview']/ul/li[2]/span").click()
        time.sleep(1)
        if data['级别'][i] == '省级':
            driver.find_element_by_xpath("//div[@class='treeview']/ul/li[4]").click()
        else:
            driver.find_element_by_xpath("//div[@class='treeview']/ul/li[5]").click()
        time.sleep(1)
        driver.find_element_by_xpath(".//*[@onclick='confirm();']").click()

        selector1 = Select(driver.find_element_by_id("province"))
        selector1.select_by_visible_text(data['省'][i])
        time.sleep(1)
        selector2 = Select(driver.find_element_by_id("city"))
        selector2.select_by_visible_text(data['地'][i])

        driver.switch_to.frame('ueditor_0') #注意这种editor一定有frame,一定要切换frame
        # list_0 = data['content'][i].split('<br>')
        # for j in range(0, len(list_0)):
        #     answer_text = driver.find_element_by_tag_name('body')
        #     answer_text.send_keys(list_0[j])
        #     answer_text.send_keys(Keys.ENTER)
        # 加载应用
        app = win32com.client.Dispatch('Word.Application')
        # 打开word，经测试要是绝对路径
        doc = app.Documents.Open(r'C:\Users\Mr.You\Desktop\rocca\zcwj\{}\{}.docx'.format(data['bag'][i], data['name'][i]))
        # 复制word的所有内容
        doc.Content.Copy()
        # 关闭word
        doc.Close()

        answer_text = driver.find_element_by_tag_name('body')
        answer_text.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        driver.switch_to.parent_frame()
        driver.find_element_by_xpath(".//*[@onclick='save(2);']").click()
        time.sleep(2)

        driver.switch_to.default_content()
    except Exception as e:
        print(data['政策名称'][i])
        driver.switch_to.frame('iframe_11')
        driver.find_element_by_xpath(".//*[@onclick='cancel();']").click()
        time.sleep(2)
        driver.switch_to.default_content()
