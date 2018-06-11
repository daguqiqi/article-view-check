#-*- coding: UTF-8 -*-

from other_packages import platform
from other_packages.Tkinter import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import os
import xlwt

def btn_click():

    article_key = e1.get()

    if article_key != "":
        if platform.system() == "Windows":
            web_f = open(os.path.split(os.path.realpath(__file__))[0] + '\websites.txt', 'r')
            chromedriver = os.path.split(os.path.realpath(__file__))[0] + '\chromedriver.exe'
            os.environ["webdriver.chrome.driver"] = chromedriver
        else:
            chromedriver = "driver/chromedriver"
            web_f = open(os.path.split(os.path.realpath(__file__))[0] + '/websites.txt', 'r')

        driver = webdriver.Chrome(chromedriver)
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
        sheet.col(0).width = 200 * 20
        sheet.col(1).width = 600 * 20
        sheet.col(2).width = 300 * 20
        sheet.col(3).width = 800 * 20
        sheet.write(0, 0, '网站')
        sheet.write(0, 1, '文章名')
        sheet.write(0, 2, '作者')
        sheet.write(0, 3, '文章链接')
        sheet.write(0, 4, '浏览量')
        try:
            index = 0
            for web in web_f:
                index = index+1
                article_name = "无结果"
                article_author = "无结果"
                article_url = "无结果"
                article_viewnum = "无结果"
                if "嘻哈" in web:
                    url="http://www.xhcjtv.com/portal/search/index/type/1/keyword/"+article_key+".html"
                    driver.get(url)
                    WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator=(By.CLASS_NAME, 'left')))
                    article_list = driver.find_elements_by_class_name("news_list")
                    if len(article_list)!=0:
                        text = article_list[0].find_element_by_class_name("news_text").find_element_by_tag_name("h3").text + ""
                        article_name = text.encode("utf-8").split("丨")[1]
                        article_author = text.encode("utf-8").split("丨")[0]
                        article_url = article_list[0].get_attribute("href")
                        article_viewnum = article_list[0].find_elements_by_class_name("time_text")[1].text
                        driver.save_screenshot(os.path.split(os.path.realpath(__file__))[0] + '/统计结果/'+web+'.png');
                if "金色" in web:
                    url = "https://www.jinse.com/search/" + article_key
                    driver.get(url)
                    WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator=(By.CLASS_NAME, 'ja-article-list')))
                    article_list = driver.find_element_by_class_name("article-main").find_elements_by_class_name("clear")
                    if len(article_list) != 0:
                        text = article_list[0].find_element_by_tag_name("h3").find_element_by_tag_name("a").get_attribute("title") + ""
                        article_name = text.encode("utf-8").split("|")[0]
                        article_author = text.encode("utf-8").split("|")[1]
                        article_url = article_list[0].find_element_by_tag_name("h3").find_element_by_tag_name("a").get_attribute("href")
                        article_viewnum = article_list[0].find_element_by_class_name("fr").text
                        driver.save_screenshot(os.path.split(os.path.realpath(__file__))[0] + '/统计结果/' + web + '.png');
                if "火星" in web:
                    url = "http://www.huoxing24.com/search/" + article_key
                    driver.get(url)
                    WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator=(By.CLASS_NAME, 'search-result-list')))
                    article_list = driver.find_elements_by_class_name("index-news-list")
                    if len(article_list) != 0:
                        article_name = article_list[0].find_element_by_class_name("list-right").find_element_by_tag_name("a").get_attribute("title")
                        article_author = article_list[0].find_element_by_class_name("portrait").find_element_by_tag_name("a").get_attribute("title")
                        article_url = article_list[0].find_element_by_class_name("list-right").find_element_by_tag_name("a").get_attribute("href")
                        article_viewnum = article_list[0].find_element_by_class_name("read-count").text
                        driver.save_screenshot(os.path.split(os.path.realpath(__file__))[0] + '/统计结果/' + web + '.png');
                if "巴比特" in web:
                    url = "http://www.8btc.com/?s=" + article_key
                    driver.get(url)
                    WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator=(By.CLASS_NAME, 'mainstay')))
                    article_list = driver.find_elements_by_tag_name("article")
                    if len(article_list) != 0:
                        article_name = article_list[0].find_element_by_class_name("article-title").find_element_by_tag_name("a").get_attribute("title")
                        article_author = article_list[0].find_element_by_class_name("article-info").find_element_by_tag_name("a").text
                        article_url = article_list[0].find_element_by_class_name("article-title").find_element_by_tag_name("a").get_attribute("href")
                        driver.get(article_url)
                        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator=(By.CLASS_NAME, 'fa-eye-span')))
                        article_viewnum = driver.find_element_by_class_name("fa-eye-span").text
                        driver.save_screenshot(os.path.split(os.path.realpath(__file__))[0] + '/统计结果/' + web + '.png');

                sheet.write(index, 0, web)
                sheet.write(index, 1, article_name)
                sheet.write(index, 2, article_author)
                sheet.write(index, 3, article_url)
                sheet.write(index, 4, article_viewnum)

            web_f.close()
            driver.quit()
            book.save(os.path.split(os.path.realpath(__file__))[0] + '/统计结果/【'+article_key.encode("utf-8")+'】统计结果.xls')
            str = "测试结束！"
            if platform.system() == "Windows":
                dir = os.path.split(os.path.realpath(__file__))[0] + '/统计结果/【'+article_key.encode("utf-8")+'】统计结果.xls'
                print str.decode('UTF-8').encode('GBK')
            else:
                dir = "open " + os.path.split(os.path.realpath(__file__))[0] + '/统计结果/【'+article_key.encode("utf-8")+'】统计结果.xls'
                print str
            os.system(dir)
        # except Exception, e:
        #     error = "如果你看到这句话那么有下面两种可能：\n1.你没有找到适合自己的ChromeDriver\n2.你输入的URL不正确\n请具体看下面的错误！"
        #     if platform.system() == "Windows":
        #         print error.decode('UTF-8').encode('GBK')
        #     else:
        #         print error
        #     print e
        finally:
            web_f.close()
            driver.quit()
    else:
        error = "请填写文章名称！"
        if platform.system() == "Windows":
            print error.decode('UTF-8').encode('GBK')
        else:
            print error

root = Tk()
root.title("文章阅读量统计工具")
root.geometry("400x150")

Label(root, text="   请输入想统计的文章名称：", font=("Arial",20)).pack()
frm = Frame(root)

frm1 = Frame(frm)
Label(frm1, text="文章名称",  font=("Arial",15)).pack(side=LEFT)
e1 = Entry(frm1)

e1.pack()

frm1.pack()

b = Button(frm, text = '开始统计', command=btn_click)
b.pack()

frm.pack()
e1.insert(0,"")
root.mainloop()

