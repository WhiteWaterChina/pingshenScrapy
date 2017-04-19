#!/usr/bin/env python
# -*- coding:cp936 -*-
import time
import os
import re
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import datetime
import wx
import wx.xrc


def sumtimesplit(strtimeList):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimeList:
        if re.search(u'��', item):
            timeList = item.split("��".decode('gbk'))
            timeOne = int(timeList[0]) * 86400
            timeTwo = int(timeList[1].split("Сʱ".decode('gbk'))[0]) * 3600
            totalTimeTemp = timeOne + timeTwo
            tempTimeFunc.append(totalTimeTemp)
        else:
            timeList = item.split("Сʱ".decode('gbk'))
            totalTimeTemp = int(timeList[0]) * 3600
            tempTimeFunc.append(totalTimeTemp)
    for item in tempTimeFunc:
        totalTime += item
    dayTime, hourtimeTemp = divmod(totalTime, 86400)
    hourTime = divmod(hourtimeTemp, 3600)[0]
    dataReturn = "%d��%dСʱ".decode('gbk') % (dayTime, hourTime)
    return dataReturn


class pingshenFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"����ϵͳ��Ϣץȡ����", pos=wx.DefaultPosition,
                          size=wx.Size(642, 280), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.test_title1 = wx.StaticText(self, wx.ID_ANY, u"�������������û���������", wx.DefaultPosition, wx.DefaultSize,
                                         wx.ST_NO_AUTORESIZE)
        self.test_title1.Wrap(-1)
        self.test_title1.SetFont(wx.Font(11, 70, 90, 90, False, "����"))
        self.test_title1.SetForegroundColour(wx.Colour(255, 0, 0))
        self.test_title1.SetBackgroundColour(wx.Colour(255, 255, 0))

        bSizer2.Add(self.test_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        gSizer2 = wx.GridSizer(2, 2, 0, 0)

        self.text_username = wx.StaticText(self, wx.ID_ANY, u"�û���", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_username.Wrap(-1)
        gSizer2.Add(self.text_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_username = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer2.Add(self.input_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.text_password = wx.StaticText(self, wx.ID_ANY, u"����", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_password.Wrap(-1)
        gSizer2.Add(self.text_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_password = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        gSizer2.Add(self.input_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer2.Add(gSizer2, 0, 0, 5)

        self.text_title2 = wx.StaticText(self, wx.ID_ANY, u"��������������Ҫץȡ����Ϣ����ֹ����(�������/��/����Ϣ����ʽΪ20170101.��λ�����º���һ��Ҫ��0����",
                                         wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_title2.Wrap(-1)
        self.text_title2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer2.Add(self.text_title2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        gSizer3 = wx.GridSizer(0, 2, 0, 0)

        self.text_startdate = wx.StaticText(self, wx.ID_ANY, u"��ʼ����", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_startdate.Wrap(-1)
        gSizer3.Add(self.text_startdate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_startdate = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer3.Add(self.input_startdate, 0, wx.ALL, 5)

        self.text_enddate = wx.StaticText(self, wx.ID_ANY, u"��������", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_enddate.Wrap(-1)
        gSizer3.Add(self.text_enddate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_enddate = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer3.Add(self.input_enddate, 0, wx.ALL, 5)

        bSizer2.Add(gSizer3, 0, 0, 5)

        m_sdbSizer1 = wx.StdDialogButtonSizer()
        self.m_sdbSizer1OK = wx.Button(self, wx.ID_OK)
        m_sdbSizer1.AddButton(self.m_sdbSizer1OK)
        self.m_sdbSizer1Cancel = wx.Button(self, wx.ID_CANCEL)
        m_sdbSizer1.AddButton(self.m_sdbSizer1Cancel)
        m_sdbSizer1.Realize()

        bSizer2.Add(m_sdbSizer1, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_sdbSizer1OK.Bind(wx.EVT_BUTTON, self.get_data)
        self.m_sdbSizer1Cancel.Bind(wx.EVT_BUTTON, self.close)

    def close(self, event):
        self.Close()


    def get_data(self, event):
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
        chromedriverPath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
        auditList = []
        nameList = []
        productNameList = []
        submitTimeList = []
        closeTimeList = []
        handleTimeList = []
        statusList = []
        reportList = []
        reportNameList = []
        totalTestTimeList = []
        start_data = int(self.input_startdate.GetValue())
        end_date = int(self.input_enddate.GetValue())
        browser = webdriver.Chrome(chromedriverPath)
        # browser = webdriver.PhantomJS()
        url = "http://218.57.146.175/techAudit/welcome.htm"
        browser.get(url)
        browser.find_element_by_id("loginName").send_keys(username)
        browser.find_element_by_id("password").send_keys(password)
        browser.find_element_by_css_selector("#loginForm > div > div.communalForm.clear > dl:nth-child(5) > dd > a").click()
        time.sleep(3)
        try:
            testEle = browser.find_element_by_css_selector("#msg")
            dlg_error = wx.MessageDialog(None, "�û������������������������ִ�г���".decode('gbk'), "�û���/�����������".decode('gbk'), wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
#            tkMessageBox.showerror("�û���/�����������".decode('gbk'), "�û������������������������ִ�г���".decode('gbk'))
            browser.close()
            if dlg_error.ShowModal() == wx.ID_OK:
                self.Close()
        except selenium.common.exceptions.NoSuchElementException:
            pass
        browser.switch_to.frame('ta')
        abc = browser.find_element_by_xpath("//div[@id='_easyui_tree_2']")
        abc.click()
        browser.implicitly_wait(10)
        browser.switch_to.default_content()
        browser.switch_to.frame(1)
        WebDriverWait(browser, 100).until(
            ec.element_to_be_clickable((By.XPATH, "//div[@class='datagrid-pager pagination']/table/tbody/tr/td[10]/a")))
        totalPagesEle = browser.find_element_by_css_selector(
            "body > div.panel.layout-panel.layout-panel-center > div > div > div > div.datagrid-pager.pagination > table > tbody > tr > td:nth-child(8) > span")
        time.sleep(5)
        totalPagesUnicode = totalPagesEle.text
        patternPages = re.compile(r'\d+', re.U)
        totalPages = int(re.findall(patternPages, totalPagesUnicode)[0])
        for pageCount in range(2, totalPages + 2):
            lineTotal = len(browser.find_elements_by_css_selector(
                "body > div.panel.layout-panel.layout-panel-center > div > div > div > div.datagrid-view > div.datagrid-view2 > div.datagrid-body > table > tbody > tr"))
            for count in range(0, lineTotal):
                auditNo = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(2) > div > span" % count).get_attribute("title")
                submitTimeTemp = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(7) > div" % count).text.split(" ")[0]
                submitTimeTempTemp = int("".join(submitTimeTemp.split('-')))
                if auditNo not in auditList and end_date >= submitTimeTempTemp >= start_data:
                    name = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(3) > div > span" % count).get_attribute("title")
                    productName = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(5) > div" % count).text
                    submitTimeTempTemp = "".join(submitTimeTemp.split('-'))
                    submitTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(submitTimeTemp, '%Y-%m-%d')))
                    closeTimeTemp = \
                        browser.find_element_by_css_selector(
                            "#datagrid-row-r1-2-%d > td:nth-child(8) > div" % count).text.split(
                            " ")[0]
                    if closeTimeTemp != u'':
                        closeTime = datetime.datetime.fromtimestamp(
                            time.mktime(time.strptime(closeTimeTemp, '%Y-%m-%d')))
                    else:
                        closeTime = '���������'.decode('gbk')
                    handleTime = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(9) > div" % count).text
                    status = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(12) > div" % count).text
                    lineData = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d" % count)
                    ActionChains(browser).double_click(lineData).perform()
                    time.sleep(2)
                    browser.switch_to.default_content()
                    browser.switch_to.frame(2)
                    try:
                        reportTemp = browser.find_element_by_css_selector(
                            "form#testUploadForm > table > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(1)")
                        reportList.append("�б���".decode('gbk'))
                        nameTemp = reportTemp.text[:-6]
                        reportNameList.append(nameTemp)
                    except selenium.common.exceptions.NoSuchElementException:
                        reportList.append("�ޱ���".decode('gbk'))
                        reportNameList.append("��".decode('gbk'))
                    try:
                        timeTestTemp = []
                        totalTd = browser.find_elements_by_tag_name("td")
                        for item in totalTd:
                            textTemp = item.text
                            if textTemp == "����".decode('gbk'):
                                timeTest = item.find_element_by_xpath("parent::tr/td[3]").text
                                timeTestTemp.append(timeTest)
                        timeTestData = sumtimesplit(timeTestTemp)
                        totalTestTimeList.append(timeTestData)
                    except selenium.common.exceptions.NoSuchElementException:
                        totalTestTimeList.append("�޲��Բ���".decode('gbk'))
                    browser.switch_to.default_content()
                    closeButton = browser.find_element_by_css_selector(
                        "#tabs > div.tabs-header.tabs-header-noborder > div.tabs-wrap > ul > li.tabs-selected > a.tabs-close")
                    closeButton.click()
                    browser.switch_to.frame(1)
                    auditList.append(auditNo)
                    nameList.append(name)
                    productNameList.append(productName)
                    submitTimeList.append(submitTime)
                    closeTimeList.append(closeTime)
                    handleTimeList.append(handleTime)
                    statusList.append(status)
            print("��ǰ����ץȡ��%dҳ���ܹ�%dҳ".decode('gbk') % (pageCount - 1, totalPages))
            inNum = browser.find_element_by_css_selector("input.pagination-num")
            inNum.clear()
            inNum.send_keys(pageCount)
            inNum.send_keys(Keys.ENTER)
            time.sleep(5)
        browser.quit()
        # ���������ݴ�������������ٷ�����ϵ
        TitleItem = ['������'.decode('gbk'), '��������'.decode('gbk'), '��Ŀ����'.decode('gbk'), '�ύʱ��'.decode('gbk'),
                     '�ر�ʱ��'.decode('gbk'), '����ʱ��'.decode('gbk'), '���Ի���ʱ��'.decode('gbk'), '״̬'.decode('gbk'),
                     '�Ƿ��б��渽��'.decode('gbk'),
                     '��������'.decode('gbk'), ]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("����ϵͳץȡ��Ϣ-%s.xlsx".decode('gbk') % timestamp)
        SheetOne = WorkBook.add_worksheet('����ϵͳץȡ��Ϣ'.decode('gbk'))
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)
        formatTwo = WorkBook.add_format()
        formatTwo.set_border(1)
        formatTwo.set_num_format('yy/mm/dd')
        SheetOne.set_column('A:J', 14)
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], formatOne)
        lineCount = 1
        for index, item in enumerate(auditList):
            SheetOne.write(lineCount, 0, auditList[index], formatOne)
            SheetOne.write(lineCount, 1, nameList[index], formatOne)
            SheetOne.write(lineCount, 2, productNameList[index], formatOne)
            SheetOne.write(lineCount, 3, submitTimeList[index], formatTwo)
            SheetOne.write(lineCount, 4, closeTimeList[index], formatTwo)
            SheetOne.write(lineCount, 5, handleTimeList[index], formatOne)
            SheetOne.write(lineCount, 6, totalTestTimeList[index], formatOne)
            SheetOne.write(lineCount, 7, statusList[index], formatOne)
            SheetOne.write(lineCount, 8, reportList[index], formatOne)
            SheetOne.write(lineCount, 9, reportNameList[index], formatOne)
            lineCount += 1
        WorkBook.close()
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        dlg_info = wx.MessageDialog(None, 'ץ��%s��������Ѿ������д�롶����ϵͳץȡ��Ϣ.xlsx���������в��ģ�����Cancel�˳�����'.decode('gbk') % len(auditList), '�����ʾ'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg_info.ShowModal()
#        tkMessageBox.showinfo('�����'.decode('gbk'), 'ץ��%s��������Ѿ������д�롶����ϵͳץȡ��Ϣ.xlsx���������в��ģ�����EXIT�˳�����'.decode('gbk') % len(auditList))


if __name__ == '__main__':
    app = wx.App()
    frame = pingshenFrame(None)
    frame.Show()
    app.MainLoop()
