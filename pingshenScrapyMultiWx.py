#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import multiprocessing
import xlsxwriter
import os
import time
import datetime
from threading import Thread
import wx
from multiprocessing import Pool
import base64
import sys
from bs4 import element

sys.setrecursionlimit(3000)

ver = "Ward Yan-20200915"
web_address = "218.57.146.175:8114"


def sumtimesplit(strtimelist):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimelist:
        if re.search(u'��', item):
            timeList = item.split("��")
            timeOne = int(timeList[0]) * 86400
            timeTwo = int(timeList[1].split("Сʱ")[0]) * 3600
            totalTimeTemp = timeOne + timeTwo
            tempTimeFunc.append(totalTimeTemp)
        else:
            timeList = item.split("Сʱ")
            totalTimeTemp = int(timeList[0]) * 3600
            tempTimeFunc.append(totalTimeTemp)
    for item in tempTimeFunc:
        totalTime += item
    dayTime, hourtimeTemp = divmod(totalTime, 86400)
    hourTime = divmod(hourtimeTemp, 3600)[0]
    dataReturn = "{}��{}Сʱ".format(dayTime, hourTime)
    return dataReturn


def get_status(status):
    switcher = {
        "0": "����",
        "1": "�ύ",
        "audit-submit": "�޸�������Ϣ",
        "shenhepeizhi-sq": "��ǰ�������",
        "shenhe-chanpinjingli": "��Ʒ�������",
        "querenxuanpei-ddy": "����Աȷ��ѡ��",
        "shenhepingshen-yf": "�з��ӿ����������",
        "shenhepingshen-csjk": "���Խӿ����������",
        "ceshi-cs": "������Ա����",
        "shenheceshibaogao-yf": "�з��ӿ���˲��Ա���",
        "shenheceshibaogao-xmjl": "��Ŀ������˲��Ա���",
        "shenheceshibaogao-csfzr": "���Ը�������˲��Ա���",
        "shenheceshibaogao-csjk": "���Խӿ�����˲��Ա���",
        "xfzl-gc": "������Աȷ���Ƿ��·�ָ��",
        "100": "�ر�",
        "101": "�쳣�ر�",
        "102": "��ͣ",
        "103": "��ֹ",
        "vm-audit": "VM���",
        "npi-audit": "NPI����",
        "leader-test-judge": "����teamleader���Ծ���",
        "exec-test": "ִ�в���",
        "leader-audit": "����teamleader��˲��Խ��",
        "vm-test-audit": "VM��˲��Խ��",
        "test_report_concordance": "���Ա�������",
        "os-comp-test": "OS�����Բ���",
        "product-verification": "������֤",
        "oqc-valication": "OQC��֤",
        "material-add": "����׷��",
        "material-assess": "��������",
        "om-maintaince": "BOMά��",
        "custom-dev": "���ƻ�����",
        "hadware-dev": "�̼��з�",
    }
    return switcher.get(status, status)


def get_detail(link, login_session):
    headers_data_all = {
        'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.9,en;q=0.8",
        'Connection': "keep-alive",
        'Host': "{}".format(web_address),
        'Referer': "http://{}/techAudit/welcome.htm".format(web_address),
        'Upgrade-Insecure-Requests': "1",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36"
    }
    get_page = login_session.get(link, headers=headers_data_all)
    data_page = get_page.text
    print("Get link:{} with return code {}".format(link, get_page.status_code))
    data_filter = BeautifulSoup(data_page, "html5lib")
    # ��ȡ������Ϣ
    attachment_temp = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr")
    if len(attachment_temp) <= 1:
        has_report = "�ޱ���"
        report_filename = "None"
    else:
        has_report = "�б���"
        filename_temp = []
        for index_report, item_report in enumerate(attachment_temp):
            if index_report != 0:
                filename = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr:nth-of-type(2) > td:nth-of-type(1)")[0].contents[0].strip()
                filename_temp.append(filename)
        filename_write = ";".join(filename_temp)
        report_filename = filename_write
    # ����Ҫ����Ϣ
    try:
        keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(9) > td:nth-of-type(1)")[0].get_text()
    except IndexError:
        try:
            keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(7) > td:nth-of-type(1)")[0].get_text()
        except IndexError:
            keywords_temp = "None"
    keywords = keywords_temp
    # ���Ի���ʱ��
    timeTestTemp = []
    try:
        test_item = data_filter.select("body > div:nth-of-type(1) > table:nth-of-type(3) > tbody > tr")
        for item_tr in test_item:
            item_td = item_tr.select("td:nth-of-type(5)")[0].get_text().strip()
            if item_td == "����":
                test_time = item_tr.select("td:nth-of-type(3)")[0].get_text().strip()
                timeTestTemp.append(test_time)
        if len(timeTestTemp) == 0:
            timeTestData = "None"
        else:
            timeTestData = sumtimesplit(timeTestTemp)
        test_time = timeTestData
    except IndexError:
        test_time = "None"
    # NPI���һ�ε�������Ϣ
    npi_comment = ""
    npi_comment_list = data_filter.find_all('td',text="NPI����")
    # npi_comment_list = data_filter.select('td[text="NPI����"]')
    if npi_comment_list is not None:
        print("Got one!")
        if len(npi_comment_list) != 0:
            npi_comment_element = npi_comment_list[-1]
            npi_comment_1 = npi_comment_element.next_sibling.next_sibling.next_sibling.next_sibling
            print(type(npi_comment_1))
            if isinstance(npi_comment_1, element.Tag):
                npi_comment = npi_comment_1.text
                print(npi_comment)
        print(1)

    return link, has_report, report_filename, keywords, test_time, npi_comment


class PingShenFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"����ϵͳ��Ϣץȡ����-{}".format(ver), pos=wx.DefaultPosition,
                          size=wx.Size(504, 785), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"�������������û���������", wx.DefaultPosition, wx.DefaultSize,
                                         wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer3, 0, wx.EXPAND, 5)

        gSizer2 = wx.GridSizer(2, 2, 0, 0)

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"�û���", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_username.Wrap(-1)

        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        gSizer2.Add(self.input_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"����", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_password.Wrap(-1)

        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        gSizer2.Add(self.input_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(gSizer2, 0, 0, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_title2 = wx.StaticText(self.m_panel1, wx.ID_ANY,
                                         u"��������������Ҫץȡ����Ϣ����ֹ����(�������/��/����Ϣ��\n��ʽΪ20170101.��λ�����º���һ��Ҫ��0����",
                                         wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.text_title2.Wrap(-1)

        self.text_title2.SetFont(
            wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_title2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer10.Add(bSizer4, 0, wx.EXPAND, 5)

        gSizer3 = wx.GridSizer(0, 2, 0, 0)

        self.text_startdate = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��ʼ����", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_startdate.Wrap(-1)

        self.text_startdate.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_startdate.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer3.Add(self.text_startdate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_startdate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                           0)
        gSizer3.Add(self.input_startdate, 0, wx.ALL, 5)

        self.text_enddate = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��������", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_enddate.Wrap(-1)

        self.text_enddate.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_enddate.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer3.Add(self.text_enddate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_enddate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                         0)
        gSizer3.Add(self.input_enddate, 0, wx.ALL, 5)

        bSizer10.Add(gSizer3, 0, 0, 5)

        bSizer13 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText9 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"�밴���°�ť����ȡ��Ŀ��Ϣ", wx.DefaultPosition,
                                           wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.m_staticText9.Wrap(-1)

        self.m_staticText9.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText9.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer13.Add(self.m_staticText9, 0, wx.ALL | wx.EXPAND, 5)

        self.button_get_productname = wx.Button(self.m_panel1, wx.ID_ANY, u"��ȡ��Ŀ����", wx.DefaultPosition, wx.DefaultSize,
                                                0)
        bSizer13.Add(self.button_get_productname, 0, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer13, 0, wx.EXPAND, 5)

        bSizer11 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText8 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��������ѡ����Ҫ��ȡ���ݵĲ�Ʒ�ͺ�!���Զ�ѡ��", wx.DefaultPosition,
                                           wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.m_staticText8.Wrap(-1)

        self.m_staticText8.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText8.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer11.Add(self.m_staticText8, 1, wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        listbox_productnameChoices = []
        self.listbox_productname = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                              listbox_productnameChoices, wx.LB_HSCROLL | wx.LB_MULTIPLE | wx.LB_SORT)
        bSizer11.Add(self.listbox_productname, 0, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer11, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.text_3 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��������ѡ����Ҫ������ļ�����ʾ����Ŀ", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        self.text_3.Wrap(-1)

        self.text_3.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_3.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_3.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer9.Add(self.text_3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.VERTICAL)

        bSizer61 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkBox_audit = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"������", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_audit.SetValue(True)
        bSizer61.Add(self.checkBox_audit, 0, wx.ALL, 5)

        self.checkBox_name = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"��������", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_name.SetValue(True)
        bSizer61.Add(self.checkBox_name, 0, wx.ALL, 5)

        self.checkBox_productName = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"��Ŀ����", wx.DefaultPosition, wx.DefaultSize,
                                                0)
        self.checkBox_productName.SetValue(True)
        bSizer61.Add(self.checkBox_productName, 0, wx.ALL, 5)

        self.checkBox_submitTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"�ύʱ��", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_submitTime.SetValue(True)
        bSizer61.Add(self.checkBox_submitTime, 0, wx.ALL, 5)

        self.checkBox_closeTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"�ر�ʱ��", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_closeTime.SetValue(True)
        bSizer61.Add(self.checkBox_closeTime, 0, wx.ALL, 5)

        bSizer6.Add(bSizer61, 0, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkBox_handleTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"����ʱ��", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_handleTime.SetValue(True)
        bSizer8.Add(self.checkBox_handleTime, 0, wx.ALL, 5)

        self.checkBox_totalTestTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"���Ի���ʱ��", wx.DefaultPosition,
                                                  wx.DefaultSize, 0)
        self.checkBox_totalTestTime.SetValue(True)
        bSizer8.Add(self.checkBox_totalTestTime, 0, wx.ALL, 5)

        self.checkBox_status = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"��ǰ״̬", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_status.SetValue(True)
        bSizer8.Add(self.checkBox_status, 0, wx.ALL, 5)

        self.checkBox_report = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"�Ƿ��б��渽��", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_report.SetValue(True)
        bSizer8.Add(self.checkBox_report, 0, wx.ALL, 5)

        self.checkBox_summary = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"����Ҫ��", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_summary.SetValue(True)
        bSizer8.Add(self.checkBox_summary, 0, wx.ALL, 5)

        bSizer6.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer10.Add(bSizer6, 0, wx.EXPAND, 5)

        bSizer21 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self.m_panel1, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer10.Add(bSizer21, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.textctrl_display = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.textctrl_display, 1, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer91, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_get_productname.Bind(wx.EVT_BUTTON, self.get_productname)
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def get_productname(self, event):
        self.updatedisplay("��ʼ��ȡ��Ŀ������Ϣ,�����ĵȴ�")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue().encode()).decode()
        # ��¼
        url_login = "http://{}/techAudit/login.htm".format(web_address)
        login_session = requests.session()
        headers_login = {
            'accept': "tapplication/json, text/javascript, */*; q=0.01",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9,en;q=0.8",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/logout.htm?service=http://{}/techAudit/welcome.htm&reload=true".format(web_address, web_address),
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest",
        }
        payload_login = "loginName={}&password={}".format(username, password)
        #post��¼
        login_session.post(url_login, headers=headers_login, data=payload_login)
        # ��ʼ��ȡ����
        headers_data = {
            'Accept': "application/json, text/javascript, */*; q=0.01",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/auditList/queryAuditList.htm".format(web_address),
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest"
        }
        # ����1��ȡ�����������
        payload_data_test = "row=1"
        url_data = "http://{}/techAudit/AuditListController/getQueryAuditList.htm?".format(web_address)
        response_data_test = login_session.post(url_data, headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        # ץȡ����ÿҳ5000�������У������������/5000������ץȡ��ҳ�����pages_number
        if int(total_item) % 1000 == 0:
            pages_number = int(total_item) // 1000
        else:
            pages_number = int(total_item) // 1000 + 1

        productname_list_all = []

        # Ȼ��ʹ��ÿҳ1000������ҳץȡ
        for item_pages_temp in range(pages_number):
            item_pages = item_pages_temp + 1
            payload_data = "page={}&rows=1000".format(item_pages)
            print(payload_data)
            response_data_get = login_session.post(url_data, headers=headers_data, data=payload_data)
            response_data = response_data_get.text
            # ��ȡproductname����Ϣ
            productname_list_temp = re.findall(r'"modeName":"(.*?)",', response_data)
            for item_productname_temp in productname_list_temp:
                if item_productname_temp not in productname_list_all:
                    productname_list_all.append(item_productname_temp)
        for item in productname_list_all:
            self.listbox_productname.Append(item)
        self.updatedisplay("ץȡ��Ŀ��Ϣ����,��������ѡ����Ҫץȡ��Ϣ����Ŀ���ƣ�Ȼ����GO��ʼץȡ��")
        diag_finish_project = wx.MessageDialog(None, "��ȡ��Ŀ��Ϣ��ɣ���ѡ����Ҫץȡ��Ϣ����Ŀ���ƣ�Ȼ����GO��ʼץȡ��", '��ʾ',
                                               wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish_project.ShowModal()

    def run_all(self):
        self.button_go.Disable()
        self.updatedisplay("��ʼץȡ�������ĵȴ�...")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        # ��ȡ�û���������
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue().encode()).decode()
        # ��ȡ��ʼ�ͽ���ʱ��
        start_time = self.input_startdate.GetValue()
        end_time = self.input_enddate.GetValue()
        # ��ȡѡ�����Ŀ
        productname_selected_list = []
        productname_selected_index_list = self.listbox_productname.GetSelections()
        for item in productname_selected_index_list:
            productname_selected_list.append(self.listbox_productname.GetString(item))
        # ��¼
        login_session = requests.session()

        url_login = "http://{}/techAudit/login.htm".format(web_address)
        headers_login = {
            'accept': "tapplication/json, text/javascript, */*; q=0.01",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9,en;q=0.8",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/logout.htm?service=http://{}/techAudit/welcome.htm&reload=true".format(web_address, web_address),
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest",
        }
        payload_login = "loginName={}&password={}".format(username, password)
        # ʹ�����ϻ�ȡ����Ϣpost��¼
        login_session.post(url_login, headers=headers_login, data=payload_login)
        # ��ʼ��ȡ����
        headers_data = {
            'Accept': "application/json, text/javascript, */*; q=0.01",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/auditList/queryAuditList.htm".format(web_address),
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest"
        }
        # ����1��ȡ�����������
        payload_data_test = "rows=1"
        url_data = "http://218.57.146.175/techAudit/AuditListController/getQueryAuditList.htm?"
        response_data_test = login_session.post(url_data, headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        # ץȡ����ÿҳ1000�������У������������/1000������ץȡ��ҳ�����pages_number
        if int(total_item) % 1000 == 0:
            pages_number = int(total_item) // 1000
        else:
            pages_number = int(total_item) // 1000 + 1
        auditno_list = []
        id_list = []
        create_date_list = []
        close_date_list = []
        projectname_list = []
        productname_list = []
        status_list = []
        url_list = []
        total_time_list = []
        # Ȼ��ʹ��ÿҳ1000������ҳץȡ
        for item_pages_temp in range(pages_number):
            item_pages = item_pages_temp + 1
            payload_data = "page={}&rows=1000".format(item_pages)
            print(payload_data)
            response_data_get = login_session.post(url_data, headers=headers_data, data=payload_data)
            response_data = response_data_get.text
            # �Ȼ�ȡautidno/id/creattime/updatetime/projectname/productname��ԭʼ����
            auditno_list_temp = re.findall(r'"auditNo":"(\w+)",', response_data)
            id_list_temp = re.findall(r'"id":(\d+),', response_data)
            create_date_list_temp = re.findall(r'"submitTime":"(\d+-\d+-\d+)', response_data)  # 2016-09-23
            close_date_list_temp_1 = re.findall(r'"closeTime":(.*?),', response_data)  # "closeTime":"2015-04-10 08:42:21", or "closeTime":null
            projectname_list_temp = re.findall(r'"projectName":"(.*?)",', response_data)
            productname_list_temp = re.findall(r'"modeName":"(.*?)",', response_data)
            status_list_temp = re.findall(r'"status":"(.*?)",', response_data)
            total_time_list_temp = re.findall(r'"handleTime":"(.*?)"', response_data)
            # ת��Ҫ��Ŀ�ʼ�ͽ���ʱ�����1970��1��1�յ�����
            start_date = time.mktime(time.strptime(start_time, '%Y%m%d'))
            end_date = time.mktime(time.strptime(end_time, '%Y%m%d'))
            # ����ر�ʱ��Ķ������
            close_date_list_temp = []
            for item_close_date in close_date_list_temp_1:
                if item_close_date == "null":
                    close_date_list_temp.append("None")
                else:
                    temp_close_date = item_close_date.strip().split(" ")[0][1:]
                    close_date_list_temp.append(temp_close_date)

            # ����һ�飬ȥ��ʱ�䲻����Ҫ��ġ�����ѡ�����Ŀ���Ƶ�
            auditno_list_every_page = []
            id_list_every_page = []
            create_date_list_every_page = []
            close_date_list_every_page = []
            projectname_list_every_page = []
            productname_list_every_page = []
            status_list_every_page = []
            url_list_every_page = []
            total_time_list_every_page = []

            for index_creattime, item_creattime in enumerate(create_date_list_temp):
                create_date_format = time.mktime(time.strptime(item_creattime, '%Y-%m-%d'))
                pruductname_every = productname_list_temp[index_creattime]
                # ʱ���ж� + ��Ʒ�����ж�
                if float(start_date) <= float(create_date_format) <= float(end_date) and pruductname_every in productname_selected_list:
                    # ����ʱ�䣬������
                    create_date_list_every_page.append(item_creattime)
                    # �����ţ���һ��
                    auditno_list_every_page.append(auditno_list_temp[index_creattime])
                    url = "http://{}/techAudit/v1Details/viewBill.htm?id=".format(web_address) + id_list_temp[index_creattime]
                    # ÿ�������url�е�Ψһ��ţ���Ϊ����ǰ���ȡ�������ݺͺ���ÿ����������ݵ�Ŧ��
                    id_list_every_page.append(id_list_temp[index_creattime])
                    # �ر�ʱ�䣬������
                    close_date_list_every_page.append(close_date_list_temp[index_creattime])
                    # ��Ŀ���ƣ��ڶ���
                    projectname_list_every_page.append(projectname_list_temp[index_creattime])
                    # ��Ʒ���ƣ�������
                    productname_list_every_page.append(productname_list_temp[index_creattime])
                    # ��ǰ״̬���ڰ���
                    status_list_every_page.append(get_status(status_list_temp[index_creattime]))
                    url_list_every_page.append(url)
                    # �ܻ���ʱ�䣬������
                    total_time_list_every_page.append(total_time_list_temp[index_creattime])


            auditno_list.extend(auditno_list_every_page)
            id_list.extend(id_list_every_page)
            create_date_list.extend(create_date_list_every_page)
            close_date_list.extend(close_date_list_every_page)
            projectname_list.extend(projectname_list_every_page)
            productname_list.extend(productname_list_every_page)
            status_list.extend(status_list_every_page)
            url_list.extend(url_list_every_page)
            total_time_list.extend(total_time_list_every_page)

        print("auditno_list:{}".format(len(auditno_list)))
        print("id_list:{}".format(len(id_list)))
        print("create_date_list:{}".format(len(create_date_list)))
        print("close_date_list:{}".format(len(close_date_list)))
        print("projectname_list:{}".format(len(projectname_list)))
        print("productname_list:{}".format(len(productname_list)))
        print("status_list:{}".format(len(status_list)))
        print("url_list:{}".format(len(url_list)))
        print("total_time_list:{}".format(len(total_time_list)))
        self.updatedisplay("��������{}������ʱ�����Ŀ����Ҫ���������Ⱥ�ץȡ����~~~~".format(len(auditno_list)))

        # ��ȡÿ���������ϸҳ����Ϣ
        dict_data_detail = {}
        for item_1 in id_list:
            dict_data_detail["{}".format(item_1)] = []
        temp_detail = []
        pool_detail = Pool()
        for index, item_2 in enumerate(url_list):
            temp_detail.append(pool_detail.apply_async(get_detail, args=(item_2, login_session)))
        pool_detail.close()
        pool_detail.join()
        # return link, has_report, report_filename, keywords, test_time
        for item_detail in temp_detail:
            data_detail_temp = item_detail.get()
            if data_detail_temp is not None:
                # if data_detail_temp[0] != "None":
                num = data_detail_temp[0].split("=")[-1]
                index_to_log = id_list.index(num)
                # ������ audit_number
                dict_data_detail["{}".format(num)].append(auditno_list[index_to_log])
                # �������� project name
                dict_data_detail["{}".format(num)].append(projectname_list[index_to_log])
                # ��Ʒ���� product name
                dict_data_detail["{}".format(num)].append(productname_list[index_to_log])
                # �ύʱ�� create time
                dict_data_detail["{}".format(num)].append(create_date_list[index_to_log])
                # �ر�ʱ�� close time
                dict_data_detail["{}".format(num)].append(close_date_list[index_to_log])
                # �ܴ���ʱ�� total time
                dict_data_detail["{}".format(num)].append(total_time_list[index_to_log])
                # ���Ի���ʱ�� test time
                dict_data_detail["{}".format(num)].append(data_detail_temp[4])
                # ״̬ status
                dict_data_detail["{}".format(num)].append(status_list[index_to_log])
                # �Ƿ��б��� has_report
                dict_data_detail["{}".format(num)].append(data_detail_temp[1])
                # �������ƣ� report filename
                dict_data_detail["{}".format(num)].append(data_detail_temp[2])
                # ����Ҫ�� keywords
                dict_data_detail["{}".format(num)].append(data_detail_temp[3])
                # NPI��ע
                dict_data_detail["{}".format(num)].append(data_detail_temp[5])

        autidno_list_write = []
        projectname_list_write = []
        productname_list_write = []
        create_date_list_write = []
        close_date_list_write = []
        total_time_list_write = []
        test_time_list_write = []
        status_list_write = []
        has_report_list_write = []
        report_filename_list_write = []
        keywords_list_write = []
        npi_comment_list_write = []

        for item_write in dict_data_detail:
            autidno_list_write.append(dict_data_detail[item_write][0])
            projectname_list_write.append(dict_data_detail[item_write][1])
            productname_list_write.append(dict_data_detail[item_write][2])
            create_date_list_write.append(dict_data_detail[item_write][3])
            close_date_list_write.append(dict_data_detail[item_write][4])
            total_time_list_write.append(dict_data_detail[item_write][5])
            test_time_list_write.append(dict_data_detail[item_write][6])
            status_list_write.append(dict_data_detail[item_write][7])
            has_report_list_write.append(dict_data_detail[item_write][8])
            report_filename_list_write.append(dict_data_detail[item_write][9])
            keywords_list_write.append(dict_data_detail[item_write][10])
            npi_comment_list_write.append(dict_data_detail[item_write][11])

        # д�뵽����xlsx�ĵ�
        TitleItem = ['������', '��������', '��Ŀ����', '�ύʱ��', '������ʱ��', '����ʱ��', '���Ի���ʱ��', '״̬', '�Ƿ��б��渽��', '��������', '����Ҫ��', 'NPI��ע']
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("����ϵͳץȡ��Ϣ-{}.xlsx".format(timestamp))
        SheetOne = WorkBook.add_worksheet('����ϵͳץȡ��Ϣ')
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)

        SheetOne.set_column('A:J', 15)
        already_write_list = []
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], formatOne)
        for index_write, item_write in enumerate(autidno_list_write):
            if len(item_write) != 0:
                if item_write not in already_write_list:
                    already_write_list.append(item_write)
                    if self.checkBox_audit.GetValue():
                        SheetOne.write(1 + index_write, 0, item_write, formatOne)
                    if self.checkBox_name.GetValue():
                        SheetOne.write(1 + index_write, 1, projectname_list_write[index_write], formatOne)
                    if self.checkBox_productName.GetValue():
                        SheetOne.write(1 + index_write, 2, productname_list_write[index_write], formatOne)
                    if self.checkBox_submitTime.GetValue():
                        SheetOne.write_datetime(1 + index_write, 3,
                                                datetime.datetime.strptime(create_date_list_write[index_write], '%Y-%m-%d'),
                                                WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                    if self.checkBox_closeTime.GetValue():
                        if re.search(r'\d+', close_date_list_write[index_write]) is None:
                            SheetOne.write(1 + index_write, 4, "���������", formatOne)
                        else:
                            SheetOne.write(1 + index_write, 4, datetime.datetime.strptime(close_date_list_write[index_write], '%Y-%m-%d'), WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                    if self.checkBox_handleTime.GetValue():
                        SheetOne.write(1 + index_write, 5, total_time_list_write[index_write], formatOne)
                    if self.checkBox_totalTestTime.GetValue():
                        SheetOne.write(1 + index_write, 6, test_time_list_write[index_write], formatOne)
                    if self.checkBox_status.GetValue():
                        SheetOne.write(1 + index_write, 7, status_list_write[index_write], formatOne)
                    if self.checkBox_report.GetValue():
                        SheetOne.write(1 + index_write, 8, has_report_list_write[index_write], formatOne)
                        SheetOne.write(1 + index_write, 9, report_filename_list_write[index_write], formatOne)
                    if self.checkBox_summary.GetValue():
                        SheetOne.write(1 + index_write, 10, keywords_list_write[index_write], formatOne)
                    SheetOne.write(1 + index_write, 11, npi_comment_list_write[index_write], formatOne)

        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay(
            "ץ��{}��������Ѿ������д�롶����ϵͳץȡ��Ϣ-{}.xlsx���������в��ģ�����EXIT�˳�����".format(len(already_write_list), timestamp))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def close(self, event):
        self.Close()

    def newthread(self):
        Thread(target=self.run_all).start()

    def onbutton(self, event):
        self.button_go.Disable()
        self.newthread()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("��ɵ�{}ҳ".format(t))
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText(t)
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = PingShenFrame(None)
    frame.Show()
    app.MainLoop()
