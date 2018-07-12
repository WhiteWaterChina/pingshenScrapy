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


ver = "20180711"


def sumtimesplit(strtimelist):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimelist:
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


def get_status(status):
    switcher = {
        "0": "����".decode('gbk'),
        "1": "�ύ".decode('gbk'),
        "shenhepeizhi-sq": "��ǰ�������".decode('gbk'),
        "shenhe-chanpinjingli": "��Ʒ�������".decode('gbk'),
        "querenxuanpei-ddy": "����Աȷ��ѡ��".decode('gbk'),
        "shenhepingshen-yf": "�з��ӿ����������".decode('gbk'),
        "shenhepingshen-csjk": "���Խӿ����������".decode('gbk'),
        "ceshi-cs": "������Ա����".decode('gbk'),
        "shenheceshibaogao-yf": "�з��ӿ���˲��Ա���".decode('gbk'),
        "shenheceshibaogao-xmjl": "��Ŀ������˲��Ա���".decode('gbk'),
        "shenheceshibaogao-csfzr": "���Ը�������˲��Ա���".decode('gbk'),
        "shenheceshibaogao-csjk": "���Խӿ�����˲��Ա���".decode('gbk'),
        "xfzl-gc": "������Աȷ���Ƿ��·�ָ��".decode('gbk'),
        "100": "�ر�".decode('gbk'),
        "101": "�쳣�ر�".decode('gbk'),
        "product-verification": "������֤".decode('gbk'),
        "vm-audit": "VM���".decode('gbk'),
        "exec-test": "ִ�в���".decode('gbk'),
        "os-comp-test": "OS�����Բ���".decode('gbk'),
        "audit-submit": "�޸Ĵ��ύ".decode('gbk'),
        "npi-audit": "NPI����".decode('gbk')
    }
    return switcher.get(status, status)


def get_detail(link, login_session):
    headers_data_all = {
        'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.8",
        'Connection': "keep-alive",
        'Host': "218.57.146.175",
        'Referer': "http://218.57.146.175/techAudit/welcome.htm",
        'Upgrade-Insecure-Requests': "1",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"
    }
    get_page = login_session.get(link, headers=headers_data_all)
    data_page = get_page.text
    print("Get link:%s with return code %s" % (link, get_page.status_code))
    data_filter = BeautifulSoup(data_page, "html.parser")
    # ��ȡ������Ϣ
    attachment_temp = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr")
    if len(attachment_temp) <= 1:
        has_report = "�ޱ���".decode('gbk')
        report_filename = "None"
    else:
        has_report = "�б���".decode('gbk')
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
            if item_td == "����".decode('gbk'):
                test_time = item_tr.select("td:nth-of-type(3)")[0].get_text().strip()
                timeTestTemp.append(test_time)
        if len(timeTestTemp) == 0:
            timeTestData = "None"
        else:
            timeTestData = sumtimesplit(timeTestTemp)
        test_time = timeTestData
    except IndexError:
        test_time = "None"
    return link, has_report, report_filename, keywords, test_time


class PingShenFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"����ϵͳ��Ϣץȡ����-Version:%s" % ver, pos=wx.DefaultPosition,
                          size=wx.Size(504, 460), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

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
        self.text_title1.SetFont(wx.Font(12, 70, 90, 90, False, wx.EmptyString))
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
                                         u"��������������Ҫץȡ����Ϣ����ֹ����(�������/��/����Ϣ����ʽΪ20170101.\n��λ�����º���һ��Ҫ��0����",
                                         wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTRE)
        self.text_title2.Wrap(-1)
        self.text_title2.SetFont(wx.Font(9, 70, 90, 90, False, wx.EmptyString))
        self.text_title2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_title2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

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

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.text_3 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��������ѡ����Ҫ������ļ�����ʾ����Ŀ", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        self.text_3.Wrap(-1)
        self.text_3.SetFont(wx.Font(12, 70, 90, 90, False, wx.EmptyString))
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

        self.checkBox_closeTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"������ʱ��", wx.DefaultPosition, wx.DefaultSize,
                                              0)
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
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

        self._thread = Thread(target=self.run, args=())
        self._thread.daemon = True

    def run(self):
        self.button_go.Disable()
        self.updatedisplay("��ʼץȡ".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue())
        start_time = self.input_startdate.GetValue()
        end_time = self.input_enddate.GetValue()
        # ��¼
        headers_base = {
            'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.8",
            'Cache-Control': "max-age=0",
            'Connection': "keep-alive",
            'Content-Length': "125",
            'Content-Type': "application/x-www-form-urlencoded",
            'Host': "218.57.146.175",
            'Origin': "http://218.57.146.175",
            'Referer': "http://218.57.146.175/inspurSSO/login",
            'Upgrade-Insecure-Requests': "1",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        }
        url_login = "http://218.57.146.175/inspurSSO/login"
        login_session = requests.session()
        headers_login = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "max-age=0",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded",
            'host': "218.57.146.175",
            'upgrade-insecure-requests': "1",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        }
        # ��ȡ��¼ҳ���lt/execution/eventid��Ϣ
        response_login_test = login_session.get(url_login, headers=headers_login).text
        data_soup_tobe_filter = BeautifulSoup(response_login_test, "html.parser")
        lt = data_soup_tobe_filter.find('input', {'name': 'lt'})['value']
        execution = data_soup_tobe_filter.find('input', {'name': 'execution'})['value']
        eventid = data_soup_tobe_filter.find('input', {'name': '_eventId'})['value']
        payload_login = {
            'username': "%s" % username,
            'password': "%s" % password,
            'lt': "%s" % lt,
            'execution': "%s" % execution,
            '_eventId': "%s" % eventid
        }
        # ʹ�����ϻ�ȡ����Ϣpost��¼
        log_in = login_session.post(url_login, headers=headers_base, data=payload_login)
        # ��ʼ��ȡ����
        headers_data = {
            'Accept': "application/json, text/javascript, */*; q=0.01",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Length': "34",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "218.57.146.175",
            'Origin': "http://218.57.146.175",
            'Referer': "http://218.57.146.175/techAudit/auditList/queryAuditList.htm",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest"
        }
        # ����1��ȡ�����������
        payload_data_test = "rows=1"
        url_data = "http://218.57.146.175/techAudit/AuditListController/getQueryAuditList.htm?"
        response_data_test = login_session.post(url_data, headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        # Ȼ��ʹ���������ֱ��1ҳ��ʾ
        payload_data = "page=1&rows={total_row}".format(total_row=total_item)
        response_data_get = login_session.post(url_data, headers=headers_data, data=payload_data)
        response_data = response_data_get.text

        # �Ȼ�ȡautidno/id/creattime/updatetime/projectname/productname��ԭʼ����
        auditno_list_temp = re.findall(r'"auditNo":"(\w+)",', response_data)
        id_list_temp = re.findall(r'"id":(\d+),', response_data)
        create_date_list_temp = re.findall(r'"submitTime":"(\d+-\d+-\d+)', response_data)  # 2016-09-23
        close_date_list_temp_1 = re.findall(r'"closeTime":(.*?),', response_data) # "closeTime":"2015-04-10 08:42:21", or "closeTime":null
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

        # ����һ�飬ȥ��ʱ�䲻����Ҫ���
        autidno_list = []
        id_list = []
        create_date_list = []
        close_date_list = []
        projectname_list = []
        productname_list = []
        status_list = []
        url_list = []
        total_time_list = []
        # report_list = []
        # report_filename_list = []
        # keywords_list = []
        # test_time_list = []
        for index_creattime, item_creattime in enumerate(create_date_list_temp):
            # print(item_creattime)
            now_date = time.mktime(time.strptime(item_creattime, '%Y-%m-%d'))
            # print(now_date)
            # print now_date
            if float(start_date) <= float(now_date) <= float(end_date):
                # ����ʱ�䣬������
                create_date_list.append(item_creattime)
                # �����ţ���һ��
                autidno_list.append(auditno_list_temp[index_creattime])
                url = "http://218.57.146.175/techAudit/details/viewBill.htm?id=" + id_list_temp[index_creattime]
                # ÿ�������url�е�Ψһ��ţ���Ϊ����ǰ���ȡ�������ݺͺ���ÿ����������ݵ�Ŧ��
                id_list.append(id_list_temp[index_creattime])
                # �ر�ʱ�䣬������
                close_date_list.append(close_date_list_temp[index_creattime])
                # ��Ŀ���ƣ��ڶ���
                projectname_list.append(projectname_list_temp[index_creattime])
                # ��Ʒ���ƣ�������
                productname_list.append(productname_list_temp[index_creattime])
                # ��ǰ״̬���ڰ���
                status_list.append(get_status(status_list_temp[index_creattime]))
                url_list.append(url)
                # �ܻ���ʱ�䣬������
                total_time_list.append(total_time_list_temp[index_creattime])
        self.updatedisplay("��������%s������ʱ��Ҫ���������Ⱥ�ץȡ����~~~~".decode('gbk') % len(autidno_list))
        # ��ҳ��ȡ
        dict_data_detail = {}
        for item_1 in id_list:
            dict_data_detail["%s" % item_1] = []
        temp_detail = []
        pool_detail = Pool()
        for index, item_2 in enumerate(url_list):
            temp_detail.append(pool_detail.apply_async(get_detail,args=(item_2, login_session)))
            # self.updatedisplay("�Ѿ���ʼץȡ%s/%s".decode('gbk') %(str(url_list.index(item_2) + 1), str(len(url_list))))
        pool_detail.close()
        pool_detail.join()
        #return link, has_report, report_filename, keywords, test_time
        for item_detail in temp_detail:
            data_detail_temp = item_detail.get()
            if data_detail_temp is not None:
                # if data_detail_temp[0] != "None":
                num = data_detail_temp[0].split("=")[-1]
                index_to_log = id_list.index(num)
                # ������ audit_number
                dict_data_detail["%s" % num].append(autidno_list[index_to_log])
                # �������� project name
                dict_data_detail["%s" % num].append(projectname_list[index_to_log])
                # ��Ʒ���� product name
                dict_data_detail["%s" % num].append(productname_list[index_to_log])
                # �ύʱ�� create time
                dict_data_detail["%s" % num].append(create_date_list[index_to_log])
                # �ر�ʱ�� close time
                dict_data_detail["%s" % num].append(close_date_list[index_to_log])
                # �ܴ���ʱ�� total time
                dict_data_detail["%s" % num].append(total_time_list[index_to_log])
                # ���Ի���ʱ�� test time
                dict_data_detail["%s" % num].append(data_detail_temp[4])
                # ״̬ status
                dict_data_detail["%s" % num].append(status_list[index_to_log])
                # �Ƿ��б��� has_report
                dict_data_detail["%s" % num].append(data_detail_temp[1])
                # �������ƣ� report filename
                dict_data_detail["%s" % num].append(data_detail_temp[2])
                # ����Ҫ�� keywords
                dict_data_detail["%s" % num].append(data_detail_temp[3])

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

        # ���������ݴ�������������ٷ�����ϵ
        TitleItem = ['������'.decode('gbk'), '��������'.decode('gbk'), '��Ŀ����'.decode('gbk'), '�ύʱ��'.decode('gbk'),
                     '������ʱ��'.decode('gbk'), '����ʱ��'.decode('gbk'), '���Ի���ʱ��'.decode('gbk'), '״̬'.decode('gbk'),
                     '�Ƿ��б��渽��'.decode('gbk'), '��������'.decode('gbk'), '����Ҫ��'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("����ϵͳץȡ��Ϣ-%s.xlsx".decode('gbk') % timestamp)
        SheetOne = WorkBook.add_worksheet('����ϵͳץȡ��Ϣ'.decode('gbk'))
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)

        SheetOne.set_column('A:J', 14)
        already_write_list = []
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], formatOne)
        for index_write, item_write in enumerate(autidno_list_write):
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
                        SheetOne.write(1 + index_write, 4, "���������".decode('gbk'), formatOne)
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
        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay(
            "ץ��%s��������Ѿ������д�롶����ϵͳץȡ��Ϣ-%s.xlsx���������в��ģ�����EXIT�˳�����".decode('gbk') % (len(autidno_list), timestamp))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def close(self, event):
        self.Close()

    def onbutton(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("��ɵ�%sҳ".decode('gbk') % t)
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText("%s".decode('gbk') % t)
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = PingShenFrame(None)
    frame.Show()
    app.MainLoop()
