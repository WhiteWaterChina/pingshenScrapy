#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import time
import datetime
import sys
from multiprocessing import Pool
import multiprocessing
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


def get_detail(link, login_session_sub):
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
    # headers_base = {
    #     'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    #     'Accept-Encoding': "gzip, deflate",
    #     'Accept-Language': "zh-CN,zh;q=0.8",
    #     'Cache-Control': "max-age=0",
    #     'Connection': "keep-alive",
    #     'Content-Length': "125",
    #     'Content-Type': "application/x-www-form-urlencoded",
    #     'Host': "218.57.146.175",
    #     'Origin': "http://218.57.146.175",
    #     'Referer': "http://218.57.146.175/inspurSSO/login",
    #     'Upgrade-Insecure-Requests': "1",
    #     'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
    # }
    # url_login = "http://218.57.146.175/inspurSSO/login"
    # login_session = requests.session()
    # # headers_login = {
    # #     'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    # #     'accept-encoding': "gzip, deflate",
    # #     'accept-language': "zh-CN,zh;q=0.8",
    # #     'cache-control': "max-age=0",
    # #     'connection': "keep-alive",
    # #     'content-type': "application/x-www-form-urlencoded",
    # #     'host': "218.57.146.175",
    # #     'upgrade-insecure-requests': "1",
    # #     'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
    # # }
    # # ��ȡ��¼ҳ���lt/execution/eventid��Ϣ
    # # response_login_test = login_session.get(url_login, headers=headers_login).text
    # # data_soup_tobe_filter = BeautifulSoup(response_login_test, "html.parser")
    # # lt = data_soup_tobe_filter.find('input', {'name': 'lt'})['value']
    # # execution = data_soup_tobe_filter.find('input', {'name': 'execution'})['value']
    # # eventid = data_soup_tobe_filter.find('input', {'name': '_eventId'})['value']
    # payload_login = {
    #     'username': "%s" % username,
    #     'password': "%s" % password,
    #     'lt': "%s" % lt,
    #     'execution': "%s" % execution,
    #     '_eventId': "%s" % eventid
    # }
    # ʹ�����ϻ�ȡ����Ϣpost��¼
    # log_in = login_session.post(url_login, headers=headers_base, data=payload_login)
    get_page = login_session_sub.get(link, headers=headers_data_all)
    data_page = get_page.text
    # print("Get link:%s with return code %s" % (link, get_page.status_code))
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


# main
if __name__ == '__main__':
    multiprocessing.freeze_support()
    input_length = len(sys.argv)
    if input_length != 5:
        print("Input length is incorrect!")
        print("Usage:%s username password start_date(like:20180101) end_date(like:20180808)" % sys.argv[0])
        sys.exit(255)

    print("��ʼץȡ".decode('gbk'))
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
    username = sys.argv[1]
    password = base64.b64encode(sys.argv[2])
    start_time = sys.argv[3]
    end_time = sys.argv[4]
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
    print("��������%s������ʱ��Ҫ���������Ⱥ�ץȡ����~~~~".decode('gbk') % len(autidno_list))
    # ��ҳ��ȡ
    dict_data_detail = {}
    for item_1 in id_list:
        dict_data_detail["%s" % item_1] = []
    temp_detail = []
    pool_detail = Pool()
    for index, item_2 in enumerate(url_list):
        temp_detail.append(pool_detail.apply_async(get_detail, args=(item_2, login_session)))
        # self.updatedisplay("�Ѿ���ʼץȡ%s/%s".decode('gbk') %(str(url_list.index(item_2) + 1), str(len(url_list))))
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
            SheetOne.write(1 + index_write, 0, item_write, formatOne)
            SheetOne.write(1 + index_write, 1, projectname_list_write[index_write], formatOne)
            SheetOne.write(1 + index_write, 2, productname_list_write[index_write], formatOne)
            SheetOne.write_datetime(1 + index_write, 3,
                                    datetime.datetime.strptime(create_date_list_write[index_write], '%Y-%m-%d'),
                                    WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            if re.search(r'\d+', close_date_list_write[index_write]) is None:
                SheetOne.write(1 + index_write, 4, "���������".decode('gbk'), formatOne)
            else:
                SheetOne.write(1 + index_write, 4, datetime.datetime.strptime(close_date_list_write[index_write], '%Y-%m-%d'), WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            SheetOne.write(1 + index_write, 5, total_time_list_write[index_write], formatOne)
            SheetOne.write(1 + index_write, 6, test_time_list_write[index_write], formatOne)
            SheetOne.write(1 + index_write, 7, status_list_write[index_write], formatOne)

            SheetOne.write(1 + index_write, 8, has_report_list_write[index_write], formatOne)
            SheetOne.write(1 + index_write, 9, report_filename_list_write[index_write], formatOne)

            SheetOne.write(1 + index_write, 10, keywords_list_write[index_write], formatOne)
    WorkBook.close()
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
    print("ץ��%s��������Ѿ������д�롶����ϵͳץȡ��Ϣ-%s.xlsx���������в��ģ�����EXIT�˳�����".decode('gbk') % (len(autidno_list), timestamp))
    time.sleep(1)
    print("Finished")
