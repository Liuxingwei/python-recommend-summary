#!/usr/bin/env python3
#encoding:utf-8
"""
自辽宁省政府集中采购网采集中标企业信息，并输出到excel中
:author: liuxingwei
"""
import selenium.webdriver
import time
import re


def main():
    """
    主控函数
    :return:
    """
    pages = 100 #采集页数
    fileName = './result/' + time.strftime('%Y%m%d', time.localtime(time.time())) + ".csv" #输出文件名，以日期作为文件名

    data = collection(pages)
    generate_table_file(data, fileName)


def collection(pages):
    """
    采集数据
    :param pages: 采集页数
    :param fileName: 输出文件名（带路径）
    :return:
    """
    data = []
    links = []
    companys = []
    url = 'http://www.lnzc.gov.cn/SitePages/AfficheListAll2.aspx'  # 辽宁省政府集中采购网网址
    opts = selenium.webdriver.ChromeOptions()
    # opts.headless = True
    driver = selenium.webdriver.Chrome(options=opts)
    driver.get(url)
    for i in range(0, pages):
        infoList = driver.find_elements_by_css_selector('.col-xs-8 .infoArea:last-child ul a')
        for info in infoList:
            links.append({'title': info.get_attribute('title'), 'href': info.get_attribute('href')})
        infoNextPage = driver.find_elements_by_css_selector('.ms-paging a:last-child img')
        if infoNextPage:
            selenium.webdriver.ActionChains(driver).move_to_element(infoNextPage[0]).click(infoNextPage[0]).perform()
        else:
            break
    for link in links:
        typePattern = re.compile('(.*废标公告|废标公告|中标公告|成交公告|结果公告|竞争性磋商)')
        typeMatch = typePattern.search(link['title'])
        if not typeMatch:
            continue
        type = typeMatch.group(1)

        voidPattern = re.compile('废标')

        if type == '中标公告':
            cType = '中标'
        elif voidPattern.search(type):
            continue
        else:
            cType = '成交'
        driver.get(link['href'])
        content = driver.page_source

        itemPattern = re.compile('项目名称：([^<]*)')
        itemMatch = itemPattern.search(content)
        itemName = itemMatch and itemMatch.group(1)
        trs = driver.find_elements_by_css_selector('.ms-rteTable-default tr')
        if not trs:
            continue
        if len(driver.find_elements_by_css_selector('.ms-rteTable-default')) > 1:
            continue
        trNumber = len(driver.find_elements_by_css_selector('.ms-rteTable-default tr'))

        addressPattern = re.compile('(、.{2}供应商地址.*?)\d、')
        addressMatch = addressPattern.search(content)
        addresses = []
        if addressMatch:
            addressString = re.sub('&#160;', '', addressMatch.group(1))
            addressString = re.sub('<p.*?>', '', addressString)
            addressString = re.sub('</p>', '<br>', addressString)
            addressString = re.sub('<div.*?>', '', addressString)
            addressString = re.sub('</div>{1,}', '<br>', addressString)
            addressString = re.sub('<.?span.*?>', '', addressString)
            addressString = re.sub('<br.*?>', '<br>', addressString)
            addressString = re.sub('(<br>){1,}', ';', addressString)
            addressString = re.sub('\s*', '', addressString)
            addressString = re.sub('；', ';', addressString)
            addressString = addressString + ';'
            addressString = re.sub('、.{2}供应商地址：;*', '', addressString)
            addressCutPattern = re.compile(';')
            addresses = addressCutPattern.split(addressString)
        m = 0 #废标数量
        for j in range(2, trNumber + 1):
            if len(driver.find_elements_by_css_selector('.ms-rteTable-default tr:nth-child(' + str(j) + ') td')) < 2:
                m += 1
                continue
            td = driver.find_element_by_css_selector('.ms-rteTable-default tr:nth-child(' + str(j) + ') td:nth-child(2)')
            companyName = td.text
            if (companyName == '废标'):
                m += 1
                continue
            if {'name': companyName} in companys:
                continue
            res = {}
            companys.append({'name': companyName})
            res.update({'link': link['href'], 'title': link['title'], 'itemName': itemName, 'companyName': companyName})
            if len(addresses) > j - 2 - m:
                res.update({'address': addresses[j - 2 - m]})
            else:
                res.update({'address': ''})
            data.append(res)

    driver.close()
    return data


def generate_table_file(data, fileName):
    """
    生成表格文件
    :param data:
    :param fileName:
    :return:
    """
    file = open(file=fileName, mode='w', encoding='utf8')
    file.write('公司名称,公司地址\n')
    for row in data:
        file.write(row['companyName'] + ',' + row['address'] + '\n')
    file.close()
    return


main()
