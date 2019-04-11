# coding:utf-8
import requests
import re
import importlib, sys

importlib.reload(sys)
import pandas as pd
import datetime
from openpyxl import load_workbook

'''爬虫表格示例'''
# 2019年节日放假安排查询
# 节日	放假时间	调休上班日期	放假天数
# 元旦	12月30日~1月1日	12月29日（周六）上班	3天
# 春节	2月4日(除夕）~2月10日	2月2日（周六）、2月3日（周日）上班	7天
# 清明节	4月5日~4月7日	与周末连休	3天
# 劳动节	5月1日~5月4日	4月28日（星期日）、5月5日（星期日）上班	4天
# 端午节	6月7日~6月9日	与周末连休	3天
# 国庆节	10月1日~10月7日	9月29日（周日）、10月12日（周六）上班	7天
# 中秋节	9月13日~9月15日	与周末连休	3天

'''适用条件'''


# 1. 放假时间列必须是以xx月xx日开始，且以~分隔的连续日期
# 2. 调休上班日期列中，包含“休”“无”那行没有上班日期安排

class spider(object):
    def __init__(self):
        print(u'开始爬取...')

    # 获取网页源码，返回页面所有信息
    def getSource(self, url):
        html = requests.get(url)
        return html.text

    # 改变url实现抓取所有年份的假日，得所有url列表
    def changePage(self, url, now_page, total_page):
        page_group = []
        for i in range(now_page, total_page + 1):
            link = url[:26] + str(i) + url[30:]
            page_group.append(link)
        return page_group

    # 把页面上假日的table截取出来，返回字符串
    def geteveyTable(self, source):
        everytable = re.findall('(<table style=.*?</table>)', source, re.S)
        return everytable

    # 取出放假日期范围，返回日期列表
    def getHoliday(self, year, string):
        string = "".join(re.findall('\d.*', string, re.S))
        if string.find('年') >= 0:
            string = string[5:]
        if string.find('~') == -1:
            string = string + '~' + string
        str1 = string.split('~')
        m1 = "".join(re.findall('\d.*(?=月\d)', str1[0], re.S))
        d1 = "".join(re.findall('(?<=月).*\d', str1[0], re.S))
        m2 = "".join(re.findall('\d.*(?=月\d)', str1[1], re.S))
        d2 = "".join(re.findall('(?<=月).*\d', str1[1], re.S))
        # 规范化日期格式为yyyy mm dd
        if int(m1) < 10:
            m1 = '0' + m1

        if int(d1) < 10:
            d1 = '0' + d1

        if int(m2) < 10:
            m2 = '0' + m2

        if int(d2) < 10:
            d2 = '0' + d2

        if int(m1) == 12:
            y1 = str(int(year) - 1)  # 如果元旦前是12月则月份要取少一年
        else:
            y1 = year
        # 根据日期范围获取所有连续日期
        date_list = []
        begin_date = datetime.datetime.strptime(y1 + m1 + d1, "%Y%m%d")
        end_date = datetime.datetime.strptime(year + m2 + d2, "%Y%m%d")
        while begin_date <= end_date:
            date_str = begin_date.strftime("%Y%m%d")
            date_list.append(date_str)
            begin_date += datetime.timedelta(days=1)

        return date_list

    # 取出调休上班日期，返回日期列表
    def getWorkday(self, year, string):
        string = "".join(re.findall('\d.*', string, re.S))
        if string.find('年') >= 0:
            string = string[5:]
        list1 = re.split('[，、]', string)
        list2 = []
        for a in list1:
            m1 = "".join(re.findall('\d.*(?=月\d)', a, re.S))
            d1 = "".join(re.findall('(?<=月).*\d', a, re.S))
            if int(m1) == 12:
                y1 = str(int(year) - 1)
            else:
                y1 = year
            if int(m1) < 10:
                m1 = '0' + m1
            if int(d1) < 10:
                d1 = '0' + d1

            list2.append(y1 + m1 + d1)

        return list2

    # 分割源码获得每年所有假日表：日期、节日名称、是否放假
    def getInfo(self, eachyear):
        info = {}
        ho = re.findall('jiad/">(.*?)</a>', eachyear, re.S)  # 节日名称列表
        year = re.search('<a href="/(.*?)_', eachyear, re.S).group(1)  # 年份字符串
        a = re.findall('<td>(.*?)</td>', eachyear, re.S)  # 除了节日名称列的其他数据列表，包括放假时间、调休上班日期、放假天数
        j = len(a)
        # 得出放假日期
        i = 0
        k = 0
        df = pd.DataFrame()
        while (i < j):
            if i > j:
                continue
            holidayname = ''.join(ho[k])
            holiday_list = list(map(int, self.getHoliday(year, a[i])))
            df1 = pd.DataFrame(holiday_list, columns=['period_id'])
            df1['holiday_name'], df1['is_holiday'] = [holidayname, 1]
            df = pd.concat([df, df1])
            i += 3
            k += 1
        # 得出调休上班日期
        m = 1
        n = 0
        df2 = pd.DataFrame()
        while (m < j):
            if m > j:
                continue
            holidayname = ''.join(ho[n])
            data = a[m]
            checkdata = a[m].find('休')  # 是否包含休
            checkdata2 = a[m].find('无')  # 是否包含无
            m += 3
            n += 1
            if checkdata >= 0 or checkdata2 >= 0:  # 如果该节日没有特殊调休安排则跳过循环
                continue
            work_list = list(map(int, self.getWorkday(year, data)))
            df_temp = pd.DataFrame(work_list, columns=['period_id'])
            df_temp['holiday_name'], df_temp['is_holiday'] = [holidayname, 0]
            df2 = pd.concat([df2, df_temp])
        return pd.concat([df, df2])

    # 将结果写入excel
    def write2excel(self, excelWriter, year, data):
        excel_header = list(data.columns.values)  # excel的标题
        data.to_excel(excelWriter, sheet_name=year, header=excel_header, index=False)
        writer.save()
        writer.close()

    # 将结果追加到excel
    def add2excel(self, excelWriter, year, data):
        book = load_workbook(excelWriter.path)
        excelWriter.book = book
        excel_header = list(data.columns.values)  # excel的标题
        data.to_excel(excelWriter, sheet_name=year, header=excel_header, index=False)
        excelWriter.save()
        excelWriter.close()


if __name__ == '__main__':
    # 初始化参数
    url = 'https://fangjia.51240.com/2018__fangjia/'  # 抓取的页面
    start_page = 2012  # 抓取开始年份
    end_page = 2019  # 抓取结束年份
    filepath = 'data/节假日安排爬虫数据.xlsx'  # 写入excel文件路径

    print(u'初始化类..')
    myspider = spider()
    all_links = myspider.changePage(url, start_page, end_page)
    df = pd.DataFrame()
    for link in all_links:
        print(u'正在处理页面..' + link)
        html = myspider.getSource(link)
        everyyear = myspider.geteveyTable(html)
        for each in everyyear:
            df = pd.concat([df, myspider.getInfo(each)])
    df.index = range(len(df))
    df.sort_values(by='period_id', inplace=True)

    # 创建ExcelWriter对象
    writer = pd.ExcelWriter(filepath, engine='openpyxl')

    # sheet名
    if start_page == end_page:
        sheetname = str(end_page)
    else:
        sheetname = str(start_page) + '-' + str(end_page)

    # 生成excel
    myspider.write2excel(writer, sheetname, df)  # 写入excel，会覆盖掉所有已存在的sheet

    # excel中增加sheet
    # myspider.add2excel(writer, sheetname, df)  # 写入excel，并新增sheet

    # 读取写入结果
    # data = pd.read_excel('data/节假日安排爬虫数据.xlsx', sheet_name=str(end_page))
    # print(data)
    print('写入Excel完成！')
