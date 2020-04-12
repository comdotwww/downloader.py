from threading import Event, Thread
from time import sleep
from traceback import format_exc

import requests
import xlrd
import threading

openXlsx = xlrd.open_workbook('dd.xlsx', 'r')  # 打开.xlsx文件
sht = openXlsx.sheets()[0]  # 打开表格中第一个sheet
row1 = sht.row_values(0)


class downloader:
    # 构造函数
    def __init__(self, url, filename):
        # 设置url
        self.url = url
        # 设置线程数
        self.num = 10
        # 文件名从url最后取
        self.name = filename
        # 用head方式去访问资源
        r = requests.head(self.url)
        # 取出资源的字节数
        self.total = int(r.headers['Content-Length'])
        print('total is %s' % self.total)

    def get_range(self):
        ranges = []
        # 比如total是50,线程数是4个。offset就是12
        offset = int(self.total / self.num)
        for i in range(self.num):
            if i == self.num - 1:
                # 最后一个线程，不指定结束位置，取到最后
                ranges.append((i * offset, ''))
            else:
                # 每个线程取得区间
                ranges.append((i * offset, (i + 1) * offset))
        # range大概是[(0,12),(12,24),(25,36),(36,'')]
        return ranges

        return ranges

    def download(self, start, end):
        headers = {'Range': 'Bytes=%s-%s' % (start, end),
                   'Accept-Encoding': '*',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko'
                   }
        # 获取数据段
        res = requests.get(self.url, headers=headers)
        # seek到指定位置
        print('%s:%s download success' % (start, end))
        self.fd.seek(start)
        self.fd.write(res.content)

    def run(self):
        # 打开文件，文件对象存在self里
        self.fd = open(self.name, 'wb')
        thread_list = []

        n = 0
        for ran in self.get_range():
            start, end = ran
            print('thread %d start:%s,end:%s' % (n, start, end))
            n += 1
            # 开线程
            thread = threading.Thread(target=self.download, args=(start, end))
            thread.start()
            thread_list.append(thread)
        for i in thread_list:
            # 设置等待
            i.join()
        print('download %s load success' % self.name)
        self.fd.close()


if __name__ == '__main__':
    begin = 114  # 设置要下载的图片的范围，对应于 Excel 中的行数
    ends = 6114
    for i in range(begin, ends):
        url = sht.cell(i, 3).value  # 依次读取每行第四列的数据，也就是 URL
        f = requests.get(url)
        suffix = url[-13:]  # 根据链接地址获取文件后缀，后缀有.pdf 和 .pdzx 两种
        name = sht.cell(i, 0).value
        author = sht.cell(i, 1).value
        filename = name + '_' + author + '_' + suffix  # 构造完整文件名称
        progress = (i + 1 - begin) / (ends - begin) * 100  # 计算下载进度
        print("下载进度：", progress, "%")  # 显示下载进度

        # 新建实例
        down = downloader(url, filename)
        # 执行run方法
        down.run()
