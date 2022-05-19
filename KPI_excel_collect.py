#!/usr/bin/python
#author:zmy
#date:2022.4.11

import xlwt
import logging
import time
import os
import re
import pandas as pd
import numpy as np

class Kqtools():

 # 设置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        filename="kpi_collect.log")
    logger = logging.getLogger(__name__)

    def __init__(self,filepath,output_filename):
        """
        :param filepath:  待处理excel表格 xls、xlsx文件所在文件夹
        :param output_filename: 输出结果报存的excel文件
        """
        #self.output_filename = r"D:\python-excel-master\data\user1.xls"
        self.filepath = filepath
        self.output_filename = output_filename

    # 用来读取列识别列表中xls或xlsx文件，将其名字添加到list中返回
    def collect_xls(list_collect):
        typedata = []
        for each_element in list_collect:
            if isinstance(each_element, list):
                Kqtools.collect_xls(each_element)
            elif each_element.endswith("xls"):
                typedata.insert(0, each_element)
            elif each_element.endswith("xlsx"):
                typedata.insert(0, each_element)
        return typedata

    # new_dict_list.append(self.fix_value(colx[i]))

    # 读取文件夹中包含的所有xls和xlsx格式表格文件
    def read_xls(path):
        name = []
        for file in os.walk(path):
            # os.walk() 返回三个参数：路径，子文件夹，路径下的文件
            for each_list in file[2]:
                file_path = file[0] + "/" + each_list
                name.insert(0, file_path)
            all_xls = Kqtools.collect_xls(name)
        return all_xls

    def excel_style(self,blod =False,bc = False):
        # --------------------样式设置---------------------
        style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        font = xlwt.Font()  # 为样式创建字体
        font.name = 'Times New Roman'
        font.bold = blod  # 黑体
        borders = xlwt.Borders()  # Create Borders
        borders.left = xlwt.Borders.THIN
        style.borders = borders
        if bc==True:
            pattern = xlwt.Pattern()  # Create the Pattern
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
            pattern.pattern_fore_colour = 22
            style.pattern = pattern
        style.font = font
        style.alignment = al
        return  style

    def handle_execl(self):

        flielist = Kqtools.read_xls(self.filepath)
        print("FlieList: %s" % flielist)
        name_list = []
        content = []
        regexp = "\-|\."
        for file in flielist:
            df = pd.read_excel(file,sheet_name = 1 )
            name_extract = re.split(regexp,file)
            name_list.append(name_extract[1])
            content.append(name_extract[1])
            content.append(df.values[24, 3])
            content.append(df.values[25, 3])
            content.append(df.values[26, 3])
            content.append(df.values[27, 3])
            content.append(df.values[28, 3])
            content.append(df.values[4, 4])
            content.append(df.values[14, 4])
            content.append(df.values[21, 4])
            content.append(df.values[40, 4])
            content.append(df.values[52, 4])
            content.append(df.values[57, 4])
            content.append(df.values[64, 4])
            content.append(df.values[70, 4])
            content.append(df.values[74, 4])
            content.append(df.values[86, 4])
        print(name_list)
        logging.info("--获取文件姓名列表: %s ", name_list)
        print(content)
        logging.info("--Sheet数据提取全量数据（一维）: %s ", content)
        #写表
        outfile = xlwt.Workbook()
        xlsheet = outfile.add_sheet("季度KPI统计结果")
        table_header = ["姓名", "问题单总数", "致命","严重", "一般", "提示", "测试设计工作","文档编写工作","测试执行工作","对外测试工作","专项测试","自动化与工具开发","周边支持工作","流程制度遵守情况","角色","自评总分"]
        headerlen = len(table_header)
        name_list_len =len(name_list)

        content_duowei=np.resize(content, (name_list_len,headerlen))
        print(content_duowei)
        logging.info("--Sheet数据提取全量数据（二维）: %s ", content_duowei)

        for i in range(headerlen):
            xlsheet.col(i).width = 0x0d00 + i * 100
            xlsheet.write(0, i, table_header[i],self.excel_style(blod=True,bc=True))
            for j in range(name_list_len):
               xlsheet.write(j+1, i, content_duowei[j,i],self.excel_style())
        print("开始写入excel文件,记得关闭查看过的生成文件...")
        logging.info("开始写入excel文件,记得关闭查看过的生成文件...")
        outfile.save(self.output_filename)
        logging.info("--恭喜！KIP数据已提取完成，输出结果Excel文件见程序当前目录。--")

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    logging.info("请把需要采集的数据放到 D:/test 目录下，采集EXCEl文件请以”XXX-姓名.xls“命名，以便程序进行采集工作")
    logging.info("数据提取任务已开始，请耐心等待！")

def run_job(data_from_excel =True):
    now = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    kqtools = Kqtools(filepath=r"D:/test", output_filename="./{}-kpi-result.xls".format(now))
    kqtools.handle_execl()

if __name__ == '__main__':
    print_hi('数据提取任务已开始，请耐心等待！')
    run_job()
