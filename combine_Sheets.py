# -*- coding: UTF-8 -*-
#__author__ = 'Luke'

import os
import re
import sys
import xlrd
import xlsxwriter
import time



def banner():
    banner = '''
combine_sheets
'''
    info = '\n作者：Luke'
    print (info)
    print (banner)







def creat_and_write(array):         #新建sheet并将array写入
    worksheet=workbook.add_worksheet('合并sheets')
    for i,p in enumerate(array):
        for j,q in enumerate(p):
            worksheet.write(i,j,q)




def process(a,b):
    bili = float(a/b)
    num_arrow = int(50 * bili)
    num_line = 50-num_arrow
    process_bar = '[' + '>' * num_arrow + '-' * num_line + ']'+\
                  '{:.2%}'.format(bili) + '\r' #输出内容，'\r'表示不换行回到最左边
    print (process_bar,end = '')       #末尾不加\n输出字符串
    # sys.stdout.write(process_bar)       #末尾不加\n输出字符串
    # sys.stdout.flush()      #刷新输出,Windows运行可不加



if __name__ == '__main__':
    banner()
    uijm = time.strftime('%m_%d~%H-%M-%S',time.localtime(time.time()))
    root = input("\n请拖入Excel文件:\n\n")
    if os.path.isfile(root) is False:
        print ('您输入的文件不存在或名称中存在空格！');sys.exit(0)
    file_sheets = xlrd.open_workbook(root).sheet_names()
    sheet_num = len(file_sheets)

    title_row = input("\n请输入Excel中相同的标题行数(默认为1):\n\n")
    if title_row == '':title_row = 1
    else:title_row = int(title_row)
    # title_row = 1;print ('\nExcel中相同的标题行数 默认设置为了1')
    
    print('\n正在写入数据到Excel中。。。');print (root)
    workbook = xlsxwriter.Workbook(os.path.splitext(root)[0] + '-合并sheet-%s.xlsx'%uijm)
    hebk = []
    xls = xlrd.open_workbook(root)
    for index,sheet in enumerate(file_sheets):
        xls_sheet = xls.sheet_by_name(sheet)
        if sheet == file_sheets[0]:#只获取第一个的title_row
            for i in range(title_row):
                hebk.append(xls_sheet.row_values(i))
        for j in range(title_row,xls_sheet.nrows):
            hebk.append(xls_sheet.row_values(j))
        process(index+1,len(file_sheets))
    creat_and_write(hebk)
    workbook.close()
    print ('\n合并完毕，文件保存在Excel文件的目录下，命名为 《合并sheet %s.xlsx》'%uijm)
    # print ('\n5秒后关闭~\n')
    # time.sleep(5)
    os.system('pause')

