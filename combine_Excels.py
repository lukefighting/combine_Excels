# -*- coding: UTF-8 -*-
#__author__ = 'Luke'

import os
import sys
import xlrd
import xlsxwriter
import time



def banner():
    banner = '''
                      _     _              ______              _ 
                     | |   (_)            |  ____|            | |
   ___ ___  _ __ ___ | |__  _ _ __   ___  | |__  __  _____ ___| |
  / __/ _ \| '_ ` _ \| '_ \| | '_ \ / _ \ |  __| \ \/ / __/ _ \ |
 | (_| (_) | | | | | | |_) | | | | |  __/ | |____ >  < (_|  __/ |
  \___\___/|_| |_| |_|_.__/|_|_| |_|\___| |______/_/\_\___\___|_|
                                      ______                     
                                     |______|                    '''
    info = '\n作者：Luke'
    print (info)
    print (banner)



def listdirfile(dir):
    file_array = []
    if os.path.exists(dir) is False:
        return False
    elif os.listdir(dir) == []:
        return False
    for filename in os.listdir(dir):
        file_array.append(root+'\\'+filename)
    return file_array

def hebk(list,sheet_num):   #输入的Excel列表，和每个Excel的sheet页
    hebk = []
    for index,excel in enumerate(list):
        xls = xlrd.open_workbook(excel)
        try:
            xls_sheet = xls.sheet_by_name(file_sheets[sheet_num])
        except:
            print ('\n\n遇到错误！！请核实<{}>这个Excel！！\n\n需保证所有Excel具有相同的sheet页数，且名字必须相同!!\n'.format(excel))
            os.system('pause')
            sys.exit()
        if excel == list[0]:
            for i in range(title_row):
                hebk.append(xls_sheet.row_values(i))
        for j in range(title_row,xls_sheet.nrows):
            hebk.append(xls_sheet.row_values(j))
        process(index+1,len(list))
    return hebk



def creat_and_write(array,sheet_num):         #新建sheet并将array写入
    worksheet=workbook.add_worksheet(file_sheets[sheet_num])
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
    uijm = time.strftime('%m_%d %H-%M',time.localtime(time.time()))
    root = input("\n请拖入包含多个Excel的文件夹:\n\n")
    file_list = listdirfile(root)       #文件路径不存在或文件夹是空的返回false
    if file_list is False:
        print ('您输入的文件路径不存在或文件夹是空的！')
        os.system('pause')
        sys.exit(0)
    file_sheets = xlrd.open_workbook(listdirfile(root)[0]).sheet_names()
    sheet_num = len(file_sheets)        #第一个Excel有几个sheet页，且默认其他的Excel也具有相同的sheet页
    # title_row = input("\n请输入Excel中相同的标题行数(默认为0):\n\n")
    # if title_row == '':
        # title_row = 0
    # else:
        # title_row = int(title_row)
    title_row = 1;print ('\nExcel中相同的标题行数 默认设置为了1')
    print('\n正在写入数据到Excel中。。。')
    workbook = xlsxwriter.Workbook('{}合并{}.xlsx'.format(root,uijm))
    for i in range(sheet_num):
        hebk_array = hebk(file_list,i)
        creat_and_write(hebk_array,i)
    workbook.close()
    print ('\n合并完毕，文件保存在输入文件夹所在的文件夹内，命名为 《{}合并{}.xlsx》'.format(os.path.basename(root),uijm))

    os.system('pause')



