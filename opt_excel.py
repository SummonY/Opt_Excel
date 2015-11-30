#!/usr/bin/python
# -*- coding:utf-8 -*-

import xlrd
import xlwt
#import xlutils
import os
import sys
from optparse import OptionParser


projectfile = '2015ProjectProcessChart.xls'


def get_all_dir(dir = ''):
    '''
    get all directory.
    '''
    dirlist = []
    for dirc in os.walk(dir):
        if dirc[0] != dir and os.path.basename(dirc) != '.ropeproject':
            dirlist.append(dirc)
    return dirlist

def get_codeandname_from_file(filename):
    '''
    get citycode and cityname from filename
    '''
    citycode = filename[0:6]
    cityname = filename[6:len(filename) - 4]
    return citycode, cityname


def get_all_files_from_dir(dirlist = []):
    '''
    get all files from directory.
    '''
    for li in dirlist:
        filelist = []
        files = os.listdir(li)
        for filename in files:
            if filename.find(".xls") != -1:
                citycode, cityname = get_codeandname_from_file(filename)
                filelist.append([citycode, cityname])


def open_excel(file='file.xls'):
    '''
    Open from Excel.
    '''
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)


def write_excel(file='file.xls'):
    '''
    Write to excel.
    '''
    try:
        data = xlwt.Workbook(file)
        return data
    except Exception, e:
        print str(e)

def excel_table_byindex(file='file.xls', colnameindex=0, by_index=0):
    '''
    get data by index.
    '''
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows     # 行数
    colnames = table.row_values(colnameindex)   # 某行数据
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list

def write_project_tosheet3(file='file.xls', readindex = 1, by_name=u'Sheet1'):
    '''
    get project list.
    '''
    datafile = open_excel(file)
    table = datafile.sheet_by_name(by_name)
    nrows = table.nrows
    list = []

    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        print row[readindex]

    return list

def find_project(readfile='file.xls', cityname=u''):
    data = open_excel(readfile)
    table = data.sheets()[0]
    nrows = table.nrows
    list = []

    linenum = 8
    for rownum in range(8, nrows):
        row = table.row_values(rownum)
        if row[1].find(cityname) == -1:
            linenum += 1
            continue
        else:
            linenum += 1
            break

    print nrows, linenum

    while linenum < nrows:
        row = table.row_values(linenum)
        if row[1] == '' or row[1].find("####") != -1:
            break;
        if row[1].find(u'小学') != -1 or row[1].find(u'中学') != -1 \
            or row[1].find(u'学校') != -1:
            linenum += 1
            continue
        projectname = "" + row[1] + ""
        list.append(projectname)
        linenum += 1

    if list.__len__() == 0:
        print '404 Not Found!'
        return [];
    else:
        return list


def write_result_to_sheet3(cityname, readfile = 'file.xls', writefile = 'file.xls',\
                           writeindex = 0, by_name=u'Sheet3'):
    '''
    get sheets result
    '''

    data = write_excel(writefile)
    worksheet = data.add_sheet(by_name, cell_overwrite_ok=True)

    valuelist = excel_table_byname(readfile, 52, 2, 3, u'表2_村饮水现状、规划及落实')
    for l in valuelist:
        if l[2] == '':
            continue
        strv = (l[0] + "!" + l[1] + "," + "%d" + ";") %(int(l[2]))
        print strv

    print '写入数据'
    # 写数据到新项目文件
    projectlist = find_project(projectfile, cityname)
    projectnum = projectlist.__len__()
    if projectnum == 0:
        return ;

    listnum = valuelist.__len__()
    print "projectlistnum = ", projectnum, "listnum = ", listnum
    totalnum = 0

    index = 0
    for pro in projectlist:
        worksheet.write(index, 1, pro)
        index += 1

    index = 0
    for l in valuelist:
        if l[2] == '':
            continue
        strv = (l[0] + "!" + l[1] + "," + "%d" + ";") %(int(l[2]))
        worksheet.write(index, 0, l[0])
        worksheet.write(index, 2, u"2015/12/30")
        worksheet.write(index, 5, u"880")
        worksheet.write(index, 6, strv)
        worksheet.write(index, 7, l[2])
        worksheet.write(index, 8, l[2])
        totalnum += int(l[2])
        index += 1
        if index >= projectnum:
            break

    # 若村多于工程，合并最后一个
    rst = ""
    tmptotal = 0
    index = projectnum - 1
    while index < listnum:
        strv = (valuelist[index][0] + "!" + valuelist[index][1] + "," + \
                "%d" + ";") %(int(valuelist[index][2]))
        rst += strv
        tmptotal += int(valuelist[index][2])
        index += 1
    worksheet.write(projectnum - 1, 6, rst)
    worksheet.write(projectnum - 1, 7, tmptotal)
    worksheet.write(projectnum - 1, 8, tmptotal)

    worksheet.write(projectnum, 1, projectnum)
    worksheet.write(projectnum, 6, listnum)

    data.save(writefile)

def excel_table_byname(file = 'file.xls', judgeindex = 0, readindex1 = 0, \
                    readindex2 = 0, by_name1=u'Sheet2'):
    '''
    get excel table by name.
    '''
    data = open_excel(file)
    table = data.sheet_by_name(by_name1)
    #table3 = data.sheet_by_name(by_name2)
    nrows = table.nrows
    #sheet3rows = table3.nrows
    list = []

    for rownum in range(8, nrows):
        row = table.row_values(rownum)
        if row[judgeindex] != '' and int(row[judgeindex]) != 0:
            zhen = row[readindex1]
            cun = row[readindex2]
            strvalue = row[judgeindex]
            list.append([zhen, cun, strvalue])

    return list


def parse_args(argv):
    '''
    set parse args
    '''
    Usage = "Usage: ./trans_excel.py [-d dir] [-f filename.xls]"
    parser = OptionParser(Usage)
    parser.add_option('-f', '--file', dest='FILE', help='operate specify file.')
    parser.add_option('-d', '--dir', dest='DIR', help='operate all dir files.')

    (options, args) = parser.parse_args()
    return options


def main(argv=None):
    '''
    main function.
    '''
    # set parse_args
    options =  parse_args(argv)

    if options.DIR:
        print "Operate all ", options.DIR
        dirlist = get_all_dir(options.DIR)
        get_all_files_from_dir(dirlist)

    if options.FILE:
        filedir = os.path.dirname(options.FILE)
        filename = os.path.basename(options.FILE)
        print filedir + os.sep + filename
        (citycode, cityname) = get_codeandname_from_file(filename)
        readfile = citycode + cityname +'.xls'
        writefile = citycode + cityname + 'Sheet3.xls'
        write_result_to_sheet3(unicode(cityname, 'utf-8'), readfile, writefile, 8,\
                               u'表3_集中供水工程基本情况')

if __name__=='__main__':
    main(sys.argv)
