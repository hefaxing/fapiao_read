#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# author: CY
# Date: 2021-08-19
# QQ: 77061066
# Version: 1.0.210820.1700
#
# pip install pdf2docx==0.5.2
# pip install pyzbar==0.1.8
# pip install pandas==1.3.0
# pip install Pillow==8.3.1
# pip install frontend==0.0.3
# pip install openpyxl==3.0.7
#
# Net Framework 4
# Adobe Acrobat
# PDFConvert.exe
#
# 获取PDF发票信息
# 相应配置查看config.ini，保证config.ini跟脚本在同一目录
# 读取方法：利用Acrobat把PDF转DOCX，解压DOCX使用xml读取word/document.xml文字，获取出所有文字类w:t节点nodeValue值，re过滤出发票内容。

import os
import ast
import re
import configparser
import fitz
import shutil
import pandas
import subprocess
from pdf2docx import Converter
from zipfile import ZipFile
from xml.dom.minidom import parseString
from PIL import Image
from pyzbar.pyzbar import decode


def subprocess_popen(CMD):
    # 启用子进程执行外部shell命令
    try:
        # 执行外部shell命令
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        P_CMD = subprocess.Popen(CMD, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)
        P_CMD.wait()
        ret1, ret2 = P_CMD.communicate()
        RET_INFO = ret1 + ret2
    except Exception as e:
        RET_INFO = 'error'
    return RET_INFO


def Is_chinese(ch):
    # 判断字符是否中文
    # ch : str ('中')

    if '\u4e00' <= ch <= '\u9fff':
        return True
    return False

def Filter_cn(filter_list):
    # 过滤中文字符，取中文字符前的字符串
    # filter_list : list (['9', '2', '4', 'D', '2', 'C', '广', '州', '市'])

    ret_str = ''
    for filter_str in filter_list:
        if Is_chinese(filter_str):
            break
        ret_str = ret_str + filter_str
    return ret_str

def Get_files(path, suffix):
    # 获取目录下所有匹配扩展名的文件
    # path : str ('C:\\Windows\\Temp')
    # suffix : str ('pdf')
    #
    #pdf_files = [os.path.join(root, file_name) for root, subdirs, file_names in os.walk(path) for file_name in file_names if file_name.endswith('.%s' % suffix)]

    pdf_files = []
    for root, subdirs, file_names in os.walk(path):
        for file_name in file_names:
            if file_name.endswith('.%s' % suffix):
                file_path = os.path.join(root, file_name)
                pdf_files.append(file_path)

    return pdf_files

def Filter_str(filter_str):
    # 过滤不必要的字符
    # filter_str : str ('采用PDF转DOCX，解压DOCX使用xml读取word/document.xml文字，获取出所有文字类w:t节点nodeValue值，re过滤出发票内容。')

    filter_str = filter_str.replace(' ', '')
    filter_str = filter_str.replace('　', '')
    filter_str = filter_str.replace(':', '')
    filter_str = filter_str.replace('：', '')
    filter_str = filter_str.replace('￥', '')
    filter_str = filter_str.replace('¥', '')
    filter_str = filter_str.replace('（', '(')
    filter_str = filter_str.replace('）', ')')
    return filter_str

def Read_QR_code(img_file):
    # 读取图片中的二维码信息
    # img_file : str ('C:\\Windows\\Temp\\temp.png')

    qr_code = {'qr_code_code': '', 'qr_code_number': '', 'qr_code_total': '', 'qr_code_date': '', 'qr_code_check_code': ''}
    img = Image.open(img_file)
    barcodes = decode(img)
    if barcodes is not None:
        qr_code_data = barcodes[0].data.decode("utf-8")
        qr_code_date = qr_code_data.split(',')
        qr_code['qr_code_code'] = qr_code_date[2]
        qr_code['qr_code_number'] = qr_code_date[3]
        qr_code['qr_code_total'] = qr_code_date[4]
        qr_code['qr_code_date'] = qr_code_date[5]
        qr_code['qr_code_check_code'] = qr_code_date[6]
    return qr_code

def From_pdf_to_png(pdf_file, temp_png_file):
    # pdf转png
    # pdf_file : str ('C:\\Windows\\Temp\\temp.pdf')
    # temp_png_file : str ('C:\\Windows\\Temp\\temp.png')

    doc = fitz.open(pdf_file)
    page = doc.loadPage(0)
    trans = fitz.Matrix(2, 2)
    pix = page.getPixmap(matrix=trans, alpha=False)
    pix.writePNG(temp_png_file)
    doc.close()
    return temp_png_file

def From_pdf_to_docx(pdf_file, docx_file):
    # pdf转docx
    # pdf_file : str ('C:\\Windows\\Temp\\temp.pdf')
    # docx_file : str ('C:\\Windows\\Temp\\temp.docx')

    cv = Converter(pdf_file)

    # 默认参数start=0, end=None表示转换所有页面
    cv.convert(docx_file)
    cv.close()
    return docx_file

def From_pdf_to_docx2(pdf_file, docx_path):
    # pdf转docx
    # pdf_file : str ('C:\\Windows\\Temp\\temp.pdf')
    # docx_path : str ('C:\\Windows\\Temp')

    s = subprocess_popen('%s\\PDFConvert.exe -f docx -i "%s" -o %s' % (python_path, pdf_file, docx_path))

    if s:
        From_pdf_to_docx(pdf_file, '%s\\%s' % (docx_path, os.path.basename(pdf_file).replace(extension_name, 'docx')))
    return docx_path


def Read_docx(docx_file):
    # 解压docx文档
    # docx_file : str ('C:\\Windows\\Temp\\temp.docx')

    zf = ZipFile(docx_file)
    # 查看所有文件名字
    #for item in zf.filelist:
    #    print(item.filename)

    # 读取word/document.xml内容
    myfile = zf.open('word/document.xml')
    xml_str = myfile.read()
    zf.close()

    #collection = DOMTree.documentElement
    #print('collection属性',collection.nodeName,collection.nodeValue,collection.nodeType)
    
    # 解析xml内容
    DOMTree = parseString(xml_str)
    collection = DOMTree.documentElement
    
    # 获取所有'w:t'节点
    w_t = collection.getElementsByTagName('w:t')
    node_text = []
    for w_t_node in w_t:
        for node in w_t_node.childNodes:
            node_text.append(Filter_str(node.nodeValue))
    
    #print(node_text)
    return node_text

def Get_fapiao_info(text_info):
    # 匹配发票信息
    # text_info : list (['发票代码', '888888888888', '发票号码', '88888888', ...])

    text_info = ''.join(text_info)
    text_info2 = text_info.split('密码区')[-1]
    #print(text_info)

    # 此re匹配报错，修改匹配规则，或如不必要，可注释对应行
    fapiao_info = {}
    fapiao_info['fapiao_code'] = re.findall("发票代码(\d+)", text_info)[0]
    fapiao_info['fapiao_number'] = re.findall("发票号码(\d+)", text_info)[0]
    fapiao_info['fapiao_date'] = re.findall("开票日期(.*?)校验码", text_info)[0]
    fapiao_info['fapiao_check_code'] = re.findall("校验码(\d+)", text_info)[0]
    fapiao_info['fapiao_buyer_name'] = re.findall("购买方名称(.*?)纳税人识别号", text_info)[0]
    fapiao_info['fapiao_buyer_name'] = fapiao_info['fapiao_buyer_name'].split('密码区')[0]
    fapiao_info['fapiao_buyer_tax_number'] = re.findall("购买方名称.*?纳税人识别号(.*?)地址", text_info)[0]
    try:
        fapiao_info['fapiao_goods'] = re.findall("服务名称(.*?)合计规格型号", text_info)[0]
    except:
        fapiao_info['fapiao_goods'] = re.findall("项目名称(.*?)合计规格型号", text_info)[0]
    fapiao_info['fapiao_s_tax_total'] = re.findall("\(大写\)(.*?)\(小写\)", text_info)[0]
    fapiao_info['fapiao_tax_total'] = re.findall("\(小写\)(\d+.\d+)", text_info)[0]
    fapiao_info['fapiao_seller_name'] = re.findall("销售方名称(.*?)纳税人识别号", text_info2)[0]
    fapiao_info['fapiao_seller_tax_number'] = Filter_cn(list(re.findall("销售方名称.*?纳税人识别号(.*?)地址", text_info2)[0]))
    fapiao_info['fapiao_address_phone'] = re.findall("销售方名称.*?电话(.*?)开户行", text_info2)[0]
    fapiao_info['fapiao_bank_name'] = re.findall("销售方名称.*?开户行及账号(.*?)备注", text_info2)[0]

    #print('发票代码 %s' % fapiao_info['fapiao_code'])
    #print('发票号码 %s' % fapiao_info['fapiao_number'])
    #print('开票日期 %s' % fapiao_info['fapiao_date'])
    #print('校验码 %s' % fapiao_info['fapiao_check_code'])
    #print('购买方名称 %s' % fapiao_info['fapiao_buyer_name'])
    #print('购买方纳税人识别号 %s' % fapiao_info['fapiao_buyer_tax_number'])
    #print('服务名称 %s' % fapiao_info['fapiao_goods'])
    #print('价税合计（大写） %s' % fapiao_info['fapiao_tax_total'])
    #print('价税合计（小写） %s' % fapiao_info['fapiao_tax_total'])
    #print('销售方名称 %s' % fapiao_info['fapiao_seller_name'])
    #print('销售方纳税人识别号 %s' % fapiao_info['fapiao_seller_tax_number'])
    #print('地址、电话 %s' % fapiao_info['fapiao_address_phone'])
    #print('开户行及账号 %s' % fapiao_info['fapiao_bank_name'])
    #print('=' * 50)
    return fapiao_info

def Save_txt(text_info, out_file):
    # 把获取的信息保存到txt文档
    # text_info : str or list ('发票代码888888888888发票号码88888888...' or ['发票代码', '888888888888', '发票号码', '88888888', ...])
    # out_file : str ('C:\\Windows\\Temp\\temp.txt')

    if isinstance(text_info, list):
        out_text = ''.join(text_info).encode('utf-8')
    elif isinstance(text_info, str):
        out_text = text_info.encode('utf-8')
    # 保存读取的信息
    with open(out_file, 'ab') as file_object:
        file_object.write(b"%s\n" % out_text)
        file_object.close()
    return out_file

def Save_xlsx(text_info, out_file):
    # 保存发票信息到EXCEL文档
    # text_info : list ([{'发票代码': '888888888888', '发票号码': '88888888', ...}, {'发票代码': '888888888888', '发票号码': '88888888', ...}])
    # out_file : str ('C:\\Windows\\Temp\\temp.xlsx')

    pf = pandas.DataFrame(text_info)

    # 指定列的顺序
    order = ["发票代码", "发票号码", "开票日期", "校验码", "购买方名称", '购买方纳税人识别号', "价税合计", "服务名称", "销售方名称", "销售方纳税人识别号", "销售方地址、电话", "销售方开户行及账号"]
    pf = pf[order]
    # 打开excel文件
    file_path = pandas.ExcelWriter(out_file)
    # 替换空单元格
    pf.fillna(' ', inplace=True)
    # 输出
    pf.to_excel(file_path, encoding='utf-8', index=False, sheet_name="sheet1")
    file_path.save()
    return out_file

def Filter_name(f_text):
    # 过滤企业名字头和尾
    # f_text : str (XXXX有限公司)

    try:
        filter_name = "START%sEND" % re.findall('.*?公司', f_text)[0]
    except:
        filter_name = "START%sEND" % f_text
    for filter_s in cn_region:
        filter_name = filter_name.replace("START%s市" % filter_s, '')
    for filter_s in cn_region:
        filter_name = filter_name.replace("START%s" % filter_s, '')
    for filter_s in company_name_filter:
        filter_name = filter_name.replace("%sEND" % filter_s, '')
    for filter_s in region_filter:
        filter_name = filter_name.replace("START%s" % filter_s, '')
    filter_name = filter_name.replace("START", '')
    filter_name = filter_name.replace("END", '')
    return filter_name

def Filter_goods(f_text):
    # 过滤服务名称，剔除第一项
    # f_text : str ('*热爱祖国*服务人民*崇尚科学*辛勤劳动*团结互助*诚实守信*遵纪守法*艰苦奋斗')

    if '*' in f_text:
        # *号分隔，截取第三个之后的
        filter_name = f_text.split('*')[2:]
        if filter_name == []:
            filter_name = f_text.split('*')[1:]
        filter_name = ''.join(filter_name)
    else:
        filter_name = f_text
    # 过滤符号
    filter_name = re.sub('\W+', '', filter_name).replace("_", '')
    # 截取前十个字符
    filter_name = filter_name[:10]
    return filter_name

def Clear_temp_file(temp_file):
    # 清理临时文件
    # temp_file : str ('C:\\Windows\\Temp\\temp.txt')

    os.remove(temp_file)

def Is_exists(is_path, is_type):
    # 判断目录是否存在
    # in_path : str ('C:\\Windows\\Temp')
    # is_type : int (0: 目录不存在抛出异常，1: 目录不存在则创建)

    is_exists = os.path.exists(is_path)

    if is_exists:
        return is_path
    
    if is_type == 0:
        raise UserWarning(u"config: %s 不存在此目录 " % is_path)
    elif is_type == 1:
        os.makedirs(is_path)

    return is_path

def New_file_name(text_info, file_format, file_join):
    # 生成新文件名字
    # text_info : list ([{'发票代码': '888888888888', '发票号码': '88888888', ...}, {'发票代码': '888888888888', '发票号码': '88888888', ...}])
    # file_format : list (['购买方名称', '发票号码', '开票日期', '服务名称', '价税合计', '销售方名称'])
    # file_join : str ('-')

    file_format_list = []
    for order_name in [fapiao_order[s_name] for s_name in file_format]:
        if order_name == 'fapiao_buyer_name':
            filter_info = Filter_name(text_info[order_name])
            if filter_info in company_name.keys():
                filter_info = company_name[buyer_name]
        elif order_name == 'fapiao_seller_name':
            filter_info = Filter_name(text_info[order_name])
        elif order_name == 'fapiao_goods':
            filter_info = Filter_goods(text_info[order_name])
        else:
            filter_info = text_info[order_name]
        file_format_list.append(filter_info)

    return file_join.join(file_format_list)

if __name__ == '__main__':
    # 获取脚本绝对路径
    python_file = os.path.abspath(__file__)
    python_path = os.path.abspath(os.path.dirname(os.path.abspath(__file__)))

    # 读取config配置文件
    cf = configparser.ConfigParser()
    cf.read("%s\\config.ini" % python_path, encoding='utf-8-sig')

    in_path = ast.literal_eval(cf.get("config", "in_path"))
    out_path = ast.literal_eval(cf.get("config", "out_path"))
    temp_path = ast.literal_eval(cf.get("config", "temp_path"))
    extension_name = ast.literal_eval(cf.get("config", "extension_name"))
    fapiao_order = ast.literal_eval(cf.get("config", "fapiao_order"))
    company_name = ast.literal_eval(cf.get("config", "company_name"))
    company_name_filter = ast.literal_eval(cf.get("config", "company_name_filter"))
    cn_region = ast.literal_eval(cf.get("config", "cn_region"))
    region_filter = ast.literal_eval(cf.get("config", "region_filter"))
    is_create_file = ast.literal_eval(cf.get("config", "is_create_file"))
    is_create_xlsx = ast.literal_eval(cf.get("config", "is_create_xlsx"))
    is_create_txt = ast.literal_eval(cf.get("config", "is_create_txt"))
    new_file_format = ast.literal_eval(cf.get("config", "new_file_format"))
    new_file_join = ast.literal_eval(cf.get("config", "new_file_join"))

    # 临时文件
    temp_docx_file = '%s\\docx_temp.docx' % temp_path
    temp_png_file = '%s\\png_temp.png' % temp_path
    temp_txt_file = '%s\\发票归总.txt' % out_path
    temp_xlsx_file = '%s\\发票归总.xlsx' % out_path

    # 判断读取目录是否存在
    Is_exists(in_path, 0)

    # 判断存放目录是否存在，不存在则创建
    Is_exists(out_path, 1)

    failed_files = []
    new_file_sum = {}
    results = []
    # 获取目录下所有需要处理的文件
    file_names = Get_files(in_path, extension_name)
    for file_name in file_names:
        try:
            From_pdf_to_png(file_name, temp_png_file)
            ret_info = Read_QR_code(temp_png_file)
            From_pdf_to_docx2(file_name, temp_path)
            temp_docx_file = '%s\\%s' % (temp_path, os.path.basename(file_name).replace(extension_name, 'docx'))
            text_info = Read_docx(temp_docx_file)
            Clear_temp_file(temp_docx_file)
            ret_info.update(Get_fapiao_info(text_info))
        except:
            failed_files.append(file_name)
            continue

        if is_create_txt:
            Save_txt(text_info, temp_txt_file)

        #print(ret_info)
        #ret_info = {'qr_code_code': '888888888888', 'qr_code_number': '88888888', 'qr_code_total': '88.88', 'qr_code_date': '20210819', 'qr_code_check_code': '88888888888888888888', 'fapiao_code': '888888888888', 'fapiao_number': '88888888', 'fapiao_check_code': '88888888888888888888', 'fapiao_buyer_name': 'XXXX有限公司', 'fapiao_buyer_tax_number': '购买方纳税人识别号', 'fapiao_goods': '*热爱祖国*服务人民*崇尚科学*辛勤劳动*团结互助*诚实守信*遵纪守法*艰苦奋斗', 'fapiao_s_tax_total': '壹佰陆拾捌圆整', 'fapiao_tax_total': '168.00', 'fapiao_seller_name': '销售方名称', 'fapiao_seller_tax_number': '销售方纳税人识别号', 'fapiao_address_phone': '销售方地址、电话', 'fapiao_bank_name': '销售方开户行及账号'}

        if is_create_file:
            # 以新文件命名：购买方名称-发票号码-开票日期-服务名称-价税合计-销售方名称
            new_file_name = New_file_name(ret_info, new_file_format, new_file_join)
            new_path = os.path.dirname(file_name.replace(in_path, out_path))
            new_file = "%s\%s.%s" % (new_path, new_file_name, extension_name)
            print(new_file)
            ret_info['new_file'] = new_file
            Is_exists(new_path, 1)

            shutil.copyfile(file_name, new_file)
            if new_file_name in  new_file_sum.keys():
                new_file_sum[new_file_name].append(file_name)
            else:
                new_file_sum[new_file_name] = [file_name]

        for fapiao_key, fapiao_value in fapiao_order.items():
            ret_info[fapiao_key] = ret_info.pop(fapiao_value)
        results.append(ret_info)

    
    Clear_temp_file(temp_png_file)

    if is_create_xlsx:
        Save_xlsx(results, temp_xlsx_file)

    print('')
    print(u'=' * 50)
    print(u'输出目录：%s' % out_path)
    print(u'=' * 50)
    print('')
    repeat_file_list = []
    for new_file, old_file in new_file_sum.items():
        if len(old_file) > 1:

            repeat_file_list.append([new_file] + old_file)

    if repeat_file_list:
        print(u'-' * 50)
        print(u'重复文件：')
        for repeat_file in repeat_file_list:
            for i in repeat_file:
                print(i)
            print('')
        print(u'-' * 50)

    print('')
    if failed_files:
        print(u'*' * 50)
        print(u'读取失败文件：')
        for failed_file in failed_files:
            print(failed_file)
        print(u'*' * 50)
