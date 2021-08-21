# 发票信息读取
PDF发票信息读取 : fapiao_pdf_read.py
* 支持发票生成新文件名字
* 支持发票信息导出到EXCEL和TXT

#### 原发票目录（windows下路径\需要转译 \\ 或者 / ）
in_path = u'C:\\Windows\\Temp'
#### 新发票目录（windows下路径\需要转译 \\ 或者 / ）
out_path = u'C:\\Windows\\Temp'
#### 临时目录（windows下路径\需要转译 \\ 或者 / ）
temp_path = 'C:\\Windows\\Temp'
#### 发票扩展名
extension_name = 'pdf'

#### 是否生成EXCEL文档 True or False
is_create_xlsx = True
#### 是否生成TXT文档 True or False
is_create_txt = False
#### 是否生成新文件名字 True or False
is_create_file = True
#### 新文件名字 （'购买方名称-发票号码-开票日期-服务名称-价税合计-销售方名称'）
new_file_format = ['购买方名称', '发票号码', '开票日期', '服务名称', '价税合计', '销售方名称']
#### 新文件分隔符，不填则无分隔符
new_file_join = '-'
