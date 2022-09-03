# author: 测试蔡坨坨
# datetime: 2022/7/2 20:47
# function: Python操作Excel表格

# xlwt是Python的第三方模块，需要先下载安装才能使用，这里我们使用pip命令下载
# pip3 install xlwt

# 1.导入Excel表格文件处理函数
import xlrd
import xlwt
from faker import Faker

# 2.创建Excel表格类型文件
# 实例化Workbook对象
# 第一个参数：encoding表示编码
# 第二个参数：style_compression设置是否压缩，0表示不压缩
work_book = xlwt.Workbook(encoding="utf-8", style_compression=0)

# 3.在Excel表格类型文件中建立一张表sheet表单
# 第一个参数：sheetname，表示sheet名
# 第二个参数：cell_overwrite_ok用于确认同一cell单元是否可以重设值，True表示可以重设
sheet = work_book.add_sheet(sheetname="用户信息表", cell_overwrite_ok=True)

# 4.自定义列名
# 用一个元组col自定义列的数量以及属性
col = ("姓名", "电话", "地址")

# 5.将列属性元组col写进sheet表单中
# 使用for循环将col元组的元组值写到sheet表单中
# 第一个参数是行，第二个参数是列，第三个参数是值
for i in range(0, 3):
    sheet.write(0, i, col[i])

# 6.创建数据并将数据写入表格
# 使用Faker模块生成10组数据
faker = Faker("zh_CN")
data_list = []
for i in range(0, 10):
    data = [faker.name(), faker.phone_number(), faker.address()]
    data_list.append(data)
print(data_list)  # [['杨雪梅', '13596272521', '湖南省宁德市高明杨街Z座 257668'], ……]

# 将数据写入Excel文件
# 先用第一个for循环进行每行写入
# 再用第二个for循环把每一行当中的列值写进入
for i in range(0, 10):
    data = data_list[i]
    for j in range(0, 3):
        sheet.write(i + 1, j, data[j])

# 7.保存Excel文件，调用save()方法
# 定义一个文件路径save_path，例如当前目录下./ 文件名为 userinfo.xls
save_path = "./userinfo.xls"
work_book.save(save_path)

# 8.读取Excel文件（ps：读取前确保文件非打开状态）
# 得到文件
file_name = xlrd.open_workbook("./userinfo.xls")
# 得到sheet页
sheet = file_name.sheets()[0]
# 获取总列数
total_rows = sheet.nrows
# 获取总行数
total_cols = sheet.ncols
print(total_rows, total_cols)  # 11 3
for i in range(1, total_rows):
    for j in range(0, total_cols):
        info = sheet.row_values(i)[j]
        print(info)
