# author: 蔡合升
# datetime: 2022/8/16 22:44
# function:

import xlwt
from xmindparser import xmind_to_dict


def resolvePath(dict, lists, title):
    # title去除首尾空格
    title = title.strip()
    # 如果title是空字符串，则直接获取value
    if len(title) == 0:
        concatTitle = dict['title'].strip()
    else:
        concatTitle = title + '\t' + dict['title'].strip()
    if dict.__contains__('topics') == False:
        lists.append(concatTitle)
    else:
        for d in dict['topics']:
            resolvePath(d, lists, concatTitle)


def xmind_to_excel(list, excel_path):
    # print(f'list是{list}')

    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet('签署模块', cell_overwrite_ok=True)

    # 第一行固定的表头标题
    row_header = ['序号', '模块', '功能点']
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i])

    # 增量索引
    index = 0

    for h in range(0, len(list)):
        lists = []
        resolvePath(list[h], lists, '')
        # print(list[h])
        # print('\n'.join(lists))
        # print(len(lists))
        # print(lists)

        for j in range(0, len(lists)):
            lists[j] = lists[j].split('\t')
            # print(lists[j])
            # print(f'这是lists[j]长度{len(lists[j])}')

            for n in range(0, len(lists[j])):
                # print(lists[j][n])
                # 第一列的序号
                sheet.write(j + index + 1, 0, j + index + 1)
                sheet.write(j + index + 1, n + 1, lists[j][n])
                # 自定义内容，比如测试点、预期结果、实际结果、操作步骤、优先级……
                if n >= 2:
                    sheet.write(0, n + 1, '自定义' + str(n - 1))
            # 遍历结束lists，给增量索引赋值，跳出for j循环，开始for h循环
            if j == len(lists) - 1:
                index += len(lists)
    f.save(excel_path)


def run(xmind_path):
    # 将XMind转化成字典
    xmind_dict = xmind_to_dict(xmind_path)
    # Excel文件与XMind文件保存在同一目录下
    excel_name = xmind_path.split('\\')[-1].split(".")[0] + '.xlsx'
    excel_path = "\\".join(xmind_path.split('\\')[:-1]) + "\\" + excel_name
    print(excel_path)
    xmind_to_excel(xmind_dict[0]['topic']['topics'], excel_path)


if __name__ == '__main__':
    xmind_path_ = r'F:\Desktop\xmind_excel\用例模板.xmind'
    run(xmind_path_)
