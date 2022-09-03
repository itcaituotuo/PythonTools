# author: 测试蔡坨坨
# datetime: 2022/8/16 22:44
# function: XMind转Excel

from typing import Any, List

import xlwt
from xmindparser import xmind_to_dict


def resolve_path(dict_, lists, title):
    """
    通过递归取出每个主分支下的所有小分支并将其作为一个列表
    :param dict_:
    :param lists:
    :param title:
    :return:
    """
    # 去除title的首尾空格
    title = title.strip()
    # 若title为空，则直接取value
    if len(title) == 0:
        concat_title = dict_["title"].strip()
    else:
        concat_title = title + "\t" + dict_["title"].strip()
    if not dict_.__contains__("topics"):
        lists.append(concat_title)
    else:
        for d in dict_["topics"]:
            resolve_path(d, lists, concat_title)


def xmind_to_excel(list_, excel_path):
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("XX模块", cell_overwrite_ok=True)

    # 第一行固定的表头标题
    row_header = ["序号", "模块", "功能点"]
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i])

    # 增量索引
    index = 0

    for h in range(0, len(list_)):
        lists: List[Any] = []
        resolve_path(list_[h], lists, "")
        # print(lists)
        # print('\n'.join(lists))  # 主分支下的小分支

        for j in range(0, len(lists)):
            # 将主分支下的小分支构成列表
            lists[j] = lists[j].split('\t')
            # print(lists[j])

            for n in range(0, len(lists[j])):
                # 生成第一列的序号
                sheet.write(j + index + 1, 0, j + index + 1)
                sheet.write(j + index + 1, n + 1, lists[j][n])
                # 自定义内容，比如：测试点/用例标题、预期结果、实际结果、操作步骤、优先级……
                # 这里为了更加灵活，除序号、模块、功能点的标题固定，其余以【自定义+序号】命名，如：自定义1，需生成Excel表格后手动修改
                if n >= 2:
                    sheet.write(0, n + 1, "自定义" + str(n - 1))
            # 遍历完lists并给增量索引赋值，跳出for j循环，开始for h循环
            if j == len(lists) - 1:
                index += len(lists)
    f.save(excel_path)


def run(xmind_path):
    """
    运行主程序
    :param xmind_path: XMind文件绝对路径
    :return:
    """
    # 将XMind转化成字典
    xmind_dict = xmind_to_dict(xmind_path)
    # print("将XMind中所有内容提取出来并转换成列表：", xmind_dict)
    # Excel文件与XMind文件保存在同一目录下
    excel_name = xmind_path.split('\\')[-1].split(".")[0] + '.xlsx'
    excel_path = "\\".join(xmind_path.split('\\')[:-1]) + "\\" + excel_name
    print(excel_path)
    # print("通过切片得到所有分支的内容：", xmind_dict[0]['topic']['topics'])
    xmind_to_excel(xmind_dict[0]['topic']['topics'], excel_path)


if __name__ == '__main__':
    xmind_path_ = r"F:\Desktop\coder\PythonTools\xmind_to_excel\用例模板.xmind"
    run(xmind_path_)
