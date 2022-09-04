# PythonTools

Python编写的小工具集合



## Python操作Excel

./operate_excel

### 安装

xlwt是Python的第三方模块，需要先下载安装才能使用，这里我们使用pip命令下载

```bash
pip3 install xlwt
```



### 使用

1. 导入Excel表格文件处理函数

   ```python
   import xlrd
   import xlwt
   from faker import Faker
   ```

2. 创建Excel表格类型文件

   ```python
   # 实例化Workbook对象
   # 第一个参数：encoding表示编码
   # 第二个参数：style_compression设置是否压缩，0表示不压缩
   work_book = xlwt.Workbook(encoding="utf-8", style_compression=0)
   ```

3. 在Excel表格类型文件中建立一张表sheet表单

   ```python
   # 第一个参数：sheetname，表示sheet名
   # 第二个参数：cell_overwrite_ok用于确认同一cell单元是否可以重设值，True表示可以重设
   sheet = work_book.add_sheet(sheetname="用户信息表", cell_overwrite_ok=True)
   ```

4. 自定义列名

   ```python
   # 用一个元组col自定义列的数量以及属性
   col = ("姓名", "电话", "地址")
   ```

5. 将列属性元组col写进sheet表单中

   ```python
   # 使用for循环将col元组的元组值写到sheet表单中
   # 第一个参数是行，第二个参数是列，第三个参数是值
   for i in range(0, 3):
       sheet.write(0, i, col[i])
   ```

6. 创建数据并将数据写入表格

   ```python
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
   ```

7. 保存Excel文件，调用save()方法

   ```python
   # 定义一个文件路径save_path，例如当前目录下./ 文件名为 userinfo.xls
   save_path = "./userinfo.xls"
   work_book.save(save_path)
   ```

8. 读取Excel文件（ps：读取前确保文件非打开状态）

   ```python
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
   ```

运行结果：借助Faker模块生成随机的个人信息，并将其写入Excel表格

![](C:/Users/DELL/AppData/Roaming/Typora/typora-user-images/image-20220903204358075.png)



### 完整代码

源码获取请关注公众号`测试蔡坨坨`，回复关键词`源码`

```python
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

```





## Python代码实现XMind测试用例快速转成Excel用例

./xmind_to_excel

### 前言

XMind和Excel是在日常测试工作中最常用的两种用例编写形式，两者也有各自的优缺点。

使用XMind编写测试用例更有利于测试思路的梳理，以及更加便捷高效，用例评审效率更高，但是由于每个人使用XMind的方式不同，设计思路也不一样，可能就不便于其他人执行和维护。

使用Excel编写测试用例由于有固定的模板，所以可能更加形式化和规范化，更利于用例管理和维护，以及让其他人更容易执行用例，但是最大的缺点就是需要花费更多的时间成本。

由于项目需要，需要提供Excel形式的测试用例，同时编写两种形式的测试用例显然加大了工作量，于是写了个Python脚本，可快速将XMind用例转换成Excel用例。





### 设计思路

Excel测试用例模板样式如下图所示：

<img src="https://caituotuo.top/my-img/202208302310296.png"  />

表头固定字段：序号、模块、功能点

为了让脚本更加灵活，后面的字段会根据XMind中每一个分支的长度自增，例如：测试点/用例标题、预期结果、实际结果、前置条件、操作步骤、优先级、编写人、执行人等



根据Excel模板编写对应的XMind测试用例：

![](https://caituotuo.top/my-img/202208302319084.png)



实现：

将XMind中的每一条分支作为一条序号的用例，再将每个字段写入Excel中的每一个单元格中

![](https://caituotuo.top/my-img/202208302321804.png)

再手动调整美化一下表格：

<img src="https://caituotuo.top/my-img/202208302325085.png" style="zoom:80%;" />





### 完整代码

GitHub：https://github.com/itcaituotuo/PythonTools

```python
# author: 小趴蔡
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
    xmind_path_ = r"F:\Desktop\coder\python_operate_files\用例模板.xmind"
    run(xmind_path_)

```





### 代码解析

#### 1. 调用xmind_to_dict()方法将XMind中所有内容取出并转成字典

```python
xmind_dict = xmind_to_dict(xmind_path)
```

```json
[{'title': '画布 1', 'topic': {'title': '需求名称', 'topics': [{'title': '模块', 'topics': [{'title': '功能点1', 'topics': [{'title': '测试点1', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}, {'title': '测试点2', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}, {'title': '测试点3'}]}, {'title': '功能点2', 'topics': [{'title': '测试点1'}, {'title': '测试点2', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}]}, {'title': '功能点3'}]}]}, 'structure': 'org.xmind.ui.logic.right'}]
```



#### 2. 通过切片得到所有分支的内容

```python
xmind_dict[0]['topic']['topics']
```

```json
[{'title': '模块', 'topics': [{'title': '功能点1', 'topics': [{'title': '测试点1', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}, {'title': '测试点2', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}, {'title': '测试点3'}]}, {'title': '功能点2', 'topics': [{'title': '测试点1'}, {'title': '测试点2', 'topics': [{'title': '预期结果', 'topics': [{'title': '实际结果'}]}]}]}, {'title': '功能点3'}]}]
```



#### 3. 通过递归取出每个主分支下的所有小分支并将其作为一个列表

```python
def resolve_path(dict_, lists, title):
    """
    通过递归取出每个主分支下的所有小分支并将其作为一个列表
    :param dict_:
    :param lists:
    :param title:
    :return:
    """
    # 去除title首尾空格
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
```

```python
    for h in range(0, len(list_)):
        lists: List[Any] = []
        resolve_path(list_[h], lists, "")
        print(lists)
        print('\n'.join(lists))  # 主分支下的小分支

        for j in range(0, len(lists)):
            # 将主分支下的小分支构成列表
            lists[j] = lists[j].split('\t')
            print(lists[j])
```

```
lists：
['模块\t功能点1\t测试点1\t预期结果\t实际结果', '模块\t功能点1\t测试点2\t预期结果\t实际结果', '模块\t功能点1\t测试点3', '模块\t功能点2\t测试点1', '模块\t功能点2\t测试点2\t预期结果\t实际结果', '模块\t功能点3']

主分支下的小分支：
模块	功能点1	测试点1	预期结果	实际结果
模块	功能点1	测试点2	预期结果	实际结果
模块	功能点1	测试点3
模块	功能点2	测试点1
模块	功能点2	测试点2	预期结果	实际结果
模块	功能点3

将主分支下的小分支构成列表：
['模块', '功能点1', '测试点1', '预期结果', '实际结果']
['模块', '功能点1', '测试点2', '预期结果', '实际结果']
['模块', '功能点1', '测试点3']
['模块', '功能点2', '测试点1']
['模块', '功能点2', '测试点2', '预期结果', '实际结果']
['模块', '功能点3']
```



#### 4. 写入Excel（生成单sheet的Excel文件、生成固定的表头标题、列序号取值、固定标题外的自定义标题）

```python
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("xx模块", cell_overwrite_ok=True)

    # 第一行固定的表头标题
    row_header = ["序号", "模块", "功能点"]
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i])
```

```python
            for n in range(0, len(lists[j])):
                # 生成第一列的序号
                sheet.write(j + index + 1, 0, j + index + 1)
                sheet.write(j + index + 1, n + 1, lists[j][n])
                # 自定义内容，比如：测试点/用例标题、预期结果、实际结果、操作步骤、优先级……
                # 这里为了更加灵活，除序号、模块、功能点的标题固定，其余以【自定义+序号】命名，如：自定义1，需生成Excel表格后手动修改
                if n >= 2:
                    sheet.write(0, n + 1, "自定义" + str(n - 1))
```





## YAML数据驱动

### 1 什么是YAML

YAML：YAML Ain't a Markup Language，翻译过来就是「YAML不是一种标记语言」。

但是在开发这种语言时，YAML的意思其实是Yet Another Markup Language「仍是一种标记语言」。

它是一种以数据为中心的标记语言，比 XML 和 JSON 更适合作为配置文件。

YAML的配置文件后缀为**.yml**或**.yaml**，如：**caituotuo.yml**或**caituotuo.yaml**。

YAML的语法和其他高级语言类似，并且可以简单表达清单、散列表，标量等数据形态。它使用空白符号缩进和大量依赖外观的特色，特别适合用来表达或编辑数据结构、各种配置文件、倾印调试内容、文件大纲等。





### 2 YAML语法

#### 2.1 基本语法

- 使用缩进表示层级关系
- 缩进不允许使用tab，只允许空格（官方说法不允许使用tab，当然如果你使用tab在某些地方也是可以的，例如在PyCharm软件上）
- 缩进的空格数不重要，只要相同层级的元素左对齐即可
- 大小写敏感
- 前面加上#表示注释

例如：

```yaml
req:
  username: 测试蔡坨坨 # 这是姓名
  gender: Boy
  ip: 上海
  blog: www.caituotuo.top
res:
  status: 1
  code: 200
```



#### 2.2 数据类型

- 对象：键值对的集合，又称为映射（mapping）/ 哈希（hashes） / 字典（dictionary）
- 数组：一组按次序排列的值，又称为序列（sequence） / 列表（list）
- 纯量（scalars）：单个的、不可再分的值，又称字面量

##### 纯量

纯量是指单个的，不可拆分的值，例如：数字、字符串、布尔值、Null、日期等，纯量直接写在键值对的value中即可。

###### 字符串：

默认情况下字符串是不需要使用单引号或双引号的

```yaml
username: 测试蔡坨坨
```

当然使用双引号或者单引号包裹字符也是可以的

```yaml
username: 'Hello world 蔡坨坨'
username: "Hello world 蔡坨坨"
```

字符串可以拆成多行，每一行会被转化成一个空格

```yaml
# 字符串可以拆成多行，每一行会被转化成一个空格 '测试 蔡坨坨'
username3: 测试
  蔡坨坨
```

###### 布尔值：

```yaml
boolean:
  - TRUE  #true,True都可以
  - FALSE  #false，False都可以
  
# {'boolean': [True, False]}
```

###### 数字：

```yaml
float:
  - 3.14
  - 6.8523015e+5  #可以使用科学计数法
int:
  - 123
  - 0b1010_0111_0100_1010_1110    #二进制表示
  
# {'float': [3.14, 685230.15], 'int': [123, 685230]}
```

###### Null：

```yaml
null:
  nodeName: 'node'
  parent: ~  #使用~表示null
  parent2: None  #使用None表示null
  parent3: null  #使用null表示null
  
# {None: {'nodeName': 'node', 'parent': None, 'parent2': 'None', 'parent3': None}}
```

###### 时间和日期：

```yaml
date:
  - 2018-02-17    #日期必须使用ISO 8601格式，即yyyy-MM-dd
datetime:
  - 2018-02-17T15:02:31+08:00    #时间使用ISO 8601格式，时间和日期之间使用T连接，最后使用+代表时区
  
# {'date': [datetime.date(2018, 2, 17)], 'datetime': [datetime.datetime(2018, 2, 17, 15, 2, 31, tzinfo=datetime.timezone(datetime.timedelta(seconds=28800)))]}
```



##### 对象

使用`key:[空格]value`的形式表示一对键值对（空格不能省略），例如：`blog: caituotuo.top`。

行内写法：

```yaml
key: {key1: value1, key2: value2, ...}
```

普通写法，使用缩进表示对象与属性的层级关系：

```yaml
key: 
    child-key: value
    child-key2: value2
```



##### 数组

以 `-` 开头的行表示构成一个数组。

普通写法：

```yaml
name:
    - 测试蔡坨坨
    - 小趴蔡
    - 蔡蔡
```

YAML 支持多维数组，可以使用行内表示：

```yaml
key: [value1, value2, ...]
```

数据结构的子成员是一个数组，则可以在该项下面缩进一个空格：

```yaml
username:
      -
        - 测试蔡坨坨
        - 小趴蔡
        - 蔡蔡
      -
        - A
        - B
        - C
        
# {'username': [['测试蔡坨坨', '小趴蔡', '蔡蔡'], ['A', 'B', 'C']]}
```

相对复杂的例子：

companies 属性是一个数组，每一个数组元素又是由 id、name、price 三个属性构成

```yaml
companies:
    -
        id: 1
        name: caituotuo
        price: 300W
    -
        id: 2
        name: 测试蔡坨坨
        price: 500W
       
# {'companies': [{'id': 1, 'name': 'caituotuo', 'price': '300W'}, {'id': 2, 'name': '测试蔡坨坨', 'price': '500W'}]}
```

数组也可以使用flow流式的方式表示：

```yaml
companies2: [ { id: 1,name: caituotuo,price: 300W },{ id: 2,name: 测试蔡坨坨,price: 500W } ]
```



##### 复合结构

以上三种数据结构可以任意组合使用，以实现不同的用户需求，例如：

```yaml
platform:
  - 公众号
  - 小红书
  - 博客
sites:
  公众号: 测试蔡坨坨
  小红书: 测试蔡坨坨
  blog: caituotuo.top
  
# {'platform': ['公众号', '小红书', '博客'], 'sites': {'公众号': '测试蔡坨坨', '小红书': '测试蔡坨坨', 'blog': 'caituotuo.top'}
```





### 3 引用

`&` 锚点和 `*` 别名，可以用来引用。

举个栗子：

& 用来建立锚点defaults，<< 表示合并到当前数据，* 用来引用锚点

```yaml
defaults: &defaults
  adapter: postgres
  host: localhost

development:
  database: myapp_development
  <<: *defaults

test:
  database: myapp_test
  <<: *defaults
```

等价于：

```yaml
defaults:
  adapter: postgres
  host: localhost

development:
  database: myapp_development
  adapter: postgres
  host: localhost

test:
  database: myapp_test
  adapter: postgres
  host: localhost
```





### 4 组织结构

一个YAML文件可以由一个或多个文档组成，文档之间使用`---`作为分隔符，且整个文档相互独立，互不干扰，如果YAML文件只包含一个文档，则`---`分隔符可以省略。

```yaml
---
website:
  name: 测试蔡坨坨
  url: caituotuo.top
---
website: { name: 测试蔡坨坨,url: www.caituotuo.top }
---
公众号: 测试蔡坨坨
---
小红书: 测试蔡坨坨
```

```python
f7 = "./files/多文档.yml"
with open(f7, "r", encoding="UTF-8") as f:
    content = yaml.safe_load_all(f)
    for i in content:
        print(i)
```

```bash
运行结果：

{'website': {'name': '测试蔡坨坨', 'url': 'caituotuo.top'}}
{'website': {'name': '测试蔡坨坨', 'url': 'www.caituotuo.top'}}
{'公众号': '测试蔡坨坨'}
{'小红书': '测试蔡坨坨'}
```





### 5 实战

#### 封装思路

将YAML相关操作封装成CommonUtil公共模块，之后直接引入调用即可。

相关功能：

1. 读取yaml文件数据
2. 将yaml数据转换成json格式
3. 可以动态设置参数

这里要说一下动态设置参数

在自动化测试中，肯定不能把所有的参数都写死，因此就会用到参数化，例如：提取前一个接口的返回值作为后一个接口的入参，这里通过Python中的Template模块进行动态参数的设置

yaml文件中通过`$变量名`的形式设置变量

```yaml
username: $username
```

给变量附上具体的值

```python
with open(yaml_path, "r", encoding="UTF-8") as f:
	text = f.read()
# Template(text).safe_substitute(key_value)
Template(text).safe_substitute({"username": "测试蔡坨坨"}) # username为变量名
```



#### 完整代码

源码获取请关注公众号：`测试蔡坨坨`，回复关键词：`源码`

```python
# author: 测试蔡坨坨
# datetime: 2022/9/4 18:04
# function: Python操作YAML文件

import os
from string import Template

import yaml


class YamlUtil:
    @staticmethod
    def yaml_util(yaml_path, key_value=None):
        """
        读取yml文件 设置动态变量
        :param yaml_path: 文件路径
        :param key_value: 动态变量 如：{"username": "测试蔡坨坨"} yaml中的变量：$username
        :return:
        """
        try:
            with open(yaml_path, "r", encoding="UTF-8") as f:
                text = f.read()
                if key_value is not None:
                    re = Template(text).safe_substitute(key_value)
                    json_data = yaml.safe_load(re)
                else:
                    json_data = yaml.safe_load(text)
            return json_data
        except FileNotFoundError:
            raise FileNotFoundError("文件不存在")
        except Exception:
            raise Exception("未知异常")

    @staticmethod
    def multiple(yaml_path):
        """
        多文档
        :param yaml_path: yaml文件路径
        :return: list
        """
        json_data = []
        try:
            with open(yaml_path, "r", encoding="UTF-8") as f:
                content = yaml.safe_load_all(f)
                for i in content:
                    json_data.append(i)
            return json_data
        except FileNotFoundError:
            raise FileNotFoundError("文件不存在")
        except Exception:
            raise Exception("未知异常")


if __name__ == '__main__':
    f1 = "./files/初体验.yml"
    print(YamlUtil().yaml_util(f1))

    f2 = "./files/纯量.yml"
    print(YamlUtil().yaml_util(f2))

    f3 = "./files/数组.yml"
    print(YamlUtil().yaml_util(f3))

    f4 = "./files/复合结构.yml"
    print(YamlUtil().yaml_util(f4))

    f5 = "./files/引用.yml"
    print(YamlUtil().yaml_util(f5))

    f6 = "./files/参数化.yml"
    print(YamlUtil().yaml_util(f6, {"username": "测试蔡坨坨"}))

    f7 = "./files/多文档.yml"
    for i in YamlUtil().multiple(f7):
        print(i)

```
