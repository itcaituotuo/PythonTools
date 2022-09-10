# author: 测试蔡坨坨
# datetime: 2022/9/4 18:04
# function: Python操作YAML文件

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
