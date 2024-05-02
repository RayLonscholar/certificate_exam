import re

# 定义要处理的字符串列表
strings = ["52（B）：RayLon", "104（D）：Ray"]

# 定义正则表达式模式
pattern = r'^\d+'

# 遍历每个字符串并处理
for string in strings:
    # 使用正则表达式替换掉前面的数字部分
    result = re.sub(pattern, '', string)
    print(result)

