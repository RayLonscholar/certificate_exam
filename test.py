import re

# 定义要处理的字符串列表
strings = ["52（B）：RayLon", "104 (D)：Ray"]

# 定义正则表达式模式
pattern = r'（(.*?)）'

# 遍历每个字符串并处理
for string in strings:
    # 使用正则表达式匹配括号内的值
    match = re.search(pattern, string)
    if match:
        # 提取括号内的值
        value_inside_brackets = match.group(1)
        print(value_inside_brackets)

