#!/usr/bin/env python3
"""创建测试用的Excel文件"""

import pandas as pd

# 创建第一个sheet的数据
df1 = pd.DataFrame({
    'A': [1, 2, 3, 4],
    'B': ['a', 'b', 'c', 'd'],
    'C': [10.5, 20.5, 30.5, 40.5]
})

# 创建第二个sheet的数据
df2 = pd.DataFrame({
    'X': [10, 20, 30],
    'Y': ['x', 'y', 'z'],
    'Z': [100, 200, 300]
})

# 创建第三个sheet的数据（包含特殊字符的sheet名）
df3 = pd.DataFrame({
    'Name': ['张三', '李四', '王五'],
    'Age': [25, 30, 35],
    'City': ['北京', '上海', '广州']
})

# 保存到Excel文件
with pd.ExcelWriter('test_excel.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Sheet1', index=False)
    df2.to_excel(writer, sheet_name='Sheet2', index=False)
    df3.to_excel(writer, sheet_name='用户信息', index=False)

print('测试Excel文件创建成功: test_excel.xlsx')
print('包含3个sheet: Sheet1, Sheet2, 用户信息') 