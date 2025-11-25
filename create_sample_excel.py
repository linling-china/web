import pandas as pd

# 创建示例数据
data = {
    '用户姓名': ['张三', '李四', '王五', '赵六', '钱七'],
    '账号': ['ZH001', 'ZH002', 'ZH003', 'ZH004', 'ZH005'],
    '资产编号': ['ZC001', 'ZC002', 'ZC003', 'ZC004', 'ZC005'],
    '计算机名': ['PC001', 'PC002', 'PC003', 'PC004', 'PC005'],
    '联系电话': ['13800138001', '13800138002', '13800138003', '13800138004', '13800138005'],
    '所在部门': ['技术部', '销售部', '人事部', '财务部', '市场部']
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存为Excel文件
df.to_excel('/workspace/sample_accounts.xlsx', index=False, engine='openpyxl')

print("示例Excel文件已创建: sample_accounts.xlsx")