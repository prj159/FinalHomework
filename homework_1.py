import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# 1. 读取数据
# 改用 read_excel，并指定文件名和工作表名
df_info = pd.read_excel('商品销售数据.xlsx', sheet_name='信息表')
df_sales = pd.read_excel('商品销售数据.xlsx', sheet_name='销售数据表')

# 2. 数据处理：合并表以获取价格，并计算销售金额
# 将两个表按'商品编号'合并
df_merged = pd.merge(df_sales, df_info[['商品编号', '商品销售价']], on='商品编号', how='left')
# 计算每单的销售金额
df_merged['销售金额'] = df_merged['订单数量'] * df_merged['商品销售价']

# 3. 提取“月份”作为X轴数据
# 将订单日期转换为datetime格式
df_merged['订单日期'] = pd.to_datetime(df_merged['订单日期'])
# 提取月份（格式为 YYYY-MM）
df_merged['月份'] = df_merged['订单日期'].dt.strftime('%Y-%m')

# 4. 按月分组求和，获取Y轴数据
monthly_sales = df_merged.groupby('月份')['销售金额'].sum().reset_index()

# 准备绘图数据
x = monthly_sales['月份']
y = monthly_sales['销售金额']

# 5. 使用plot()绘制折线图
plt.figure(figsize=(10, 6)) # 设置画布大小

# 设置中文字体：Mac 用户请使用 Arial Unicode MS
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS'] 
plt.rcParams['axes.unicode_minus'] = False # 解决负号显示问题

plt.plot(x, y, marker='o', linestyle='-', color='#1f77b4', linewidth=2, label='月度销售额')

# 6. 设置图表和x轴、y轴标题
plt.title('商品每月销售金额变化趋势', fontsize=16)
plt.xlabel('月份', fontsize=12)
plt.ylabel('销售金额 (元)', fontsize=12)

# 7. 优化图表显示（修改了这里）

# --- 修改点 A：设置 X 轴 ---
plt.xticks(rotation=45) # 保持日期倾斜

# --- 修改点 B：解决“过于偏上”的问题 ---
# 删除了原来强制从0开始且步长为50万的代码
# 让 Matplotlib 自动根据数据范围调整，或者手动设置一个更紧凑的范围
# 下面这行代码会让 Y 轴下限变成“最低销售额的 0.9 倍”，这样折线就会居中显示了
plt.ylim(bottom=y.min() * 0.9, top=y.max() * 1.1)

# 格式化 Y 轴：不使用科学计数法，显示完整数字
plt.gca().yaxis.set_major_formatter(ticker.FormatStrFormatter('%d'))

# --- 修改点 C：给每个点添加数据标签 ---
for a, b in zip(x, y):
    # a 是月份，b 是金额
    # %.0f 表示不保留小数
    # ha='center' 横向居中, va='bottom' 纵向在点上方
    plt.text(a, b, '%.0f' % b, ha='center', va='bottom', fontsize=10, color='black')

# 添加网格
plt.grid(True, linestyle='--', alpha=0.5)

# 显示图表
plt.tight_layout()
plt.show()

# 打印分析结果以便查看数值
print(monthly_sales)