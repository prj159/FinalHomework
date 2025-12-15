import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 设置中文字体，防止绘图乱码 (尝试使用常见中文字体)
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'PingFang SC'] 
plt.rcParams['axes.unicode_minus'] = False

# 1. 读取数据
df_info = pd.read_excel('商品销售数据.xlsx', sheet_name='信息表')
df_sales = pd.read_excel('商品销售数据.xlsx', sheet_name='销售数据表')

# 2. 数据预处理与合并
# 将销售表和信息表通过 '商品编号' 进行合并
df_merged = pd.merge(df_sales, df_info, on='商品编号', how='left')

# 计算销售金额 = 订单数量 * 商品销售价
df_merged['销售金额'] = df_merged['订单数量'] * df_merged['商品销售价']

# 将订单日期转换为日期格式，并提取月份 (例如 '2022-01')
df_merged['订单日期'] = pd.to_datetime(df_merged['订单日期'])
df_merged['月份'] = df_merged['订单日期'].dt.to_period('M')

# 3. 按照“月份”和“商品大类”对销售金额进行分组汇总
# reset_index() 将结果转回 DataFrame 格式方便查看
monthly_category_sales = df_merged.groupby(['月份', '商品大类'])['销售金额'].sum().reset_index()

# 打印前几行查看汇总结果
print("分组汇总结果前5行：")
print(monthly_category_sales.head())

# 4. 数据透视：为了方便绘制分组柱状图，我们将数据转换一下形状
# 行索引为月份，列名为商品大类，值为销售金额
pivot_data = monthly_category_sales.pivot(index='月份', columns='商品大类', values='销售金额')

# 5. 使用 bar() 绘制柱状图
# 使用 Pandas 内置的 plot(kind='bar') 底层调用的就是 matplotlib 的 bar()
ax = pivot_data.plot(kind='bar', figsize=(12, 6), width=0.8)

plt.title('不同商品大类的月度销售额对比', fontsize=16)
plt.xlabel('月份', fontsize=12)
plt.ylabel('销售金额 (元)', fontsize=12)
plt.xticks(rotation=45) # x轴标签旋转45度以免重叠
plt.legend(title='商品大类')
plt.grid(axis='y', linestyle='--', alpha=0.5)
plt.tight_layout()

# 显示图表
plt.tight_layout()
plt.show()