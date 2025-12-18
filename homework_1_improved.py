"""
改进版的商品销售趋势分析代码
主要改进点：
1. 添加异常处理
2. 模块化设计
3. 更好的错误信息
4. 自动保存图表
5. 数据验证和清理
6. 更灵活的配置
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import warnings
from typing import Tuple, Optional
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 忽略警告
warnings.filterwarnings('ignore')


class SalesDataAnalyzer:
    """销售数据分析器类"""
    
    def __init__(self, excel_file: str = '商品销售数据.xlsx'):
        """
        初始化分析器
        
        Args:
            excel_file: Excel数据文件路径
        """
        self.excel_file = excel_file
        self.df_info = None
        self.df_sales = None
        self.df_merged = None
        self.monthly_sales = None
        
    def load_data(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        加载Excel数据
        
        Returns:
            信息表和销售表的DataFrame
            
        Raises:
            FileNotFoundError: 文件不存在时
            ValueError: 工作表不存在或格式错误时
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(self.excel_file):
                raise FileNotFoundError(f"数据文件 '{self.excel_file}' 不存在")
            
            logger.info(f"正在加载数据文件: {self.excel_file}")
            
            # 读取Excel文件
            self.df_info = pd.read_excel(self.excel_file, sheet_name='信息表')
            self.df_sales = pd.read_excel(self.excel_file, sheet_name='销售数据表')
            
            logger.info(f"信息表加载完成: {len(self.df_info)} 行, {len(self.df_info.columns)} 列")
            logger.info(f"销售表加载完成: {len(self.df_sales)} 行, {len(self.df_sales.columns)} 列")
            
            return self.df_info, self.df_sales
            
        except Exception as e:
            logger.error(f"加载数据失败: {e}")
            raise
    
    def validate_and_clean_data(self) -> None:
        """验证和清理数据"""
        if self.df_info is None or self.df_sales is None:
            raise ValueError("请先加载数据")
        
        # 检查缺失值
        info_missing = self.df_info.isnull().sum().sum()
        sales_missing = self.df_sales.isnull().sum().sum()
        
        if info_missing > 0:
            logger.warning(f"信息表有 {info_missing} 个缺失值")
        if sales_missing > 0:
            logger.warning(f"销售表有 {sales_missing} 个缺失值")
        
        # 检查关键列是否存在
        required_info_cols = ['商品编号', '商品销售价']
        required_sales_cols = ['订单日期', '商品编号', '订单数量']
        
        for col in required_info_cols:
            if col not in self.df_info.columns:
                raise ValueError(f"信息表缺少必要列: {col}")
        
        for col in required_sales_cols:
            if col not in self.df_sales.columns:
                raise ValueError(f"销售表缺少必要列: {col}")
        
        # 确保价格是数字类型
        if not pd.api.types.is_numeric_dtype(self.df_info['商品销售价']):
            self.df_info['商品销售价'] = pd.to_numeric(self.df_info['商品销售价'], errors='coerce')
            logger.info("已将'商品销售价'转换为数值类型")
    
    def process_data(self) -> pd.DataFrame:
        """
        处理数据：合并、计算金额、提取月份
        
        Returns:
            处理后的合并DataFrame
        """
        if self.df_info is None or self.df_sales is None:
            raise ValueError("请先加载数据")
        
        # 合并数据
        self.df_merged = pd.merge(
            self.df_sales, 
            self.df_info[['商品编号', '商品销售价']], 
            on='商品编号', 
            how='left'
        )
        
        # 检查是否有商品找不到价格
        missing_prices = self.df_merged['商品销售价'].isnull().sum()
        if missing_prices > 0:
            logger.warning(f"有 {missing_prices} 条记录找不到对应商品价格")
            # 用平均价格填充缺失值
            avg_price = self.df_info['商品销售价'].mean()
            self.df_merged['商品销售价'] = self.df_merged['商品销售价'].fillna(avg_price)
        
        # 计算销售金额
        self.df_merged['销售金额'] = self.df_merged['订单数量'] * self.df_merged['商品销售价']
        
        # 确保订单日期是datetime类型
        if not pd.api.types.is_datetime64_any_dtype(self.df_merged['订单日期']):
            self.df_merged['订单日期'] = pd.to_datetime(self.df_merged['订单日期'])
        
        # 提取月份
        self.df_merged['月份'] = self.df_merged['订单日期'].dt.strftime('%Y-%m')
        
        # 按月分组求和
        self.monthly_sales = self.df_merged.groupby('月份')['销售金额'].sum().reset_index()
        self.monthly_sales = self.monthly_sales.sort_values('月份')
        
        logger.info(f"数据合并完成: {len(self.df_merged)} 行")
        logger.info(f"按月汇总完成: {len(self.monthly_sales)} 个月份")
        
        return self.df_merged
    
    def create_sales_trend_chart(
        self, 
        save_path: Optional[str] = None,
        show_chart: bool = True
    ) -> plt.Figure:
        """
        创建销售趋势图表
        
        Args:
            save_path: 图表保存路径，如为None则不保存
            show_chart: 是否显示图表
            
        Returns:
            matplotlib图表对象
        """
        if self.monthly_sales is None:
            raise ValueError("请先处理数据")
        
        # 设置中文字体
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 准备数据
        x = self.monthly_sales['月份']
        y = self.monthly_sales['销售金额']
        
        # 创建图表
        fig, ax = plt.subplots(figsize=(12, 7))
        
        # 绘制折线图
        ax.plot(x, y, marker='o', linestyle='-', color='#1f77b4', 
                linewidth=2, markersize=8, label='月度销售额')
        
        # 设置标题和标签
        ax.set_title('商品每月销售金额变化趋势', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('月份', fontsize=12)
        ax.set_ylabel('销售金额 (元)', fontsize=12)
        
        # 设置X轴刻度
        ax.set_xticks(range(len(x)))
        ax.set_xticklabels(x, rotation=45, ha='right')
        
        # 设置Y轴格式
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        
        # 自适应Y轴范围
        y_min, y_max = y.min(), y.max()
        y_padding = (y_max - y_min) * 0.1
        ax.set_ylim(y_min - y_padding, y_max + y_padding)
        
        # 在每个点上添加数据标签
        for i, (month, amount) in enumerate(zip(x, y)):
            label = f'{amount:,.0f}'
            ax.text(i, amount, label, 
                   ha='center', va='bottom', 
                   fontsize=10, color='black',
                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8))
        
        # 添加总计信息
        total_sales = y.sum()
        avg_monthly = y.mean()
        ax.text(0.13, 0.98, 
                f'总销售额: ¥{total_sales:,.0f}\n月均销售额: ¥{avg_monthly:,.0f}',
                transform=ax.transAxes,
                fontsize=11,
                verticalalignment='top',
                bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', alpha=0.8))
        
        # 添加网格
        ax.grid(True, linestyle='--', alpha=0.3, axis='y')
        
        # 添加图例
        ax.legend(loc='upper left')
        
        # 调整布局
        plt.tight_layout()
        
        # 保存图表
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            logger.info(f"图表已保存至: {save_path}")
        
        # 展示图表
        if show_chart:
            plt.show()
        
        return fig
    
    def generate_report(self) -> dict:
        """
        生成数据分析报告
        
        Returns:
            包含关键指标的字典
        """
        if self.df_merged is None or self.monthly_sales is None:
            raise ValueError("请先处理数据")
        
        report = {
            'basic_info': {
                'total_orders': len(self.df_merged),
                'total_months': len(self.monthly_sales),
                'total_sales': self.monthly_sales['销售金额'].sum(),
                'avg_monthly_sales': self.monthly_sales['销售金额'].mean(),
                'max_monthly_sales': self.monthly_sales['销售金额'].max(),
                'min_monthly_sales': self.monthly_sales['销售金额'].min(),
                'best_month': self.monthly_sales.loc[self.monthly_sales['销售金额'].idxmax(), '月份'],
                'worst_month': self.monthly_sales.loc[self.monthly_sales['销售金额'].idxmin(), '月份']
            },
            'growth_info': {
                'sales_growth_rate': self.calculate_growth_rate(),
                'monthly_details': self.monthly_sales.to_dict('records')
            }
        }
        
        return report
    
    def calculate_growth_rate(self) -> float:
        """计算总增长率"""
        if len(self.monthly_sales) < 2:
            return 0
        
        first_month = self.monthly_sales['销售金额'].iloc[0]
        last_month = self.monthly_sales['销售金额'].iloc[-1]
        
        if first_month == 0:
            return 0
        
        return (last_month - first_month) / first_month * 100
    
    def print_report(self) -> None:
        """打印分析报告"""
        report = self.generate_report()
        
        print("\n" + "="*60)
        print("                   销售数据分析报告")
        print("="*60)
        
        print(f"\n📊 基础统计:")
        print(f"   总订单数: {report['basic_info']['total_orders']:,} 单")
        print(f"   统计月份: {report['basic_info']['total_months']} 个月")
        print(f"   总销售额: ¥{report['basic_info']['total_sales']:,.0f}")
        print(f"   月均销售额: ¥{report['basic_info']['avg_monthly_sales']:,.0f}")
        
        print(f"\n📈 月度表现:")
        print(f"   最高月销售额: ¥{report['basic_info']['max_monthly_sales']:,.0f} ({report['basic_info']['best_month']})")
        print(f"   最低月销售额: ¥{report['basic_info']['min_monthly_sales']:,.0f} ({report['basic_info']['worst_month']})")
        
        growth_rate = report['growth_info']['sales_growth_rate']
        print(f"\n📈 增长率:")
        print(f"   总增长率: {growth_rate:+.1f}%")
        
        print(f"\n📅 月度详细数据:")
        print(self.monthly_sales.to_string(index=False))
        
        print("\n" + "="*60)


def main():
    """主函数"""
    try:
        # 创建分析器实例
        analyzer = SalesDataAnalyzer('商品销售数据.xlsx')
        
        # 1. 加载数据
        df_info, df_sales = analyzer.load_data()
        
        # 2. 验证数据
        analyzer.validate_and_clean_data()
        
        # 3. 处理数据
        analyzer.process_data()
        
        # 4. 打印报告
        analyzer.print_report()
        
        # 5. 创建图表
        analyzer.create_sales_trend_chart(
            save_path='sales_trend_chart.png',
            show_chart=True
        )
        
        logger.info("分析完成！")
        
    except Exception as e:
        logger.error(f"程序运行失败: {e}")
        print(f"\n❌ 错误: {e}")
        print("请检查数据文件或联系开发人员。")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
