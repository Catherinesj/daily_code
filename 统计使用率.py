# -*- coding: utf-8 -*-
"""
耗材借出核心统计（修复日期异常版）
仅统计：借出次数、总借出数量、总借出时长(天)
兼容：20261.19、2626.3.2 等异常日期格式
"""
import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class SimpleConsumableAnalyzer:
    """精简版耗材借出分析器：修复日期异常"""
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.clean_data = None
        self.start_date = None
        self.result = None
        self.today = datetime.now()  # 今日日期，用于未归还时长计算

    def load_clean_data(self):
        """加载并清洗数据，重点修复异常日期"""
        print("正在加载并清洗数据...")
        # 读取Excel并设置列名
        raw_df = pd.read_excel(self.file_path)
        columns = ['序号', '借用日期', '耗材名称', '单位', '数量', '空列', 
                  '计划归还日期', '实际归还日期', '学号', '借用人', '用途', '备注']
        df = raw_df.iloc[3:].copy() if len(raw_df) > 3 else raw_df.copy()
        df.columns = columns
        df = df.reset_index(drop=True)

        # 核心过滤：仅保留有效借出记录
        df = df.dropna(how='all')  # 删空行
        df = df[df['耗材名称'].notna() & (df['耗材名称'] != '')]  # 有耗材名称
        df = df[df['借用日期'].notna()]  # 有借用日期
        df['数量'] = pd.to_numeric(df['数量'], errors='coerce').fillna(1)  # 数量转数值

        # ========== 核心修复：增强版日期解析（兼容异常格式） ==========
        def parse_date_fix(s):
            """
            修复异常日期：
            1. 20261.19 → 2026.1.19
            2. 2626.3.2 → 2026.3.2
            3. 兼容 2026.2 / 2026.3.2 标准格式
            """
            if pd.isna(s):
                return None
            s = str(s).strip()
            
            # 第一步：修复年份异常
            if '.' in s:
                parts = s.split('.')
                # 处理 20261.19 → 拆分为 2026 / 1 / 19
                if len(parts[0]) == 5 and parts[0].startswith('202'):
                    year_part = parts[0][:4]  # 取前4位：2026
                    month_day_part = parts[0][4:] + '.' + '.'.join(parts[1:])  # 1.19
                    return parse_date_fix(month_day_part.replace(year_part, ''))
                # 处理 2626.3.2 → 替换为 2026.3.2
                if parts[0] == '2626':
                    parts[0] = '2026'
                    s = '.'.join(parts)
            
            # 第二步：标准日期解析
            try:
                if '.' in s:
                    parts = s.split('.')
                    if len(parts) == 3:  # YYYY.MM.DD
                        year, month, day = map(int, parts)
                    elif len(parts) == 2:  # YYYY.MM → 补1号
                        year, month = map(int, parts)
                        day = 1
                    else:
                        return None
                    
                    # 年份合法性校验（仅允许2020-2030）
                    if 2020 <= year <= 2030:
                        return datetime(year, month, day)
                    else:
                        return None
            except:
                return None
        
        # 转换借还日期（使用修复后的解析函数）
        df['借用日期_标准'] = df['借用日期'].apply(parse_date_fix)
        df['实际归还日期_标准'] = df['实际归还日期'].apply(parse_date_fix)
        
        # 最终过滤：仅保留借用日期有效的记录（剔除无法修复的异常日期）
        df = df[df['借用日期_标准'].notna()]

        # 计算单条记录的借出时长（天）
        def cal_duration(row):
            borrow = row['借用日期_标准']
            return_date = row['实际归还日期_标准']
            # 已归还：实际归还日期 - 借用日期；未归还：今日 - 借用日期
            end_date = return_date if return_date is not None else self.today
            return max(0, (end_date - borrow).days)
        
        df['单条时长(天)'] = df.apply(cal_duration, axis=1)
        # 标准化耗材名称（统一FPGA不同写法）
        df['耗材名称_标准'] = df['耗材名称'].apply(self._standardize_name)
        self.clean_data = df

        print(f"数据清洗完成！有效借出记录共 {len(self.clean_data)} 条")
        if len(self.clean_data) > 0:
            print(f"数据时间范围：{self.clean_data['借用日期_标准'].min().strftime('%Y-%m-%d')} 至 {self.clean_data['借用日期_标准'].max().strftime('%Y-%m-%d')}")

    def _standardize_name(self, name):
        """标准化耗材名称，避免同设备不同名称"""
        if pd.isna(name):
            return "未知耗材"
        name = str(name).strip()
        # 统一FPGA相关名称
        if "fpga" in name.lower():
            if "(zynq)" in name.lower() or "zynq" in name.lower():
                return "FPGA (ZYNQ)"
            else:
                return "FPGA (AWC-C4)"
        return name

    def set_start_time(self, start_date_str: str = None):
        """设置统计开始时间，格式YYYY-MM-DD"""
        if start_date_str is None:
            while True:
                s = input("\n请输入统计开始时间（格式：YYYY-MM-DD）：").strip()
                try:
                    self.start_date = datetime.strptime(s, '%Y-%m-%d')
                    break
                except:
                    print("日期格式错误！示例：2025-12-01")
        else:
            self.start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        print(f"统计开始时间已设置：{self.start_date.strftime('%Y年%m月%d日')}")

    def analyze(self):
        """核心分析：仅统计借出次数、总借出数量、总借出时长"""
        if self.clean_data is None:
            self.load_clean_data()
        if self.start_date is None:
            self.set_start_time()

        # 筛选指定时间后的记录
        filter_df = self.clean_data[self.clean_data['借用日期_标准'] >= self.start_date].copy()
        if len(filter_df) == 0:
            print(f"\n{self.start_date.strftime('%Y年%m月%d日')} 后无有效借出记录！")
            return

        # 按耗材分组统计核心指标
        stats_df = filter_df.groupby('耗材名称_标准').agg(
            借出次数=('借用日期_标准', 'count'),
            总借出数量=('数量', 'sum'),
            总借出时长_天=('单条时长(天)', 'sum')
        ).reset_index()

        # 排序并添加排名
        stats_df = stats_df.sort_values('借出次数', ascending=False).reset_index(drop=True)
        stats_df.index = stats_df.index + 1  # 排名从1开始
        stats_df.rename_axis('排名', inplace=True)
        # 数值取整
        stats_df = stats_df.astype({'总借出数量': int, '总借出时长_天': int})
        self.result = stats_df

        # 打印结果
        print(f"\n===== 耗材借出核心统计（{self.start_date.strftime('%Y年%m月%d日')} 至今）=====")
        print(stats_df)
        # 整体汇总
        total_count = stats_df['借出次数'].sum()
        total_num = stats_df['总借出数量'].sum()
        total_dur = stats_df['总借出时长_天'].sum()
        print(f"\n整体汇总：总借出次数 {total_count} 次 | 总借出数量 {total_num} 个 | 总借出时长 {total_dur} 天")

    def save_excel(self, save_path: str = "耗材借出核心统计.xlsx"):
        """导出Excel报告"""
        if self.result is None:
            print("请先执行分析（analyze()），再导出！")
            return
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            self.result.to_excel(writer, sheet_name='耗材核心统计')
            # 整体汇总表
            total_df = pd.DataFrame({
                '统计指标': ['统计开始时间', '总借出次数', '总借出数量', '总借出时长(天)', '统计生成时间'],
                '数值': [
                    self.start_date.strftime('%Y-%m-%d'),
                    self.result['借出次数'].sum(),
                    self.result['总借出数量'].sum(),
                    self.result['总借出时长_天'].sum(),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            })
            total_df.to_excel(writer, sheet_name='整体汇总', index=False)
        print(f"\nExcel报告已导出：{save_path}")

# ==================== 运行配置（仅需修改这2处）====================
if __name__ == "__main__":
    # 1. 你的Excel文件路径
    EXCEL_FILE_PATH = r"D:\Desktop\1.xlsx"
    # 2. 统计开始时间（None则手动输入，也可直接指定如'2026-01-01'）
    START_DATE = None  # 示例：START_DATE = '2026-01-01'

    # 执行统计
    analyzer = SimpleConsumableAnalyzer(EXCEL_FILE_PATH)
    analyzer.load_clean_data()
    analyzer.set_start_time(START_DATE)
    analyzer.analyze()
    # 导出Excel（修改保存路径）
    if analyzer.result is not None:
        analyzer.save_excel(save_path=r"D:\Desktop\耗材借出核心统计.xlsx")