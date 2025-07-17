import camelot
import pandas as pd
import re
import logging
from abc import ABC, abstractmethod
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
import pdfplumber
import numpy as np
from datetime import datetime

class BankParser(ABC):
    """银行解析器抽象基类"""
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.logger = logging.getLogger(self.__class__.__name__)
        self.transactions = []

    @abstractmethod
    def parse(self) -> list[dict]:
        """解析PDF文件并提取交易记录"""
        pass

    def save_to_excel(self, output_path: str):
        """将提取的交易记录保存到Excel文件"""
        if not self.transactions:
            self.logger.warning("没有提取到任何交易记录，无法保存到Excel。")
            return

        df = pd.DataFrame(self.transactions)
        if df.empty:
            self.logger.warning("交易记录为空，无法生成Excel文件。")
            return

        # 按账户类型分组
        grouped = df.groupby('账户类型')

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for account_type, group in grouped:
                # 移除'账户类型'列，因为它在工作表名称中已经体现
                group = group.drop(columns=['账户类型'])
                group.to_excel(writer, sheet_name=account_type, index=False)

                # 格式化工作表
                worksheet = writer.sheets[account_type]
                # 设置表头字体为粗体
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)
                # 自动调整列宽
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value)) for cell in column_cells)
                    worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2

        self.logger.info(f"交易记录已成功保存到 {output_path}")


class HSBCParser(BankParser):
    def __init__(self, file_path: str):
        super().__init__(file_path)
        # 初始化表格容器
        self.hkd_current_tables = []
        self.hkd_savings_tables = []
        self.foreign_savings_tables = []
        
        # 用于跟踪已处理的5列表格数量
        self.five_col_table_count = 0
        
    def parse(self) -> list[dict]:
        """解析汇丰银行账单PDF并提取交易记录"""
        self.logger.info(f"开始解析汇丰银行账单: {self.file_path}")
        
        try:
            # 使用pdfplumber打开PDF文件
            with pdfplumber.open(self.file_path) as pdf:
                # 遍历每一页
                for page_num, page in enumerate(pdf.pages, 1):
                    self.logger.info(f"处理第 {page_num} 页")
                    
                    # 提取当前页的表格
                    tables = page.extract_tables()
                    
                    # 处理每个表格
                    for table_num, table in enumerate(tables, 1):
                        if not table or len(table) <= 1:  # 跳过空表格或只有一行的表格
                            continue
                            
                        # 转换为DataFrame以便处理
                        df = pd.DataFrame(table)
                        
                        # 检查表格的列数（使用第一行判断）
                        num_cols = len(df.columns)
                        
                        # 根据列数和表头内容分类表格
                        if num_cols == 5:
                            # 检查表头是否包含关键词
                            header = df.iloc[0].astype(str).str.lower()
                            has_date = any('日期' in str(col) for col in header)
                            has_balance = any('结余' in str(col) for col in header)
                            
                            if has_date and has_balance:
                                # 是有效的5列表格，判断是港币往来还是港币储蓄
                                if self.five_col_table_count == 0:
                                    self.logger.info(f"找到港币往来表格: 页 {page_num}, 表格 {table_num}")
                                    self.hkd_current_tables.append(df)
                                else:
                                    self.logger.info(f"找到港币储蓄表格: 页 {page_num}, 表格 {table_num}")
                                    self.hkd_savings_tables.append(df)
                                    
                                self.five_col_table_count += 1
                                
                        elif num_cols == 6:
                            # 检查表头是否包含关键词
                            header = df.iloc[0].astype(str).str.lower()
                            has_currency = any('货币' in str(col) for col in header)
                            has_balance = any('结余' in str(col) for col in header)
                            
                            if has_currency and has_balance:
                                self.logger.info(f"找到外币储蓄表格: 页 {page_num}, 表格 {table_num}")
                                self.foreign_savings_tables.append(df)
            
            # 开始数据整理
            self._process_tables()
            
            self.logger.info(f"解析完成，共提取 {len(self.transactions)} 条交易记录")
            
            return self.transactions
            
        except Exception as e:
            self.logger.error(f"解析汇丰银行账单时出错: {str(e)}")
            raise
    
    def _process_tables(self):
        """处理所有提取的表格"""
        # 处理港币往来表格
        hkd_current_df = self._merge_tables(self.hkd_current_tables, '港币往来')
        
        # 处理港币储蓄表格
        hkd_savings_df = self._merge_tables(self.hkd_savings_tables, '港币储蓄')
        
        # 处理外币储蓄表格
        foreign_savings_df = self._merge_tables(self.foreign_savings_tables, '外币储蓄')
        
        # 提取交易记录
        hkd_current_transactions = self._extract_transactions(hkd_current_df, '港币往来', 'HKD')
        hkd_savings_transactions = self._extract_transactions(hkd_savings_df, '港币储蓄', 'HKD')
        foreign_savings_transactions = self._extract_foreign_transactions(foreign_savings_df)
        
        # 合并所有交易记录
        self.transactions = hkd_current_transactions + hkd_savings_transactions + foreign_savings_transactions
    
    def _merge_tables(self, tables, account_type):
        """合并同一账户类型的多个表格"""
        if not tables:
            self.logger.warning(f"没有找到{account_type}表格")
            return pd.DataFrame()
            
        # 合并所有表格
        merged_df = pd.concat(tables, ignore_index=True)
        
        # 设置列名（根据账户类型）
        if account_type in ['港币往来', '港币储蓄']:
            merged_df.columns = ['日期', '进支详情', '存入', '提取', '结余']
        else:  # 外币储蓄
            merged_df.columns = ['货币', '日期', '进支详情', '存入', '提取', '结余']
        
        # 移除重复的表头行（第一行是表头，其他行如果与表头相同则是重复的表头）
        header = merged_df.iloc[0].astype(str)
        merged_df = merged_df[~merged_df.apply(lambda row: row.equals(header), axis=1).shift(fill_value=False)]
        
        # 移除空行（所有值都为空或NaN）
        merged_df = merged_df.dropna(how='all')
        
        # 重置索引
        merged_df = merged_df.reset_index(drop=True)
        
        # 移除第一行（表头）
        merged_df = merged_df.iloc[1:].reset_index(drop=True)
        
        # 替换NaN为空字符串
        merged_df = merged_df.fillna('')
        
        return merged_df
    
    def _extract_transactions(self, df, account_type, currency):
        """从DataFrame中提取交易记录"""
        if df.empty:
            return []
            
        transactions = []
        current_date = None
        
        for _, row in df.iterrows():
            # 提取日期（如果有）
            if row['日期'] and row['日期'] != '':
                try:
                    # 尝试解析日期
                    current_date = self._parse_date(row['日期'])
                except:
                    # 如果解析失败，保持当前日期不变
                    pass
            
            # 跳过没有交易详情的行
            if not row['进支详情'] or row['进支详情'] == '':
                continue
                
            # 创建交易记录
            transaction = {
                '账户类型': account_type,
                '日期': current_date,
                '交易描述': row['进支详情'],
                '货币': currency,
                '存入金额': self._parse_amount(row['存入']),
                '提取金额': self._parse_amount(row['提取']),
                '结余': self._parse_amount(row['结余'])
            }
            
            transactions.append(transaction)
            
        return transactions
    
    def _extract_foreign_transactions(self, df):
        """从外币储蓄表格中提取交易记录"""
        if df.empty:
            return []
            
        transactions = []
        current_currency = None
        current_date = None
        
        for _, row in df.iterrows():
            # 提取货币（如果有）
            if row['货币'] and row['货币'] != '':
                current_currency = row['货币']
                
            # 提取日期（如果有）
            if row['日期'] and row['日期'] != '':
                try:
                    current_date = self._parse_date(row['日期'])
                except:
                    # 如果解析失败，保持当前日期不变
                    pass
            
            # 跳过没有交易详情的行
            if not row['进支详情'] or row['进支详情'] == '':
                continue
                
            # 创建交易记录
            transaction = {
                '账户类型': '外币储蓄',
                '日期': current_date,
                '交易描述': row['进支详情'],
                '货币': current_currency,
                '存入金额': self._parse_amount(row['存入']),
                '提取金额': self._parse_amount(row['提取']),
                '结余': self._parse_amount(row['结余'])
            }
            
            transactions.append(transaction)
            
        return transactions
    
    def _parse_date(self, date_str):
        """解析日期字符串"""
        # 移除非数字和分隔符
        date_str = re.sub(r'[^0-9/\-.]', '', str(date_str))
        
        # 尝试不同的日期格式
        date_formats = ['%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%Y/%m/%d', '%Y-%m-%d', '%Y.%m.%d']
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
            except:
                continue
                
        # 如果所有格式都失败，返回原始字符串
        return date_str
    
    def _parse_amount(self, amount_str):
        """解析金额字符串"""
        if not amount_str or amount_str == '':
            return None
            
        # 移除货币符号、逗号和空格
        amount_str = re.sub(r'[^0-9.\-]', '', str(amount_str))
        
        try:
            return float(amount_str)
        except:
            return None

class ESunBankParser(BankParser):
    def parse(self) -> list[dict]:
        self.logger.warning("玉山银行解析器尚未实现。")
        return []

class GenericParser(BankParser):
    def parse(self) -> list[dict]:
        self.logger.warning("通用解析器尚未实现。")
        return []

BANK_PARSERS = {
    '汇丰银行': HSBCParser,
    '玉山银行': ESunBankParser,
    '其他': GenericParser,
}

def get_available_parsers():
    return list(BANK_PARSERS.keys())

def get_bank_parser(bank_name: str):
    """根据银行名称获取对应的解析器类"""
    return BANK_PARSERS.get(bank_name, GenericParser)