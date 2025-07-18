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

    def print_transactions(self):
        """打印提取的交易记录，用于调试"""
        if not self.transactions:
            self.logger.warning("没有提取到任何交易记录，无法打印。")
            return
            
        self.logger.info(f"共提取到 {len(self.transactions)} 条交易记录：")
        for i, transaction in enumerate(self.transactions, 1):
            self.logger.info(f"===== 交易记录 {i} =====")
            for key, value in transaction.items():
                self.logger.info(f"{key}: {value}")
                
        # 检查可能的问题
        self.logger.info("===== 交易记录检查 =====")
        # 检查日期格式
        date_issues = [i for i, t in enumerate(self.transactions, 1) 
                      if not isinstance(t.get('日期'), str) or not re.match(r'\d{4}-\d{2}-\d{2}', str(t.get('日期', '')))]  
        if date_issues:
            self.logger.warning(f"发现 {len(date_issues)} 条记录的日期格式不正确: {date_issues}")
            
        # 检查金额
        amount_issues = [i for i, t in enumerate(self.transactions, 1) 
                        if (t.get('存入金额') is None and t.get('提取金额') is None)]  
        if amount_issues:
            self.logger.warning(f"发现 {len(amount_issues)} 条记录没有金额信息: {amount_issues}")
            
        # 检查描述
        desc_issues = [i for i, t in enumerate(self.transactions, 1) 
                      if not t.get('交易描述')]  
        if desc_issues:
            self.logger.warning(f"发现 {len(desc_issues)} 条记录没有交易描述: {desc_issues}")

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
                    
                    # 提取页面文本用于调试
                    page_text = page.extract_text()
                    self.logger.info(f"页面 {page_num} 文本预览: {page_text[:200]}...")
                    
                    # 提取当前页的表格
                    tables = page.extract_tables()
                    self.logger.info(f"页面 {page_num} 找到 {len(tables)} 个表格")
                    
                    # 如果pdfplumber无法提取表格，尝试使用文本分析方法提取
                    if not tables:
                        self.logger.info(f"尝试从页面 {page_num} 文本中提取表格")
                        text_tables = self._extract_tables_from_text(page_text)
                        if text_tables:
                            self.logger.info(f"从页面 {page_num} 文本中成功提取 {len(text_tables)} 个表格")
                            # 将文本提取的表格转换为DataFrame格式
                            for table_num, table in enumerate(text_tables, 1):
                                df = pd.DataFrame(table)
                                # 记录表格信息用于调试
                                self.logger.info(f"页面 {page_num} 文本表格 {table_num} 信息: {len(df)}行, {len(df.columns)}列")
                                if not df.empty:
                                    self.logger.info(f"文本表格 {table_num} 第一行内容: {df.iloc[0].tolist()}")
                                
                                # 检查表格的列数
                                num_cols = len(df.columns)
                                
                                # 根据列数和内容特征分类表格
                                header = df.iloc[0].astype(str).str.lower()
                                header_str = ' '.join([str(x).lower() for x in header])
                                self.logger.info(f"文本表格 {table_num} 表头: {header_str}")
                                
                                # 更宽松的条件：只要表头中包含一些银行账单常见词即可
                                bank_keywords = ['日期', '详情', '金额', '余额', '结余', '存入', '提取', 'date', 'details', 'amount', 'balance']
                                has_bank_keyword = any(keyword in header_str for keyword in bank_keywords)
                                
                                # 根据列数和表头内容正确分类表格
                                if num_cols == 6 or '货币' in header_str:
                                    self.logger.info(f"找到外币储蓄表格: 页 {page_num}, 文本表格 {table_num}")
                                    self.foreign_savings_tables.append(df)
                                elif num_cols == 5 and (has_bank_keyword or self._looks_like_transaction_table(df)):
                                    if self.five_col_table_count == 0:
                                        self.logger.info(f"找到港币往来表格: 页 {page_num}, 文本表格 {table_num}")
                                        self.hkd_current_tables.append(df)
                                    else:
                                        self.logger.info(f"找到港币储蓄表格: 页 {page_num}, 文本表格 {table_num}")
                                        self.hkd_savings_tables.append(df)
                                        
                                    self.five_col_table_count += 1
                            continue  # 已处理文本提取的表格，继续下一页
                    
                    # 处理每个表格
                    for table_num, table in enumerate(tables, 1):
                        if not table or len(table) <= 1:  # 跳过空表格或只有一行的表格
                            self.logger.info(f"跳过页面 {page_num} 表格 {table_num}: 表格为空或只有一行")
                            continue
                            
                        # 转换为DataFrame以便处理
                        df = pd.DataFrame(table)
                        
                        # 记录表格信息用于调试
                        self.logger.info(f"页面 {page_num} 表格 {table_num} 信息: {len(df)}行, {len(df.columns)}列")
                        if not df.empty:
                            self.logger.info(f"表格 {table_num} 第一行内容: {df.iloc[0].tolist()}")
                        
                        # 检查表格的列数（使用第一行判断）
                        num_cols = len(df.columns)
                        
                        # 放宽表格识别条件
                        # 方法1: 根据列数直接分类
                        if num_cols == 5:
                            # 尝试检查表头是否包含关键词，但放宽条件
                            header = df.iloc[0].astype(str).str.lower()
                            header_str = ' '.join([str(x).lower() for x in header])
                            self.logger.info(f"表格 {table_num} 表头: {header_str}")
                            
                            # 更宽松的条件：只要表头中包含一些银行账单常见词即可
                            bank_keywords = ['日期', '详情', '金额', '余额', '结余', '存入', '提取', 'date', 'details', 'amount', 'balance']
                            has_bank_keyword = any(keyword in header_str for keyword in bank_keywords)
                            
                            if has_bank_keyword or self._looks_like_transaction_table(df):
                                if self.five_col_table_count == 0:
                                    self.logger.info(f"找到港币往来表格: 页 {page_num}, 表格 {table_num}")
                                    self.hkd_current_tables.append(df)
                                else:
                                    self.logger.info(f"找到港币储蓄表格: 页 {page_num}, 表格 {table_num}")
                                    self.hkd_savings_tables.append(df)
                                    
                                self.five_col_table_count += 1
                            else:
                                self.logger.info(f"表格 {table_num} 不符合港币账户表格特征")
                                
                        elif num_cols == 6:
                            # 尝试检查表头是否包含关键词，但放宽条件
                            header = df.iloc[0].astype(str).str.lower()
                            header_str = ' '.join([str(x).lower() for x in header])
                            self.logger.info(f"表格 {table_num} 表头: {header_str}")
                            
                            # 更宽松的条件：只要表头中包含一些外币账单常见词即可
                            foreign_keywords = ['货币', '日期', '详情', '金额', '余额', '结余', '存入', '提取', 'currency', 'date', 'details', 'amount', 'balance']
                            has_foreign_keyword = any(keyword in header_str for keyword in foreign_keywords)
                            
                            if has_foreign_keyword or self._looks_like_transaction_table(df):
                                self.logger.info(f"找到外币储蓄表格: 页 {page_num}, 表格 {table_num}")
                                self.foreign_savings_tables.append(df)
                            else:
                                self.logger.info(f"表格 {table_num} 不符合外币储蓄表格特征")
                        else:
                            # 尝试分析其他列数的表格
                            self.logger.info(f"发现 {num_cols} 列表格，尝试分析")
                            if self._looks_like_transaction_table(df):
                                self.logger.info(f"表格 {table_num} 看起来像交易表格，尝试处理")
                                # 根据内容特征判断表格类型
                                if self._contains_currency_column(df):
                                    self.logger.info(f"表格 {table_num} 包含货币列，可能是外币储蓄表格")
                                    self.foreign_savings_tables.append(df)
                                else:
                                    if self.five_col_table_count == 0:
                                        self.logger.info(f"表格 {table_num} 可能是港币往来表格")
                                        self.hkd_current_tables.append(df)
                                    else:
                                        self.logger.info(f"表格 {table_num} 可能是港币储蓄表格")
                                        self.hkd_savings_tables.append(df)
                                    self.five_col_table_count += 1
        
            # 开始数据整理
            self._process_tables()
            
            self.logger.info(f"解析完成，共提取 {len(self.transactions)} 条交易记录")
            
            # 打印交易记录用于调试
            self.print_transactions()
            
            return self.transactions
            
        except Exception as e:
            self.logger.error(f"解析汇丰银行账单时出错: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            raise
    
    def _looks_like_transaction_table(self, df):
        """判断表格是否看起来像交易记录表格"""
        if df.empty or len(df) <= 1:
            return False
            
        # 检查是否包含日期格式的数据
        date_pattern = r'\d{1,2}[/\-.年]\d{1,2}[/\-.月]\d{2,4}|\d{2,4}[/\-.年]\d{1,2}[/\-.月]\d{1,2}'
        has_date = False
        
        # 检查是否包含金额格式的数据
        amount_pattern = r'\d{1,3}(,\d{3})*(\.[\d]{2})?'
        has_amount = False
        
        # 检查前5行数据
        for i in range(min(5, len(df))):
            row_str = ' '.join([str(x) for x in df.iloc[i].tolist()])
            if re.search(date_pattern, row_str):
                has_date = True
            if re.search(amount_pattern, row_str):
                has_amount = True
                
        return has_date and has_amount
    
    def _contains_currency_column(self, df):
        """检查表格是否包含货币列"""
        if df.empty:
            return False
            
        # 常见货币代码
        currency_codes = ['USD', 'EUR', 'GBP', 'JPY', 'CNY', 'AUD', 'CAD', 'CHF', 'HKD']
        
        # 检查前5行的第一列是否包含货币代码
        for i in range(min(5, len(df))):
            if df.iloc[i, 0] in currency_codes:
                return True
                
        return False
    
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
            
        # 检查表格列数是否一致，并根据列数重新分类
        filtered_tables = []
        expected_cols = 6 if account_type == '外币储蓄' else 5
        
        for i, df in enumerate(tables):
            if len(df.columns) == expected_cols:
                filtered_tables.append(df)
            else:
                self.logger.warning(f"跳过一个列数不匹配的{account_type}表格：预期{expected_cols}列，实际{len(df.columns)}列")
        
        if not filtered_tables:
            self.logger.warning(f"筛选后没有符合条件的{account_type}表格")
            return pd.DataFrame()
        
        # 合并所有表格
        merged_df = pd.concat(filtered_tables, ignore_index=True)
        
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

    def _extract_tables_from_text(self, text):
        """从文本中提取表格结构"""
        tables = []
        
        # 分析页面内容，查找可能的表格区域
        lines = text.split('\n')
        
        # 查找港币往来表格
        hkd_current_table = self._find_table_in_text(lines, ['日期', '进支详情', '存入', '提取', '结余'], 5)
        if hkd_current_table:
            self.logger.info("从文本中找到港币往来表格")
            tables.append(hkd_current_table)
        
        # 查找港币储蓄表格
        hkd_savings_table = self._find_table_in_text(lines, ['日期', '进支详情', '存入', '提取', '结余'], 5, 
                                                  skip_if_found=len(tables) > 0)
        if hkd_savings_table:
            self.logger.info("从文本中找到港币储蓄表格")
            tables.append(hkd_savings_table)
        
        # 查找外币储蓄表格
        foreign_savings_table = self._find_table_in_text(lines, ['货币', '日期', '进支详情', '存入', '提取', '结余'], 6)
        if foreign_savings_table:
            self.logger.info("从文本中找到外币储蓄表格")
            tables.append(foreign_savings_table)
        
        return tables
    
    def _find_table_in_text(self, lines, header_keywords, expected_cols, skip_if_found=False):
        """在文本行中查找特定表格"""
        if skip_if_found:
            return None
            
        # 查找可能的表头行
        header_line_idx = -1
        for i, line in enumerate(lines):
            # 检查这一行是否包含所有关键词
            line_lower = line.lower()
            if all(keyword.lower() in line_lower for keyword in header_keywords):
                header_line_idx = i
                break
        
        if header_line_idx == -1:
            # 尝试更宽松的匹配：至少包含一半以上的关键词
            min_keywords = len(header_keywords) // 2 + 1
            for i, line in enumerate(lines):
                line_lower = line.lower()
                matched_keywords = sum(1 for keyword in header_keywords if keyword.lower() in line_lower)
                if matched_keywords >= min_keywords:
                    header_line_idx = i
                    break
        
        if header_line_idx == -1:
            return None
        
        # 找到表头后，提取表格内容
        table_rows = []
        
        # 添加表头
        header_row = []
        for keyword in header_keywords:
            header_row.append(keyword)
        table_rows.append(header_row)
        
        # 从表头下一行开始提取数据行
        i = header_line_idx + 1
        while i < len(lines):
            line = lines[i].strip()
            if not line:  # 跳过空行
                i += 1
                continue
                
            # 检查是否已经到达表格末尾（通常是遇到了另一个表格的表头或页脚）
            if any(keyword.lower() in line.lower() for keyword in ['页', 'page', '总计', 'total']):
                break
                
            # 尝试将行拆分为单元格
            cells = self._split_line_into_cells(line, expected_cols)
            if cells and len(cells) >= expected_cols - 1:  # 允许少一列的情况
                # 确保单元格数量正确
                while len(cells) < expected_cols:
                    cells.append('')  # 补充缺少的列
                table_rows.append(cells[:expected_cols])  # 只取需要的列数
            
            i += 1
        
        return table_rows if len(table_rows) > 1 else None
    
    def _split_line_into_cells(self, line, expected_cols):
        """将文本行拆分为表格单元格"""
        # 方法1：尝试使用空格拆分
        cells = line.split()
        
        # 如果拆分后的单元格数量接近预期列数，可能是有效的数据行
        if len(cells) >= expected_cols - 1 and len(cells) <= expected_cols + 2:
            # 尝试合并可能被错误拆分的单元格
            return self._merge_cells_if_needed(cells, expected_cols)
        
        # 方法2：尝试使用固定宽度拆分
        # 假设每列的宽度大约是行长度除以预期列数
        if len(line) > expected_cols * 3:  # 确保行足够长
            col_width = len(line) // expected_cols
            cells = []
            for i in range(expected_cols):
                start = i * col_width
                end = start + col_width if i < expected_cols - 1 else len(line)
                cell = line[start:end].strip()
                cells.append(cell)
            return cells
        
        # 方法3：使用正则表达式查找日期、金额等模式
        import re
        
        # 查找日期模式
        date_match = re.search(r'\d{1,2}[/\-.年]\d{1,2}[/\-.月]\d{2,4}|\d{2,4}[/\-.年]\d{1,2}[/\-.月]\d{1,2}', line)
        if date_match:
            date_str = date_match.group(0)
            # 将行分为日期前、日期和日期后三部分
            parts = [line[:date_match.start()].strip(), date_str, line[date_match.end():].strip()]
            
            # 进一步处理日期后的部分，查找金额
            rest = parts[2]
            amount_matches = re.findall(r'\d{1,3}(,\d{3})*(\.[0-9]{2})?', rest)
            
            if amount_matches and len(amount_matches) >= 1:
                # 根据找到的金额进一步拆分
                cells = [parts[0], parts[1]]  # 日期前的部分和日期
                
                # 提取交易描述（金额之前的部分）
                desc_end = rest.find(amount_matches[0][0]) if amount_matches[0][0] in rest else len(rest)
                cells.append(rest[:desc_end].strip())
                
                # 添加金额列
                for i, match in enumerate(amount_matches):
                    if i < expected_cols - 3:  # 确保不超过预期列数
                        cells.append(match[0])
                
                return cells
        
        return None
    
    def _merge_cells_if_needed(self, cells, expected_cols):
        """合并可能被错误拆分的单元格"""
        if len(cells) == expected_cols:
            return cells
            
        # 如果单元格数量多于预期，尝试合并中间的描述单元格
        if len(cells) > expected_cols:
            # 假设前1列是日期，后3列是金额相关
            date_cols = 1
            amount_cols = 3
            
            # 计算需要合并的单元格数量
            merge_count = len(cells) - expected_cols
            
            # 合并中间的描述单元格
            merged_cells = cells[:date_cols]
            merged_description = ' '.join(cells[date_cols:date_cols+merge_count+1])
            merged_cells.append(merged_description)
            merged_cells.extend(cells[date_cols+merge_count+1:])
            
            return merged_cells
            
        # 如果单元格数量少于预期，添加空单元格
        if len(cells) < expected_cols:
            while len(cells) < expected_cols:
                cells.append('')
            return cells
            
        return cells

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