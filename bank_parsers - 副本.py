import camelot
import pandas as pd
import re
import logging
from abc import ABC, abstractmethod
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

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
    """汇丰银行解析器"""
    HKD_CURRENT_KEYWORDS = ['HKD Current', '港元往来', '港币往来']
    HKD_SAVINGS_KEYWORDS = ['HKD Savings', '港元储蓄', '港币储蓄']
    FOREIGN_SAVINGS_KEYWORDS = ['Foreign Currency Savings', '外币储蓄']
    CURRENCY_CODES = ['USD', 'EUR', 'GBP', 'AUD', 'CAD', 'JPY', 'CHF', 'NZD', 'SGD']

    def __init__(self, file_path: str):
        super().__init__(file_path)
        self.date_pattern = re.compile(r'\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))\b')
        self.amount_pattern = re.compile(r'([\d,]+\.\d{2})')

    def should_filter_transaction(self, description: str) -> bool:
        """判断是否应过滤掉特定交易描述"""
        filter_keywords = [
            'Opening Balance', 'Closing Balance',
            '账户结余', '期初余额', '期末余额',
            'Total No. of Deposits', '存入次数总计', 'Total No. of Withdrawals', '提取次数总计',
            'Total Deposit Amount', '存入总额', 'Total Withdrawal Amount', '提取总额'
        ]
        return any(keyword.lower() in description.lower() for keyword in filter_keywords)


    def _extract_currency_from_text(self, text: str) -> str | None:
        """从文本中提取货币代码"""
        for code in self.CURRENCY_CODES:
            if code in text:
                return code
        if 'HKD' in text or '港币' in text:
            return 'HKD'
        return None

    def find_account_sections_on_page(self, tables: list, page_num: int) -> dict[str, list[tuple[int, int]]]:
        """在单个页面中查找不同账户类型所在的表格和行范围"""
        account_sections = {'港币往来': [], '港币储蓄': [], '外币储蓄': []}
        for i, table_df in enumerate(tables):
            current_account_type = None
            start_row = -1
            last_foreign_start = -1

            for row_index, row in table_df.iterrows():
                row_text = ' '.join(map(str, row.tolist()))
                new_account_type = None
                if any(k.lower() in row_text.lower() for k in self.HKD_CURRENT_KEYWORDS):
                    new_account_type = '港币往来'
                elif any(k.lower() in row_text.lower() for k in self.HKD_SAVINGS_KEYWORDS):
                    new_account_type = '港币储蓄'
                elif any(k.lower() in row_text.lower() for k in self.FOREIGN_SAVINGS_KEYWORDS) or any(c in row_text for c in self.CURRENCY_CODES):
                    new_account_type = '外币储蓄'

                if new_account_type:
                    if current_account_type and start_row != -1:
                        if current_account_type == '外币储蓄' and new_account_type == '外币储蓄':
                            # 合并连续的外币储蓄部分
                            continue
                        else:
                            end_row = row_index - 1
                            account_sections[current_account_type].append((i, start_row, end_row))
                            self.logger.info(f"在第{page_num}页的表格{i+1}中找到 {current_account_type} 部分 (行 {start_row}-{end_row})")
                    current_account_type = new_account_type
                    start_row = row_index
                    if current_account_type == '外币储蓄':
                        last_foreign_start = row_index

            if current_account_type and start_row != -1:
                account_sections[current_account_type].append((i, start_row, len(table_df) - 1))
                self.logger.info(f"在第{page_num}页的表格{i+1}中找到 {current_account_type} 部分 (行 {start_row}-{len(table_df)-1})")

        return account_sections

    def extract_account_transactions(self, tables: list, page_num: int, account_type: str, sections: list[tuple[int, int, int]]) -> list[dict]:
        """从指定账户部分的表格中提取交易记录"""
        page_transactions = []
        self.logger.info(f"--- 解析 '{account_type}' 部分 ({len(sections)}个) ---")

        for table_index, start_row, end_row in sections:
            table_df = tables[table_index]
            account_df = table_df.iloc[start_row:end_row + 1].copy()
            transactions = self.extract_transactions_from_table(account_df, page_num, account_type)
            page_transactions.extend(transactions)
            self.logger.info(f"从该部分提取到 {len(transactions)} 条交易记录")

        return page_transactions

    def _is_structured_table(self, df: pd.DataFrame) -> bool:
        """判断是否为结构化表格"""
        if df.shape[1] < 3:
            return False
        header_text = ' '.join(map(str, df.columns)).lower() + ' ' + ' '.join(map(str, df.iloc[0].tolist())).lower()
        keywords = ['date', 'transaction', 'deposit', 'withdrawal', 'balance', '日期', '详情', '存入', '提取', '结余']
        if sum(keyword in header_text for keyword in keywords) >= 3:
            return True
        return False

    def extract_transactions_from_table(self, table_df: pd.DataFrame, page_num: int, target_account_type: str) -> list[dict]:
        self.logger.info(f'=== 解析表格 {page_num} (目标账户类型: {target_account_type}) ===')
        self.logger.info(f'表格形状: {table_df.shape}')
        self.logger.info(f'--- 表格 {page_num} 完整原始数据 ---\n{table_df.to_string()}')

        if self._is_structured_table(table_df):
            self.logger.info("检测到结构化表格，但解析器未实现。")
            pass

        self.logger.info("使用混合模式解析器。")
        transactions = []
        inherited_date = None
        inherited_currency = None

        for index, row in table_df.iterrows():
            row_text = ' '.join(map(str, row.tolist()))
            account_type_in_row = self._identify_account_type_from_context(table_df, index, row_text)

            if account_type_in_row is None:
                account_type_in_row = target_account_type

            if account_type_in_row != target_account_type:
                continue

            try:
                self.logger.info(f"找到{target_account_type}交易行 {index}: {row_text[:100]}...")
                parsed_transactions, inherited_date, inherited_currency = self._parse_mixed_format_row(
                    row_text, index, table_df, inherited_date, inherited_currency
                )
                if parsed_transactions:
                    for trans in parsed_transactions:
                        trans['账户类型'] = target_account_type
                        transactions.append(trans)
                        self.logger.info(f"创建{target_account_type}交易: {trans['日期']} - {trans['交易描述']} - {trans['货币']} - 余额: {trans.get('结余', 'N/A')}")
                    self.logger.info(f"从该行提取到 {len(parsed_transactions)} 条{target_account_type}交易")
            except Exception as e:
                self.logger.error(f"解析行 {index} 时出错: {e}", exc_info=True)

        return transactions

    def _identify_account_type_from_context(self, df: pd.DataFrame, current_index: int, row_text: str) -> str | None:
        context_text = row_text
        for i in range(max(0, current_index - 1), max(0, current_index - 3), -1):
            if i < len(df):
                context_text = ' '.join(map(str, df.iloc[i].tolist())) + ' ' + context_text

        if any(keyword.lower() in context_text.lower() for keyword in self.HKD_CURRENT_KEYWORDS):
            return '港币往来'
        if any(keyword.lower() in context_text.lower() for keyword in self.HKD_SAVINGS_KEYWORDS):
            return '港币储蓄'
        if any(keyword.lower() in context_text.lower() for keyword in self.FOREIGN_SAVINGS_KEYWORDS) or any(code in context_text for code in self.CURRENCY_CODES):
            return '外币储蓄'
        return None

    def _parse_mixed_format_row(self, row_text: str, row_index: int, df: pd.DataFrame, inherited_date: str | None, inherited_currency: str | None) -> tuple[list[dict], str | None, str | None]:
        transactions = []
        
        # 尝试从当前行提取货币，如果没有则继承
        current_currency = self._extract_currency_from_text(row_text)
        if not current_currency:
            current_currency = inherited_currency
        else:
            # 如果提取到新的货币，清理一下行文本，避免干扰描述提取
            for code in self.CURRENCY_CODES:
                row_text = row_text.replace(code, '')

        # 尝试从当前行提取日期，如果没有则继承
        dates = self.date_pattern.findall(row_text)
        current_date = dates[0] if dates else inherited_date

        amounts = [float(a.replace(',', '')) for a in self.amount_pattern.findall(row_text)]
        
        # 清理描述文本
        description_text = self.date_pattern.sub('', row_text)
        description_text = self.amount_pattern.sub('', description_text).strip()
        descriptions = [d.strip() for d in description_text.split('\n') if d.strip() and not d.replace('.', '', 1).isdigit()]
        
        final_description = ' '.join(descriptions)
        if not final_description or self.should_filter_transaction(final_description):
            return [], current_date, current_currency

        # 特殊处理 B/F BALANCE 和 CREDIT INTEREST 在同一行的情况
        if 'B/F BALANCE' in row_text and 'CREDIT INTEREST' in row_text:
            if len(amounts) >= 2:
                # B/F Balance
                transactions.append({
                    '日期': dates[0] if dates else current_date,
                    '交易描述': 'B/F BALANCE 承前转结',
                    '货币': 'HKD', # 通常是港币
                    '存入': None,
                    '提取': None,
                    '结余': amounts[-2] # 倒数第二个是它的余额
                })
                # Credit Interest
                transactions.append({
                    '日期': dates[1] if len(dates) > 1 else (dates[0] if dates else current_date),
                    '交易描述': 'CREDIT INTEREST 利息收入',
                    '货币': 'HKD',
                    '存入': amounts[0] if len(amounts) > 2 else None, # 第一个通常是利息金额
                    '提取': None,
                    '结余': amounts[-1] # 最后一个是最终余额
                })
            return transactions, (dates[-1] if dates else current_date), current_currency

        if amounts:
            balance = amounts[-1]
            deposit = None
            withdrawal = None

            # 如果金额多于一个，尝试确定是存入还是提取
            if len(amounts) > 1:
                # 简单的逻辑：检查关键字
                if any(k in final_description.lower() for k in ['deposit', 'credit', '存入', '利息']):
                    deposit = amounts[0]
                else:
                    withdrawal = amounts[0]
            
            # 标准化 'B/F BALANCE' 描述
            if 'B/F BALANCE' in final_description and '承前转结' not in final_description:
                final_description = 'B/F BALANCE 承前转结'

            transaction_record = {
                '日期': current_date,
                '交易描述': final_description,
                '货币': current_currency,
                '存入': deposit,
                '提取': withdrawal,
                '结余': balance
            }
            transactions.append(transaction_record)

        return transactions, current_date, current_currency

    def parse(self) -> list[dict]:
        """解析PDF文件并提取交易记录"""
        try:
            tables = camelot.read_pdf(self.file_path, pages='all', flavor='stream', edge_tol=500, row_tol=10)
        except Exception as e:
            self.logger.error(f"使用Camelot读取PDF文件失败: {e}")
            return []

        self.logger.info(f"在PDF中找到 {len(tables)} 个表格。")
        all_transactions = []
        processed_transactions = set()

        for i, table in enumerate(tables):
            page_num = i + 1
            self.logger.info(f"\n=== 第{page_num}页 调试信息 ===")
            self.logger.info(f"找到 {len(tables)} 个表格")
            self.logger.info(f"\n--- 表格 {i+1} 原始数据 ---")
            self.logger.info(f"表格形状: {table.df.shape}")
            self.logger.info(f"表格完整数据:\n{table.df.to_string()}")
            # self.logger.info(f"前5行数据:\n{table.df.head().to_string()}")  # 注释掉或删除这一行

            account_sections_on_page = self.find_account_sections_on_page([table.df], page_num)

            for account_type, sections in account_sections_on_page.items():
                if sections:
                    self.logger.info(f"在表格 {i+1} 中找到 {len(sections)} 个 '{account_type}' 部分")
                    transactions = self.extract_account_transactions([table.df], page_num, account_type, sections)
                    for trans in transactions:
                        # 创建一个唯一的标识符来检查重复
                        transaction_id = (trans.get('日期'), trans.get('交易描述'), trans.get('结余'), trans.get('账户类型'))
                        if transaction_id not in processed_transactions:
                            all_transactions.append(trans)
                            processed_transactions.add(transaction_id)

        self.transactions = all_transactions
        return self.transactions

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