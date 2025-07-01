import re
import logging
from abc import ABC, abstractmethod
from datetime import datetime

logger = logging.getLogger(__name__)

class BankParser(ABC):
    """银行账单解析器的抽象基类"""
    
    @abstractmethod
    def parse(self, pdf):
        """解析PDF文件并提取交易记录"""
        pass
    
    def extract_tables(self, pdf):
        """从PDF中提取表格数据"""
        tables = []
        for page in pdf.pages:
            try:
                # 尝试提取表格
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
            except Exception as e:
                logger.warning(f"提取表格时出错: {str(e)}")
        return tables
    
    def clean_text(self, text):
        """清理文本，移除多余的空白字符"""
        if text is None:
            return ""
        return re.sub(r'\s+', ' ', str(text)).strip()
    
    def parse_amount(self, amount_str):
        """解析金额字符串为浮点数"""
        if not amount_str or amount_str == "--" or amount_str == "-":
            return 0.0
        
        # 移除非数字字符（保留小数点和负号）
        amount_str = re.sub(r'[^\d.-]', '', str(amount_str))
        
        try:
            return float(amount_str)
        except ValueError:
            return 0.0
    
    def parse_date(self, date_str):
        """解析日期字符串为标准格式"""
        if not date_str:
            return ""
        
        # 清理日期字符串
        date_str = self.clean_text(date_str)
        
        # 尝试多种常见的日期格式
        date_formats = [
            "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日",
            "%Y.%m.%d", "%d-%m-%Y", "%d/%m/%Y",
            "%Y%m%d"
        ]
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%Y-%m-%d")
            except ValueError:
                continue
        
        # 如果无法解析，返回原始字符串
        return date_str


class ICBCParser(BankParser):
    """工商银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        # 提取文本用于辅助识别
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() or ""
        
        # 检查是否是信用卡账单
        is_credit_card = "信用卡" in full_text or "Credit Card" in full_text
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "Transaction Date" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "Transaction Date", "日期"])
            desc_idx = self._find_column_index(headers, ["交易描述", "Transaction Description", "摘要", "交易说明"])
            amount_idx = self._find_column_index(headers, ["交易金额", "Transaction Amount", "金额"])
            balance_idx = self._find_column_index(headers, ["账户余额", "Balance", "余额"])
            income_idx = self._find_column_index(headers, ["收入金额", "Income", "收入", "贷方金额"])
            expense_idx = self._find_column_index(headers, ["支出金额", "Expense", "支出", "借方金额"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 交易金额
                if amount_idx is not None and amount_idx < len(row):
                    transaction["交易金额"] = self.parse_amount(row[amount_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 收入和支出金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 如果没有明确的收入/支出字段，但有交易金额，则根据金额正负判断
                if "交易金额" in transaction and "收入金额" not in transaction and "支出金额" not in transaction:
                    amount = transaction["交易金额"]
                    if amount > 0:
                        transaction["收入金额"] = amount
                        transaction["支出金额"] = 0.0
                    else:
                        transaction["收入金额"] = 0.0
                        transaction["支出金额"] = abs(amount)
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


class CCBParser(BankParser):
    """建设银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "交易时间" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "交易时间", "日期"])
            desc_idx = self._find_column_index(headers, ["交易描述", "摘要", "交易说明", "交易内容"])
            income_idx = self._find_column_index(headers, ["收入金额", "收入", "贷方金额", "贷方"])
            expense_idx = self._find_column_index(headers, ["支出金额", "支出", "借方金额", "借方"])
            balance_idx = self._find_column_index(headers, ["账户余额", "余额"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 收入金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                # 支出金额
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


class ABCParser(BankParser):
    """农业银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "交易时间" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "交易时间", "日期"])
            desc_idx = self._find_column_index(headers, ["交易描述", "摘要", "交易说明", "交易内容"])
            income_idx = self._find_column_index(headers, ["收入金额", "收入", "贷方金额", "贷方"])
            expense_idx = self._find_column_index(headers, ["支出金额", "支出", "借方金额", "借方"])
            balance_idx = self._find_column_index(headers, ["账户余额", "余额"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 收入金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                # 支出金额
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


class BOCParser(BankParser):
    """中国银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "交易时间" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "交易时间", "日期"])
            desc_idx = self._find_column_index(headers, ["交易描述", "摘要", "交易说明", "交易内容"])
            income_idx = self._find_column_index(headers, ["收入金额", "收入", "贷方金额", "贷方"])
            expense_idx = self._find_column_index(headers, ["支出金额", "支出", "借方金额", "借方"])
            balance_idx = self._find_column_index(headers, ["账户余额", "余额"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 收入金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                # 支出金额
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


class CMBParser(BankParser):
    """招商银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        # 提取文本用于辅助识别
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() or ""
        
        # 检查是否是信用卡账单
        is_credit_card = "信用卡" in full_text or "Credit Card" in full_text
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "交易时间" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "交易时间", "日期"])
            desc_idx = self._find_column_index(headers, ["交易描述", "摘要", "交易说明", "交易内容"])
            income_idx = self._find_column_index(headers, ["收入金额", "收入", "贷方金额", "贷方"])
            expense_idx = self._find_column_index(headers, ["支出金额", "支出", "借方金额", "借方"])
            balance_idx = self._find_column_index(headers, ["账户余额", "余额"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 收入金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                # 支出金额
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


# 通用解析器，用于未特别实现的银行
class GenericParser(BankParser):
    """通用银行账单解析器"""
    
    def parse(self, pdf):
        transactions = []
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        for table in tables:
            # 跳过空表格
            if not table or len(table) <= 1:
                continue
            
            # 查找表头行
            header_row = None
            for i, row in enumerate(table):
                row_text = " ".join([str(cell) for cell in row if cell])
                if "交易日期" in row_text or "交易时间" in row_text or "日期" in row_text:
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            # 解析表头
            headers = [self.clean_text(cell) for cell in table[header_row]]
            
            # 查找关键列索引
            date_idx = self._find_column_index(headers, ["交易日期", "交易时间", "日期", "Date"])
            desc_idx = self._find_column_index(headers, ["交易描述", "摘要", "交易说明", "交易内容", "Description"])
            income_idx = self._find_column_index(headers, ["收入金额", "收入", "贷方金额", "贷方", "Credit"])
            expense_idx = self._find_column_index(headers, ["支出金额", "支出", "借方金额", "借方", "Debit"])
            amount_idx = self._find_column_index(headers, ["交易金额", "金额", "Amount"])
            balance_idx = self._find_column_index(headers, ["账户余额", "余额", "Balance"])
            
            # 处理数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                
                # 跳过空行
                if not row or all(not cell for cell in row):
                    continue
                
                # 确保行长度与表头一致
                if len(row) < len(headers):
                    row.extend([None] * (len(headers) - len(row)))
                
                # 提取交易数据
                transaction = {}
                
                # 交易日期
                if date_idx is not None and date_idx < len(row):
                    transaction["交易日期"] = self.parse_date(row[date_idx])
                
                # 交易描述
                if desc_idx is not None and desc_idx < len(row):
                    transaction["交易描述"] = self.clean_text(row[desc_idx])
                
                # 交易金额
                if amount_idx is not None and amount_idx < len(row):
                    transaction["交易金额"] = self.parse_amount(row[amount_idx])
                
                # 收入金额
                if income_idx is not None and income_idx < len(row):
                    transaction["收入金额"] = self.parse_amount(row[income_idx])
                
                # 支出金额
                if expense_idx is not None and expense_idx < len(row):
                    transaction["支出金额"] = self.parse_amount(row[expense_idx])
                
                # 账户余额
                if balance_idx is not None and balance_idx < len(row):
                    transaction["账户余额"] = self.parse_amount(row[balance_idx])
                
                # 如果没有明确的收入/支出字段，但有交易金额，则根据金额正负判断
                if "交易金额" in transaction and "收入金额" not in transaction and "支出金额" not in transaction:
                    amount = transaction["交易金额"]
                    if amount > 0:
                        transaction["收入金额"] = amount
                        transaction["支出金额"] = 0.0
                    else:
                        transaction["收入金额"] = 0.0
                        transaction["支出金额"] = abs(amount)
                
                # 添加到交易列表
                if transaction.get("交易日期"):
                    transactions.append(transaction)
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name in header:
                    return i
        return None


# 银行解析器映射
BANK_PARSERS = {
    "工商银行": ICBCParser(),
    "建设银行": CCBParser(),
    "农业银行": ABCParser(),
    "中国银行": BOCParser(),
    "招商银行": CMBParser(),
    # 其他银行使用通用解析器
    "交通银行": GenericParser(),
    "浦发银行": GenericParser(),
    "民生银行": GenericParser(),
    "中信银行": GenericParser(),
    "光大银行": GenericParser(),
    "华夏银行": GenericParser(),
    "广发银行": GenericParser(),
    "平安银行": GenericParser(),
    "邮储银行": GenericParser(),
}


def get_bank_parser(bank_name):
    """获取指定银行的解析器"""
    return BANK_PARSERS.get(bank_name, GenericParser())


def get_supported_banks():
    """获取支持的银行列表"""
    return list(BANK_PARSERS.keys())