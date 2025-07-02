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
        for page_num, page in enumerate(pdf.pages):
            try:
                # 尝试提取表格
                page_tables = page.extract_tables()
                if page_tables:
                    # 记录页码信息，方便调试
                    for table in page_tables:
                        if table and len(table) > 1:  # 只添加非空表格
                            tables.append({
                                'page': page_num + 1,
                                'data': table
                            })
            except Exception as e:
                logger.warning(f"提取第{page_num+1}页表格时出错: {str(e)}")
        
        # 如果没有找到表格，尝试使用文本分析
        if not tables:
            for page_num, page in enumerate(pdf.pages):
                try:
                    text = page.extract_text() or ""
                    # 这里可以添加文本分析逻辑，从纯文本中提取表格数据
                    # ...
                except Exception as e:
                    logger.warning(f"分析第{page_num+1}页文本时出错: {str(e)}")
        
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
class ESunBankParser(BankParser):
    """玉山银行账单解析器 - 增强版"""
    
    def parse(self, pdf):
        transactions = []
        account_info = {}
        account_type_transactions = {}  # 用于存储不同账户类型的交易记录
        
        # 提取表格
        tables = self.extract_tables(pdf)
        
        # 提取文本用于辅助识别
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n\n"
        
        # 记录提取的文本，用于调试
        logger.debug(f"提取的文本内容: {full_text[:500]}...")
        
        # 尝试提取账户信息 - 同时支持英文和繁体中文
        account_match = re.search(r'Account No\.\s*:\s*([\d\w]+)|帐戶號碼\s*:\s*([\d\w]+)', full_text)
        if account_match:
            account_info['账户号码'] = account_match.group(1) or account_match.group(2)
        
        # 查找所有账户类型
        account_types = ["Savings Account", "Current Account", "Time Deposit", "Foreign Currency", 
                       "活期存款", "定期存款", "外幣帳戶", "支票帳戶"]
        
        # 查找所有 "Account Transaction History" 关键字位置
        # 使用更灵活的正则表达式，处理可能存在的空格问题
        history_matches = list(re.finditer(r'Account\s+Transaction\s+History|帳戶\s*交易\s*歷史', full_text, re.IGNORECASE))
        
        if history_matches:
            logger.info(f"找到 {len(history_matches)} 个 'Account Transaction History' 关键字")
            
            # 处理每个Account Transaction History部分
            for i, history_match in enumerate(history_matches):
                start_pos = history_match.end()
                end_pos = len(full_text)
                
                # 如果有下一个匹配，则当前部分到下一个匹配开始
                if i < len(history_matches) - 1:
                    end_pos = history_matches[i+1].start()
                
                section_text = full_text[start_pos:end_pos]
                
                # 尝试识别该部分的账户类型
                detected_account_type = None
                for acc_type in account_types:
                    # 在当前部分或前面一小段文本中查找账户类型
                    search_text = full_text[max(0, start_pos-200):start_pos] + section_text[:200]
                    if acc_type.lower() in search_text.lower():
                        detected_account_type = acc_type
                        logger.info(f"在第 {i+1} 个 'Account Transaction History' 部分检测到账户类型: {detected_account_type}")
                        break
                
                if not detected_account_type:
                    detected_account_type = f"未知账户类型_{i+1}"
                    logger.warning(f"无法识别第 {i+1} 个 'Account Transaction History' 部分的账户类型，使用默认值: {detected_account_type}")
                
                # 处理该部分的表格或文本
                section_transactions = self._process_account_section(tables, section_text, detected_account_type)
                
                if section_transactions:
                    # 添加账户信息和类型到每个交易记录
                    for transaction in section_transactions:
                        transaction.update(account_info)
                        transaction['账户类型'] = detected_account_type
                    
                    # 添加到对应账户类型的交易记录列表
                    if detected_account_type not in account_type_transactions:
                        account_type_transactions[detected_account_type] = []
                    account_type_transactions[detected_account_type].extend(section_transactions)
                    
                    # 同时添加到总交易记录列表
                    transactions.extend(section_transactions)
        
        # 如果没有找到任何Account Transaction History部分，尝试处理整个文档
        if not transactions:
            logger.warning("未找到 'Account Transaction History' 部分，尝试处理整个文档")
            
            # 尝试从表格中提取
            table_transactions = self._process_tables(tables)
            if table_transactions:
                # 添加账户信息到每个交易记录
                for transaction in table_transactions:
                    transaction.update(account_info)
                    # 如果没有指定账户类型，使用默认值
                    if '账户类型' not in transaction:
                        transaction['账户类型'] = "未知账户类型"
                
                transactions.extend(table_transactions)
            
            # 如果表格提取失败，尝试从文本中提取
            if not transactions:
                logger.warning("未能从表格中提取到交易记录，尝试使用文本分析")
                text_transactions = self._extract_transactions_from_text(full_text)
                
                # 添加账户信息到每个交易记录
                for transaction in text_transactions:
                    transaction.update(account_info)
                    # 如果没有指定账户类型，使用默认值
                    if '账户类型' not in transaction:
                        transaction['账户类型'] = "未知账户类型"
                
                transactions.extend(text_transactions)
        
        # 返回所有交易记录和按账户类型分组的交易记录
        return transactions, account_type_transactions
    
    def _extract_transactions_from_text(self, text):
        """从文本中提取交易记录 - 增强版"""
        transactions = []
        
        # 记录提取的文本，用于调试
        logger.debug(f"正在从文本中提取交易记录，文本长度: {len(text)}")
        
        # 尝试查找 "Account Transaction History" 部分
        history_match = re.search(r'Account Transaction History|帳戶交易歷史', text, re.IGNORECASE)
        if history_match:
            # 获取匹配位置之后的文本
            start_pos = history_match.end()
            relevant_text = text[start_pos:]
            logger.debug(f"找到 'Account Transaction History' 关键字，提取后续文本进行分析")
        else:
            relevant_text = text
        
        # 尝试多种模式匹配交易记录
        patterns = [
            # 标准格式: 日期 货币 描述 金额
            r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+(\w{3})\s+([^\n]+?)\s+([-+]?\d[\d,.]+)',
            # 无货币格式: 日期 描述 金额
            r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+([^\n]+?)\s+([-+]?\d[\d,.]+)',
            # 繁体中文日期格式
            r'(\d{4}年\d{2}月\d{2}日)\s+([^\n]+?)\s+([-+]?\d[\d,.]+)',
            # 表格式文本: 多行匹配
            r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+(\w{3})?\s*([^\n]+?)\s+([-+]?\d[\d,.]+)\s+([-+]?\d[\d,.]+)?',
            # 日期 + 收入/支出 格式
            r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+([^\n]+?)\s+(\d[\d,.]+)\s+(\d[\d,.]+)',
            # 日期 + 货币 + 收入/支出 格式
            r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+(\w{3})\s+([^\n]+?)\s+(\d[\d,.]+)\s+(\d[\d,.]+)'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, relevant_text)
            if matches:
                for match in matches:
                    transaction = {}
                    
                    if len(match) >= 3:  # 至少有日期、描述、金额
                        date = match[0]
                        transaction["交易日期"] = self.parse_date(date)
                        
                        if len(match) >= 4 and match[1] and len(match[1]) == 3:  # 有货币字段
                            currency = match[1]
                            desc = match[2]
                            amount = match[3]
                            transaction["货币"] = currency
                            transaction["交易描述"] = desc.strip()
                            
                            # 尝试判断是收入还是支出
                            amount_value = self.parse_amount(amount)
                            if amount_value < 0:
                                transaction["支出金额"] = abs(amount_value)
                                transaction["收入金额"] = 0.0
                            else:
                                transaction["收入金额"] = amount_value
                                transaction["支出金额"] = 0.0
                            
                            # 如果有第5个元素，可能是余额
                            if len(match) >= 5 and match[4]:
                                transaction["账户余额"] = self.parse_amount(match[4])
                        else:  # 无货币字段
                            desc = match[1] if len(match) == 3 else match[2]
                            amount = match[2] if len(match) == 3 else match[3]
                            transaction["交易描述"] = desc.strip()
                            
                            # 尝试判断是收入还是支出
                            amount_value = self.parse_amount(amount)
                            if amount_value < 0:
                                transaction["支出金额"] = abs(amount_value)
                                transaction["收入金额"] = 0.0
                            else:
                                transaction["收入金额"] = amount_value
                                transaction["支出金额"] = 0.0
                        
                        transactions.append(transaction)
        
        # 尝试查找表格式的文本
        table_patterns = [
            r'Transaction Date\s+Currency\s+Description\s+.*?\n((?:[^\n]+\n)+)',
            r'交易日期\s+幣別\s+摘要\s+.*?\n((?:[^\n]+\n)+)',
            r'日期\s+摘要\s+.*?金額\s+.*?\n((?:[^\n]+\n)+)',
            r'Account Transaction History.*?\n((?:[^\n]+\n)+)',
            r'帳戶交易歷史.*?\n((?:[^\n]+\n)+)'
        ]
        
        for pattern in table_patterns:
            match = re.search(pattern, relevant_text)
            if match:
                table_text = match.group(1)
                # 按行分割
                lines = table_text.strip().split('\n')
                for line in lines:
                    # 尝试从每行提取交易信息
                    line_match = re.search(r'(\d{4}[/.-]\d{2}[/.-]\d{2})\s+(\w{3})?\s*([^\d]+)\s+([-+]?\d[\d,.]+)', line)
                    if line_match:
                        date = line_match.group(1)
                        currency = line_match.group(2) if line_match.group(2) else ""
                        desc = line_match.group(3)
                        amount = line_match.group(4)
                        
                        transaction = {
                            "交易日期": self.parse_date(date),
                            "交易描述": desc.strip()
                        }
                        
                        if currency:
                            transaction["货币"] = currency
                        
                        # 尝试判断是收入还是支出
                        amount_value = self.parse_amount(amount)
                        if amount_value < 0:
                            transaction["支出金额"] = abs(amount_value)
                            transaction["收入金额"] = 0.0
                        else:
                            transaction["收入金额"] = amount_value
                            transaction["支出金额"] = 0.0
                        
                        transactions.append(transaction)
        
        if transactions:
            logger.info(f"通过文本分析成功提取到{len(transactions)}条交易记录")
        else:
            logger.warning("文本分析未能提取到任何交易记录")
        
        return transactions
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引 - 增强版，支持部分匹配"""
        for name in possible_names:
            for i, header in enumerate(headers):
                # 完全匹配
                if name == header:
                    return i
                # 部分匹配
                if name in header:
                    return i
                # 忽略大小写的部分匹配
                if name.lower() in header.lower():
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
    "玉山银行": ESunBankParser(),
    # 其他银行使用通用解析器
    "渣打银行": GenericParser(),
    "汇丰银行": GenericParser(),
    "南洋银行": GenericParser(),
    "恒生银行": GenericParser(),
    "中银香港": GenericParser(),
    "东亚银行": GenericParser(),
    "大新银行": GenericParser(),
}


def get_bank_parser(bank_name):
    """获取指定银行的解析器"""
    return BANK_PARSERS.get(bank_name, GenericParser())


def get_supported_banks():
    """获取支持的银行列表"""
    return list(BANK_PARSERS.keys())

    def _process_account_section(self, tables, section_text, account_type):
        """处理特定账户类型的部分"""
        transactions = []
        
        # 首先尝试从表格中提取
        # 查找与该部分相关的表格
        relevant_tables = []
        for table_info in tables:
            table = table_info['data']
            page_num = table_info['page']
            
            # 简单判断表格是否属于当前部分（可以根据实际情况优化）
            table_text = "\n".join([" ".join([str(cell) for cell in row if cell]) for row in table])
            if account_type.lower() in table_text.lower() or any(keyword in table_text for keyword in ["Transaction Date", "交易日期", "Date", "Currency"]):
                relevant_tables.append(table_info)
        
        # 处理相关表格
        if relevant_tables:
            for table_info in relevant_tables:
                table_transactions = self._process_table(table_info, account_type)
                if table_transactions:
                    transactions.extend(table_transactions)
        
        # 如果表格提取失败，尝试从文本中提取
        if not transactions:
            text_transactions = self._extract_transactions_from_text(section_text)
            if text_transactions:
                # 添加账户类型
                for transaction in text_transactions:
                    transaction['账户类型'] = account_type
                transactions.extend(text_transactions)
        
        return transactions
    
    def _process_table(self, table_info, account_type):
        """处理单个表格"""
        transactions = []
        table = table_info['data']
        page_num = table_info['page']
        
        # 跳过空表格
        if not table or len(table) <= 1:
            return transactions
        
        # 查找表头行
        header_row = None
        for i, row in enumerate(table):
            row_text = " ".join([str(cell) for cell in row if cell])
            # 增加更多可能的表头标识
            if ("Transaction" in row_text and "Date" in row_text) or \
               "交易日期" in row_text or \
               "交易日" in row_text or \
               "Transaction Date" in row_text or \
               "Date" in row_text or \
               "Currency" in row_text or \
               "Description" in row_text or \
               "Withdrawal" in row_text or \
               "Deposit" in row_text or \
               "Balance" in row_text or \
               "幣別" in row_text or \
               "摘要" in row_text or \
               "提款" in row_text or \
               "存款" in row_text or \
               "餘額" in row_text or \
               "Account Transaction History" in row_text or \
               "帳戶交易歷史" in row_text:
                header_row = i
                break
        
        if header_row is None:
            # 如果找不到表头，尝试使用第一行作为表头
            if len(table) > 1:
                header_row = 0
            else:
                return transactions
        
        # 解析表头
        headers = [self.clean_text(cell) for cell in table[header_row]]
        
        # 查找关键列索引 - 增加繁体中文支持
        date_idx = self._find_column_index(headers, [
            "Transaction Date", "交易日期", "日期", "Date", "Transaction", "交易日", "Value Date", "入帳日"
        ])
        currency_idx = self._find_column_index(headers, [
            "Currency", "币别", "货币", "幣別", "幣種", "币种", "Ccy"
        ])
        desc_idx = self._find_column_index(headers, [
            "Description", "摘要", "交易描述", "Value Date", "摘要", "交易說明", "備註", "附註", "Particulars"
        ])
        withdrawal_idx = self._find_column_index(headers, [
            "Withdrawal", "支出", "借方金额", "支出", "提款", "支出金額", "借方", "付款", "Debit", "Dr", "Withdrawal Amount"
        ])
        deposit_idx = self._find_column_index(headers, [
            "Deposit", "存入", "贷方金额", "收入", "存款", "收入金額", "贷方", "收款", "Credit", "Cr", "Deposit Amount"
        ])
        balance_idx = self._find_column_index(headers, [
            "Balance", "余额", "账户余额", "餘額", "帳戶餘額", "Ledger Balance"
        ])
        remark_idx = self._find_column_index(headers, [
            "Remark", "备注", "註記", "備註", "附註", "Notes"
        ])
        
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
            
            # 添加账户类型
            transaction['账户类型'] = account_type
            
            # 交易日期
            if date_idx is not None and date_idx < len(row):
                transaction["交易日期"] = self.parse_date(row[date_idx])
            
            # 货币
            if currency_idx is not None and currency_idx < len(row):
                transaction["货币"] = self.clean_text(row[currency_idx])
            
            # 交易描述
            if desc_idx is not None and desc_idx < len(row):
                transaction["交易描述"] = self.clean_text(row[desc_idx])
            
            # 支出金额
            if withdrawal_idx is not None and withdrawal_idx < len(row):
                withdrawal = self.parse_amount(row[withdrawal_idx])
                transaction["支出金额"] = withdrawal
            
            # 存入金额
            if deposit_idx is not None and deposit_idx < len(row):
                deposit = self.parse_amount(row[deposit_idx])
                transaction["收入金额"] = deposit
            
            # 账户余额
            if balance_idx is not None and balance_idx < len(row):
                transaction["账户余额"] = self.parse_amount(row[balance_idx])
            
            # 备注
            if remark_idx is not None and remark_idx < len(row):
                transaction["备注"] = self.clean_text(row[remark_idx])
            
            # 添加到交易列表
            if transaction.get("交易日期"):
                transactions.append(transaction)
        
        return transactions
    
    def _process_tables(self, tables):
        """处理所有表格"""
        transactions = []
        
        for table_info in tables:
            table_transactions = self._process_table(table_info, "未知账户类型")
            if table_transactions:
                transactions.extend(table_transactions)
        
        return transactions