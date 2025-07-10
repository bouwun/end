import re
import logging
from abc import ABC, abstractmethod
from datetime import datetime
import pandas as pd
import pdfplumber
import camelot  # 添加camelot导入

logger = logging.getLogger(__name__)

class BankParser(ABC):
    """银行账单解析器的抽象基类"""
    
    @abstractmethod
    def parse(self, pdf):
        """解析PDF文件并提取交易记录"""
        pass
    
    @abstractmethod
    def save_to_excel(self, transactions, output_path, account_info=None):
        """保存交易记录到Excel文件"""
        pass
    
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
        
        # 汇丰银行常用的日期格式
        date_formats = [
            "%d %b %Y",    # 01 Jan 2024
            "%d-%b-%Y",    # 01-Jan-2024
            "%d/%m/%Y",    # 01/01/2024
            "%Y-%m-%d",    # 2024-01-01
            "%d %B %Y",    # 01 January 2024
        ]
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%Y-%m-%d")
            except ValueError:
                continue
        
        # 如果无法解析，返回原始字符串
        return date_str


class HSBCParser(BankParser):
    """汇丰银行账单解析器"""
    
    def __init__(self):
        super().__init__()
        self.bank_name = "汇丰银行"
        self.target_account_types = ["港币往来", "港币储蓄", "外币储蓄"]
        self.account_keywords = {
            "港币往来": ["港币往来", "HKD CURRENT", "HKD Current"],
            "港币储蓄": ["港币储蓄", "HKD SAVINGS", "HKD Savings"],
            "外币储蓄": ["外币储蓄", "FOREIGN CURRENCY SAVINGS", "Foreign Currency Savings"]
        }
    
    def _find_column_name(self, headers, candidates):
        """在表头中查找匹配的列名"""
        for candidate in candidates:
            for header in headers:
                if candidate.lower() in str(header).lower():
                    return header
        return None
    
    def should_filter_transaction(self, transaction_details):
        """判断是否应该过滤掉某个交易记录"""
        if not transaction_details:
            return True
        
        # 过滤掉的关键词
        filter_keywords = [
            "B/F BALANCE", "BALANCE B/F", "BALANCE BROUGHT FORWARD",
            "结余", "余额结转", "期初余额", "OPENING BALANCE"
        ]
        
        transaction_upper = transaction_details.upper()
        for keyword in filter_keywords:
            if keyword in transaction_upper:
                return True
        
        return False
    
    def extract_currency_from_details(self, transaction_details, account_type):
        """从交易详情中提取货币类型"""
        if not transaction_details:
            # 根据账户类型推断货币
            if "港币" in account_type:
                return "HKD"
            elif "外币" in account_type:
                return "USD"  # 默认美元，可根据实际情况调整
            else:
                return "HKD"
        
        # 从交易详情中查找货币代码
        currency_pattern = r'\\b([A-Z]{3})\\b'
        currency_matches = re.findall(currency_pattern, transaction_details.upper())
        
        # 常见货币代码
        common_currencies = ['USD', 'HKD', 'CNY', 'EUR', 'GBP', 'JPY', 'AUD', 'CAD', 'SGD']
        
        for currency in currency_matches:
            if currency in common_currencies:
                return currency
        
        # 如果没有找到，根据账户类型推断
        if "港币" in account_type:
            return "HKD"
        elif "外币" in account_type:
            return "USD"  # 默认美元
        else:
            return "HKD"


    def find_account_pages(self, pdf_obj, account_type):
        """查找指定账户类型所在的页面"""
        account_pages = []
        keywords = self.account_keywords.get(account_type, [])
        
        try:
            for page_num, page in enumerate(pdf_obj.pages, 1):
                text = page.extract_text() or ""
                
                # 检查页面是否包含当前账户类型的关键词
                for keyword in keywords:
                    keyword_pos = text.find(keyword)
                    if keyword_pos != -1:
                        # 找到关键词后，检查是否有结束标记
                        end_markers = [
                            "Total No. of Deposits:",
                            "Total No. of Deposits",
                            "Total No of Deposits:",
                            "Total No of Deposits",
                            "Total Deposits:",
                            "Total Deposits"
                        ]
                        
                        # 查找结束位置
                        end_pos = len(text)
                        
                        for marker in end_markers:
                            marker_pos = text.find(marker, keyword_pos + len(keyword))
                            if marker_pos != -1:
                                end_pos = marker_pos
                                break
                        
                        # 计算文本区域的行号范围（用于后续表格过滤）
                        text_before_start = text[:keyword_pos]
                        text_in_section = text[keyword_pos:end_pos]
                        
                        start_line = text_before_start.count('\n')
                        end_line = start_line + text_in_section.count('\n')
                        
                        account_pages.append({
                            'page': page_num,
                            'account_type': account_type,
                            'keyword': keyword,
                            'start_pos': keyword_pos,
                            'end_pos': end_pos,
                            'start_line': start_line,
                            'end_line': end_line,
                            'text_section': text_in_section
                        })
                        
                        logger.info(f"在第{page_num}页找到{account_type}账户: {keyword} (行{start_line}-{end_line})")
                        break  # 找到一个关键词就够了
        
        except Exception as e:
            logger.error(f"查找{account_type}账户页面时出错: {str(e)}")
        
        return account_pages
    
    def extract_account_transactions(self, pdf_path, page_num, account_type, page_info):
        """从指定页面提取指定账户类型的交易记录"""
        transactions = []
        
        try:
            # 使用camelot提取表格
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num),
                flavor='stream',
                table_areas=None,
                columns=None,
                edge_tol=1000,
                row_tol=20,
                column_tol=10,
                split_text=True
            )
        
            # 添加调试输出
            print(f"\n=== 第{page_num}页 {account_type} 调试信息 ===")
            print(f"找到 {len(tables)} 个表格")
        
            # 处理每个表格
            for table_idx, table in enumerate(tables):
                df = table.df
                
                # 输出原始表格数据
                print(f"\n--- 表格 {table_idx + 1} 原始数据 ---")
                print("表格形状:", df.shape)
                print("前10行数据:")
                print(df.head(10).to_string())
                
                # 检查表格是否包含交易数据的特征
                has_transaction_features = self.check_transaction_features(df)
                print(f"包含交易特征: {has_transaction_features}")
                
                if not has_transaction_features:
                    continue
                
                # 检查表格内容是否属于当前账户类型的文本区域
                table_text = df.to_string()
                is_in_section = self.is_table_in_account_section(table_text, account_type, page_info)
                print(f"属于{account_type}区域: {is_in_section}")
                
                if not is_in_section:
                    continue
                
                # 提取表格中的交易记录
                table_transactions = self.extract_transactions_from_table(df, account_type, table_idx + 1)
                print(f"提取到 {len(table_transactions)} 条交易记录")
                
                # 输出提取的交易记录
                for i, trans in enumerate(table_transactions[:3]):  # 只显示前3条
                    print(f"交易 {i+1}: {trans}")
                
                transactions.extend(table_transactions)
        
        except Exception as e:
            logger.error(f"从第{page_num}页提取{account_type}交易记录时出错: {str(e)}")
            print(f"错误: {str(e)}")
        
        return transactions
    
    def check_transaction_features(self, df):
        """检查表格是否包含交易特征"""
        transaction_indicators = [
            "Date", "Transaction", "Details", "Deposit", "Withdrawal", "Balance",
            "日期", "交易", "存款", "取款", "余额", "CCY", "B/F BALANCE", "CREDIT INTEREST"
        ]
        
        found_indicators = []
        for idx, row in df.iterrows():
            row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()]).upper()
            for indicator in transaction_indicators:
                if indicator.upper() in row_text and indicator not in found_indicators:
                    found_indicators.append(indicator)
        
        return len(found_indicators) >= 3
    
    def is_table_in_account_section(self, table_text, account_type, page_info):
        """检查表格是否在指定账户类型的文本区域内"""
        # 检查表格是否包含明显的排除标志
        exclude_keywords = ["Statement Summary", "账单摘要", "Fee Schedule", "收费表", "Portfolio Summary", "资产摘要"]
        table_text_upper = table_text.upper()
        
        for exclude_keyword in exclude_keywords:
            if exclude_keyword.upper() in table_text_upper:
                return False
        
        # 检查表格是否包含当前账户类型的关键词
        current_keywords = self.account_keywords.get(account_type, [])
        for keyword in current_keywords:
            if keyword.upper() in table_text_upper:
                return True
        
        # 如果不包含当前账户类型关键词，则检查是否在文本区域内
        text_section = page_info.get('text_section', '')
        table_lines = table_text.split('\n')
        for line in table_lines[:5]:  # 检查表格前5行
            line = line.strip()
            if len(line) > 10 and line in text_section:
                return True
        
        return False
    
    def parse(self, pdf_path_or_obj):
        """解析汇丰银行PDF账单"""
        all_transactions = []
        
        try:
            # 获取PDF路径和对象
            if isinstance(pdf_path_or_obj, str):
                pdf_path = pdf_path_or_obj
                with pdfplumber.open(pdf_path) as pdf:
                    pdf_obj = pdf
            else:
                pdf_obj = pdf_path_or_obj
                pdf_path = getattr(pdf_obj, 'stream', None)
                if hasattr(pdf_path, 'name'):
                    pdf_path = pdf_path.name
                else:
                    return all_transactions
            
            # 遍历账户类型
            for account_type in self.target_account_types:
                account_pages = self.find_account_pages(pdf_obj, account_type)
                
                if not account_pages:
                    continue
                
                # 从找到的页面提取交易记录
                for page_info in account_pages:
                    page_num = page_info['page']
                    transactions = self.extract_account_transactions(pdf_path, page_num, account_type, page_info)
                    
                    if transactions:
                        all_transactions.extend(transactions)
        
        except Exception as e:
            logger.error(f"解析汇丰银行PDF时出错: {str(e)}")
        
        return all_transactions
    
    def extract_transactions_from_table(self, df, account_type, table_idx):
        """从表格中提取指定账户类型的交易记录"""
        transactions = []
        
        try:
            # 添加调试输出
            print(f"\n=== 解析表格 {table_idx} (目标账户类型: {account_type}) ===")
            print(f"表格形状: {df.shape}")
            
            # 初始化状态跟踪变量
            last_currency = "HKD"
            last_date = None
            
            # 查找包含交易数据的行，并识别实际的账户类型
            for idx, row in df.iterrows():
                row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()])
                
                # 检查是否包含日期模式和交易信息
                if re.search(r'\d{1,2}\s+[A-Za-z]{3}', row_text) and ('BALANCE' in row_text or 'CREDIT' in row_text or 'DEBIT' in row_text or 'DEPOSIT' in row_text):
                    
                    # 识别该行数据实际属于哪个账户类型
                    actual_account_type = self._identify_account_type_from_context(row_text, idx, df)
                    
                    # 只处理属于目标账户类型的数据
                    if actual_account_type == account_type:
                        print(f"\n找到{account_type}交易行 {idx}: {row_text[:200]}...")
                        
                        # 解析混合格式的交易数据
                        parsed_transactions, last_currency, last_date = self._parse_mixed_format_row(row_text, actual_account_type, last_currency, last_date)
                        transactions.extend(parsed_transactions)
                        
                        print(f"从该行提取到 {len(parsed_transactions)} 条{actual_account_type}交易")
                    else:
                        print(f"\n跳过非{account_type}数据 (实际类型: {actual_account_type}): {row_text[:100]}...")
        
        except Exception as e:
            logger.error(f"从表格{table_idx}提取{account_type}交易记录时出错: {str(e)}")
            print(f"解析错误: {str(e)}")
        
        return transactions
    
    def _identify_account_type_from_context(self, row_text, row_idx, df):
        """根据上下文识别交易数据实际属于的账户类型"""
        try:
            # 向上查找最近的账户类型标识
            for i in range(row_idx, -1, -1):
                context_row = " ".join([str(cell) for cell in df.iloc[i].values if pd.notna(cell) and str(cell).strip()])
                
                # 检查港币往来
                if any(keyword in context_row for keyword in self.account_keywords.get("港币往来", [])):
                    return "港币往来"
                
                # 检查港币储蓄
                if any(keyword in context_row for keyword in self.account_keywords.get("港币储蓄", [])):
                    return "港币储蓄"
                
                # 检查外币储蓄
                if any(keyword in context_row for keyword in self.account_keywords.get("外币储蓄", [])):
                    return "外币储蓄"
                
                # 如果找到了货币代码，可能是外币储蓄
                if re.search(r'\b(USD|GBP|EUR|JPY|CNY)\b', context_row):
                    return "外币储蓄"
            
            # 如果在当前行中包含货币代码，判断为外币储蓄
            if re.search(r'\b(USD|GBP|EUR|JPY|CNY)\b', row_text):
                return "外币储蓄"
            
            # 默认返回港币往来（如果无法确定）
            return "港币往来"
            
        except Exception as e:
            logger.error(f"识别账户类型时出错: {str(e)}")
            return "港币往来"  # 默认值
    
    def _parse_mixed_format_row(self, row_text, account_type, last_currency="HKD", last_date=None):
        """解析混合格式的交易行数据"""
        transactions = []
        current_currency = last_currency  # 继承上一行的货币类型
        current_date = last_date  # 继承上一行的日期
        
        try:
            # 按换行符分割文本
            parts = row_text.split('\n')
            
            # 检测货币类型
            detected_currency = current_currency  # 默认使用继承的货币类型
            if account_type == "外币储蓄":
                # 查找货币代码
                for part in parts:
                    currency_match = re.search(r'\b(USD|GBP|EUR|JPY|CNY)\b', part)
                    if currency_match:
                        detected_currency = currency_match.group(1)
                        current_currency = detected_currency  # 更新当前货币类型
                        break
            
            # 查找日期和交易详情的模式
            i = 0
            while i < len(parts):
                part = parts[i].strip()
                
                # 跳过货币代码行
                if re.match(r'^(USD|GBP|EUR|JPY|CNY)$', part):
                    i += 1
                    continue
                
                # 查找日期模式
                date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', part)
                if date_match:
                    date_value = date_match.group(1)
                    current_date = date_value  # 更新当前日期
                    
                    # 提取交易描述 - 改进逻辑
                    transaction_desc = ""
                    if date_match.end() < len(part):
                        remaining_text = part[date_match.end():].strip()
                        # 移除多余的文字如"提取"、"结余"
                        remaining_text = re.sub(r'\s*(提取|结余)\s*', ' ', remaining_text).strip()
                        transaction_desc = remaining_text
                    
                    # 查找下一行的交易详情
                    if i + 1 < len(parts):
                        next_part = parts[i + 1].strip()
                        # 检查是否是交易描述而不是金额或货币代码
                        if (next_part and 
                            not re.match(r'^[\d,]+\.\d{2}$', next_part) and 
                            not re.search(r'\d{1,2}\s+[A-Za-z]{3}', next_part) and
                            not re.match(r'^(USD|GBP|EUR|JPY|CNY)$', next_part)):
                            # 常见的交易描述关键词
                            if any(keyword in next_part for keyword in ['BALANCE', 'CREDIT', 'INTEREST', 'DEPOSIT', 'WITHDRAWAL', 'TRANSFER']):
                                if transaction_desc:
                                    transaction_desc += " " + next_part
                                else:
                                    transaction_desc = next_part
                                i += 1
                    
                    # 清理和标准化交易描述
                    if transaction_desc:
                        # 标准化常见术语
                        transaction_desc = re.sub(r'B/F\s*BALANCE', 'B/F BALANCE 承前结余', transaction_desc)
                        transaction_desc = re.sub(r'CREDIT\s*INTEREST', 'CREDIT INTEREST 利息收入', transaction_desc)
                        transaction_desc = re.sub(r'^DEPOSIT$', 'DEPOSIT 存款', transaction_desc)
                        
                        # 移除多余的空格
                        transaction_desc = re.sub(r'\s+', ' ', transaction_desc).strip()
                    
                    # 查找金额信息
                    deposit_amount = 0.0
                    withdrawal_amount = 0.0
                    balance_value = 0.0
                    
                    # 在当前行和后续行查找金额
                    amount_text = row_text[row_text.find(date_value):]
                    amounts = re.findall(r'([\d,]+\.\d{2})', amount_text)
                    
                    if amounts:
                        # 根据交易类型分配金额
                        if 'CREDIT' in transaction_desc or '利息' in transaction_desc:
                            deposit_amount = self.parse_amount(amounts[0])
                            if len(amounts) > 1:
                                balance_value = self.parse_amount(amounts[-1])
                        elif 'DEPOSIT' in transaction_desc and 'B/F' not in transaction_desc:
                            deposit_amount = self.parse_amount(amounts[0])
                            if len(amounts) > 1:
                                balance_value = self.parse_amount(amounts[-1])
                        elif 'WITHDRAWAL' in transaction_desc or 'DEBIT' in transaction_desc:
                            withdrawal_amount = self.parse_amount(amounts[0])
                            if len(amounts) > 1:
                                balance_value = self.parse_amount(amounts[-1])
                        else:
                            # B/F BALANCE 或其他情况
                            balance_value = self.parse_amount(amounts[0])
                    
                    # 创建交易记录
                    if transaction_desc and transaction_desc not in ['提取', '结余', '存入', '承前转结']:
                        transaction = {
                            "账户类型": account_type,
                            "银行名称": self.bank_name,
                            "Date": date_value,
                            "Transaction Details": transaction_desc,
                            "Deposit": deposit_amount,
                            "Withdrawal": withdrawal_amount,
                            "Balance": balance_value,
                            "Currency": detected_currency
                        }
                        
                        transactions.append(transaction)
                        print(f"创建{account_type}交易: {date_value} - {transaction_desc} - {detected_currency} - 余额: {balance_value}")
                
                else:
                    # 当前行没有日期，但可能有交易信息
                    # 使用继承的日期和货币类型
                    if current_date and part and not re.match(r'^(USD|GBP|EUR|JPY|CNY)$', part):
                        # 检查是否包含交易关键词
                        if any(keyword in part for keyword in ['BALANCE', 'CREDIT', 'INTEREST', 'DEPOSIT', 'WITHDRAWAL', 'TRANSFER']):
                            transaction_desc = part
                            
                            # 标准化交易描述
                            transaction_desc = re.sub(r'B/F\s*BALANCE', 'B/F BALANCE 承前结余', transaction_desc)
                            transaction_desc = re.sub(r'CREDIT\s*INTEREST', 'CREDIT INTEREST 利息收入', transaction_desc)
                            transaction_desc = re.sub(r'^DEPOSIT$', 'DEPOSIT 存款', transaction_desc)
                            transaction_desc = re.sub(r'\s+', ' ', transaction_desc).strip()
                            
                            # 查找金额信息
                            deposit_amount = 0.0
                            withdrawal_amount = 0.0
                            balance_value = 0.0
                            
                            amounts = re.findall(r'([\d,]+\.\d{2})', part)
                            if amounts:
                                if 'CREDIT' in transaction_desc or '利息' in transaction_desc:
                                    deposit_amount = self.parse_amount(amounts[0])
                                    if len(amounts) > 1:
                                        balance_value = self.parse_amount(amounts[-1])
                                elif 'DEPOSIT' in transaction_desc and 'B/F' not in transaction_desc:
                                    deposit_amount = self.parse_amount(amounts[0])
                                    if len(amounts) > 1:
                                        balance_value = self.parse_amount(amounts[-1])
                                elif 'WITHDRAWAL' in transaction_desc or 'DEBIT' in transaction_desc:
                                    withdrawal_amount = self.parse_amount(amounts[0])
                                    if len(amounts) > 1:
                                        balance_value = self.parse_amount(amounts[-1])
                                else:
                                    balance_value = self.parse_amount(amounts[0])
                            
                            # 创建交易记录（使用继承的日期和货币类型）
                            if transaction_desc and transaction_desc not in ['提取', '结余', '存入', '承前转结']:
                                transaction = {
                                    "账户类型": account_type,
                                    "银行名称": self.bank_name,
                                    "Date": current_date,  # 使用继承的日期
                                    "Transaction Details": transaction_desc,
                                    "Deposit": deposit_amount,
                                    "Withdrawal": withdrawal_amount,
                                    "Balance": balance_value,
                                    "Currency": current_currency  # 使用继承的货币类型
                                }
                                
                                transactions.append(transaction)
                                print(f"创建{account_type}交易(继承): {current_date} - {transaction_desc} - {current_currency} - 余额: {balance_value}")
                
                i += 1
        
        except Exception as e:
            logger.error(f"解析混合格式行时出错: {str(e)}")
            print(f"解析错误: {str(e)}")
        
        return transactions, current_currency, current_date  # 返回更新后的状态
    
    def _parse_multi_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析包含多个交易的行"""
        transactions = []
        
        try:
            detail_parts = details_text.split("\n") if details_text else []
            deposit_parts = deposit_text.split("\n") if deposit_text else []
            withdrawal_parts = withdrawal_text.split("\n") if withdrawal_text else []
            balance_parts = balance_text.split("\n") if balance_text else []
            
            i = 0
            while i < len(detail_parts):
                detail_part = detail_parts[i].strip()
                
                if not detail_part:
                    i += 1
                    continue
                
                # 跳过货币代码行
                if re.match(r'^[A-Z]{3}$', detail_part):
                    i += 1
                    continue
                
                # 解析日期和交易描述
                date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', detail_part)
                if date_match:
                    date_value = self.parse_date(date_match.group(1))
                    if not date_value:
                        date_value = date_match.group(1)
                    
                    transaction_desc = detail_part[date_match.end():].strip()
                    
                    # 查找对应的金额和余额
                    deposit_amount = 0.0
                    withdrawal_amount = 0.0
                    balance_value = 0.0
                    
                    if i < len(deposit_parts) and deposit_parts[i].strip():
                        deposit_amount = self.parse_amount(deposit_parts[i])
                    
                    if i < len(withdrawal_parts) and withdrawal_parts[i].strip():
                        withdrawal_amount = self.parse_amount(withdrawal_parts[i])
                    
                    if i < len(balance_parts) and balance_parts[i].strip():
                        balance_value = self.parse_amount(balance_parts[i])
                    
                    # 特殊处理利息收入
                    if "CREDIT INTEREST" in transaction_desc or "利息收入" in transaction_desc:
                        amount_in_desc = re.search(r'([\d,]+\.\d{2})', transaction_desc)
                        if amount_in_desc and deposit_amount == 0.0:
                            deposit_amount = self.parse_amount(amount_in_desc.group(1))
                    
                    # 创建交易记录
                    transaction = {
                        "账户类型": account_type,
                        "银行名称": self.bank_name,
                        "Date": date_value,
                        "Transaction Details": transaction_desc,
                        "Deposit": deposit_amount,
                        "Withdrawal": withdrawal_amount,
                        "Balance": balance_value,
                        "Currency": self.extract_currency_from_details(transaction_desc, account_type)
                    }
                    
                    transactions.append(transaction)
                
                i += 1
        
        except Exception as e:
            logger.error(f"解析多交易行时出错: {str(e)}")
        
        return transactions
    
    def _parse_single_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析单个交易行"""
        try:
            transaction_details = self.clean_text(details_text)
            
            # 添加调试输出
            print(f"\n解析单行交易:")
            print(f"原始详情: '{details_text}'")
            print(f"清理后详情: '{transaction_details}'")
            print(f"存款: '{deposit_text}', 取款: '{withdrawal_text}', 余额: '{balance_text}'")
            
            if self.should_filter_transaction(transaction_details):
                print("交易被过滤")
                return None
            
            # 改进日期匹配逻辑
            date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', transaction_details)
            if not date_match:
                print("未找到日期匹配")
                return None
            
            date_value = self.parse_date(date_match.group(1))
            if not date_value:
                date_value = date_match.group(1)
            
            # 提取交易描述（去除日期部分）
            transaction_desc = transaction_details[date_match.end():].strip()
            if not transaction_desc:
                transaction_desc = transaction_details
            
            print(f"提取的日期: '{date_value}'")
            print(f"提取的描述: '{transaction_desc}'")
            
            transaction = {
                "账户类型": account_type,
                "银行名称": self.bank_name,
                "Date": date_value,
                "Transaction Details": transaction_desc,
                "Deposit": self.parse_amount(deposit_text),
                "Withdrawal": self.parse_amount(withdrawal_text),
                "Balance": self.parse_amount(balance_text),
                "Currency": self.extract_currency_from_details(transaction_details, account_type)
            }
            
            print(f"生成的交易记录: {transaction}")
            return transaction
        
        except Exception as e:
            logger.error(f"解析单交易行时出错: {str(e)}")
            print(f"解析错误: {str(e)}")
            return None
    
    def split_table_by_account_type(self, df, target_account_type):
        """根据账户类型分割表格，返回目标账户类型的行范围列表"""
        account_sections = []
        
        try:
            target_keywords = self.account_keywords.get(target_account_type, [])
            
            for idx, row in df.iterrows():
                row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()])
                
                for keyword in target_keywords:
                    if keyword in row_text:
                        end_idx = len(df) - 1
                        
                        # 查找下一个账户类型的开始位置
                        for next_idx in range(idx + 1, len(df)):
                            next_row_text = " ".join([str(cell) for cell in df.iloc[next_idx].values if pd.notna(cell) and str(cell).strip()])
                            
                            found_other_account = False
                            for other_account in self.target_account_types:
                                if other_account != target_account_type:
                                    other_keywords = self.account_keywords.get(other_account, [])
                                    for other_keyword in other_keywords:
                                        if other_keyword in next_row_text:
                                            end_idx = next_idx - 1
                                            found_other_account = True
                                            break
                                if found_other_account:
                                    break
                            
                            if found_other_account or "Total No. of Deposits" in next_row_text:
                                break
                        
                        account_sections.append((idx, end_idx))
                        break
        
        except Exception as e:
            logger.error(f"分割表格时出错: {str(e)}")
        
        return account_sections
    
    def parse(self, pdf_path_or_obj):
        """解析汇丰银行PDF账单"""
        all_transactions = []
        
        try:
            # 获取PDF路径和对象
            if isinstance(pdf_path_or_obj, str):
                pdf_path = pdf_path_or_obj
                with pdfplumber.open(pdf_path) as pdf:
                    pdf_obj = pdf
            else:
                pdf_obj = pdf_path_or_obj
                pdf_path = getattr(pdf_obj, 'stream', None)
                if hasattr(pdf_path, 'name'):
                    pdf_path = pdf_path.name
                else:
                    return all_transactions
            
            # 遍历账户类型
            for account_type in self.target_account_types:
                account_pages = self.find_account_pages(pdf_obj, account_type)
                
                if not account_pages:
                    continue
                
                # 从找到的页面提取交易记录
                for page_info in account_pages:
                    page_num = page_info['page']
                    transactions = self.extract_account_transactions(pdf_path, page_num, account_type, page_info)
                    
                    if transactions:
                        all_transactions.extend(transactions)
        
        except Exception as e:
            logger.error(f"解析汇丰银行PDF时出错: {str(e)}")
        
        return all_transactions
    

    
    def _identify_account_type_from_context(self, row_text, row_idx, df):
        """根据上下文识别交易数据实际属于的账户类型"""
        try:
            # 向上查找最近的账户类型标识
            for i in range(row_idx, -1, -1):
                context_row = " ".join([str(cell) for cell in df.iloc[i].values if pd.notna(cell) and str(cell).strip()])
                
                # 检查港币往来
                if any(keyword in context_row for keyword in self.account_keywords.get("港币往来", [])):
                    return "港币往来"
                
                # 检查港币储蓄
                if any(keyword in context_row for keyword in self.account_keywords.get("港币储蓄", [])):
                    return "港币储蓄"
                
                # 检查外币储蓄
                if any(keyword in context_row for keyword in self.account_keywords.get("外币储蓄", [])):
                    return "外币储蓄"
                
                # 如果找到了货币代码，可能是外币储蓄
                if re.search(r'\b(USD|GBP|EUR|JPY|CNY)\b', context_row):
                    return "外币储蓄"
            
            # 如果在当前行中包含货币代码，判断为外币储蓄
            if re.search(r'\b(USD|GBP|EUR|JPY|CNY)\b', row_text):
                return "外币储蓄"
            
            # 默认返回港币往来（如果无法确定）
            return "港币往来"
            
        except Exception as e:
            logger.error(f"识别账户类型时出错: {str(e)}")
            return "港币往来"  # 默认值
    

    
    def _parse_multi_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析包含多个交易的行"""
        transactions = []
        
        try:
            detail_parts = details_text.split("\n") if details_text else []
            deposit_parts = deposit_text.split("\n") if deposit_text else []
            withdrawal_parts = withdrawal_text.split("\n") if withdrawal_text else []
            balance_parts = balance_text.split("\n") if balance_text else []
            
            i = 0
            while i < len(detail_parts):
                detail_part = detail_parts[i].strip()
                
                if not detail_part:
                    i += 1
                    continue
                
                # 跳过货币代码行
                if re.match(r'^[A-Z]{3}$', detail_part):
                    i += 1
                    continue
                
                # 解析日期和交易描述
                date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', detail_part)
                if date_match:
                    date_value = self.parse_date(date_match.group(1))
                    if not date_value:
                        date_value = date_match.group(1)
                    
                    transaction_desc = detail_part[date_match.end():].strip()
                    
                    # 查找对应的金额和余额
                    deposit_amount = 0.0
                    withdrawal_amount = 0.0
                    balance_value = 0.0
                    
                    if i < len(deposit_parts) and deposit_parts[i].strip():
                        deposit_amount = self.parse_amount(deposit_parts[i])
                    
                    if i < len(withdrawal_parts) and withdrawal_parts[i].strip():
                        withdrawal_amount = self.parse_amount(withdrawal_parts[i])
                    
                    if i < len(balance_parts) and balance_parts[i].strip():
                        balance_value = self.parse_amount(balance_parts[i])
                    
                    # 特殊处理利息收入
                    if "CREDIT INTEREST" in transaction_desc or "利息收入" in transaction_desc:
                        amount_in_desc = re.search(r'([\d,]+\.\d{2})', transaction_desc)
                        if amount_in_desc and deposit_amount == 0.0:
                            deposit_amount = self.parse_amount(amount_in_desc.group(1))
                    
                    # 创建交易记录
                    transaction = {
                        "账户类型": account_type,
                        "银行名称": self.bank_name,
                        "Date": date_value,
                        "Transaction Details": transaction_desc,
                        "Deposit": deposit_amount,
                        "Withdrawal": withdrawal_amount,
                        "Balance": balance_value,
                        "Currency": self.extract_currency_from_details(transaction_desc, account_type)
                    }
                    
                    transactions.append(transaction)
                
                i += 1
        
        except Exception as e:
            logger.error(f"解析多交易行时出错: {str(e)}")
        
        return transactions
    
    def _parse_single_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析单个交易行"""
        try:
            transaction_details = self.clean_text(details_text)
            
            # 添加调试输出
            print(f"\n解析单行交易:")
            print(f"原始详情: '{details_text}'")
            print(f"清理后详情: '{transaction_details}'")
            print(f"存款: '{deposit_text}', 取款: '{withdrawal_text}', 余额: '{balance_text}'")
            
            if self.should_filter_transaction(transaction_details):
                print("交易被过滤")
                return None
            
            # 改进日期匹配逻辑
            date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', transaction_details)
            if not date_match:
                print("未找到日期匹配")
                return None
            
            date_value = self.parse_date(date_match.group(1))
            if not date_value:
                date_value = date_match.group(1)
            
            # 提取交易描述（去除日期部分）
            transaction_desc = transaction_details[date_match.end():].strip()
            if not transaction_desc:
                transaction_desc = transaction_details
            
            print(f"提取的日期: '{date_value}'")
            print(f"提取的描述: '{transaction_desc}'")
            
            transaction = {
                "账户类型": account_type,
                "银行名称": self.bank_name,
                "Date": date_value,
                "Transaction Details": transaction_desc,
                "Deposit": self.parse_amount(deposit_text),
                "Withdrawal": self.parse_amount(withdrawal_text),
                "Balance": self.parse_amount(balance_text),
                "Currency": self.extract_currency_from_details(transaction_details, account_type)
            }
            
            print(f"生成的交易记录: {transaction}")
            return transaction
        
        except Exception as e:
            logger.error(f"解析单交易行时出错: {str(e)}")
            print(f"解析错误: {str(e)}")
            return None
    
    def split_table_by_account_type(self, df, target_account_type):
        """根据账户类型分割表格，返回目标账户类型的行范围列表"""
        account_sections = []
        
        try:
            target_keywords = self.account_keywords.get(target_account_type, [])
            
            for idx, row in df.iterrows():
                row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()])
                
                for keyword in target_keywords:
                    if keyword in row_text:
                        end_idx = len(df) - 1
                        
                        # 查找下一个账户类型的开始位置
                        for next_idx in range(idx + 1, len(df)):
                            next_row_text = " ".join([str(cell) for cell in df.iloc[next_idx].values if pd.notna(cell) and str(cell).strip()])
                            
                            found_other_account = False
                            for other_account in self.target_account_types:
                                if other_account != target_account_type:
                                    other_keywords = self.account_keywords.get(other_account, [])
                                    for other_keyword in other_keywords:
                                        if other_keyword in next_row_text:
                                            end_idx = next_idx - 1
                                            found_other_account = True
                                            break
                                if found_other_account:
                                    break
                            
                            if found_other_account or "Total No. of Deposits" in next_row_text:
                                break
                        
                        account_sections.append((idx, end_idx))
                        break
        
        except Exception as e:
            logger.error(f"分割表格时出错: {str(e)}")
        
        return account_sections
    
    def parse(self, pdf_path_or_obj):
        """解析汇丰银行PDF账单"""
        all_transactions = []
        
        try:
            # 获取PDF路径和对象
            if isinstance(pdf_path_or_obj, str):
                pdf_path = pdf_path_or_obj
                with pdfplumber.open(pdf_path) as pdf:
                    pdf_obj = pdf
            else:
                pdf_obj = pdf_path_or_obj
                pdf_path = getattr(pdf_obj, 'stream', None)
                if hasattr(pdf_path, 'name'):
                    pdf_path = pdf_path.name
                else:
                    return all_transactions
            
            # 遍历账户类型
            for account_type in self.target_account_types:
                account_pages = self.find_account_pages(pdf_obj, account_type)
                
                if not account_pages:
                    continue
                
                # 从找到的页面提取交易记录
                for page_info in account_pages:
                    page_num = page_info['page']
                    transactions = self.extract_account_transactions(pdf_path, page_num, account_type, page_info)
                    
                    if transactions:
                        all_transactions.extend(transactions)
        
        except Exception as e:
            logger.error(f"解析汇丰银行PDF时出错: {str(e)}")
        
        return all_transactions
    
    def save_to_excel(self, transactions, output_path, account_info=None):
        """保存汇丰银行交易记录到Excel文件，按用户期望的格式"""
        try:
            if not transactions:
                return
            
            # 确保所有交易记录都有必要的字段
            for trans in transactions:
                if '银行' not in trans and account_info:
                    trans['银行'] = account_info.get('bank_name', self.bank_name)
                if '文件名' not in trans and account_info:
                    trans['文件名'] = account_info.get('file_name', '')
            
            # 创建DataFrame并重新排列列顺序以匹配用户期望
            df = pd.DataFrame(transactions)
            
            # 定义期望的列顺序
            expected_columns = [
                '账户类型', '银行名称', 'Date', 'Transaction Details', 
                'Deposit', 'Withdrawal', 'Balance', 'Currency', '银行', '文件名'
            ]
            
            # 确保所有期望的列都存在
            for col in expected_columns:
                if col not in df.columns:
                    df[col] = ''
            
            # 重新排列列顺序
            df = df[expected_columns]
            
            # 按账户类型分组保存
            if '账户类型' in df.columns and not df['账户类型'].isna().all():
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    # 保存汇总表（所有数据）
                    df.to_excel(writer, sheet_name='汇总', index=False)
                    
                    # 按账户类型分别保存
                    for account_type, group in df.groupby('账户类型'):
                        if not group.empty:
                            sheet_name = str(account_type).replace('/', '_').replace('\\', '_')[:31]
                            group.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # 如果没有账户类型分组，直接保存
                df.to_excel(output_path, index=False)
            
            logger.info(f"成功保存Excel文件: {output_path}")
            
        except Exception as e:
            logger.error(f"保存汇丰银行Excel文件时出错: {str(e)}")
            raise


# 玉山银行解析器（占位符）
class ESunBankParser(BankParser):
    """玉山银行账单解析器"""
    
    def parse(self, pdf):
        return []
    
    def save_to_excel(self, transactions, output_path, account_info=None):
        try:
            df = pd.DataFrame(transactions)
            df.to_excel(output_path, index=False)
        except Exception as e:
            logger.error(f"保存玉山银行Excel文件时出错: {str(e)}")


# 通用解析器
class GenericParser(BankParser):
    """通用银行账单解析器"""
    
    def parse(self, pdf):
        return []
    
    def save_to_excel(self, transactions, output_path, account_info=None):
        try:
            df = pd.DataFrame(transactions)
            df.to_excel(output_path, index=False)
        except Exception as e:
            logger.error(f"保存Excel文件时出错: {str(e)}")


# 银行解析器映射
BANK_PARSERS = {
    "玉山银行": ESunBankParser(),
    "汇丰银行": HSBCParser(),
    "渣打银行": GenericParser(),
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