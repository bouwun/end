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
    """汇丰银行账单解析器 - 专门提取三种账户类型的交易数据"""
    
    def __init__(self):
        super().__init__()
        self.bank_name = "汇丰银行"
        
        # 定义目标账户类型和对应关键词
        self.target_account_types = ["港币往来", "港币储蓄", "外币储蓄"]
        self.account_keywords = {
            "港币往来": ["港币往来", "HKD Current", "HKD CURRENT"],
            "港币储蓄": ["港币储蓄", "HKD Savings", "HKD SAVINGS"],
            "外币储蓄": ["外币储蓄", "Foreign Currency Savings", "FOREIGN CURRENCY SAVINGS"]
        }
        
        # 定义需要过滤的关键词
        self.filter_keywords = ["Total", "Exchange", "The", "HSBC", "多谢", "Thank", "Statement", "Balance Brought Forward", "Balance Carried Forward"]
    
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
            # 使用Camelot提取表格
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num),
                flavor='stream',
                table_areas=None,
                columns=None,
                edge_tol=500,
                row_tol=10,
                column_tol=0
            )
            
            logger.info(f"Camelot在第{page_num}页找到{len(tables)}个表格")
            
            # 如果没有找到表格，尝试更宽松的参数
            if len(tables) == 0:
                logger.info("尝试使用更宽松的参数重新提取表格")
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
            
            # 处理每个表格
            for table_idx, table in enumerate(tables):
                df = table.df
                logger.info(f"处理{account_type}的表格{table_idx + 1}，形状: {df.shape}")
                
                # 打印表格内容
                logger.info(f"表格{table_idx + 1}内容:")
                logger.info(f"表格DataFrame:\n{df.to_string()}")
                
                # 检查表格是否包含交易数据的特征
                table_text = df.to_string()
                logger.info(f"表格{table_idx + 1}文本内容:\n{table_text}")
                
                has_transaction_features = self.check_transaction_features(table_text)
                
                if not has_transaction_features:
                    logger.info(f"表格{table_idx + 1}不包含交易数据特征，跳过")
                    continue
                
                # 更精确的表格过滤：检查表格内容是否属于当前账户类型的文本区域
                if not self.is_table_in_account_section(table_text, account_type, page_info):
                    logger.info(f"表格{table_idx + 1}不在{account_type}账户的文本区域内，跳过")
                    continue
                
                logger.info(f"表格{table_idx + 1}通过所有过滤条件，开始提取交易记录")
                
                # 提取表格中的交易记录
                table_transactions = self.extract_transactions_from_table(df, account_type, table_idx + 1)
                transactions.extend(table_transactions)
                
                logger.info(f"从表格{table_idx + 1}提取了{len(table_transactions)}条交易记录")
        
        except Exception as e:
            logger.error(f"从第{page_num}页提取{account_type}交易记录时出错: {str(e)}")
            # 尝试使用pdfplumber作为回退
            transactions = self.extract_transactions_fallback(pdf_path, page_num, account_type, page_info)
        
        return transactions
    
    def is_table_in_account_section(self, table_text, account_type, page_info):
        """检查表格是否在指定账户类型的文本区域内"""
        # 检查表格是否包含明显的排除标志
        exclude_keywords = ["Statement Summary", "账单摘要", "Fee Schedule", "收费表", "Portfolio Summary", "资产摘要"]
        table_text_upper = table_text.upper()
        
        for exclude_keyword in exclude_keywords:
            if exclude_keyword.upper() in table_text_upper:
                logger.info(f"表格包含排除关键词'{exclude_keyword}'，跳过")
                return False
        
        # 检查表格是否包含当前账户类型的关键词
        current_keywords = self.account_keywords.get(account_type, [])
        has_current_keyword = False
        for keyword in current_keywords:
            if keyword.upper() in table_text_upper:
                has_current_keyword = True
                logger.info(f"表格包含{account_type}关键词'{keyword}'")
                break
        
        # 如果包含当前账户类型的关键词，则接受（即使同时包含其他账户类型）
        if has_current_keyword:
            logger.info(f"表格包含{account_type}关键词，接受处理")
            return True
        
        # 如果不包含当前账户类型关键词，则检查是否在文本区域内
        text_section = page_info.get('text_section', '')
        
        # 简单的文本重叠检查：如果表格中的某些关键内容出现在账户文本区域中
        table_lines = table_text.split('\n')
        for line in table_lines[:5]:  # 检查表格前5行
            line = line.strip()
            if len(line) > 10 and line in text_section:  # 长度大于10的行且在文本区域中
                logger.info(f"表格内容与{account_type}文本区域重叠，接受")
                return True
        
        logger.info(f"表格不属于{account_type}账户区域，跳过")
        return False
    
    def check_transaction_features(self, table_text):
        """检查表格是否包含交易数据的特征"""
        # 检查表格文本中是否包含交易指示词
        transaction_indicators = [
            "Date", "日期", "Transaction", "交易", "Details", "详情", 
            "Deposit", "存款", "Withdrawal", "取款", "Balance", "余额",
            "Debit", "借方", "Credit", "贷方", "Amount", "金额"
        ]
        
        table_text_upper = table_text.upper()
        
        # 至少需要包含2个交易指示词才认为是交易表格
        indicator_count = 0
        found_indicators = []
        for indicator in transaction_indicators:
            if indicator.upper() in table_text_upper:
                indicator_count += 1
                found_indicators.append(indicator)
                if indicator_count >= 2:
                    logger.info(f"表格包含交易特征关键词: {found_indicators}")
                    return True
        
        logger.info(f"表格不包含足够的交易特征关键词，仅找到: {found_indicators}")
        return False
    
    def extract_transactions_from_table(self, df, account_type, table_idx):
        """从表格中提取指定账户类型的交易记录"""
        transactions = []
        
        try:
            # 使用split_table_by_account_type方法分割表格
            account_sections = self.split_table_by_account_type(df, account_type)
            
            if not account_sections:
                logger.warning(f"在表格{table_idx}中未找到{account_type}账户数据")
                return transactions
            
            # 处理每个账户部分
            for section_start, section_end in account_sections:
                logger.info(f"处理{account_type}账户数据：第{section_start}行到第{section_end}行")
                
                # 提取该部分的数据
                section_df = df.iloc[section_start:section_end+1].copy().reset_index(drop=True)
                
                # 查找表头行
                header_row_idx = None
                for idx, row in section_df.iterrows():
                    row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()]).upper()
                    header_keywords = ["DATE", "TRANSACTION", "DETAILS", "DEPOSIT", "WITHDRAWAL", "BALANCE", "日期", "交易", "存款", "取款", "余额", "CCY"]
                    if any(keyword in row_text for keyword in header_keywords):
                        header_row_idx = idx
                        logger.info(f"在{account_type}部分找到表头行{idx}: {row.values}")
                        break
                
                if header_row_idx is None:
                    logger.warning(f"{account_type}部分未找到明确表头，跳过")
                    continue
                
                # 设置列名
                headers = [self.clean_text(str(cell)) for cell in section_df.iloc[header_row_idx].values]
                section_df.columns = headers
                data_df = section_df.iloc[header_row_idx + 1:].reset_index(drop=True)
                
                # 查找关键列
                date_col = self._find_column_name(headers, ["Date", "日期", "交易日期", "Transaction Date", "CCY Date Transaction Details", "货币 日期 进支详情"])
                details_col = self._find_column_name(headers, ["Transaction Details", "交易详情", "Details", "Description", "说明", "摘要", "CCY Date Transaction Details", "货币 日期 进支详情"])
                deposit_col = self._find_column_name(headers, ["Deposit", "存款", "贷方", "Credit", "收入", "存入"])
                withdrawal_col = self._find_column_name(headers, ["Withdrawal", "取款", "借方", "Debit", "支出", "提取"])
                balance_col = self._find_column_name(headers, ["Balance", "余额", "结余"])
                
                logger.info(f"{account_type}列名 - Date: {date_col}, Details: {details_col}, Deposit: {deposit_col}, Withdrawal: {withdrawal_col}, Balance: {balance_col}")
                
                # 处理数据行
                for idx, row in data_df.iterrows():
                    # 获取原始数据
                    raw_details = str(row[details_col]) if details_col and details_col in row.index else ""
                    raw_deposit = str(row[deposit_col]) if deposit_col and deposit_col in row.index else ""
                    raw_withdrawal = str(row[withdrawal_col]) if withdrawal_col and withdrawal_col in row.index else ""
                    raw_balance = str(row[balance_col]) if balance_col and balance_col in row.index else ""
                    
                    logger.info(f"处理第{idx}行原始数据: Details='{raw_details}', Deposit='{raw_deposit}', Withdrawal='{raw_withdrawal}', Balance='{raw_balance}'")
                    
                    # 检查是否包含换行符，表示多个交易
                    if "\n" in raw_details or "\n" in raw_balance:
                        transactions.extend(self._parse_multi_transaction_row(
                            raw_details, raw_deposit, raw_withdrawal, raw_balance, account_type
                        ))
                    else:
                        # 处理单个交易
                        transaction = self._parse_single_transaction_row(
                            raw_details, raw_deposit, raw_withdrawal, raw_balance, account_type
                        )
                        if transaction:
                            transactions.append(transaction)
        
        except Exception as e:
            logger.error(f"从表格{table_idx}提取{account_type}交易记录时出错: {str(e)}")
        
        return transactions
    
    def _parse_multi_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析包含多个交易的行"""
        transactions = []
        
        try:
            # 分割各个字段
            detail_parts = details_text.split("\n") if details_text else []
            deposit_parts = deposit_text.split("\n") if deposit_text else []
            withdrawal_parts = withdrawal_text.split("\n") if withdrawal_text else []
            balance_parts = balance_text.split("\n") if balance_text else []
            
            logger.info(f"分割后的数据: Details={detail_parts}, Deposit={deposit_parts}, Withdrawal={withdrawal_parts}, Balance={balance_parts}")
            
            # 处理每个交易部分
            i = 0
            while i < len(detail_parts):
                detail_part = detail_parts[i].strip()
                
                if not detail_part:
                    i += 1
                    continue
                
                # 检查是否为货币代码行（如USD, GBP等）
                if re.match(r'^[A-Z]{3}$', detail_part):
                    currency = detail_part
                    i += 1
                    continue
                
                # 尝试解析日期和交易描述
                date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', detail_part)
                if date_match:
                    date_value = self.parse_date(date_match.group(1))
                    if not date_value:
                        date_value = date_match.group(1)
                    
                    # 提取交易描述（日期后的部分）
                    transaction_desc = detail_part[date_match.end():].strip()
                    
                    # 查找对应的金额和余额
                    deposit_amount = 0.0
                    withdrawal_amount = 0.0
                    balance_value = 0.0
                    
                    # 从存款列查找金额
                    if i < len(deposit_parts) and deposit_parts[i].strip():
                        deposit_amount = self.parse_amount(deposit_parts[i])
                    
                    # 从取款列查找金额
                    if i < len(withdrawal_parts) and withdrawal_parts[i].strip():
                        withdrawal_amount = self.parse_amount(withdrawal_parts[i])
                    
                    # 从余额列查找余额
                    if i < len(balance_parts) and balance_parts[i].strip():
                        balance_value = self.parse_amount(balance_parts[i])
                    
                    # 特殊处理利息收入
                    if "CREDIT INTEREST" in transaction_desc or "利息收入" in transaction_desc:
                        # 在交易描述中查找金额
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
                    logger.info(f"添加{account_type}多行交易记录: {transaction}")
                
                i += 1
        
        except Exception as e:
            logger.error(f"解析多交易行时出错: {str(e)}")
        
        return transactions
    
    def _parse_single_transaction_row(self, details_text, deposit_text, withdrawal_text, balance_text, account_type):
        """解析单个交易行"""
        try:
            transaction_details = self.clean_text(details_text)
            
            # 过滤不需要的记录
            if self.should_filter_transaction(transaction_details):
                return None
            
            # 尝试解析日期
            date_match = re.search(r'(\d{1,2}\s+[A-Za-z]{3})', transaction_details)
            if not date_match:
                return None
            
            date_value = self.parse_date(date_match.group(1))
            if not date_value:
                date_value = date_match.group(1)
            
            # 创建交易记录
            transaction = {
                "账户类型": account_type,
                "银行名称": self.bank_name,
                "Date": date_value,
                "Transaction Details": transaction_details,
                "Deposit": self.parse_amount(deposit_text),
                "Withdrawal": self.parse_amount(withdrawal_text),
                "Balance": self.parse_amount(balance_text),
                "Currency": self.extract_currency_from_details(transaction_details, account_type)
            }
            
            logger.info(f"添加{account_type}单行交易记录: {transaction}")
            return transaction
        
        except Exception as e:
            logger.error(f"解析单交易行时出错: {str(e)}")
            return None
    
    def split_table_by_account_type(self, df, target_account_type):
        """根据账户类型分割表格，返回目标账户类型的行范围列表"""
        account_sections = []
        
        try:
            # 查找目标账户类型的起始行
            target_keywords = self.account_keywords.get(target_account_type, [])
            
            for idx, row in df.iterrows():
                row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()])
                
                # 检查是否包含目标账户类型关键词
                for keyword in target_keywords:
                    if keyword in row_text:
                        logger.info(f"在第{idx}行找到{target_account_type}账户标题: {keyword}")
                        
                        # 查找该账户部分的结束位置
                        end_idx = len(df) - 1  # 默认到表格末尾
                        
                        # 查找下一个账户类型的开始位置或"Total No. of Deposits"标记
                        for next_idx in range(idx + 1, len(df)):
                            next_row_text = " ".join([str(cell) for cell in df.iloc[next_idx].values if pd.notna(cell) and str(cell).strip()])
                            
                            # 检查是否遇到其他账户类型
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
                            
                            if found_other_account:
                                break
                            
                            # 检查是否遇到"Total No. of Deposits"标记
                            if "Total No. of Deposits" in next_row_text:
                                end_idx = next_idx
                                break
                        
                        account_sections.append((idx, end_idx))
                        logger.info(f"{target_account_type}账户数据范围: 第{idx}行到第{end_idx}行")
                        break
        
        except Exception as e:
            logger.error(f"分割表格时出错: {str(e)}")
        
        return account_sections
    
    def should_filter_transaction(self, transaction_details):
        """判断是否过滤交易记录"""
        if not transaction_details or transaction_details.strip() == "":
            return True
        
        # 检查是否包含过滤关键词
        for keyword in self.filter_keywords:
            if keyword.lower() in transaction_details.lower():
                return True
        
        return False
    
    def extract_currency_from_details(self, transaction_details, account_type):
        """从交易详情中提取货币类型"""
        # 港币账户默认为HKD
        if "港币" in account_type:
            return "HKD"
        
        # 外币账户需要从详情中提取
        if "外币" in account_type:
            # 查找CCY字段
            ccy_match = re.search(r'CCY[:\s]*([A-Z]{3})', transaction_details)
            if ccy_match:
                return ccy_match.group(1)
            
            # 查找常见货币代码
            currency_codes = ["USD", "EUR", "GBP", "JPY", "CNY", "SGD", "AUD", "CAD"]
            for code in currency_codes:
                if code in transaction_details.upper():
                    return code
            
            # 默认为USD
            return "USD"
        
        return "HKD"  # 默认货币
    
    def extract_transactions_fallback(self, pdf_path, page_num, account_type, page_info):
        """使用pdfplumber作为回退方案"""
        transactions = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if page_num <= len(pdf.pages):
                    page = pdf.pages[page_num - 1]
                    tables = page.extract_tables()
                    
                    for table in tables:
                        if not table:
                            continue
                        
                        # 类似的处理逻辑...
                        # 这里可以复用之前的fallback逻辑
                        pass
        
        except Exception as e:
            logger.error(f"使用pdfplumber提取交易记录时出错: {str(e)}")
        
        return transactions
    
    def parse(self, pdf_path_or_obj):
        """解析汇丰银行PDF账单 - 使用两层循环结构"""
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
                    logger.error("无法获取PDF文件路径，无法使用Camelot")
                    return all_transactions
            
            # 第一层循环：遍历账户类型
            for account_type in self.target_account_types:
                logger.info(f"开始处理账户类型: {account_type}")
                
                # 查找当前账户类型的页面和位置
                account_pages = self.find_account_pages(pdf_obj, account_type)
                
                if not account_pages:
                    logger.info(f"未找到{account_type}账户")
                    continue
                
                # 第二层循环：从找到的页面提取交易记录
                for page_info in account_pages:
                    page_num = page_info['page']
                    logger.info(f"从第{page_num}页提取{account_type}账户的交易记录")
                    
                    # 提取该页面该账户类型的交易记录
                    transactions = self.extract_account_transactions(pdf_path, page_num, account_type, page_info)
                    
                    if transactions:
                        all_transactions.extend(transactions)
                        logger.info(f"从第{page_num}页的{account_type}账户提取了{len(transactions)}条交易记录")
                    else:
                        logger.info(f"第{page_num}页的{account_type}账户未提取到交易记录")
                
                logger.info(f"{account_type}账户总共提取了{len([t for t in all_transactions if t.get('账户类型') == account_type])}条交易记录")
        
        except Exception as e:
            logger.error(f"解析汇丰银行PDF时出错: {str(e)}")
        
        return all_transactions
    
    def save_to_excel(self, transactions, output_path, account_info=None):
        """保存汇丰银行交易记录到Excel文件"""
        try:
            if not transactions:
                logger.warning("没有交易记录可保存")
                return
            
            # 创建DataFrame
            df = pd.DataFrame(transactions)
            
            # 按账户类型分组保存
            if '账户类型' in df.columns:
                # 创建ExcelWriter对象
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    # 按账户类型分组
                    for account_type, group in df.groupby('账户类型'):
                        # 清理工作表名称（Excel工作表名称不能包含某些特殊字符）
                        sheet_name = str(account_type).replace('/', '_').replace('\\', '_')[:31]
                        group.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"保存{account_type}账户的{len(group)}条记录到工作表'{sheet_name}'")
                    
                    # 创建汇总工作表
                    df.to_excel(writer, sheet_name='汇总', index=False)
                    logger.info(f"保存汇总的{len(df)}条记录到工作表'汇总'")
            else:
                # 如果没有账户类型列，直接保存
                df.to_excel(output_path, index=False)
                logger.info(f"保存{len(df)}条交易记录到{output_path}")
            
            logger.info(f"成功保存汇丰银行交易记录到: {output_path}")
            
        except Exception as e:
            logger.error(f"保存汇丰银行Excel文件时出错: {str(e)}")
            raise
    
    def _find_column_name(self, headers, possible_names):
        """查找列名"""
        for name in possible_names:
            for header in headers:
                if name.upper() in str(header).upper():
                    return header
        return None


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