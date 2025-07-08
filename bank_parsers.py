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
        self.bank_name = "汇丰银行"
        # 修改为更简洁的中文账户类型标识
        self.target_account_types = [
            "港币往来",
            "港币储蓄", 
            "外币储蓄"
        ]
        
        # 同时保留英文关键词用于匹配
        self.account_keywords = {
            "港币往来": ["HKD Current", "港币往来", "HKD 往来"],
            "港币储蓄": ["HKD Savings", "港币储蓄", "HKD 储蓄"],
            "外币储蓄": ["Foreign Currency Savings", "外币储蓄", "外汇储蓄"]
        }
    
    def identify_account_sections(self, pdf):
        """识别PDF中的账户部分"""
        account_sections = []
        
        try:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                
                # 添加调试信息：输出每页的文本内容（前500字符）
                # logger.info(f"第{page_num + 1}页文本内容（前500字符）: {text[:500]}")
                
                # 查找目标账户类型的位置
                for account_type in self.target_account_types:
                    # 获取该账户类型的关键词列表
                    keywords = self.account_keywords.get(account_type, [account_type])
                    
                    # 尝试匹配任一关键词
                    found = False
                    for keyword in keywords:
                        if keyword in text:
                            account_sections.append({
                                'account_type': account_type,
                                'page': page_num + 1,
                                'page_obj': page
                            })
                            logger.info(f"在第{page_num + 1}页找到账户类型: {account_type} (匹配关键词: {keyword})")
                            found = True
                            break
        except Exception as e:
            logger.error(f"识别账户部分时出错: {str(e)}")
        

        return account_sections
        
    
    def extract_transactions_from_section_with_camelot(self, pdf_path, page_num, account_type):
        """使用Camelot从特定页面提取交易记录"""
        transactions = []
        
        try:
            # 使用Camelot提取指定页面的表格
            # 对于无框线表格，使用stream模式
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num),
                flavor='stream',  # 改为stream模式，适合无边框的表格
                table_areas=None,  # 自动检测表格区域
                columns=None,      # 自动检测列
                edge_tol=500,      # 增加边缘容忍度
                row_tol=10,        # 行容忍度
                column_tol=0       # 列容忍度
            )
            
            logger.info(f"Camelot在第{page_num}页找到{len(tables)}个表格")
            
            # 如果stream模式没有找到表格，尝试使用更宽松的参数
            if len(tables) == 0:
                logger.info("尝试使用更宽松的参数重新提取表格")
                tables = camelot.read_pdf(
                    pdf_path, 
                    pages=str(page_num),
                    flavor='stream',
                    table_areas=None,
                    columns=None,
                    edge_tol=1000,     # 更大的边缘容忍度
                    row_tol=20,        # 更大的行容忍度
                    column_tol=10,     # 更大的列容忍度
                    split_text=True    # 启用文本分割
                )
            
            for table_idx, table in enumerate(tables):
                df = table.df
                logger.info(f"处理表格{table_idx + 1}，形状: {df.shape}")
                logger.info(f"表格置信度: {table.accuracy}")
                
                # 输出表格的前几行用于调试
                logger.info(f"表格{table_idx + 1}前3行:\n{df.head(3).to_string()}")
                
                # 对于无框线表格，可能需要更灵活的表头识别
                header_row_idx = None
                for idx, row in df.iterrows():
                    row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell) and str(cell).strip()]).upper()
                    # 扩展关键词匹配
                    header_keywords = ["DATE", "TRANSACTION", "DETAILS", "DEPOSIT", "WITHDRAWAL", "BALANCE", "日期", "交易", "存款", "取款", "余额"]
                    if any(keyword in row_text for keyword in header_keywords):
                        header_row_idx = idx
                        logger.info(f"找到表头行{idx}: {row.values}")
                        break
                
                if header_row_idx is None:
                    # 如果没有找到明确的表头，尝试使用第一行作为表头
                    logger.warning(f"表格{table_idx + 1}未找到明确表头，尝试使用第一行")
                    header_row_idx = 0
                
                # 使用表头行作为列名
                headers = [self.clean_text(str(cell)) for cell in df.iloc[header_row_idx].values]
                logger.info(f"解析后的表头: {headers}")
                
                # 重新设置DataFrame的列名
                df.columns = headers
                
                # 删除表头行及之前的行
                df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
                
                # 查找关键列（更灵活的匹配）
                date_col = self._find_column_name(headers, ["Date", "日期", "交易日期", "Transaction Date"])
                details_col = self._find_column_name(headers, ["Transaction Details", "交易详情", "Details", "Description", "说明", "摘要"])
                deposit_col = self._find_column_name(headers, ["Deposit", "存款", "贷方", "Credit", "收入"])
                withdrawal_col = self._find_column_name(headers, ["Withdrawal", "取款", "借方", "Debit", "支出"])
                balance_col = self._find_column_name(headers, ["Balance", "余额", "结余"])
                
                logger.info(f"列名 - Date: {date_col}, Details: {details_col}, Deposit: {deposit_col}, Withdrawal: {withdrawal_col}, Balance: {balance_col}")
                
                # 处理数据行
                for idx, row in df.iterrows():
                    # 跳过空行或主要为空的行
                    non_empty_cells = [cell for cell in row.values if pd.notna(cell) and str(cell).strip()]
                    if len(non_empty_cells) < 2:  # 至少要有2个非空单元格
                        continue
                    
                    # 提取交易数据
                    transaction = {
                        "账户类型": account_type,
                        "银行名称": self.bank_name
                    }
                    
                    # Date
                    if date_col and date_col in row.index:
                        date_value = self.parse_date(row[date_col])
                        if date_value:  # 只有有效日期才继续处理
                            transaction["Date"] = date_value
                        else:
                            continue  # 跳过无效日期的行
                    else:
                        continue  # 没有日期列就跳过
                    
                    # Transaction Details
                    if details_col and details_col in row.index:
                        transaction["Transaction Details"] = self.clean_text(str(row[details_col]))
                    
                    # Deposit
                    if deposit_col and deposit_col in row.index:
                        transaction["Deposit"] = self.parse_amount(row[deposit_col])
                    else:
                        transaction["Deposit"] = 0.0
                    
                    # Withdrawal
                    if withdrawal_col and withdrawal_col in row.index:
                        transaction["Withdrawal"] = self.parse_amount(row[withdrawal_col])
                    else:
                        transaction["Withdrawal"] = 0.0
                    
                    # Balance
                    if balance_col and balance_col in row.index:
                        transaction["Balance"] = self.parse_amount(row[balance_col])
                    else:
                        transaction["Balance"] = 0.0
                    
                    # 添加交易记录
                    transactions.append(transaction)
                    logger.info(f"添加交易记录: {transaction}")
        
        except Exception as e:
            logger.error(f"使用Camelot从第{page_num}页提取交易记录时出错: {str(e)}")
            # 如果Camelot失败，回退到pdfplumber
            logger.info("回退到pdfplumber方法")
            return self.extract_transactions_from_section_fallback(page_num, account_type)
        
        return transactions
    
    def extract_transactions_from_section_fallback(self, page, account_type):
        """回退方法：使用pdfplumber提取表格"""
        transactions = []
        
        try:
            # 提取表格数据
            tables = page.extract_tables()
            logger.info(f"pdfplumber在{account_type}部分找到{len(tables)}个表格")
            
            for table_idx, table in enumerate(tables):
                if not table or len(table) <= 1:
                    continue
                
                logger.info(f"处理表格{table_idx + 1}，包含{len(table)}行")
                
                # 输出表格的前几行用于调试
                for i, row in enumerate(table[:3]):
                    logger.info(f"表格{table_idx + 1}第{i + 1}行: {row}")
                
                # 查找包含目标字段的表头行
                header_row = None
                for i, row in enumerate(table):
                    row_text = " ".join([str(cell) for cell in row if cell]).upper()
                    # 检查是否包含我们需要的字段
                    if any(field in row_text for field in ["Date", "Transaction Details", "Deposit", "Withdrawal", "Balance"]):
                        header_row = i
                        logger.info(f"找到表头行{i + 1}: {row}")
                        break
                
                if header_row is None:
                    logger.warning(f"表格{table_idx + 1}未找到有效表头")
                    continue
                
                # 解析表头
                headers = [self.clean_text(cell) for cell in table[header_row]]
                logger.info(f"解析后的表头: {headers}")
                
                # 查找关键列索引
                date_idx = self._find_column_index(headers, ["Date", "日期"])
                details_idx = self._find_column_index(headers, ["Transaction Details", "交易详情", "Details"])
                deposit_idx = self._find_column_index(headers, ["Deposit", "存款", "贷方"])
                withdrawal_idx = self._find_column_index(headers, ["Withdrawal", "取款", "借方"])
                balance_idx = self._find_column_index(headers, ["Balance", "余额"])
                
                logger.info(f"列索引 - Date: {date_idx}, Details: {details_idx}, Deposit: {deposit_idx}, Withdrawal: {withdrawal_idx}, Balance: {balance_idx}")
                
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
                    transaction = {
                        "账户类型": account_type,
                        "银行名称": self.bank_name
                    }
                    
                    # Date
                    if date_idx is not None and date_idx < len(row):
                        transaction["Date"] = self.parse_date(row[date_idx])
                    
                    # Transaction Details
                    if details_idx is not None and details_idx < len(row):
                        transaction["Transaction Details"] = self.clean_text(row[details_idx])
                    
                    # Deposit
                    if deposit_idx is not None and deposit_idx < len(row):
                        transaction["Deposit"] = self.parse_amount(row[deposit_idx])
                    else:
                        transaction["Deposit"] = 0.0
                    
                    # Withdrawal
                    if withdrawal_idx is not None and withdrawal_idx < len(row):
                        transaction["Withdrawal"] = self.parse_amount(row[withdrawal_idx])
                    else:
                        transaction["Withdrawal"] = 0.0
                    
                    # Balance
                    if balance_idx is not None and balance_idx < len(row):
                        transaction["Balance"] = self.parse_amount(row[balance_idx])
                    else:
                        transaction["Balance"] = 0.0
                    
                    # 只添加有有效日期的记录
                    if transaction.get("Date") and transaction["Date"] != "":
                        transactions.append(transaction)
                        logger.info(f"添加交易记录: {transaction}")
        
        except Exception as e:
            logger.error(f"从{account_type}部分提取交易记录时出错: {str(e)}")
        
        return transactions
    
    def parse(self, pdf_path_or_obj):
        """解析汇丰银行PDF账单"""
        all_transactions = []
        
        try:
            # 如果传入的是文件路径，先用pdfplumber打开进行文本识别
            if isinstance(pdf_path_or_obj, str):
                pdf_path = pdf_path_or_obj
                with pdfplumber.open(pdf_path) as pdf:
                    # 识别账户部分
                    account_sections = self.identify_account_sections(pdf)
            else:
                # 如果传入的是pdfplumber对象，需要获取文件路径
                pdf = pdf_path_or_obj
                pdf_path = getattr(pdf, 'stream', None)
                if hasattr(pdf_path, 'name'):
                    pdf_path = pdf_path.name
                else:
                    logger.error("无法获取PDF文件路径，无法使用Camelot")
                    return all_transactions
                
                account_sections = self.identify_account_sections(pdf)
            
            if not account_sections:
                logger.warning("未找到目标账户类型")
                return all_transactions
            
            # 从每个账户部分提取交易记录
            for section in account_sections:
                transactions = self.extract_transactions_from_section_with_camelot(
                    pdf_path,
                    section['page'], 
                    section['account_type']
                )
                all_transactions.extend(transactions)
                logger.info(f"从{section['account_type']}提取了{len(transactions)}条交易记录")
        
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
    
    def _find_column_index(self, headers, possible_names):
        """查找列索引"""
        for name in possible_names:
            for i, header in enumerate(headers):
                if name.upper() in str(header).upper():
                    return i
        return None

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