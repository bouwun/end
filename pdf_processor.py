import os
import re
import logging
from datetime import datetime
import pdfplumber
from fuzzywuzzy import fuzz
from bank_parsers import get_bank_parser, get_supported_banks

logger = logging.getLogger(__name__)

class PDFProcessor:
    """PDF处理器，用于处理银行账单PDF文件"""
    
    def __init__(self):
        self.supported_banks = get_supported_banks()
        self.bank_keywords = {
            "工商银行": ["中国工商银行", "工行", "ICBC", "Industrial and Commercial Bank of China"],
            "建设银行": ["中国建设银行", "建行", "CCB", "China Construction Bank"],
            "农业银行": ["中国农业银行", "农行", "ABC", "Agricultural Bank of China"],
            "中国银行": ["中行", "BOC", "Bank of China"],
            "交通银行": ["交行", "BOCOM", "Bank of Communications"],
            "招商银行": ["招行", "CMB", "China Merchants Bank"],
            "浦发银行": ["上海浦东发展银行", "浦发", "SPDB", "Shanghai Pudong Development Bank"],
            "民生银行": ["中国民生银行", "民生", "CMBC", "China Minsheng Bank"],
            "中信银行": ["中信", "CITIC", "China CITIC Bank"],
            "光大银行": ["中国光大银行", "光大", "CEB", "China Everbright Bank"],
            "华夏银行": ["华夏", "HXB", "Huaxia Bank"],
            "广发银行": ["广发", "CGB", "China Guangfa Bank"],
            "平安银行": ["平安", "PAB", "Ping An Bank"],
            "邮储银行": ["中国邮政储蓄银行", "邮储", "PSBC", "Postal Savings Bank of China"],
            "玉山银行": ["中国玉山银行", "玉山", "EBS", "Eastern Bank of Shandong"],
            "渣打银行": ["渣打", "SDB", "Shanghaidi Bank"],
            "汇丰银行": ["中国汇丰银行", "汇丰", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "南洋银行": ["中国南洋银行", "南洋", "NBC", "National Bank of China"],
            "恒生银行": ["中国恒生银行", "恒生", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "中银香港": ["中国中银香港", "中银", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "渣打银行": ["渣打", "SDB", "Shanghaidi Bank"],
            "汇丰银行": ["中国汇丰银行", "汇丰", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "南洋银行": ["中国南洋银行", "南洋", "NBC", "National Bank of China"],
            "恒生银行": ["中国恒生银行", "恒生", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "中银香港": ["中国中银香港", "中银", "HSBC", "HongKong and Shanghai Banking Corporation"],
            "东亚银行": ["东亚银行", "东亚", "Eastern Asia Bank"]
        }
    
    def detect_bank_type(self, pdf_path, bank_mapping=None):
        """检测PDF文件的银行类型"""
        try:
            # 提取PDF文件的前几页文本用于识别
            text = self.extract_text_from_pdf(pdf_path, max_pages=2)
            
            # 首先检查用户自定义的映射
            if bank_mapping:
                for bank_name, keywords in bank_mapping.items():
                    for keyword in keywords:
                        if keyword.lower() in text.lower():
                            return bank_name
            
            # 使用模糊匹配找出最可能的银行
            best_match = None
            best_score = 0
            
            for bank_name, keywords in self.bank_keywords.items():
                # 计算文本与每个关键词的匹配度
                for keyword in keywords:
                    # 直接检查关键词是否在文本中（不区分大小写）
                    if keyword.lower() in text.lower():
                        return bank_name
                    
                    # 使用模糊匹配计算相似度
                    score = fuzz.partial_ratio(keyword.lower(), text.lower())
                    if score > best_score:
                        best_score = score
                        best_match = bank_name
            
            # 如果最佳匹配分数超过阈值，返回匹配的银行名称
            if best_score > 80:
                return best_match
            
            # 尝试从文件名识别
            filename = os.path.basename(pdf_path).lower()
            for bank_name, keywords in self.bank_keywords.items():
                for keyword in keywords:
                    if keyword.lower() in filename:
                        return bank_name
            
            # 无法识别
            return "未知"
        
        except Exception as e:
            logger.error(f"检测银行类型时出错: {str(e)}")
            return "未知"
    
    def extract_text_from_pdf(self, pdf_path, max_pages=None, page_numbers=None):
        """从PDF文件中提取文本
        
        Args:
            pdf_path: PDF文件路径
            max_pages: 最大页数限制
            page_numbers: 指定页码列表，如果提供则忽略max_pages
            
        Returns:
            提取的文本字符串
        """
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # 确定要处理的页面
                if page_numbers:
                    pages_to_extract = [pdf.pages[i] for i in page_numbers if i < len(pdf.pages)]
                else:
                    pages_to_extract = pdf.pages[:max_pages] if max_pages else pdf.pages
                
                # 提取文本
                for page in pages_to_extract:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n\n"  # 添加页面分隔符
                
            return text
        except Exception as e:
            logger.error(f"提取PDF文本时出错: {str(e)}")
            return ""
    
    def process_pdf(self, pdf_path, bank_parser):
        """处理PDF文件并提取交易记录"""
        try:
            # 打开PDF文件
            with pdfplumber.open(pdf_path) as pdf:
                # 调用银行特定的解析器处理PDF
                transactions = bank_parser.parse(pdf)
                
                # 标准化交易记录
                standardized_transactions = self.standardize_transactions(transactions)
                
                return standardized_transactions
        
        except Exception as e:
            logger.exception(f"处理PDF文件时出错: {str(e)}")
            raise
    
    def standardize_transactions(self, transactions):
        """标准化交易记录格式"""
        standardized = []
        
        for trans in transactions:
            # 创建标准化的交易记录
            std_trans = {}
            
            # 复制原始字段
            for key, value in trans.items():
                std_trans[key] = value
            
            # 确保关键字段存在
            if "交易日期" in std_trans and isinstance(std_trans["交易日期"], str):
                # 尝试将日期字符串转换为日期对象
                try:
                    # 尝试多种常见的日期格式
                    date_formats = [
                        "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日",
                        "%Y.%m.%d", "%d-%m-%Y", "%d/%m/%Y"
                    ]
                    
                    date_str = std_trans["交易日期"]
                    for fmt in date_formats:
                        try:
                            date_obj = datetime.strptime(date_str, fmt)
                            std_trans["交易日期"] = date_obj.strftime("%Y-%m-%d")
                            break
                        except ValueError:
                            continue
                except Exception:
                    # 如果转换失败，保留原始字符串
                    pass
            
            # 处理金额字段，确保是数值类型
            for amount_field in ["交易金额", "收入金额", "支出金额", "账户余额"]:
                if amount_field in std_trans:
                    try:
                        # 移除金额中的非数字字符（保留小数点和负号）
                        amount_str = str(std_trans[amount_field])
                        amount_str = re.sub(r'[^\d.-]', '', amount_str)
                        std_trans[amount_field] = float(amount_str) if amount_str else 0.0
                    except (ValueError, TypeError):
                        # 如果转换失败，设为0
                        std_trans[amount_field] = 0.0
            
            # 如果没有明确的收入/支出金额，但有交易金额，则根据金额正负判断
            if "交易金额" in std_trans and "收入金额" not in std_trans and "支出金额" not in std_trans:
                amount = float(std_trans["交易金额"])
                if amount > 0:
                    std_trans["收入金额"] = amount
                    std_trans["支出金额"] = 0.0
                else:
                    std_trans["收入金额"] = 0.0
                    std_trans["支出金额"] = abs(amount)
            
            standardized.append(std_trans)
        
        return standardized