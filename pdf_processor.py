import os
import logging
from fuzzywuzzy import fuzz
from bank_parsers import get_supported_banks

logger = logging.getLogger(__name__)

class PDFProcessor:
    """PDF处理器，用于银行类型检测"""
    
    def __init__(self):
        self.supported_banks = get_supported_banks()
        self.bank_keywords = {
            "玉山银行": ["玉山银行", "玉山", "E.SUN", "E. SUN", "E.SUN Bank", "E. SUN BANK", "E.SUN Commercial Bank", "E. SUN COMMERCIAL BANK", "ESUN", "ESUNHKHH", "玉山銀行"],
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
            # 导入pdfplumber用于临时文本提取
            import pdfplumber
            
            # 提取PDF文件的前几页文本用于识别
            text = ""
            with pdfplumber.open(pdf_path) as pdf:
                # 只读取前2页用于银行类型识别
                pages_to_extract = pdf.pages[:2]
                for page in pages_to_extract:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n\n"
            
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