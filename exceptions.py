class PDFProcessingError(Exception):
    """PDF处理过程中的错误基类"""
    pass

class BankDetectionError(PDFProcessingError):
    """银行类型检测错误"""
    pass

class TableExtractionError(PDFProcessingError):
    """表格提取错误"""
    pass

class TransactionParsingError(PDFProcessingError):
    """交易记录解析错误"""
    pass

class OutputGenerationError(Exception):
    """输出生成错误"""
    pass