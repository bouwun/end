class PDFProcessingError(Exception):
    """PDF处理过程中的错误基类"""
    pass

class BankDetectionError(PDFProcessingError):
    """银行类型检测错误"""
    pass