import os
import json
import logging
from datetime import datetime

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".trea_py_bank_config.json")

# 日志文件目录
LOG_DIR = os.path.join(os.path.expanduser("~"), "trea_py_bank_logs")

def setup_logging():
    """设置日志"""
    # 创建日志目录
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    
    # 设置日志文件名
    log_file = os.path.join(LOG_DIR, f"trea_py_bank_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    # 配置日志
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # 清除现有的处理器
    for handler in logger.handlers[:]: 
        logger.removeHandler(handler)
    
    # 文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 设置格式
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    # 记录启动信息
    logger.info("=== 应用程序启动 ===")
    logger.info(f"日志文件: {log_file}")
    
    return logger

def get_config():
    """获取配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"读取配置文件失败: {str(e)}")
    
    # 返回默认配置
    return {
        "last_pdf_dir": os.path.expanduser("~"),
        "last_output_dir": os.path.expanduser("~"),
        "last_output_file": "",
        "bank_mapping": {}
    }

def save_config(config):
    """保存配置"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        logging.error(f"保存配置文件失败: {str(e)}")
        return False