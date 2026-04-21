"""自定义异常类"""


class ProcessingError(Exception):
    """处理过程中的通用错误"""
    pass


class ValidationError(Exception):
    """数据验证错误"""
    pass


class FileFormatError(Exception):
    """文件格式错误"""
    pass


class ColumnDetectionError(ValidationError):
    """列名检测错误"""
    pass
