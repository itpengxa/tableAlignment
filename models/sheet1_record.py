"""Sheet1数据模型 - 组织信息"""

from dataclasses import dataclass
from datetime import datetime


@dataclass
class Sheet1Record:
    """Sheet1记录数据模型"""
    row_index: int  # 行号（用于排序输出）
    organization: str
    date: datetime
    daily_amount: float
    attachment_amount: float
    difference: float

    def __post_init__(self):
        """数据验证和转换"""
        if self.organization:
            self.organization = str(self.organization).strip()
        if self.daily_amount:
            self.daily_amount = float(self.daily_amount)
        if self.attachment_amount:
            self.attachment_amount = float(self.attachment_amount)
        if self.difference:
            self.difference = float(self.difference)
