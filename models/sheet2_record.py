"""Sheet2数据模型 - 明细记录"""

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class Sheet2Record:
    """Sheet2记录数据模型"""
    row_index: int  # 原始行号（用于排序输出）
    organization: Optional[str]
    date: datetime
    name: str
    daily_amount: float

    def __post_init__(self):
        """数据验证和转换"""
        if self.organization:
            self.organization = str(self.organization).strip()
            if self.organization == 'None' or self.organization == '':
                self.organization = None
        if self.name:
            self.name = str(self.name).strip()
        if self.daily_amount:
            self.daily_amount = float(self.daily_amount)
