"""Excel文件读取器"""

from datetime import datetime
from pathlib import Path
from typing import List, Tuple

import pandas as pd

from models import Sheet1Record, Sheet2Record
from utils.exceptions import FileFormatError, ValidationError, ColumnDetectionError


class ExcelReader:
    """Excel文件读取器"""

    # 支持的列名模式（新增向后兼容）
    SUPPORTED_COLUMN_PATTERNS = {
        'date_columns': ['日期', '入账日期'],
        'amount_columns': ['当日金额', '金额']
    }

    def __init__(self, file_path: Path):
        """初始化读取器

        Args:
            file_path: Excel文件路径

        Raises:
            FileNotFoundError: 文件不存在
            FileFormatError: 文件格式错误
        """
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        if file_path.suffix.lower() != '.xlsx':
            raise FileFormatError(f"只支持.xlsx格式，当前文件: {file_path.suffix}")

        self.file_path = file_path

    def read_excel(self) -> Tuple[List[Sheet1Record], List[Sheet2Record]]:
        """读取Excel文件的两个sheet

        Returns:
            Tuple[List[Sheet1Record], List[Sheet2Record]]: Sheet1和Sheet2的数据记录

        Raises:
            ValidationError: 数据验证错误
        """
        try:
            # 读取sheet1
            sheet1_df = pd.read_excel(self.file_path, sheet_name=0)
            sheet1_records = self._parse_sheet1(sheet1_df)

            # 读取sheet2
            sheet2_df = pd.read_excel(self.file_path, sheet_name=1)
            sheet2_records = self._parse_sheet2(sheet2_df)

            return sheet1_records, sheet2_records

        except Exception as e:
            raise ValidationError(f"读取Excel文件失败: {str(e)}")

    def _parse_sheet1(self, df: pd.DataFrame) -> List[Sheet1Record]:
        """解析Sheet1数据

        Args:
            df: Sheet1的DataFrame

        Returns:
            List[Sheet1Record]: Sheet1记录列表

        Raises:
            ValidationError: 数据格式错误
        """
        records = []

        # 检查必需的列
        required_columns = ['组织', '日期', '当日金额', '附件金', '差异']
        if len(df.columns) < len(required_columns):
            raise ValidationError(f"Sheet1格式错误，需要{len(required_columns)}列，实际{len(df.columns)}列")

        for idx, row in df.iterrows():
            try:
                # 跳过空行
                if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]):
                    continue

                # 解析数据
                organization = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
                date = self._parse_date(row.iloc[1])
                daily_amount = float(row.iloc[2]) if not pd.isna(row.iloc[2]) else 0.0
                attachment_amount = float(row.iloc[3]) if not pd.isna(row.iloc[3]) else 0.0
                difference = float(row.iloc[4]) if not pd.isna(row.iloc[4]) else 0.0

                record = Sheet1Record(
                    row_index=idx,
                    organization=organization,
                    date=date,
                    daily_amount=daily_amount,
                    attachment_amount=attachment_amount,
                    difference=difference
                )
                records.append(record)

            except (ValueError, IndexError, TypeError) as e:
                raise ValidationError(f"Sheet1第{idx + 2}行数据格式错误: {str(e)}")

        return records

    def _parse_sheet2(self, df: pd.DataFrame) -> List[Sheet2Record]:
        """解析Sheet2数据（增强版本 - 支持动态列名识别）

        自动检测并适配不同的Excel列名格式：
        - 旧格式：组织、日期、姓名、当日金额
        - 新格式：组织、入账日期、姓名、金额

        Args:
            df: Sheet2的DataFrame

        Returns:
            List[Sheet2Record]: Sheet2记录列表

        Raises:
            ValidationError: 数据格式错误
            ColumnDetectionError: 无法检测到匹配的列名
        """
        records = []

        try:
            # 动态检测列名映射
            column_mapping = self._detect_column_mapping(df)
        except ColumnDetectionError as e:
            raise ValidationError(f"Sheet2格式错误: {str(e)}")
        except Exception as e:
            raise ValidationError(f"Sheet2格式错误: 无法识别列名 - {str(e)}")

        # 验证必需的最小列数
        min_columns = max(
            column_mapping['organization'],
            column_mapping['date'],
            column_mapping['name'],
            column_mapping['amount']
        ) + 1

        if len(df.columns) < min_columns:
            raise ValidationError(f"Sheet2格式错误，需要至少{min_columns}列，实际{len(df.columns)}列")

        for idx, row in df.iterrows():
            try:
                # 跳过空行（检查关键列是否都为NA）
                if (pd.isna(row.iloc[column_mapping['organization']]) and
                    pd.isna(row.iloc[column_mapping['date']]) and
                    pd.isna(row.iloc[column_mapping['name']])):
                    continue

                # 使用动态列映射解析数据
                organization = row.iloc[column_mapping['organization']] if not pd.isna(row.iloc[column_mapping['organization']]) else None
                date = self._parse_date(row.iloc[column_mapping['date']])
                name = row.iloc[column_mapping['name']] if not pd.isna(row.iloc[column_mapping['name']]) else ""
                daily_amount = float(row.iloc[column_mapping['amount']]) if not pd.isna(row.iloc[column_mapping['amount']]) else 0.0

                # 确保姓名为字符串
                if not name or pd.isna(name):
                    name = ""

                record = Sheet2Record(
                    row_index=idx,
                    organization=organization,
                    date=date,
                    name=name,
                    daily_amount=daily_amount
                )
                records.append(record)

            except (ValueError, IndexError, TypeError) as e:
                raise ValidationError(f"Sheet2第{idx + 2}行数据格式错误: {str(e)}")

        return records

    def _detect_column_mapping(self, df: pd.DataFrame) -> dict:
        """动态检测列名映射

        根据SUPPORTED_COLUMN_PATTERNS自动识别DataFrame中的列名

        Args:
            df: DataFrame

        Returns:
            dict: 列名到索引的映射字典
                {
                    'organization': int,    # 组织列索引
                    'date': int,            # 日期列索引
                    'name': int,            # 姓名列索引
                    'amount': int,          # 金额列索引
                    'date_column_name': str,# 使用的日期列名
                    'amount_column_name': str # 使用的金额列名
                }

        Raises:
            ColumnDetectionError: 无法检测到匹配的列名
        """
        columns = list(df.columns)
        column_mapping = {}

        # 必需的基础列（固定索引）
        column_mapping['organization'] = 0  # 组织列
        column_mapping['name'] = 2          # 姓名列

        # 检测日期列
        date_column_idx = None
        date_column_name = None
        for i, col in enumerate(columns):
            if col in self.SUPPORTED_COLUMN_PATTERNS['date_columns']:
                date_column_idx = i
                date_column_name = col
                break

        if date_column_idx is None:
            raise ColumnDetectionError(
                f"无法找到日期列，支持的列名: {self.SUPPORTED_COLUMN_PATTERNS['date_columns']}"
            )

        column_mapping['date'] = date_column_idx
        column_mapping['date_column_name'] = date_column_name

        # 检测金额列
        amount_column_idx = None
        amount_column_name = None
        for i, col in enumerate(columns):
            if col in self.SUPPORTED_COLUMN_PATTERNS['amount_columns']:
                amount_column_idx = i
                amount_column_name = col
                break

        if amount_column_idx is None:
            raise ColumnDetectionError(
                f"无法找到金额列，支持的列名: {self.SUPPORTED_COLUMN_PATTERNS['amount_columns']}"
            )

        column_mapping['amount'] = amount_column_idx
        column_mapping['amount_column_name'] = amount_column_name

        return column_mapping

    @staticmethod
    def _parse_date(date_value) -> datetime:
        """解析日期值

        Args:
            date_value: 日期值（可能是字符串、datetime、Timestamp等）

        Returns:
            datetime: 解析后的日期

        Raises:
            ValueError: 日期格式错误
        """
        if pd.isna(date_value):
            raise ValueError("日期不能为空")

        # 已经是datetime类型
        if isinstance(date_value, datetime):
            return date_value

        # pandas Timestamp
        if hasattr(date_value, 'to_pydatetime'):
            return date_value.to_pydatetime()

        # 字符串格式
        if isinstance(date_value, str):
            # 尝试多种日期格式
            formats = ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y']
            for fmt in formats:
                try:
                    return datetime.strptime(date_value.strip(), fmt)
                except ValueError:
                    continue
            raise ValueError(f"无法识别的日期格式: {date_value}")

        raise ValueError(f"不支持的日期类型: {type(date_value)}")
