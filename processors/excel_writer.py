"""Excel文件写入器"""

import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import List

from models import Sheet2Record
from utils.exceptions import ProcessingError


class ExcelWriter:
    """Excel文件写入器"""

    def __init__(self, output_dir: Path = None):
        """初始化写入器

        Args:
            output_dir: 输出目录（默认为当前目录）
        """
        self.output_dir = output_dir or Path.cwd()

    def write_results(self, original_file: Path,
                     sheet2_records: List[Sheet2Record]) -> Path:
        """写入处理结果到新的Excel文件

        Args:
            original_file: 原始文件路径
            sheet2_records: 处理后的Sheet2记录

        Returns:
            Path: 输出文件路径

        Raises:
            ProcessingError: 写入失败
        """
        try:
            # 生成输出文件名
            output_file = self._generate_output_filename(original_file)

            # 读取原始文件
            original_df = pd.read_excel(original_file, sheet_name=None)

            # 更新Sheet2数据
            if len(original_df) < 2:
                raise ProcessingError("Excel文件必须包含至少两个sheet")

            # 获取原始sheet名称
            sheet_names = list(original_df.keys())
            sheet2_name = sheet_names[1]

            # 创建新的DataFrame
            result_df = self._create_sheet2_dataframe(sheet2_records, original_df[sheet2_name])

            # 更新原始数据
            original_df[sheet2_name] = result_df

            # 写入Excel文件
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in original_df.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            return output_file

        except Exception as e:
            raise ProcessingError(f"写入Excel文件失败: {str(e)}")

    def _generate_output_filename(self, original_file: Path) -> Path:
        """生成输出文件名

        格式：output_原文件名.xlsx

        Args:
            original_file: 原始文件路径

        Returns:
            Path: 输出文件路径
        """
        output_name = f"output_{original_file.name}"
        return self.output_dir / output_name

    def _create_sheet2_dataframe(self, records: List[Sheet2Record],
                                original_df: pd.DataFrame) -> pd.DataFrame:
        """创建Sheet2的DataFrame

        Args:
            records: Sheet2记录列表
            original_df: 原始DataFrame（用于保持格式）

        Returns:
            pd.DataFrame: 更新后的DataFrame
        """
        # 按原始行号排序
        sorted_records = sorted(records, key=lambda x: x.row_index)

        # 准备数据
        data = []
        for record in sorted_records:
            row_data = [
                record.organization or '',  # None转换为空字符串
                record.date,
                record.name,
                record.daily_amount
            ]
            data.append(row_data)

        # 创建DataFrame，保持原始列名
        columns = original_df.columns.tolist()
        result_df = pd.DataFrame(data, columns=columns)

        return result_df
