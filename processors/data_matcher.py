"""核心数据匹配算法"""

from collections import defaultdict
from datetime import datetime
from typing import Dict, List, Optional, Set, Tuple

from models import Sheet1Record, Sheet2Record


class DataMatcher:
    """核心数据匹配器"""

    def __init__(self, tolerance: float = 10.0):
        """初始化匹配器

        Args:
            tolerance: 金额匹配容差值（默认10元）
        """
        self.tolerance = tolerance

    def match_data(self, sheet1_records: List[Sheet1Record],
                   sheet2_records: List[Sheet2Record]) -> List[Sheet2Record]:
        """主匹配方法 - 根据Sheet1的组织信息匹配Sheet2的明细记录

        算法策略：
        1. 按日期分组所有记录
        2. 对每个sheet1记录，在同日期的sheet2记录中寻找匹配
        3. 优先匹配已有组织的记录
        4. 使用无组织记录填充金额差额
        5. 确保同一人名只出现在一个组织中

        Args:
            sheet1_records: Sheet1记录列表（组织信息）
            sheet2_records: Sheet2记录列表（明细记录）

        Returns:
            List[Sheet2Record]: 匹配后的Sheet2记录（组织字段已更新）
        """
        # 按日期分组sheet2记录
        sheet2_by_date = self._group_by_date(sheet2_records)

        # 跟踪已处理的人名（确保不冲突）
        processed_names: Set[str] = set()

        # 处理结果
        result_records: List[Sheet2Record] = []

        # 对sheet1按日期排序处理
        sorted_sheet1 = sorted(sheet1_records, key=lambda x: x.date)

        for sheet1_record in sorted_sheet1:
            # 获取同日期所有sheet2记录
            same_date_records = sheet2_by_date.get(sheet1_record.date.date(), [])

            if not same_date_records:
                continue

            # 分离无组织和有组织记录
            no_org_records = [r for r in same_date_records
                            if r.organization is None or r.organization == '']
            yes_org_records = [r for r in same_date_records
                             if r.organization is not None and r.organization == sheet1_record.organization]

            # 移除已处理的有组织记录（避免重复）
            yes_org_records = [r for r in yes_org_records
                             if r.name not in processed_names]

            # 标记有组织记录为已处理
            for record in yes_org_records:
                processed_names.add(record.name)

            # 计算有组织记录总额
            total_amount = sum(r.daily_amount for r in yes_org_records)

            # 目标金额（带容差）
            target_amount = sheet1_record.daily_amount

            # 如果已经有组织记录总额在容差范围内，不需要处理无组织记录
            if abs(total_amount - target_amount) <= self.tolerance:
                result_records.extend(yes_org_records)
                result_records.extend(no_org_records)
                continue

            # 如果总额不足，尝试从无组织记录中补充
            if total_amount < target_amount:
                remaining_amount = target_amount - total_amount

                # 贪心匹配无组织记录
                matched_no_org = self._greedy_match(no_org_records,
                                                  remaining_amount,
                                                  processed_names,
                                                  sheet1_record.organization)

                # 更新有组织记录列表
                yes_org_records.extend(matched_no_org)
                for record in matched_no_org:
                    processed_names.add(record.name)

                # 更新无组织记录列表
                matched_names = {r.name for r in matched_no_org}
                no_org_records = [r for r in no_org_records
                                if r.name not in matched_names]

            # 收集结果
            result_records.extend(yes_org_records)
            result_records.extend(no_org_records)

        # 按原始序号排序
        result_records.sort(key=lambda x: x.row_index)

        return result_records

    def _group_by_date(self, records: List[Sheet2Record]) -> Dict[datetime.date, List[Sheet2Record]]:
        """按日期分组记录

        Args:
            records: Sheet2记录列表

        Returns:
            Dict[datetime.date, List[Sheet2Record]]: 按日期分组的记录
        """
        grouped = defaultdict(list)
        for record in records:
            date_key = record.date.date()
            grouped[date_key].append(record)
        return grouped

    def _greedy_match(self, no_org_records: List[Sheet2Record],
                      target_amount: float,
                      processed_names: Set[str],
                      target_organization: str) -> List[Sheet2Record]:
        """贪心算法匹配无组织记录

        策略：
        1. 优先选择金额接近目标剩余额的记录
        2. 确保人名没有冲突
        3. 在容差范围内停止

        Args:
            no_org_records: 无组织记录列表
            target_amount: 目标剩余金额
            processed_names: 已处理人名集合
            target_organization: 目标组织

        Returns:
            List[Sheet2Record]: 匹配成功的记录（组织字段已更新）
        """
        matched_records: List[Sheet2Record] = []
        current_total = 0.0

        # 复制列表，避免修改原数据
        available_records = [r for r in no_org_records
                           if r.name not in processed_names]

        while available_records and current_total < target_amount:
            # 寻找最合适的记录
            best_record = self._find_best_match(available_records,
                                              target_amount - current_total)

            if best_record is None:
                break

            # 检查添加该记录后是否超过容差范围
            new_total = current_total + best_record.daily_amount
            if abs(new_total - target_amount) > self.tolerance + 100:  # 额外缓冲
                break

            # 更新记录的组织信息
            best_record.organization = target_organization
            matched_records.append(best_record)
            current_total = new_total

            # 从可用记录中移除
            available_records.remove(best_record)

            # 如果达到容差范围，停止匹配
            if abs(current_total - target_amount) <= self.tolerance:
                break

        return matched_records

    def _find_best_match(self, records: List[Sheet2Record],
                        remaining_amount: float) -> Optional[Sheet2Record]:
        """寻找最合适的匹配记录

        策略：
        - 优先选择金额最接近剩余额的记录
        - 考虑正负差异，选择最小化差异的记录

        Args:
            records: 可用记录列表
            remaining_amount: 剩余金额

        Returns:
            Optional[Sheet2Record]: 最佳匹配记录，如果没有合适的则返回None
        """
        if not records:
            return None

        best_record = None
        min_diff = float('inf')

        for record in records:
            # 计算金额差异（绝对值）
            diff = abs(record.daily_amount - remaining_amount)

            # 更新最佳记录
            if diff < min_diff:
                min_diff = diff
                best_record = record

        return best_record

    @staticmethod
    def _has_name_conflict(name: str, all_records: List[Sheet2Record]) -> bool:
        """检查人名冲突

        确保同一人名只出现在一个组织中

        Args:
            name: 人名
            all_records: 所有记录

        Returns:
            bool: 如果存在冲突返回True
        """
        organizations = set()
        for record in all_records:
            if record.name == name and record.organization:
                organizations.add(record.organization)

        return len(organizations) > 1
