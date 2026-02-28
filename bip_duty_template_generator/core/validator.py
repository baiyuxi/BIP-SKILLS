#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据验证模块
负责验证职务数据是否符合模板要求
"""

import os
import sys
from typing import List, Dict, Any, Tuple

# 添加父目录到路径以便导入config和utils
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from config.settings import Config
from utils.helpers import clean_value, truncate_text


class ValidationError(Exception):
    """验证错误"""
    pass


class DutyValidator:
    """职务数据验证器"""

    def __init__(self, template_structure: Dict[str, Any]):
        """
        初始化验证器

        Args:
            template_structure: 模板结构信息
        """
        self.structure = template_structure
        self.required_fields = self._get_required_fields()
        self.field_mapping = self._get_field_mapping()

    def _get_required_fields(self) -> List[Dict[str, Any]]:
        """
        获取必填字段列表(从模板的*标记中读取)

        Returns:
            List[Dict]: 必填字段列表
        """
        return self.structure.get('worksheets', {}).get('duty', {}).get('required_fields', [])

    def _get_field_mapping(self) -> Dict[str, int]:
        """获取字段映射"""
        # 使用配置中的字段映射
        field_mapping = {}
        for field_key, field_config in Config.FIELD_CONFIG.items():
            field_mapping[field_key] = field_config['col']
        return field_mapping
    
    def validate_record(self, record: List[Any], row_index: int = 0) -> Tuple[bool, List[str]]:
        """
        验证单条记录
        
        Args:
            record: 记录数据
            row_index: 行索引(从0开始,用于错误提示)
            
        Returns:
            Tuple[bool, List[str]]: (是否通过验证, 错误信息列表)
        """
        errors = []
        
        # 验证必填字段
        for field in self.required_fields:
            field_code = field.get('code', '')
            field_name = field.get('name', '')
            column = field.get('column', 0)
            
            if field_code in self.field_mapping:
                col_idx = self.field_mapping[field_code]
                if col_idx < len(record):
                    value = clean_value(record[col_idx])
                    if value is None or value == '':
                        errors.append(f"第{row_index + 1}行:必填字段'{field_name}'为空")
        
        # 验证职务描述长度
        duties_col = Config.FIELD_CONFIG['duties']['col']
        if duties_col < len(record):
            duties = record[duties_col]
            if duties and len(str(duties)) > Config.MAX_DESCRIPTION_LENGTH:
                errors.append(f"第{row_index + 1}行:职务描述超过{Config.MAX_DESCRIPTION_LENGTH}字限制")
        
        # 验证职级范围(最低职等 <= 最高职等)
        maxrank_col = Config.FIELD_CONFIG['maxrank_id_name']['col']
        minrank_col = Config.FIELD_CONFIG['minrank_id_name']['col']
        
        if (maxrank_col < len(record) and minrank_col < len(record)):
            maxrank = clean_value(record[maxrank_col])
            minrank = clean_value(record[minrank_col])
            
            if maxrank and minrank:
                try:
                    # 尝试提取数字进行比较
                    max_num = self._extract_rank_number(maxrank)
                    min_num = self._extract_rank_number(minrank)
                    
                    if max_num is not None and min_num is not None:
                        if max_num < min_num:
                            errors.append(f"第{row_index + 1}行:最高职等({maxrank})不能小于最低职等({minrank})")
                except Exception:
                    pass  # 如果无法比较,跳过验证
        
        return (len(errors) == 0, errors)
    
    def _extract_rank_number(self, rank_str: str) -> int:
        """
        从职级字符串中提取数字
        
        Args:
            rank_str: 职级字符串(如"M10"、"10"、"Rank5")
            
        Returns:
            int: 提取的数字
        """
        import re
        match = re.search(r'\d+', str(rank_str))
        if match:
            return int(match.group())
        return 0
    
    def validate_all(self, records: List[List[Any]]) -> Tuple[bool, List[str]]:
        """
        验证所有记录
        
        Args:
            records: 记录列表
            
        Returns:
            Tuple[bool, List[str]]: (是否全部通过验证, 所有错误信息)
        """
        all_errors = []
        all_passed = True
        
        for idx, record in enumerate(records):
            passed, errors = self.validate_record(record, idx)
            if not passed:
                all_passed = False
                all_errors.extend(errors)
        
        return (all_passed, all_errors)
    
    def validate_and_fix(self, records: List[List[Any]]) -> Tuple[List[List[Any]], List[str]]:
        """
        验证并尝试修复记录
        
        Args:
            records: 原始记录列表
            
        Returns:
            Tuple[List[List[Any]], List[str]]: (修复后的记录列表, 警告信息)
        """
        warnings = []
        fixed_records = []
        
        for idx, record in enumerate(records):
            fixed_record = record.copy()
            
            # 截断过长的职务描述
            duties_col = Config.FIELD_CONFIG['duties']['col']
            if duties_col < len(fixed_record):
                duties = fixed_record[duties_col]
                if duties and len(str(duties)) > Config.MAX_DESCRIPTION_LENGTH:
                    original = str(duties)
                    fixed_record[duties_col] = truncate_text(original, Config.MAX_DESCRIPTION_LENGTH)
                    warnings.append(f"第{idx + 1}行:职务描述已从{len(original)}字截断至{Config.MAX_DESCRIPTION_LENGTH}字")
            
            # 确保filter_flag不设置'√'(避免过滤数据)
            filter_flag_col = Config.FIELD_CONFIG['filter_flag']['col']
            if filter_flag_col < len(fixed_record):
                if fixed_record[filter_flag_col] == '√':
                    fixed_record[filter_flag_col] = None
                    warnings.append(f"第{idx + 1}行:已清除过滤标识,确保数据正常导入")
            
            fixed_records.append(fixed_record)
        
        return (fixed_records, warnings)
    
    def check_column_count(self, records: List[List[Any]], expected_count: int) -> List[str]:
        """
        检查列数是否匹配
        
        Args:
            records: 记录列表
            expected_count: 期望的列数
            
        Returns:
            List[str]: 调整警告信息
        """
        warnings = []
        
        for idx, record in enumerate(records):
            if len(record) < expected_count:
                # 补齐缺失的列
                record.extend([''] * (expected_count - len(record)))
                warnings.append(f"第{idx + 1}行:列数不足,已自动补齐到{expected_count}列")
            elif len(record) > expected_count:
                # 截断多余的列
                del record[expected_count:]
                warnings.append(f"第{idx + 1}行:列数过多,已自动截断到{expected_count}列")
        
        return warnings
