#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数模块
提供通用的辅助函数
"""

import re
from typing import List, Dict, Any, Optional
from datetime import datetime


def extract_duty_names(text: str) -> List[str]:
    """
    从文本中提取职务名称
    
    Args:
        text: 用户输入的文本
        
    Returns:
        List[str]: 提取到的职务名称列表
        
    Examples:
        >>> extract_duty_names("包括 职务：产品经理、研发经理、需求分析、软件测试")
        ['产品经理', '研发经理', '需求分析', '软件测试']
        
        >>> extract_duty_names("我需要导入总经理和副总经理")
        ['总经理', '副总经理']
    """
    duties = []
    
    # 模式1: "职务：A、B、C" 或 "职务:A、B、C"
    pattern1 = r'职务[：:]\s*([^。\n]+)'
    match1 = re.search(pattern1, text)
    if match1:
        content = match1.group(1)
        # 分割并清理
        duties.extend([d.strip() for d in re.split('[、,;，;；]', content) if d.strip()])
        return duties
    
    # 模式2: "包括职务A和B" 或 "包括职务：A、B"
    pattern2 = r'包括.*?职务[：:]?\s*([^。\n]+)'
    match2 = re.search(pattern2, text)
    if match2:
        content = match2.group(1)
        duties.extend([d.strip() for d in re.split('[和、,;，;；]', content) if d.strip()])
        return duties
    
    # 模式3: 查找常见的职务关键词
    duty_keywords = [
        '总经理', '副总经理', '总监', '经理', '主管',
        '产品经理', '研发经理', '技术经理',
        '需求分析', '需求分析师', '软件测试', '测试工程师',
        '开发工程师', '程序员', '架构师',
        '设计师', 'UI设计', '前端开发', '后端开发',
        '行政经理', '人事经理', '财务经理'
    ]
    
    for keyword in duty_keywords:
        if keyword in text:
            duties.append(keyword)
    
    return list(set(duties))  # 去重


def generate_duty_code(duty_name: str, index: int = 1) -> str:
    """
    生成职务编码
    
    Args:
        duty_name: 职务名称
        index: 序号
        
    Returns:
        str: 职务编码
        
    Examples:
        >>> generate_duty_code('产品经理', 1)
        'PM0227001'
    """
    # 生成前缀(首字母大写)
    prefix = ''.join([c for c in duty_name if c.isalpha() or c == '经理' or c == '总监'])
    if not prefix or len(prefix) < 2:
        prefix = 'DUTY'
    
    # 如果是中文,使用职务类型的简写
    duty_prefix_map = {
        '产品经理': 'PM',
        '研发经理': 'RDM',
        '技术经理': 'TM',
        '总经理': 'GM',
        '副总经理': 'DGM',
        '需求分析': 'REQ',
        '需求分析师': 'REQ',
        '软件测试': 'TEST',
        '测试工程师': 'TEST',
        '架构师': 'ARCH',
        '总应用架构师': 'SARCH',
        '应用架构师': 'ARCH',
        '行政经理': 'ADM',
    }
    
    if duty_name in duty_prefix_map:
        prefix = duty_prefix_map[duty_name]
    
    # 生成时间戳后缀
    timestamp = datetime.now().strftime('%m%d%H%M%S')
    
    return f"{prefix}{timestamp}{index:03d}"


def truncate_text(text: str, max_length: int = 50) -> str:
    """
    截断文本到指定长度
    
    Args:
        text: 原始文本
        max_length: 最大长度
        
    Returns:
        str: 截断后的文本
    """
    if not text:
        return ''
    if len(text) <= max_length:
        return text
    return text[:max_length - 1] + '…'


def clean_value(value: Any) -> Optional[str]:
    """
    清理值,去除空白和None
    
    Args:
        value: 原始值
        
    Returns:
        Optional[str]: 清理后的值或None
    """
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return str(value)


def merge_dicts(*dicts: Dict[str, Any]) -> Dict[str, Any]:
    """
    合并多个字典,后面的覆盖前面的
    
    Args:
        *dicts: 要合并的字典
        
    Returns:
        Dict[str, Any]: 合并后的字典
    """
    result = {}
    for d in dicts:
        result.update(d)
    return result


def validate_email(email: str) -> bool:
    """
    验证邮箱格式
    
    Args:
        email: 邮箱地址
        
    Returns:
        bool: 是否有效
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    """
    安全除法,避免除零错误
    
    Args:
        numerator: 分子
        denominator: 分母
        default: 默认值
        
    Returns:
        float: 除法结果或默认值
    """
    if denominator == 0:
        return default
    return numerator / denominator
