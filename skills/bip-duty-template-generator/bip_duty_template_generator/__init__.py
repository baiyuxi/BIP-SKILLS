#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用友BIP职务档案导入模板生成器 - 包初始化
"""

from .generate_duty_template import (
    DutyTemplateGenerator,
    generate_template,
    parse_text_input,
    validate_duty_data
)

__version__ = '2.0.0'
__all__ = [
    'DutyTemplateGenerator',
    'generate_template',
    'parse_text_input',
    'validate_duty_data'
]
