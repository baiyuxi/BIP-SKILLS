#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理模块
管理技能的各种配置参数
"""

import os
from typing import List, Dict

class Config:
    """配置类"""

    # 基础路径配置
    SKILL_DIR = os.path.dirname(os.path.dirname(__file__))
    
    # 模板相关路径
    REF_DIR = os.path.join(SKILL_DIR, 'references')
    TEMPLATE_PATH = os.path.join(REF_DIR, '模板_职务.xlsx')
    
    # 输出路径配置
    OUTPUT_DIR = os.path.dirname(SKILL_DIR)  # 项目根目录
    
    # 默认配置
    DEFAULT_ORGANIZATION = "企业账号级##global00"
    DEFAULT_ORG_NAME = "企业账号级"
    
    # 字段配置(用于映射字段名称和列位置)
    # 注意:required字段由模板中的*标记决定,不从这里读取
    FIELD_CONFIG = {
        'error_msg': {'col': 0, 'name': 'ERROR_MSG'},
        'filter_flag': {'col': 1, 'name': '过滤标识'},
        'code': {'col': 2, 'name': '职务编号'},
        'name': {'col': 3, 'name': '职务名称'},
        'org_id_name': {'col': 4, 'name': '所属组织'},
        'jobtype_id_name': {'col': 5, 'name': '职务类别'},
        'jobgrade_id_name': {'col': 6, 'name': '职级'},
        'maxrank_id_name': {'col': 7, 'name': '最高职等'},
        'minrank_id_name': {'col': 8, 'name': '最低职等'},
        'duties': {'col': 9, 'name': '职责'},
        'memo': {'col': 10, 'name': '备注'},
    }
    
    # 职务描述字数限制
    MAX_DESCRIPTION_LENGTH = 50
    
    # 数据行起始行号
    DATA_START_ROW = 9
    
    # 工作表配置
    WORKSHEETS = {
        'readme': 'ReadMe',
        'duty': '职务',
        'ref': '参照',
        'template_info': 'importTemplateInfo',
        'dropdown': 'importDropDownContent'
    }
    
    @classmethod
    def get_available_organizations(cls) -> List[str]:
        """获取可用的组织列表(从模板中读取)"""
        return []
    
    @classmethod
    def get_output_filename(cls) -> str:
        """生成输出文件名"""
        from datetime import datetime
        return f"模板_职务_{datetime.now().strftime('%m%d%H%M%S')}.xlsx"
    
    @classmethod
    def get_output_path(cls, filename: str = None) -> str:
        """获取输出文件完整路径"""
        if filename is None:
            filename = cls.get_output_filename()
        return os.path.join(cls.OUTPUT_DIR, filename)
