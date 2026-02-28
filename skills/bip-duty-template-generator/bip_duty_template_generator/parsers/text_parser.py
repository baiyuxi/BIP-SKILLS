#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文本解析模块
负责解析用户输入的文本数据
"""

import os
import sys
from typing import List, Dict, Any, Optional

# 添加父目录到路径以便导入config和utils
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from config.settings import Config
from utils.helpers import extract_duty_names, generate_duty_code, truncate_text


class TextParser:
    """文本解析器"""

    def __init__(self, template_structure: Dict[str, Any]):
        """
        初始化文本解析器
        
        Args:
            template_structure: 模板结构信息
        """
        self.structure = template_structure
        self.total_columns = self._get_total_columns()
    
    def _get_total_columns(self) -> int:
        """获取模板总列数"""
        duty_info = self.structure.get('worksheets', {}).get('duty', {})
        return duty_info.get('total_columns', 11)
    
    def parse(self, text: str, default_org: Optional[str] = None) -> List[List[Any]]:
        """
        解析用户输入的文本

        Args:
            text: 用户输入的文本
            default_org: 默认所属组织(已废弃,保持参数兼容性)

        Returns:
            List[List[Any]]: 解析后的职务数据列表
        """
        # 提取职务名称
        duty_names = extract_duty_names(text)

        if not duty_names:
            return []

        # 生成职务数据
        duty_data = []
        for idx, duty_name in enumerate(duty_names, start=1):
            record = self._create_record(duty_name, idx)
            duty_data.append(record)

        return duty_data

    def _create_record(self, duty_name: str, index: int) -> List[Any]:
        """
        创建单条职务记录

        Args:
            duty_name: 职务名称
            index: 序号

        Returns:
            List[Any]: 职务记录
        """
        # 生成职务编码
        code = generate_duty_code(duty_name, index)

        # 生成默认描述
        description = self._generate_description(duty_name)

        # 创建记录(根据字段配置)
        record = [None] * self.total_columns

        # 设置字段值
        record[Config.FIELD_CONFIG['error_msg']['col']] = None  # ERROR_MSG
        record[Config.FIELD_CONFIG['filter_flag']['col']] = None  # 过滤标识(不设置'√')

        # 必填字段: 职务编号(code)
        # 注意: 这里自动生成编号,但如果用户需要可以覆盖
        record[Config.FIELD_CONFIG['code']['col']] = code

        # 必填字段: 职务名称(name)
        record[Config.FIELD_CONFIG['name']['col']] = duty_name

        # 必填字段: 所属组织(org_id_name) - 由交互询问
        record[Config.FIELD_CONFIG['org_id_name']['col']] = None

        # 必填字段: 职务类别(jobtype_id_name) - 由交互询问
        record[Config.FIELD_CONFIG['jobtype_id_name']['col']] = None

        # 必填字段: 职级(jobgrade_id_name) - 由交互询问
        record[Config.FIELD_CONFIG['jobgrade_id_name']['col']] = None

        # 非必填字段
        record[Config.FIELD_CONFIG['maxrank_id_name']['col']] = None  # 最高职等
        record[Config.FIELD_CONFIG['minrank_id_name']['col']] = None  # 最低职等
        record[Config.FIELD_CONFIG['duties']['col']] = description  # 职责
        record[Config.FIELD_CONFIG['memo']['col']] = None  # 备注

        return record
    
    def _generate_description(self, duty_name: str) -> str:
        """
        生成职务描述
        
        Args:
            duty_name: 职务名称
            
        Returns:
            str: 职务描述
        """
        # 职务描述映射
        descriptions = {
            '产品经理': '负责产品规划与设计,制定产品路线图,协调研发团队',
            '研发经理': '负责研发团队管理,推动技术创新,确保项目按时交付',
            '需求分析': '负责收集与分析用户需求,编写需求文档,协调各方资源',
            '软件测试': '负责软件测试工作,制定测试计划,保障产品质量',
            '总经理': '负责公司整体运营,制定战略目标,领导团队实现公司目标',
            '副总经理': '协助总经理处理公司日常事务,分管具体业务部门',
            '技术经理': '负责技术团队管理,推动技术方案落地,保障系统稳定性',
            '架构师': '负责系统架构设计,制定技术规范,解决核心技术难题',
            '测试工程师': '负责执行测试用例,记录缺陷,保障产品质量',
            '前端开发': '负责前端页面开发,优化用户体验,实现交互功能',
            '后端开发': '负责后端接口开发,优化系统性能,保障数据安全',
        }
        
        # 如果有预设描述,使用预设描述
        if duty_name in descriptions:
            desc = descriptions[duty_name]
        else:
            # 生成通用描述
            desc = f'负责{duty_name}相关工作,完成业务目标'
        
        # 截断到最大长度
        return truncate_text(desc, Config.MAX_DESCRIPTION_LENGTH)
