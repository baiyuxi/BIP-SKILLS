#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用友BIP职务档案导入模板生成器 - 主入口

重构后的主模块,整合所有子模块,提供简洁的API
"""

import sys
import os
from typing import List, Dict, Any, Optional, Union
from openpyxl import load_workbook

# 添加技能目录到Python路径
skill_dir = os.path.dirname(__file__)
if skill_dir not in sys.path:
    sys.path.insert(0, skill_dir)

from config.settings import Config
from core.template_reader import TemplateReader
from core.template_writer import TemplateWriter
from core.validator import DutyValidator, ValidationError
from parsers.text_parser import TextParser
from ui.interactive import InteractiveInput
from ui.interactive_ask import FieldAsker


class MissingFieldsError(Exception):
    """缺少必填字段时抛出的异常，用于AI环境交互"""
    
    def __init__(self, message: str, questions: List[Dict[str, Any]], current_data: List[List[Any]]):
        super().__init__(message)
        self.questions = questions
        self.current_data = current_data


class DutyTemplateGenerator:
    """职务模板生成器"""
    
    def __init__(self, template_path: Optional[str] = None):
        """
        初始化生成器

        Args:
            template_path: 模板文件路径,如果为None则使用默认路径
        """
        self.template_path = template_path
        self.template_reader = None
        self.template_writer = None
        self.validator = None
        self.parser = None
        self.interactive = None
        self.field_asker = None
        self.structure = None
        
    def initialize(self):
        """初始化各个组件"""
        # 创建模板读取器
        self.template_reader = TemplateReader(self.template_path)

        # 加载最新的模板文件
        print(f"📂 正在读取模板文件: {self.template_reader.template_path}")
        self.template_reader.load()

        # 获取模板结构
        self.structure = self.template_reader.get_structure()

        # 获取必填字段(从模板*标记)
        required_fields = self.template_reader.get_required_fields()
        print(f"✓ 模板已加载,识别到 {len(required_fields)} 个必填字段")

        # 创建验证器
        self.validator = DutyValidator(self.structure)

        # 创建解析器
        self.parser = TextParser(self.structure)

        # 创建交互管理器
        available_orgs = self.template_reader.get_organizations()
        self.interactive = InteractiveInput(available_orgs)

        # 创建字段询问器
        self.field_asker = FieldAsker(self.structure, available_orgs)

        return self
    
    def parse_user_input(
        self,
        user_input: str,
        default_org: Optional[str] = None,
        input_format: str = "text"
    ) -> List[List[Any]]:
        """
        解析用户输入
        
        Args:
            user_input: 用户输入(文本/JSON等)
            default_org: 默认所属组织
            input_format: 输入格式(text/json)
            
        Returns:
            List[List[Any]]: 解析后的职务数据
        """
        if input_format == "text":
            return self.parser.parse(user_input, default_org)
        else:
            raise ValueError(f"不支持的输入格式: {input_format}")
    
    def validate_and_fix(
        self,
        records: List[List[Any]],
        auto_fix: bool = True
    ) -> tuple[bool, List[List[Any]], List[str]]:
        """
        验证并修复数据
        
        Args:
            records: 职务数据列表
            auto_fix: 是否自动修复
            
        Returns:
            tuple: (是否通过验证, 修复后的数据, 错误/警告信息)
        """
        if auto_fix:
            # 自动修复
            fixed_records, warnings = self.validator.validate_and_fix(records)
            
            # 检查列数
            expected_cols = self.parser.total_columns
            col_warnings = self.validator.check_column_count(fixed_records, expected_cols)
            warnings.extend(col_warnings)
            
            # 验证
            passed, errors = self.validator.validate_all(fixed_records)
            return (passed, fixed_records, warnings + errors)
        else:
            # 只验证不修复
            passed, errors = self.validator.validate_all(records)
            return (passed, records, errors)
    
    def generate(
        self,
        user_input: Optional[str] = None,
        user_data: Optional[List[List[Any]]] = None,
        output_path: Optional[str] = None,
        interactive: bool = False,
        default_org: Optional[str] = None,
        ask_missing: bool = False,
        user_answers: Optional[Dict[str, str]] = None
    ) -> str:
        """
        生成模板

        Args:
            user_input: 用户输入(文本)
            user_data: 用户数据(已解析)
            output_path: 输出路径
            interactive: 是否启用交互式输入
            default_org: 默认所属组织
            ask_missing: 是否询问缺失的必填字段（返回问题而非直接询问）
            user_answers: 用户对缺失字段问题的回答

        Returns:
            str: 生成的模板文件路径
        """
        # 步骤1: 初始化
        if not self.template_reader:
            self.initialize()

        # 步骤2: 获取数据
        if user_data is None:
            if user_input is None:
                user_input = input("请输入职务信息: ")

            # 解析用户输入（不传入 default_org，由后续步骤处理）
            duty_data = self.parse_user_input(user_input, default_org=None)

            if not duty_data:
                raise ValueError("未能从输入中解析出职务数据")
        else:
            duty_data = user_data

        # 步骤3: 如果用户已提供答案，先应用答案
        if ask_missing and self.field_asker and user_answers:
            duty_data = self.field_asker.apply_user_answers(duty_data, user_answers)

        # 步骤4: 检查缺失的必填字段（AI环境模式）- 在验证之前检查
        # 只有当 ask_missing=True 且没有提供 user_answers 时才询问
        if ask_missing and self.field_asker and not user_answers:
            # 先检查是否有缺失字段
            missing_fields = self.field_asker.check_missing_fields(duty_data)
            
            if missing_fields:
                # 有缺失字段，返回需要询问的问题
                questions = self.field_asker.generate_questions_for_ask_tool(duty_data)
                if questions:
                    # 抛出特殊异常，包含需要询问的问题
                    raise MissingFieldsError(
                        message="需要询问用户必填字段",
                        questions=questions,
                        current_data=duty_data
                    )
        
        # 步骤5: 验证数据（当有缺失字段但不询问时，使用auto_fix）
        auto_fix = not ask_missing
        passed, fixed_data, messages = self.validate_and_fix(duty_data, auto_fix=auto_fix)

        # 打印警告信息
        for msg in messages:
            print(f"⚠️  {msg}")

        if not passed:
            raise ValidationError(f"数据验证失败:\n" + "\n".join(messages))

        # 步骤7: 交互式确认(如果启用)
        if interactive and self.template_reader:
            field_mapping = self.template_reader.get_field_mapping()
            field_names = self.structure.get('worksheets', {}).get('duty', {}).get('headers', {}).get('field_names', [])

            confirmed_data = self.interactive.confirm_all_required_fields(
                fixed_data,
                self.structure,
                field_mapping,
                field_names
            )
            duty_data = confirmed_data

        # 步骤8: 创建模板写入器并复制结构
        self.template_writer = TemplateWriter(self.structure)
        self.template_reader.copy_all_sheets(self.template_writer.workbook)

        # 移除默认的Sheet页签(如果存在)
        if 'Sheet' in self.template_writer.workbook.sheetnames:
            self.template_writer.workbook.remove(self.template_writer.workbook['Sheet'])

        # 步骤9: 更新职务数据
        self.template_writer.update_duty_data(duty_data)

        # 步骤10: 保存文件
        output_path = self.template_writer.save(output_path)

        # 步骤11: 清理资源
        self.template_reader.close()
        self.template_writer.close()

        return output_path
    
    def __enter__(self):
        """上下文管理器入口"""
        self.initialize()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        if self.template_reader:
            self.template_reader.close()
        if self.template_writer:
            self.template_writer.close()


def generate_template(
    user_input: Optional[str] = None,
    user_data: Optional[List[List[Any]]] = None,
    output_path: Optional[str] = None,
    interactive: bool = False,
    default_org: Optional[str] = None,
    template_path: Optional[str] = None,
    ask_missing: bool = True,
    user_answers: Optional[Dict[str, str]] = None
) -> str:
    """
    快速生成职务模板的便捷函数

    Args:
        user_input: 用户输入(文本)
        user_data: 用户数据(已解析)
        output_path: 输出路径
        interactive: 是否启用交互式输入
        default_org: 默认所属组织（已废弃，请使用 ask_missing + user_answers）
        template_path: 模板路径
        ask_missing: 是否询问缺失的必填字段，返回问题而非直接生成
        user_answers: 用户对缺失字段问题的回答（字典格式：{field_code: value}）

    Returns:
        str: 生成的模板文件路径

    Raises:
        MissingFieldsError: 当 ask_missing=True 且存在缺失字段时抛出，包含需要询问的问题
    """
    with DutyTemplateGenerator(template_path) as generator:
        return generator.generate(
            user_input=user_input,
            user_data=user_data,
            output_path=output_path,
            interactive=interactive,
            default_org=default_org,
            ask_missing=ask_missing,
            user_answers=user_answers
        )


def parse_text_input(text: str, template_path: Optional[str] = None) -> List[List[Any]]:
    """
    解析文本输入(仅解析,不生成模板)
    
    Args:
        text: 用户输入的文本
        template_path: 模板路径
        
    Returns:
        List[List[Any]]: 解析后的职务数据
    """
    with DutyTemplateGenerator(template_path) as generator:
        return generator.parse_user_input(text)


def validate_duty_data(data: List[List[Any]], template_path: Optional[str] = None) -> tuple[bool, List[str]]:
    """
    验证职务数据
    
    Args:
        data: 职务数据
        template_path: 模板路径
        
    Returns:
        tuple: (是否通过验证, 错误信息列表)
    """
    with DutyTemplateGenerator(template_path) as generator:
        passed, _, errors = generator.validate_and_fix(data, auto_fix=False)
        return (passed, errors)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='用友BIP职务档案导入模板生成器')
    parser.add_argument('--input', '-i', help='用户输入(文本格式)')
    parser.add_argument('--output', '-o', help='输出文件路径')
    parser.add_argument('--interactive', '-t', action='store_true', help='启用交互式输入')
    parser.add_argument('--ask-missing', '-a', action='store_true', help='询问缺失的必填字段')
    parser.add_argument('--org', help='默认所属组织')

    args = parser.parse_args()

    try:
        output = generate_template(
            user_input=args.input,
            output_path=args.output,
            interactive=args.interactive,
            default_org=args.org,
            ask_missing=args.ask_missing
        )
        print(f"✅ 模板生成成功: {output}")
    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
