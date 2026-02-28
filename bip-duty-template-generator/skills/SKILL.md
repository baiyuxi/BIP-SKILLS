---
name: "bip-duty-template-generator    用友BIP职务档案导入模板生成器"
description: "Generate Excel template for importing duty (job profile) data into Yonyou BIP system. Call this skill when users need to import duty data, generate duty import template, create job profile template, or work with duty/archive/import related tasks in BIP system."
description_zh: "生成用友BIP系统职务档案导入模板。当用户需要导入职务数据、生成职务导入模板、创建职务档案模板或处理BIP系统中与职务/档案/导入相关的任务时调用此技能。"
trigger_keywords: ["导入职务", "职务导入", "职务档案", "job profile", "duty archive", "import template", "职务数据", "档案导入", "BIP导入", "职务模板"]
---

# 用友BIP职务档案导入模板生成器

## 功能说明

本技能用于快速生成符合用友BIP系统要求的职务档案导入Excel模板文件。

## 技术架构 (v2.0)

### 模块划分

```
bip-duty-template-generator/
├── config/
│   └── settings.py          # 配置管理
├── core/
│   ├── template_reader.py  # 模板读取与解析
│   ├── template_writer.py  # 模板写入
│   └── validator.py        # 数据验证
├── parsers/
│   └── text_parser.py      # 文本解析
├── ui/
│   └── interactive.py      # 用户交互
├── utils/
│   └── helpers.py          # 工具函数
├── references/
│   └── 模板_职务.xlsx      # 标准模板
└── generate_duty_template.py # 主入口
```

### 核心特性

1. **模块化设计**: 各功能模块职责清晰,易于维护和扩展
2. **配置驱动**: 通过Config类统一管理配置,避免硬编码
3. **智能解析**: 支持从自然语言文本中提取职务信息
4. **数据验证**: 完整的必填字段检查和格式验证
5. **自动修复**: 自动截断过长描述、修复列数等常见问题

## 输入要求

用户可以通过以下方式提供职务档案数据:

1. **自然语言文本**: "我需要导入职务档案,包括职务:产品经理、研发经理、需求分析、软件测试"
2. **已解析数据**: 直接提供格式化的职务数据列表
3. **JSON数据**: 支持JSON格式输入(需扩展JSONParser)

技能会从参考模板中自动读取必填字段定义。

## 输出内容

生成一个Excel文件,包含以下工作表:

1. **ReadMe** - 模板使用指南
2. **职务** - 职务档案数据表
3. **参照** - 参照数据表

## 文件路径

输出文件路径动态生成,格式为:`模板_职务_{月日时分秒}.xlsx`,例如:`模板_职务_0227185046.xlsx`

## 必填字段

必填字段从参考模板 `references/模板_职务.xlsx` 中动态读取,根据模板中的字段定义(带*号标识)确定。

当前必填字段:
- 职务编号 (code)
- 职务名称 (name)
- 所属组织 (org_id_name)

## 使用方法

### 方式1: 使用便捷函数

```python
from generate_duty_template import generate_template, MissingFieldsError

# 从文本输入生成模板（启用询问模式）
try:
    output_path = generate_template(
        user_input="我需要导入职务档案,包括职务:产品经理、研发经理、需求分析、软件测试",
        ask_missing=True  # 启用询问模式
    )
except MissingFieldsError as e:
    # e.questions 包含需要询问的问题列表
    # 使用 ask_followup_question 工具向用户展示选项
    # 用户选择后，传入 user_answers 重新生成
    questions = e.questions
    # ...
    # 假设用户选择了"企业账号级##global00"
    output_path = generate_template(
        user_input="我需要导入职务档案,包括职务:产品经理、研发经理、需求分析、软件测试",
        ask_missing=True,
        user_answers={"org_id_name": "企业账号级##global00"}
    )

# 从已解析数据生成模板
user_data = [
    [None, None, 'PM0227001', '产品经理', '企业账号级##global00', None, None, None, None, '负责产品规划与设计', None],
    [None, None, 'RDM0227002', '研发经理', '企业账号级##global00', None, None, None, None, '负责研发团队管理', None],
]
output_path = generate_template(user_data=user_data)
```

### 方式2: 使用生成器类

```python
from generate_duty_template import DutyTemplateGenerator

with DutyTemplateGenerator() as generator:
    # 解析输入
    duty_data = generator.parse_user_input(
        "包括职务:产品经理、研发经理"
    )
    
    # 验证数据
    passed, fixed_data, messages = generator.validate_and_fix(duty_data)
    
    # 生成模板
    output_path = generator.generate(user_data=fixed_data)
```

### 方式3: 仅解析/验证

```python
# 仅解析,不生成模板
from generate_duty_template import parse_text_input

duty_data = parse_text_input("产品经理、研发经理、需求分析")

# 仅验证,不生成模板
from generate_duty_template import validate_duty_data

passed, errors = validate_duty_data(duty_data)
if not passed:
    print("验证失败:", errors)
```

## 功能特点

### 1. 智能文本解析

支持从自然语言中提取职务名称:

```python
# 示例输入:
"我需要导入职务档案,包括职务:产品经理、研发经理、需求分析、软件测试"

# 自动提取出:
['产品经理', '研发经理', '需求分析', '软件测试']
```

### 2. 自动职务编码生成

根据职务名称自动生成规范编码:

- 产品经理 -> PM0227185046001
- 研发经理 -> RDM0227185046001
- 需求分析 -> REQ0227185046001
- 软件测试 -> TEST0227185046001

### 3. 智能职务描述

根据职务名称自动生成职责描述,并确保不超过50字:

```python
产品经理 -> "负责产品规划与设计,制定产品路线图,协调研发团队"
研发经理 -> "负责研发团队管理,推动技术创新,确保项目按时交付"
需求分析 -> "负责收集与分析用户需求,编写需求文档,协调各方资源"
软件测试 -> "负责软件测试工作,制定测试计划,保障产品质量"
```

### 4. 数据验证与自动修复

- **必填字段检查**: 确保所有必填字段非空
- **描述长度限制**: 自动截断超过50字的职务描述
- **列数调整**: 自动补齐或截断列数以匹配模板
- **职级范围验证**: 确保最高职等 >= 最低职等
- **过滤标识处理**: 确保filter_flag不设置为'√',避免数据被过滤

### 5. 错误处理

```python
try:
    output = generate_template(user_input="...")
except ValidationError as e:
    print(f"数据验证失败: {e}")
except FileNotFoundError as e:
    print(f"模板文件不存在: {e}")
except Exception as e:
    print(f"生成失败: {e}")
```

## 注意事项

1. 生成的模板结构与用友BIP系统要求完全一致
2. 必填字段从参考模板中动态读取,确保与系统要求一致
3. "过滤标识"列不设置为'√',保持为空,确保数据正常导入
4. 参照工作表包含组织数据,用户可以选择其中的组织
5. 职务工作表包含字段定义、说明和用户数据行
6. 数据行从第9行开始写入,确保与模板结构一致
7. 每个职务的描述文字不超过50字,表达科技公司对不同级别职务的工作要求
8. 配置统一通过Config类管理,避免硬编码

## 检查策略

1. **检查输出的Excel文件格式与模板Excel完全一致**
   - 验证工作表名称:ReadMe、职务、参照
   - 验证职务表结构:第1-8行为结构定义,第9行开始为用户数据
   - 验证必填字段标识:带*号的字段为必填项
   - 验证列数:职务表应包含11列(ERROR_MSG、filter_flag、code、name、org_id_name、jobtype_id_name、jobgrade_id_name、maxrank_id_name、minrank_id_name、duties、memo)

2. **检查必填字段是否全部非空**
   - 职务编号(code):不能为空
   - 职务名称(name):不能为空
   - 所属组织(org_id_name):不能为空
   - 验证逻辑:遍历所有数据行,检查必填字段列的值是否为空或None
   - 如果发现必填字段为空,抛出ValueError异常并输出详细错误信息

3. **检查数据与列的对应关系**
   - 验证用户数据的列数与模板列数一致
   - 验证数据行从第9行开始写入
   - 验证每个字段的值写入正确的列位置
   - 验证职务编码格式:按"职务类型前缀+月日时分秒+序号"格式生成

4. **检查职务描述**
   - 每个职务的描述文字不超过50字
   - 描述内容表达科技公司对不同级别职务的工作要求
   - 描述文字写入"职责"列(J列)

5. **检查filter_flag处理**
   - 确保filter_flag列(第2列)不设置为'√'
   - 如果检测到'√',自动清除为None
   - 确保数据不会被导入过滤器过滤掉

## 示例数据

```python
# 示例1: 从文本输入
user_input = "我需要导入职务档案的数据,包括职务:产品经理、研发经理、需求分析、软件测试"

# 示例2: 已解析数据格式
user_data = [
    [None, None, 'PM0227185046001', '产品经理', '企业账号级##global00', '', '', '', '', '负责产品规划与设计', ''],
    [None, None, 'RDM0227185046001', '研发经理', '企业账号级##global00', '', '', '', '', '负责研发团队管理', ''],
    [None, None, 'REQ0227185046001', '需求分析', '企业账号级##global00', '', '', '', '', '负责收集与分析用户需求', ''],
    [None, None, 'TEST0227185046001', '软件测试', '企业账号级##global00', '', '', '', '', '负责软件测试工作', ''],
]
```

## 错误处理

- 当模板文件不存在时,抛出FileNotFoundError异常
- 当必填字段为空时,抛出ValidationError异常并输出详细错误信息
- 当数据行列数与模板列数不一致时,自动调整数据行长度并输出警告
- 当职务描述超过50字时,自动截断并输出警告
- 当文件保存失败时,输出错误信息并退出

## 版本历史

### v2.0 (2024-02-27) - 重大重构
- 完全重构技能架构,采用模块化设计
- 修复filter_flag错误(不再设置为'√')
- 智能文本解析,支持从自然语言提取职务信息
- 自动职务编码生成
- 完整的数据验证和自动修复机制
- 配置驱动,消除硬编码
- 支持上下文管理器,自动资源管理

### v1.0
- 初始版本
- 基本模板生成功能
