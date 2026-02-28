# bip-duty-template-generator

用友BIP职务档案导入模板生成器 - 基于自然语言描述自动生成符合用友BIP系统要求的职务导入Excel模板。

## 功能特性

- 支持自然语言输入描述职务信息
- 自动解析职务名称、职级等信息
- 智能处理必填字段（职务名称、所属组织）
- 完整保留模板格式、样式、列宽
- 支持交互式输入缺失字段

## 环境要求

- Python 3.7+
- openpyxl

## 安装

```bash
pip install openpyxl
```

## 使用方法

### 命令行方式

```bash
cd bip_duty_template_generator
python generate_duty_template.py --input "包括职务:产品经理、研发经理"
```

### Python脚本调用

```python
import sys
import os

# 添加模块路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 切换到模块目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from generate_duty_template import generate_template

# 生成模板
output = generate_template(
    user_input="我需要导入职务,包括职务:产品经理、研发经理"
)
print(f"生成成功: {output}")
```

### 交互式输入

```python
from generate_duty_template import DutyTemplateGenerator

with DutyTemplateGenerator() as generator:
    output = generator.generate(
        user_input="包括职务:产品经理",
        interactive=True,
        ask_missing=True
    )
```

## 输出说明

生成的Excel文件包含以下工作表：
1. **ReadMe** - 使用指南
2. **职务** - 职务数据（从第9行开始）
3. **参照** - 参照数据

## 必填字段说明

当前模板中的必填字段：
- 职务名称 (name) - 从用户输入中解析
- 所属组织 (org_id_name) - 从"参照"页签中选择

## 目录结构

```
bip-duty-template-generator/
├── bip_duty_template_generator/  # 核心代码
│   ├── core/          # 核心逻辑
│   ├── config/        # 配置
│   ├── parsers/       # 解析器
│   ├── utils/         # 工具函数
│   ├── references/    # 模板文件
│   └── *.py           # 入口脚本
├── README.md
└── SKILL.md
```

## License

MIT License
