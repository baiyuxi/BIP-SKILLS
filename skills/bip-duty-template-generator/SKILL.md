---
name: bip-duty-template-generator
description: Generate Excel template for importing duty (job profile) data into Yonyou BIP system.
---

# 用友BIP职务档案导入模板生成器

当用户需要导入职务数据、生成职务导入模板、创建职务档案模板或处理BIP系统中与职务/档案/导入相关的任务时，使用此技能。

## 功能说明

本技能用于快速生成符合用友BIP系统要求的职务档案导入Excel模板文件。

## 使用场景

- 用户说"导入职务"
- 用户说"生成职务模板"  
- 用户需要批量创建职务档案
- 用户询问如何导入职务数据到BIP系统

## 使用方法

### 1. 调用生成脚本

```bash
cd skills/bip-duty-template-generator
python generate.py "我需要导入职务:产品经理、研发经理"
```

### 2. 通过 Python 代码调用

```python
from generate import generate

# 生成模板
output_file = generate(
    user_input="我需要导入职务:产品经理、研发经理、需求分析",
    org_name="企业账号级##global00"  # 可选，默认使用第一个组织
)

print(f"生成成功: {output_file}")
```

## 输入格式

支持从自然语言中解析职务名称：

- "我需要导入职务:产品经理、研发经理"
- "生成职务模板，包括：前端开发、后端开发"
- "包括职务:产品经理、UI设计、软件测试"

## 输出说明

生成一个Excel文件，包含：
- **ReadMe** - 模板使用指南
- **职务** - 职务档案数据表（从第9行开始写入数据）
- **参照** - 组织参照数据

## 必填字段

当前模板中的必填字段：
- 职务编号 (code) - 自动生成
- 职务名称 (name) - 从用户输入解析
- 所属组织 (org_id_name) - 需要指定

## 注意事项

1. 生成的模板结构与用友BIP系统要求完全一致
2. 职务编码自动生成，格式为：类型前缀+月日时分秒+序号
3. 每个职务的描述不超过50字
4. filter_flag列不设置为'√'，确保数据正常导入
