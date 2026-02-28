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

## 交互流程

### 步骤1：解析用户输入

从用户输入中解析需要创建的职务列表：
- "我需要导入职务:产品经理、研发经理" → 解析为 ["产品经理", "研发经理"]
- "生成职务模板，包括：前端开发、后端开发" → 解析为 ["前端开发", "后端开发"]

### 步骤2：调用脚本检查必填字段

**先调用脚本检查模板的必填字段要求**：

```bash
cd skills/bip-duty-template-generator
python scripts/generate.py "用户输入的职务列表"
```

脚本会输出：
- ✓ 检测到的必填字段列表
- ✗ 如果缺少必填信息，会抛出错误并提示

**根据脚本输出，分情况处理**：

**情况A：没有提示必填字段缺失**
- 直接生成模板

**情况B：提示"备注字段为必填项"**
```
使用 ask_followup_question 工具询问：
问题：请为以下职务填写备注信息（每个职务1-2句话描述）
    产品经理、研发经理
```

**情况C：提示缺少组织信息**
```
使用 ask_followup_question 工具询问：
问题：请选择职务所属组织
选项：
1. 企业账号级##global00
2. [从参照表中获取其他组织]
```

### 步骤3：生成MD格式的表格文档，使用 ask_followup_question 让用户确认最终数据

**重要**：最终生成excel文件之前，必须先让用户看到导入数据的数据清单，并由用户确认数据。

### 步骤4：收集完整信息后调用脚本生成模板

当所有必填信息收集完成后，使用完整参数调用脚本：

```bash
cd skills/bip-duty-template-generator
python scripts/generate.py "我需要导入职务:产品经理、研发经理" "企业账号级##global00" '{"产品经理": "负责产品规划", "研发经理": "负责研发管理"}'
```

参数说明：
- 第1个参数：用户输入的职务列表
- 第2个参数：所属组织名称
- 第3个参数：备注信息（JSON格式字典，如需要）

```bash
cd skills/bip-duty-template-generator
python generate.py "我需要导入职务:产品经理、研发经理" "企业账号级##global00"
```

## 使用方法

### 1. 命令行调用

```bash
cd skills/bip-duty-template-generator

# 基本用法（如果备注不是必填）
python scripts/generate.py "用户输入的职务列表" "所属组织"

# 完整用法（如果备注是必填）
python scripts/generate.py "用户输入的职务列表" "所属组织" '{"职务名": "备注内容"}'
```

### 2. Python 代码调用

```python
from generate import generate

# 生成模板（如果备注不是必填）
output_file = generate(
    user_input="我需要导入职务:产品经理、研发经理、需求分析",
    org_name="企业账号级##global00"
)

# 生成模板（如果备注是必填）
output_file = generate(
    user_input="我需要导入职务:产品经理、研发经理",
    org_name="企业账号级##global00",
    remarks={
        "产品经理": "负责产品规划与设计",
        "研发经理": "负责研发团队管理"
    }
)

print(f"生成成功: {output_file}")
```

## 关键提示词（供 AI 使用）

**核心原则：先检查，再询问，最后生成**

1. **第一步**：调用脚本 `python scripts/generate.py "职务列表"` 检查必填字段
2. **第二步**：根据脚本输出，判断需要哪些必填信息
3. **第三步**：如果缺少必填信息，使用 ask_followup_question 询问用户
4. **第四步**：信息收集完整后，使用完整参数调用脚本生成模板

**询问时注意事项**：
- 使用 ask_followup_question 工具
- 提供明确的选项（如组织列表）
- 不要假设或使用默认值
- 如果备注必填，让用户为每个职务提供简要描述

## 输出说明

生成一个Excel文件，包含：
- **ReadMe** - 模板使用指南
- **职务** - 职务档案数据表（从第9行开始写入数据）
- **参照** - 组织参照数据

## 注意事项

1. 生成的模板结构与用友BIP系统要求完全一致
2. 职务编码自动生成，格式为：类型前缀+月日时分秒+序号
3. 每个职务的描述不超过50字
4. filter_flag列不设置为'√'，确保数据正常导入
