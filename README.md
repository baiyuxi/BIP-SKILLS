# BIP-SKILLS

用友BIP技能库 - 存放各类BIP系统相关的AI技能。

## 技能列表

### 1. bip-duty-template-generator
职务档案导入模板生成器 - 基于自然语言描述自动生成符合用友BIP系统要求的职务导入Excel模板。

**功能**：
- 自然语言输入描述职务信息
- 自动解析职务名称、职级等信息
- 智能处理必填字段
- 完整保留模板格式

**使用方式**：
```bash
cd bip-duty-template-generator/bip_duty_template_generator
python generate_duty_template.py --input "包括职务:产品经理、研发经理"
```

---

## 添加新技能

如需添加新技能，在仓库根目录创建新文件夹即可，结构如下：
```
BIP-SKILLS/
├── 技能名称/
│   ├── 代码文件/
│   ├── README.md
│   └── SKILL.md
└── README.md
```

## 许可证

MIT License
