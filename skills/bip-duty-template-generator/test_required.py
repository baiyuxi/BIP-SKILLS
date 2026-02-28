#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import sys
sys.path.insert(0, 'scripts')
from generate import check_required_fields

wb = load_workbook("assets/模板_职务.xlsx")
required_fields = check_required_fields(wb)

print("检测到的必填字段:")
for field, col in required_fields.items():
    print(f"  {field}: 列{col}")
