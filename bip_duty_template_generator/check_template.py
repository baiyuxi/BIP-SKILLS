#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查参考模板的结构
"""

import zipfile
import xml.etree.ElementTree as ET

def check_template_structure(template_path):
    """检查模板结构"""
    
    with zipfile.ZipFile(template_path, 'r') as zip_file:
        with zip_file.open('xl/worksheets/sheet2.xml') as sheet_file:
            tree = ET.parse(sheet_file)
            root = tree.getroot()
            
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            rows = root.findall('.//ns:row', ns)
            
            print('模板结构:')
            for row in rows:
                row_num = row.get('r')
                cells_data = []
                cells = row.findall('ns:c', ns)
                for cell in cells:
                    v = cell.find('ns:v', ns)
                    if v is not None and v.text:
                        cells_data.append(v.text)
                    else:
                        cells_data.append('')
                
                if cells_data:
                    print(f'第{row_num}行: {cells_data}')

if __name__ == '__main__':
    template_path = '/Users/baiyuxi/Documents/YONYOUWORK/AI_BIP_PM/.trae/skills/bip-duty-template-generator/references/模板_职务.xlsx'
    
    check_template_structure(template_path)
