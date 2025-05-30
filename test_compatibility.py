#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试JSON转Excel转换器的兼容性
"""

import json
import pandas as pd
from json_to_excel_converter import JSONToExcelConverter
import tkinter as tk


def test_json_formats():
    """测试不同格式的JSON文件"""
    
    # 创建一个隐藏的Tk窗口（仅用于测试）
    root = tk.Tk()
    root.withdraw()
    
    converter = JSONToExcelConverter(root)
    
    # 测试数据：63格式
    test_data_63 = {
        "departments": [
            {
                "title": "慢阻肺门诊",
                "symptom_text": "呼吸困难",
                "diagnosis_text": "慢性阻塞性肺疾病",
                "": [
                    {
                        "campus_id": 1,
                        "department_list": [
                            {
                                "title": "慢阻肺门诊",
                                "department_id": "3001",
                                "position": "门诊楼3楼"
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    # 测试数据：193格式
    test_data_193 = {
        "departments": [
            {
                "title": "PICC门诊",
                "symptom_text": "需要静脉输液",
                "diagnosis_text": "长期静脉治疗",
                "data": [
                    {
                        "campus_id": 2,
                        "department_list": [
                            {
                                "title": "PICC门诊",
                                "params": {
                                    "areaId": "16",
                                    "areaName": "北城院区",
                                    "departId": "218"
                                }
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    # 设置hospital_id
    converter.hospital_id = "63"
    
    # 测试63格式
    print("测试63格式JSON...")
    rows_63 = converter.extract_department_data(test_data_63)
    print(f"提取到 {len(rows_63)} 条记录")
    if rows_63:
        print("示例记录：", rows_63[0])
    
    # 设置hospital_id
    converter.hospital_id = "193"
    
    # 测试193格式
    print("\n测试193格式JSON...")
    rows_193 = converter.extract_department_data(test_data_193)
    print(f"提取到 {len(rows_193)} 条记录")
    if rows_193:
        print("示例记录：", rows_193[0])
    
    # 验证两种格式都能正确解析
    assert len(rows_63) == 1, "63格式解析失败"
    assert len(rows_193) == 1, "193格式解析失败"
    
    # 验证63格式的特定字段
    assert rows_63[0]['department_id'] == '3001'
    assert rows_63[0]['position'] == '门诊楼3楼'
    assert rows_63[0]['area_id'] == ''  # 63格式没有area_id
    
    # 验证193格式的特定字段
    assert rows_193[0]['department_id'] == '218'
    assert rows_193[0]['area_id'] == '16'
    assert rows_193[0]['area_name'] == '北城院区'
    assert rows_193[0]['position'] == ''  # 193格式没有position
    
    print("\n✅ 所有测试通过！两种格式都能正确解析。")
    
    # 创建测试Excel
    print("\n创建测试Excel文件...")
    all_rows = rows_63 + rows_193
    df = pd.DataFrame(all_rows)
    
    # 按照指定顺序排列列
    column_order = ['hospital_id', 'campus_id', 'department_title', 'department_id', 
                    'area_id', 'area_name', 'position', 'symptom_text', 'diagnosis_text']
    df = df[column_order]
    
    # 保存到Excel
    df.to_excel('test_compatibility_output.xlsx', index=False)
    print("测试Excel文件已保存为: test_compatibility_output.xlsx")
    
    root.destroy()


if __name__ == "__main__":
    test_json_formats() 