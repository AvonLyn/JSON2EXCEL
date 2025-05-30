#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量转换JSON文件到Excel
"""

import os
import json
import glob
import pandas as pd
from datetime import datetime
import re
import argparse


def extract_hospital_id(filename):
    """从文件名提取hospital_id"""
    basename = os.path.basename(filename)
    match = re.match(r'^(\d+)', basename)
    return match.group(1) if match else None


def extract_department_data(json_data, hospital_id):
    """从JSON数据中提取科室信息"""
    rows = []
    
    # 获取departments数组
    departments = json_data.get('departments', [])
    
    for dept in departments:
        # 获取基本信息
        symptom_text = dept.get('symptom_text', '')
        diagnosis_text = dept.get('diagnosis_text', '')
        
        # 获取data数组（注意原始JSON中有个空字符串的key）
        data_list = None
        for key in ['data', '']:  # 检查'data'和空字符串key
            if key in dept and isinstance(dept[key], list):
                data_list = dept[key]
                break
        
        if data_list:
            # 遍历每个院区
            for campus_data in data_list:
                campus_id = campus_data.get('campus_id', '')
                department_list = campus_data.get('department_list', [])
                
                # 遍历每个具体科室
                for dept_item in department_list:
                    # 检查是否有params结构（193_triage.json格式）
                    if 'params' in dept_item:
                        params = dept_item.get('params', {})
                        row = {
                            'hospital_id': hospital_id or '',
                            'campus_id': campus_id,
                            'department_title': dept_item.get('title', ''),
                            'department_id': params.get('departId', ''),
                            'area_id': params.get('areaId', ''),
                            'area_name': params.get('areaName', ''),
                            'position': '',  # 193格式中没有position
                            'symptom_text': symptom_text,
                            'diagnosis_text': diagnosis_text
                        }
                    else:
                        # 原格式（63_triage.json格式）
                        row = {
                            'hospital_id': hospital_id or '',
                            'campus_id': campus_id,
                            'department_title': dept_item.get('title', ''),
                            'department_id': dept_item.get('department_id', ''),
                            'area_id': '',  # 63格式中没有area_id
                            'area_name': '',  # 63格式中没有area_name
                            'position': dept_item.get('position', ''),
                            'symptom_text': symptom_text,
                            'diagnosis_text': diagnosis_text
                        }
                    rows.append(row)
    
    return rows


def process_json_file(json_file, output_dir):
    """处理单个JSON文件"""
    try:
        # 提取hospital_id
        hospital_id = extract_hospital_id(json_file)
        if not hospital_id:
            print(f"⚠️  警告: 无法从文件名提取hospital_id: {json_file}")
        
        # 读取JSON文件
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 提取数据
        rows = extract_department_data(data, hospital_id)
        
        if not rows:
            print(f"⚠️  警告: {json_file} 中没有找到科室数据")
            return False
        
        # 创建DataFrame
        df = pd.DataFrame(rows)
        
        # 按照指定顺序排列列
        column_order = ['hospital_id', 'campus_id', 'department_title', 'department_id', 
                       'area_id', 'area_name', 'position', 'symptom_text', 'diagnosis_text']
        df = df[column_order]
        
        # 生成输出文件名
        base_name = os.path.splitext(os.path.basename(json_file))[0]
        output_file = os.path.join(output_dir, 
                                  f"{base_name}_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # 导出到Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='科室数据', index=False)
            
            # 获取worksheet对象以调整列宽
            worksheet = writer.sheets['科室数据']
            
            # 自动调整列宽
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                # 设置最大宽度为50
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
        
        print(f"✅ 成功: {json_file} → {output_file} ({len(rows)}条记录)")
        return True
        
    except Exception as e:
        print(f"❌ 错误: 处理 {json_file} 时发生错误: {str(e)}")
        return False


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='批量转换JSON文件到Excel')
    parser.add_argument('input_pattern', nargs='?', default='*_triage.json',
                       help='输入文件模式 (默认: *_triage.json)')
    parser.add_argument('-o', '--output', default='output',
                       help='输出目录 (默认: output)')
    
    args = parser.parse_args()
    
    # 创建输出目录
    if not os.path.exists(args.output):
        os.makedirs(args.output)
    
    # 查找所有匹配的JSON文件
    json_files = glob.glob(args.input_pattern)
    
    if not json_files:
        print(f"没有找到匹配的JSON文件: {args.input_pattern}")
        return
    
    print(f"找到 {len(json_files)} 个JSON文件")
    print(f"输出目录: {args.output}")
    print("-" * 50)
    
    # 处理每个文件
    success_count = 0
    for json_file in json_files:
        if process_json_file(json_file, args.output):
            success_count += 1
    
    print("-" * 50)
    print(f"处理完成: 成功 {success_count}/{len(json_files)} 个文件")


if __name__ == "__main__":
    main() 