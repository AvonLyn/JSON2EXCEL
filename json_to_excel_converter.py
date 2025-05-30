#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JSON to Excel Converter
将医院科室JSON数据转换为Excel表格的GUI工具
支持单文件和批处理模式，动态解析参数
"""

import json
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from datetime import datetime
import threading
import glob


class JSONToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON转Excel工具 - 医院科室数据转换器")
        self.root.geometry("900x700")
        
        # 存储当前处理的数据
        self.current_data = None
        self.current_filename = None
        self.hospital_id = None
        self.baseurl = None
        self.url_params = []  # 从baseurl中提取的参数列表
        
        # 批处理模式
        self.batch_mode = tk.BooleanVar(value=False)
        self.input_dir_var = tk.StringVar(value=".")  # 添加输入目录变量
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        """创建GUI界面组件"""
        # 顶部框架 - 模式选择
        mode_frame = ttk.Frame(self.root, padding="10")
        mode_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Label(mode_frame, text="处理模式:").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(mode_frame, text="单文件模式", variable=self.batch_mode, 
                       value=False, command=self.on_mode_change).grid(row=0, column=1, padx=5)
        ttk.Radiobutton(mode_frame, text="批处理模式", variable=self.batch_mode, 
                       value=True, command=self.on_mode_change).grid(row=0, column=2, padx=5)
        
        # 文件选择框架
        self.file_frame = ttk.Frame(self.root, padding="10")
        self.file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 单文件模式控件
        self.single_file_widgets = []
        label1 = ttk.Label(self.file_frame, text="选择JSON文件:")
        label1.grid(row=0, column=0, padx=5)
        self.file_path_var = tk.StringVar()
        entry1 = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=50)
        entry1.grid(row=0, column=1, padx=5)
        btn1 = ttk.Button(self.file_frame, text="浏览", command=self.select_file)
        btn1.grid(row=0, column=2, padx=5)
        btn2 = ttk.Button(self.file_frame, text="解析", command=self.parse_json)
        btn2.grid(row=0, column=3, padx=5)
        self.single_file_widgets = [label1, entry1, btn1, btn2]
        
        # 批处理模式控件（初始隐藏）
        self.batch_widgets = []
        # 输入目录行
        label_input = ttk.Label(self.file_frame, text="输入目录:")
        entry_input = ttk.Entry(self.file_frame, textvariable=self.input_dir_var, width=30)
        btn_input = ttk.Button(self.file_frame, text="选择目录", command=self.select_input_dir)
        # 文件模式行
        label2 = ttk.Label(self.file_frame, text="文件模式:")
        self.pattern_var = tk.StringVar(value="*_triage.json")
        entry2 = ttk.Entry(self.file_frame, textvariable=self.pattern_var, width=30)
        # 输出目录行
        label3 = ttk.Label(self.file_frame, text="输出目录:")
        self.output_dir_var = tk.StringVar(value="output")
        entry3 = ttk.Entry(self.file_frame, textvariable=self.output_dir_var, width=30)
        btn3 = ttk.Button(self.file_frame, text="选择目录", command=self.select_output_dir)
        # 开始按钮
        btn4 = ttk.Button(self.file_frame, text="开始批处理", command=self.start_batch_process)
        self.batch_widgets = [label_input, entry_input, btn_input, label2, entry2, label3, entry3, btn3, btn4]
        
        # 中部框架 - 数据预览
        middle_frame = ttk.Frame(self.root, padding="10")
        middle_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(middle_frame, text="数据预览:").grid(row=0, column=0, sticky=tk.W)
        
        # 创建Treeview用于显示表格数据
        self.tree_frame = ttk.Frame(middle_frame)
        self.tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        self.tree_scroll_y = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建Treeview
        self.tree = ttk.Treeview(self.tree_frame, 
                                yscrollcommand=self.tree_scroll_y.set,
                                xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        
        # 底部框架 - 导出和状态
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        self.export_btn = ttk.Button(bottom_frame, text="导出Excel", command=self.export_excel)
        self.export_btn.grid(row=0, column=0, padx=5)
        
        # 进度条（批处理用）
        self.progress = ttk.Progressbar(bottom_frame, length=300, mode='determinate')
        self.progress.grid(row=0, column=1, padx=20)
        self.progress.grid_remove()  # 初始隐藏
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        ttk.Label(bottom_frame, textvariable=self.status_var).grid(row=1, column=0, columnspan=2, pady=5)
        
        # 配置grid权重
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        middle_frame.grid_columnconfigure(0, weight=1)
        middle_frame.grid_rowconfigure(1, weight=1)
        
        # 初始化界面状态
        self.on_mode_change()
    
    def on_mode_change(self):
        """切换模式时更新界面"""
        if self.batch_mode.get():
            # 批处理模式
            for widget in self.single_file_widgets:
                widget.grid_remove()
            
            # 批处理模式布局
            # 输入目录行
            self.batch_widgets[0].grid(row=0, column=0, padx=5, sticky=tk.E)  # 输入目录标签
            self.batch_widgets[1].grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))  # 输入目录输入
            self.batch_widgets[2].grid(row=0, column=2, padx=5)  # 选择输入目录按钮
            
            # 文件模式行
            self.batch_widgets[3].grid(row=1, column=0, padx=5, sticky=tk.E)  # 文件模式标签
            self.batch_widgets[4].grid(row=1, column=1, padx=5, sticky=(tk.W, tk.E))  # 文件模式输入
            
            # 输出目录行
            self.batch_widgets[5].grid(row=2, column=0, padx=5, sticky=tk.E)  # 输出目录标签
            self.batch_widgets[6].grid(row=2, column=1, padx=5, sticky=(tk.W, tk.E))  # 输出目录输入
            self.batch_widgets[7].grid(row=2, column=2, padx=5)  # 选择输出目录按钮
            
            # 开始按钮
            self.batch_widgets[8].grid(row=3, column=0, columnspan=3, pady=10)  # 开始批处理按钮
            
            # 配置列权重
            self.file_frame.grid_columnconfigure(1, weight=1)
            
            self.export_btn.config(state='disabled')
        else:
            # 单文件模式
            for widget in self.batch_widgets:
                widget.grid_remove()
            
            for i, widget in enumerate(self.single_file_widgets):
                widget.grid(row=0, column=i, padx=5)
            
            self.export_btn.config(state='normal')
            self.progress.grid_remove()
    
    def select_file(self):
        """选择JSON文件"""
        filename = filedialog.askopenfilename(
            title="选择JSON文件",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.current_filename = os.path.basename(filename)
            # 从文件名提取hospital_id
            match = re.match(r'^(\d+)', self.current_filename)
            if match:
                self.hospital_id = match.group(1)
                self.status_var.set(f"已选择文件: {self.current_filename}, Hospital ID: {self.hospital_id}")
            else:
                self.hospital_id = None
                self.status_var.set(f"警告: 无法从文件名提取Hospital ID")
    
    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir_var.set(directory)
    
    def select_input_dir(self):
        """选择输入目录"""
        directory = filedialog.askdirectory(title="选择输入目录")
        if directory:
            self.input_dir_var.set(directory)
    
    def extract_url_params(self, url_pattern):
        """从URL模式中提取参数名"""
        # 查找所有 {param} 格式的参数
        params = re.findall(r'\{(\w+)\}', url_pattern)
        return params
    
    def parse_json(self):
        """解析JSON文件并显示预览"""
        if not self.file_path_var.get():
            messagebox.showwarning("警告", "请先选择JSON文件")
            return
            
        try:
            # 读取JSON文件
            with open(self.file_path_var.get(), 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 解析baseurl
            self.baseurl, self.url_params = self.extract_baseurl(data)
            
            # 解析数据
            self.current_data = self.extract_department_data(data, self.hospital_id, self.baseurl)
            
            # 显示预览
            self.display_preview()
            
            self.status_var.set(f"解析成功: 共{len(self.current_data)}条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"解析失败: {str(e)}")
            self.status_var.set("解析失败")
    
    def extract_baseurl(self, json_data):
        """提取baseurl信息"""
        baseurl_list = json_data.get('baseurl', [])
        if baseurl_list and len(baseurl_list) > 0:
            # 获取第一个baseurl对象
            baseurl_obj = baseurl_list[0]
            # 查找url_pattern或类似的键
            for key in ['url_pattern', 'url', 'base_url', 'baseurl']:
                if key in baseurl_obj:
                    url = baseurl_obj[key]
                    # 提取URL中的参数
                    params = self.extract_url_params(url)
                    return url, params
        return None, []
    
    def extract_department_data(self, json_data, hospital_id, baseurl, url_params=None):
        """从JSON数据中提取科室信息"""
        rows = []
        
        # 如果没有提供url_params，使用实例变量
        if url_params is None:
            url_params = self.url_params
        
        # 获取departments数组
        departments = json_data.get('departments', [])
        
        for dept in departments:
            # 获取基本信息
            dept_title = dept.get('title', '')
            symptom_text = dept.get('symptom_text', '')
            diagnosis_text = dept.get('diagnosis_text', '')
            
            # 查找包含科室列表的key（可能是'data'、空字符串或其他）
            data_list = None
            for key, value in dept.items():
                if isinstance(value, list) and key not in ['title', 'symptom_text', 'diagnosis_text']:
                    data_list = value
                    break
            
            if data_list:
                # 遍历每个院区
                for campus_data in data_list:
                    campus_id = campus_data.get('campus_id', '')
                    department_list = campus_data.get('department_list', [])
                    
                    # 遍历每个具体科室
                    for dept_item in department_list:
                        row = {
                            'hospital_id': hospital_id or '',
                            'baseurl': baseurl or '',
                            'campus_id': campus_id,
                            'department_title': dept_item.get('title', ''),
                            'symptom_text': symptom_text,
                            'diagnosis_text': diagnosis_text
                        }
                        
                        # 收集所有URL参数到一个字典中
                        url_params_dict = {}
                        
                        # 动态提取参数
                        if 'params' in dept_item:
                            # 193格式：参数在params对象中
                            params = dept_item.get('params', {})
                            for param in url_params:
                                # 尝试不同的键名变体
                                value = params.get(param) or params.get(param.lower()) or params.get(param.upper())
                                url_params_dict[param] = value or ''
                        else:
                            # 63格式：直接在dept_item中查找
                            for param in url_params:
                                if param == 'departId':
                                    # 特殊处理departId -> department_id的映射
                                    value = dept_item.get('department_id', '')
                                elif param == 'title' or param == 'departName':
                                    # title通常对应department_title
                                    value = dept_item.get('title', '')
                                else:
                                    # 其他参数直接查找
                                    value = dept_item.get(param) or dept_item.get(param.lower()) or ''
                                url_params_dict[param] = value
                            
                            # 添加position字段（如果存在）到参数字典中
                            if 'position' in dept_item:
                                url_params_dict['position'] = dept_item['position']
                        
                        # 将参数字典转换为JSON字符串，保存在单个单元格中
                        row['url_params_json'] = json.dumps(url_params_dict, ensure_ascii=False)
                        
                        rows.append(row)
        
        return rows
    
    def display_preview(self):
        """在Treeview中显示数据预览"""
        # 清除现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.current_data:
            return
        
        # 设置列（动态列）
        if self.current_data:
            columns = list(self.current_data[0].keys())
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            # 设置列标题和宽度
            for col in columns:
                self.tree.heading(col, text=col)
                # 根据列名设置不同的宽度
                if col in ['hospital_id', 'campus_id']:
                    width = 80
                elif col in ['symptom_text', 'diagnosis_text', 'baseurl']:
                    width = 200
                elif col == 'url_params_json':
                    width = 300  # 为JSON列设置更大的宽度
                else:
                    width = 120
                self.tree.column(col, width=width)
            
            # 插入数据（只显示前100条以提高性能）
            for i, row in enumerate(self.current_data[:100]):
                values = [row.get(col, '') for col in columns]
                self.tree.insert('', 'end', values=values)
            
            if len(self.current_data) > 100:
                self.tree.insert('', 'end', values=['...' for _ in columns])
    
    def export_excel(self):
        """导出数据到Excel文件"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先解析JSON文件")
            return
        
        # 选择保存位置
        filename = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"{self.hospital_id}_departments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not filename:
            return
        
        try:
            self.save_to_excel(self.current_data, filename)
            messagebox.showinfo("成功", f"Excel文件已保存: {filename}")
            self.status_var.set(f"导出成功: {len(self.current_data)}条记录")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            self.status_var.set("导出失败")
    
    def save_to_excel(self, data, filename):
        """保存数据到Excel文件"""
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 导出到Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='科室数据', index=False)
            
            # 获取worksheet对象以调整列宽
            worksheet = writer.sheets['科室数据']
            
            # 自动调整列宽
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                # 设置最大宽度
                if col == 'baseurl':
                    worksheet.column_dimensions[chr(65 + idx)].width = 80
                elif col == 'url_params_json':
                    worksheet.column_dimensions[chr(65 + idx)].width = 100  # 为JSON列设置较大宽度
                else:
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
    
    def start_batch_process(self):
        """开始批处理"""
        pattern = self.pattern_var.get()
        output_dir = self.output_dir_var.get()
        input_dir = self.input_dir_var.get()
        
        if not pattern:
            messagebox.showwarning("警告", "请输入文件模式")
            return
        
        # 预先查找匹配的文件
        json_files = glob.glob(os.path.join(input_dir, pattern))
        
        if not json_files:
            messagebox.showwarning("警告", f"在目录 {input_dir} 中没有找到匹配 {pattern} 的文件")
            return
        
        # 显示找到的文件列表，让用户确认
        file_list = "\n".join([os.path.basename(f) for f in json_files[:10]])
        if len(json_files) > 10:
            file_list += f"\n... 和其他 {len(json_files) - 10} 个文件"
        
        msg = f"找到 {len(json_files)} 个文件:\n\n{file_list}\n\n是否继续处理？"
        if not messagebox.askyesno("确认", msg):
            return
        
        # 在新线程中执行批处理
        thread = threading.Thread(target=self.batch_process, args=(json_files, output_dir))
        thread.daemon = True
        thread.start()
    
    def batch_process(self, json_files, output_dir):
        """批处理函数"""
        try:
            # 创建输出目录
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 更新进度条
            self.root.after(0, self.progress.grid)
            self.root.after(0, lambda: self.progress.config(maximum=len(json_files)))
            
            success_count = 0
            error_files = []
            
            for i, json_file in enumerate(json_files):
                try:
                    # 更新状态
                    self.root.after(0, lambda f=json_file: self.status_var.set(f"正在处理: {os.path.basename(f)}"))
                    
                    # 提取hospital_id
                    hospital_id = self.extract_hospital_id(json_file)
                    
                    # 读取和解析JSON
                    with open(json_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    # 提取baseurl
                    baseurl, url_params = self.extract_baseurl(data)
                    
                    # 提取数据
                    rows = self.extract_department_data(data, hospital_id, baseurl, url_params)
                    
                    if rows:
                        # 生成输出文件名
                        base_name = os.path.splitext(os.path.basename(json_file))[0]
                        output_file = os.path.join(output_dir, 
                                                  f"{base_name}_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                        
                        # 保存到Excel
                        self.save_to_excel(rows, output_file)
                        success_count += 1
                    else:
                        error_files.append((json_file, "没有找到科室数据"))
                    
                except Exception as e:
                    error_files.append((json_file, str(e)))
                    print(f"处理 {json_file} 时出错: {str(e)}")
                
                # 更新进度
                self.root.after(0, lambda v=i+1: self.progress.config(value=v))
            
            # 完成
            self.root.after(0, lambda: self.status_var.set(f"批处理完成: 成功 {success_count}/{len(json_files)} 个文件"))
            
            # 构建结果消息
            msg = f"批处理完成\n成功处理 {success_count}/{len(json_files)} 个文件\n输出目录: {output_dir}"
            if error_files:
                msg += "\n\n以下文件处理失败:"
                for file, error in error_files[:5]:
                    msg += f"\n{os.path.basename(file)}: {error}"
                if len(error_files) > 5:
                    msg += f"\n... 和其他 {len(error_files) - 5} 个文件"
            
            self.root.after(0, lambda: messagebox.showinfo("完成", msg))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"批处理失败: {str(e)}"))
        finally:
            self.root.after(0, self.progress.grid_remove)
    
    def extract_hospital_id(self, filename):
        """从文件名提取hospital_id"""
        basename = os.path.basename(filename)
        match = re.match(r'^(\d+)', basename)
        return match.group(1) if match else None


def main():
    """主函数"""
    root = tk.Tk()
    app = JSONToExcelConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main() 