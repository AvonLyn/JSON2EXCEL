# JSON to Excel Converter - 医院科室数据转换工具

这是一个将医院科室JSON数据转换为Excel表格的GUI工具。

## 功能特点

- 图形化操作界面，简单易用
- 自动从文件名提取hospital_id
- 实时预览转换后的数据
- 支持导出为标准Excel格式(.xlsx)
- 自动调整Excel列宽
- **兼容多种JSON格式**（支持63_triage.json和193_triage.json格式）
- **批处理模式集成**（可在GUI中切换单文件/批处理模式）
- **动态参数解析**（根据baseurl自动确定需要提取的参数）

## 新功能说明

### 动态参数解析

程序会自动解析JSON文件中的`baseurl`字段，提取URL模板中的参数，并动态创建Excel列。

例如，如果baseurl为：
```
http://lyrmyy.ylzpay.com/hospitalPortal-web/linyi/linyirenminyiyuan/appoint/schedulingSource?areaId={areaId}&areaName={areaName}&departId={departId}&departName={title}
```

程序会自动识别并提取以下参数：
- areaId
- areaName
- departId
- title

这些参数会作为Excel的列动态生成，无需手动配置。

### 批处理模式

在GUI界面中可以切换到批处理模式：
1. 选择"批处理模式"单选按钮
2. 选择输入目录（包含JSON文件的目录）
3. 输入文件匹配模式（如 `*_triage.json`）
4. 选择输出目录
5. 点击"开始批处理"

程序会：
- 在指定的输入目录中搜索匹配的文件
- 显示找到的文件列表供确认
- 自动处理所有匹配的文件
- 显示进度条和处理结果
- 报告处理失败的文件及原因

## 支持的JSON格式

### 格式1（如63_triage.json）
```json
{
    "title": "慢阻肺门诊",
    "": [  // 空字符串key
        {
            "campus_id": 1,
            "department_list": [{
                "title": "慢阻肺门诊",
                "department_id": "3001",
                "position": "门诊楼3楼"
            }]
        }
    ]
}
```

### 格式2（如193_triage.json）
```json
{
    "title": "PICC门诊",
    "data": [  // "data"作为key
        {
            "campus_id": 2,
            "department_list": [{
                "title": "PICC门诊",
                "params": {
                    "areaId": "16",
                    "areaName": "北城院区",
                    "departId": "218"
                }
            }]
        }
    ]
}
```

## 安装依赖

确保已安装Python 3.6或更高版本，然后运行：

```bash
pip install -r requirements.txt
```

## 使用方法

### 单文件模式

1. 运行程序：
```bash
python json_to_excel_converter.py
```

2. 在GUI界面中：
   - 选择"单文件模式"（默认）
   - 点击"浏览"按钮选择JSON文件
   - 点击"解析"按钮解析JSON文件
   - 预览区域会显示解析后的数据
   - 点击"导出Excel"按钮保存为Excel文件

### 批处理模式

1. 运行程序后选择"批处理模式"
2. 输入文件匹配模式（如 `*.json` 或 `193*.json`）
3. 选择或输入输出目录
4. 点击"开始批处理"

## Excel输出格式

输出的Excel文件包含以下列：
- **hospital_id**: 医院ID（从文件名提取）
- **baseurl**: 基础URL（每行都相同）
- **campus_id**: 院区ID
- **department_title**: 科室名称
- **symptom_text**: 症状文本
- **diagnosis_text**: 诊断文本
- **url_params_json**: URL参数的JSON对象（包含所有从baseurl中提取的参数及其值）

### url_params_json 列说明

所有从baseurl模板中提取的参数（如areaId、areaName、departId等）现在会被合并成一个JSON对象，保存在`url_params_json`列中。

例如，对于以下baseurl：
```
http://lyrmyy.ylzpay.com/hospitalPortal-web/linyi/linyirenminyiyuan/appoint/schedulingSource?areaId={areaId}&areaName={areaName}&departId={departId}&departName={title}
```

`url_params_json`列的内容可能是：
```json
{
  "areaId": "16",
  "areaName": "北城院区",
  "departId": "218",
  "title": "PICC门诊"
}
```

这种格式的优点：
- 所有URL参数集中在一个单元格中，便于管理
- 保持了数据的结构化，便于程序解析
- 减少了Excel列数，使表格更加简洁

## 注意事项

- JSON文件名应以数字开头（如：63_triage.json），程序会自动提取这个数字作为hospital_id
- 如果文件名不符合格式，hospital_id将为空
- 程序会自动识别JSON格式和参数结构，无需手动配置
- 批处理模式下，每个文件的输出会带有时间戳，避免覆盖

## 示例

输入文件：
- `63_triage.json` - 格式1
- `193_triage.json` - 格式2

输出文件：
- `63_departments_20231125_143025.xlsx`
- `193_departments_20231125_143025.xlsx`

其中每一行代表一个具体的科室信息，hospital_id和baseurl列的值在同一个文件中保持一致。

## 批量处理（命令行）

除了GUI中的批处理模式，还可以使用独立的批量转换脚本：

```bash
# 处理所有 *_triage.json 文件
python batch_converter.py

# 处理特定模式的文件
python batch_converter.py "193*.json"

# 指定输出目录
python batch_converter.py -o my_output_dir

# 查看帮助
python batch_converter.py -h
```

批量处理的特点：
- 自动查找匹配的JSON文件
- 批量转换为Excel文件
- 生成带时间戳的输出文件名
- 显示处理进度和结果统计 