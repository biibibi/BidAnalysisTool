# 废标项清单更新流程

## 概述
本文档说明如何从ReferenceList.xlsx文件更新bid_analysis.html中的废标项检测清单。

## 必需文件
以下文件是更新流程的核心组件，请勿删除：

### 1. 源数据文件
- `ref/ReferenceList.xlsx` - Excel格式的废标项清单源文件

### 2. 处理脚本
- `read_reference_list.py` - 读取和分析Excel文件
- `generate_violation_items.py` - 将Excel数据转换为JavaScript格式
- `update_html.py` - 更新HTML文件中的violationItems数组

### 3. 目标文件
- `frontend/bid_analysis.html` - 需要更新的网页文件

## 更新流程

### 步骤1: 更新Excel文件
1. 编辑 `ref/ReferenceList.xlsx` 文件
2. 确保数据格式符合要求：
   - 列名：序号、类别、检测项目、要求说明、A角检查、B角复核、交叉检查
   - 类别包括：常规检查、技术常规检查、商务常规检查、技术定制检查

### 步骤2: 运行处理脚本
```bash
# 激活虚拟环境
.\venv\Scripts\activate

# 1. 读取Excel文件（可选，用于验证数据）
python read_reference_list.py

# 2. 生成JavaScript代码
python generate_violation_items.py

# 3. 更新HTML文件
python update_html.py
```

### 步骤3: 验证更新
1. 检查控制台输出确认更新成功
2. 启动服务器测试页面功能
3. 验证废标项数量和内容是否正确

## 类别映射
Excel中的类别会自动映射到网页中的分类：
- 常规检查 → 资质要求 (category-qualification)
- 技术常规检查 → 技术要求 (category-technical) 
- 商务常规检查 → 商务要求 (category-commercial)
- 技术定制检查 → 其他要求 (category-format)

## 注意事项
1. 更新前请备份原始文件
2. 确保Excel文件格式正确
3. 检查生成的JavaScript代码语法是否正确
4. 验证页面显示效果

## 最后更新
本次更新将废标项从37项更新为基于ReferenceList.xlsx的37项，采用了Excel中的分类结构。