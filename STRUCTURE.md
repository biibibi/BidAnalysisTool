# 项目结构

```
BidAnalysis/
├── README.md                   # 项目说明文档
├── LICENSE                     # MIT许可证
├── .gitignore                 # Git忽略文件配置
├── start.bat                  # Windows启动脚本
├── start.sh                   # Linux/macOS启动脚本
├── 课题方案书.md                # 项目方案文档
├── analyze_doc_structure.py   # 文档结构分析工具
├── examine_database.py        # 数据库检查工具
├── clear_files_table.py       # 数据库清理工具
├── bid_analysis.db           # SQLite数据库文件
├── 
├── backend/                   # 后端代码
│   ├── app.py                # Flask主应用
│   ├── qwen_service.py       # Qwen大模型服务
│   ├── file_handler.py       # 文件处理服务
│   ├── database.py           # 数据库管理
│   ├── run.py               # 启动脚本
│   ├── requirements.txt     # Python依赖
│   ├── .env.template       # 环境变量模板
│   ├── .env               # 环境变量配置（需要配置）
│   └── ai_agents/           # AI智能代理模块
│       ├── __init__.py
│       ├── agent_manager.py     # 代理管理器
│       ├── base_agent.py        # 基础代理类
│       ├── project_info_agent.py # 项目信息提取代理
│       ├── image_extraction_agent.py # Word图片提取与AI命名（替代原 word_image_separator 脚本）
│       ├── authorization_letter_agent.py # 授权委托书/身份证明多模态核验
│       ├── word_splitter.py     # Word文档拆分工具
│       └── wordtoc_agent.py     # Word目录提取代理
├── 
├── frontend/                 # 前端代码
│   ├── index.html          # 主页面 - 招标文件分析
│   └── bid_analysis.html   # 投标文件合规性检测页面
├── 
├── static/                  # 静态资源
│   └── icon/              # 图标文件
│       ├── Unicom.jpeg
│       ├── Unicom1.png
│       └── yuanjing.png
├── 
├── uploads/                # 文件上传目录
├── 
└── ref/                   # 参考文档
    ├── ReferenceList.xlsx # 参考资料
    ├── generate_violation_items.py # 违规项生成工具
    ├── read_reference_list.py      # 参考清单读取工具
    ├── update_html.py              # HTML更新工具
    └── UPDATE_PROCESS.md           # 更新流程说明
```

## 主要组件说明

### 后端 (backend/)
- **app.py**: Flask web服务器，提供RESTful API接口
- **qwen_service.py**: 封装Qwen大模型调用逻辑，支持多模型路由
- **file_handler.py**: 处理文件上传、内容提取和格式转换
- **database.py**: SQLite数据库操作，管理文件记录和分析结果
- **run.py**: 应用启动入口
- **ai_agents/**: AI智能代理模块
  - **agent_manager.py**: 统一管理各种AI代理
  - **base_agent.py**: 所有代理的基础类
  - **project_info_agent.py**: 项目信息提取和匹配
  - **image_extraction_agent.py**: Word图片提取、上下文分析、AI命名及（可选）OCR
  - **authorization_letter_agent.py**: 授权委托书与证件多模态核验（项目编号/名称比对、人员与证件有效性）
  - **word_splitter.py**: Word文档智能拆分
  - **wordtoc_agent.py**: Word文档目录提取

### 变更说明
- 原 `word_image_separator.py` 已删除，相关功能由 `image_extraction_agent.py` 统一实现。
- 如果需要生成“图片占位符版本文档”，可在后续扩展于 `image_extraction_agent` 中添加可选导出（待实现）。

### 前端 (frontend/)
- **index.html**: 主页面，招标文件分析功能
- **bid_analysis.html**: 投标文件合规性检测页面

### 核心功能
- **招标文件分析**: 智能提取废标条款和关键要求
- **投标文件检测**: 合规性检查和风险评估
- **项目信息匹配**: 确保投标文件与招标文件项目信息一致
- **文档处理**: 支持PDF、Word等多种格式

### 配置文件
- **.env.template**: 环境变量模板，包含API密钥等配置
- **requirements.txt**: Python依赖包列表
- **bid_analysis.db**: SQLite数据库文件

### 启动脚本
- **start.bat**: Windows一键启动脚本
- **start.sh**: Linux/macOS一键启动脚本

### 工具脚本
- **examine_database.py**: 数据库内容检查工具
- **clear_files_table.py**: 数据库清理工具
- **analyze_doc_structure.py**: 文档结构分析工具
