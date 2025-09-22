# 投标文件合规性检查工具 (BidAn### 🚀 最新更新 (v1.4)
## 🌟 功能特点

- **🤖 智能文档分析**: 支持Word（.docx）、PDF（.pdf）等格式的深度解析
- **🏗️ AI Agent架构**: 模块化设计，包含项目信息提取、文档拆分、图像处理等专业代理
- **📋 多维度分析**: 废标条款识别、项目信息匹配、合规性检查、风险评估
- **🖥️ 双页面设计**: 招标文件分析和投标文件检测分离，专业化操作界面
- **🔄 文档处理工作流**: 完整的文档预处理流程，包含目录提取、文档拆分、图像分离
- **📊 智能进度追踪**: 实时显示文档处理进度，完成后自动跳转到分析页面
- **🧹 智能清理工具**: 支持选择性清理数据库表和临时文件，灵活的数据管理
- **📁 结构化存储**: 按文件ID组织的temp目录，包含目录markdown和拆分文档
- **💾 数据持久化**: SQLite数据库完整存储分析历史和文件记录
- **🔒 本地化部署**: 数据本地存储，安全可控，支持离线使用
- **⚡ 一键启动**: Windows/Linux启动脚本，自动环境配置
- **🎨 现代化界面**: ### 🗺️ 版本规划

### 当前版本 (v1.4)
- ✅ 完整的文档处理工作流程
- ✅ 智能进度追踪和自动跳转
- ✅ 文档处理协调器架构
- ✅ 临时文件管理和清理工具
- ✅ 类型安全代码优化
- ✅ 增强的用户体验设计

### 未来版本
- 🔄 v1.5: 批量文件处理和任务队列
- 🔄 v1.6: 用户认证和权限管理
- 🔄 v1.7: Docker容器化部署
- 🔄 v1.8: 更多文件格式支持
- 🔄 v2.0: 微服务架构和云原生部署主题切换，响应式布局档处理工作流** - 集成目录提取、文档拆分、图像分离的完整文档预处理流程
- 📊 **智能进度追踪** - 文档处理完成后自动跳转到分析页面，提供实时进度反馈
- 🛠️ **文档处理协调器** - 新增`DocumentProcessor`类统一协调多个AI代理的工作流程  
- 📁 **临时文件管理** - 完善的temp目录结构，支持文档拆分结果的组织化存储
- 🧹 **智能清理工具** - 全新的数据库和文件清理工具，支持选择性清理不同类型的数据
- 🔧 **类型安全优化** - 修复所有类型注解错误，提升代码质量和IDE支持
- 📈 **增强的用户体验** - 文档处理过程的可视化进度条和详细状态提示
- 🗂️ **结构化temp目录** - 按文件ID组织的临时文件夹，包含目录markdown和拆分文档)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.12 (Windows COM)](https://img.shields.io/badge/python-3.12%20preferred-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-2.3.3-green.svg)](https://flask.palletsprojects.com/)
[![GitHub release](https://img.shields.io/github/release/biibibi/BidAnalysisTool.svg)](https://github.com/biibibi/BidAnalysisTool/releases)
[![GitHub stars](https://img.shields.io/github/stars/biibibi/BidAnalysisTool.svg)](https://github.com/biibibi/BidAnalysisTool/stargazers)

## 📖 项目介绍

**BidAnalysisTool** 是一个基于Qwen大模型的智能投标文件合规性检查工具，采用先进的AI Agent架构设计，专为招投标行业打造的智能化解决方案。

### 🎯 核心能力

1. **智能分析招标文件** - 自动提取废标条款、评分标准和关键要求
2. **合规性检查** - 验证投标文件是否符合招标要求，识别潜在的废标风险  
3. **项目信息提取** - 智能识别项目基本信息、技术要求和商务条件
4. **风险评估** - 提供详细的合规性评估报告和改进建议
5. **双文件预览** - 支持招标文件和投标文件同时预览对比

### 🚀 最新更新 (v1.3)

- ✨ **完善的AI代理架构** - 智能文档拆分、图像分离、目录提取等高级功能
- 📊 **双页面设计** - 招标文件分析页面和投标文件检测页面分离，专业化操作
- � **多模型支持** - 支持Qwen和豆包(字节跳动)双模型路由，提升分析效果
- 📁 **智能文档处理** - Word文档自动拆分、图像重命名、目录智能提取
- �️ **数据库优化** - 完善的SQLite数据管理，支持分析历史记录和结果查询
- 🎨 **界面美化** - 现代化Bootstrap 5设计，支持多主题切换
- � **项目信息匹配** - 智能检测投标文件与招标文件项目信息一致性
- � **结构化存储** - 规范化项目结构，排除测试文件，提升代码质量

## 🌟 功能特点

- **🤖 智能文档分析**: 支持Word（.docx）、PDF（.pdf）等格式的深度解析
- **🏗️ AI Agent架构**: 模块化设计，包含项目信息提取、文档拆分、图像处理等专业代理
- **📋 多维度分析**: 废标条款识别、项目信息匹配、合规性检查、风险评估
- **🖥️ 双页面设计**: 招标文件分析和投标文件检测分离，专业化操作界面
- **� 智能匹配**: 自动检测投标文件与招标文件的项目信息一致性
- **📊 详细分析报告**: 结构化分析结果，包含废标条款、建议和风险评级
- **💾 数据持久化**: SQLite数据库完整存储分析历史和文件记录
- **🔒 本地化部署**: 数据本地存储，安全可控，支持离线使用
- **⚡ 一键启动**: Windows/Linux启动脚本，自动环境配置
- **🎨 现代化界面**: Bootstrap 5设计，支持多主题切换，响应式布局

## 🏗️ 系统架构

```
┌─────────────────────────────────────────────────────────────┐
│                    Frontend (Web UI)                        │
│  ┌─────────────────┐  ┌─────────────────────────────────┐   │
│  │   File Upload   │  │      Dual Preview Area         │   │
│  │   & Management  │  │   (2x1 Grid Layout)           │   │
│  └─────────────────┘  └─────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                           │ HTTP API
┌─────────────────────────────────────────────────────────────┐
│                    Backend (Flask)                          │
│  ┌─────────────────┐  ┌─────────────────────────────────┐   │
│  │  File Handler   │  │       AI Agent Manager         │   │
│  │ (PDF/Word)      │  │  ┌─────────────────────────┐   │   │
│  └─────────────────┘  │  │ Project Info Agent      │   │   │
│                       │  │ Tender Analysis Agent   │   │   │
│  ┌─────────────────┐  │  │ Compliance Check Agent  │   │   │
│  │  Qwen Service   │  │  └─────────────────────────┘   │   │
│  │ (AI Analysis)   │  └─────────────────────────────────┘   │
│  └─────────────────┘                                        │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐   │
│  │          Database (SQLite)                          │   │
│  │  • Files Metadata  • Analysis Results             │   │
│  │  • Project Info    • Risk Assessment              │   │
│  └─────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### 🔧 核心组件

- **前端界面层**: 现代化Web界面，支持文件上传和双文件预览
- **API接口层**: RESTful API，提供标准化的服务接口
- **业务逻辑层**: 文件处理、分析调度、结果管理
- **AI服务层**: Qwen大模型集成，智能文档分析
- **数据存储层**: SQLite数据库，存储文件信息和分析结果

## 📸 界面预览

![主界面](static/screenshot.png)

## 🚀 快速开始

### 方式一：一键启动（推荐）

#### Windows用户
```bash
# 双击运行或在命令行执行
start.bat
```

#### Linux/macOS用户
```bash
chmod +x start.sh
./start.sh
```

启动脚本会自动：
1. 使用 Python 3.12 创建虚拟环境 `.venv312`（若未安装将提示安装）
2. 安装所需依赖并确保 `pywin32`（COM 支持）就绪
3. 检查环境配置并从模板生成 `.env`
4. 启动后端服务（确保 Word 相关 Agent 可用）

### 方式二：手动安装

#### 1. 克隆项目
```bash
git clone https://github.com/biibibi/BidAnalysisTool.git
cd BidAnalysisTool
```

#### 2. 配置Python环境

Windows 强烈推荐 Python 3.12（启用 Word COM 自动化）：

```bash
# 创建 Python 3.12 虚拟环境（Windows）
py -3.12 -m venv .venv312

# 激活虚拟环境（Windows）
.venv312\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
```

#### 3. 安装依赖
```bash
pip install -r backend/requirements.txt
# Windows 上启用 COM 支持（Word 自动化所需）
pip install pywin32==306
```

#### 4. 配置API密钥

创建环境配置文件：
```bash
cp backend/.env.template backend/.env
```

编辑 `backend/.env` 文件，配置您选择的AI模型服务（可选其一或同时配置）：

```env
# 全局模型提供方选择 (qwen 或 doubao)
LLM_PROVIDER=qwen

# Qwen(阿里云百炼) 配置
DASHSCOPE_API_KEY=your_qwen_api_key

# 豆包(字节跳动Ark) 配置  
ARK_API_KEY=your_ark_api_key
DOUBAO_MODEL_ID=ep-20241215xxx  # 您的推理接入点ID
ARK_BASE_URL=https://ark.cn-beijing.volces.com/api/v3
```

**获取API密钥：**

**Qwen(阿里云百炼)：**
1. 访问 [阿里云百炼控制台](https://dashscope.console.aliyun.com/)
2. 注册并登录阿里云账号
3. 在API密钥管理页面创建新的API密钥
4. 复制密钥并设置到 `DASHSCOPE_API_KEY`

**豆包(字节跳动Ark)：**
1. 访问 [火山引擎Ark控制台](https://console.volcengine.com/ark/)
2. 创建API Key和推理接入点
3. 将API Key设置到 `ARK_API_KEY`
4. 将推理接入点ID设置到 `DOUBAO_MODEL_ID`

#### 5. 启动服务
```bash
cd backend
python run.py
```

#### 6. 访问应用

后端服务启动后，在浏览器中打开：
- 主界面：`frontend/index.html` 或 `frontend/bid_analysis.html`
- API文档：`http://localhost:5000/api/health`

## 📖 使用说明

### 基本操作流程

1. **📁 文件上传**
   - 支持拖拽上传或点击选择文件
   - 支持格式：PDF (.pdf)、Word (.docx/.doc)、文本 (.txt)
   - 单独上传招标文件或投标文件
   - 文件大小限制：50MB

2. **👁️ 文件预览**
   - **单文件模式**：上传一个文件后，显示对应的预览按钮
   - **双文件模式**：同时上传招标和投标文件，显示2x1网格预览布局
   - 点击预览按钮在新窗口查看文件内容
   - 支持在线预览和文件下载

3. **📊 招标文件分析**
   - 上传招标文件后点击"招标文件分析"
   - AI智能提取：
     - 项目基本信息
     - 废标条款和要求
     - 评分标准
     - 技术规格要求
     - 商务条件

4. **✅ 投标文件检查**
   - 同时上传招标文件和投标文件
   - 点击"投标文件分析"进行文档处理和合规性检查
   - 系统会先进行文档预处理：
     - 提取文档目录结构
     - 按章节拆分文档
     - 生成结构化的markdown目录文件
   - 处理完成后自动跳转到分析页面
   - 生成详细的检查报告和风险评估
   - 提供改进建议

### 界面功能

- **🎨 主题切换**：支持5种主题（默认蓝色、高贵紫色、自然绿色、活力橙色、深色模式）
- **📱 响应式设计**：自适应桌面和移动设备
- **💡 智能提示**：完整的使用指南和帮助信息
- **📈 进度显示**：实时显示分析进度

## 🔧 API接口文档

### 文件上传
```http
POST /api/upload
Content-Type: multipart/form-data

参数：
- file: 上传的文件（.pdf或.docx）

响应：
{
    "success": true,
    "file_id": "uuid",
    "filename": "原文件名",
    "message": "文件上传成功"
}
```

### 招标文件分析
```http
POST /api/analyze/tender
Content-Type: application/json

{
    "file_id": "uuid"
}

响应：
{
    "success": true,
    "analysis_id": "uuid",
    "message": "分析任务已提交"
}
```

### 投标文件分析
```http
POST /api/analyze/bid
Content-Type: application/json

{
    "file_id": "uuid",
    "tender_analysis_id": "uuid"  // 可选，基于特定招标文件分析结果
}
```

### 获取分析结果
```http
GET /api/analysis/{analysis_id}

响应：
{
    "success": true,
    "status": "completed",
    "result": {
        "project_info": {...},
        "compliance_check": {...},
        "risk_assessment": {...}
    }
}
```

### 文档处理工作流
```http
POST /api/process-bid-document
Content-Type: application/json

{
    "file_id": "uuid"
}

响应：
{
    "success": true,
    "work_dir": "/temp/file_id/",
    "toc_result": {
        "success": true,
        "toc_count": 12,
        "md_path": "/temp/file_id/file_id_目录.md"
    },
    "split_result": {
        "success": true,
        "split_count": 6,
        "split_dir": "/temp/file_id/split_documents/"
    }
}
```

### 处理状态查询
```http
GET /api/process-bid-document/status/{file_id}

响应：
{
    "exists": true,
    "has_toc": true,
    "has_splits": true,
    "split_count": 6,
    "split_files": ["01_营业执照.docx", "02_投标函.docx", ...]
}
```

### 项目信息提取
```http
POST /api/extract-project-info
Content-Type: application/json

{
    "file_id": "uuid",
    "document_type": "tender"  // 或 "bid" 或 "auto"
}
```

### 健康检查
```http
GET /api/health

响应：
{
    "status": "healthy",
    "timestamp": "2025-01-XX XX:XX:XX"
}
```

## 📁 项目结构

```
BidAnalysisTool/
├── backend/                    # 后端服务
│   ├── app.py                 # Flask应用主文件
│   ├── run.py                 # 启动脚本
│   ├── database.py            # 数据库管理
│   ├── file_handler.py        # 文件处理服务
│   ├── qwen_service.py        # 多模型AI服务
│   ├── requirements.txt       # Python依赖
│   ├── .env.template         # 环境配置模板
│   └── ai_agents/            # AI代理模块
│       ├── __init__.py
│       ├── agent_manager.py   # 代理管理器
│       ├── base_agent.py      # 基础代理类
│       ├── document_processor.py # 文档处理协调器
│       ├── project_info_agent.py # 项目信息提取代理
│       ├── word_splitter.py   # Word文档拆分代理
│       ├── word_image_separator.py # 图像分离代理
│       └── wordtoc_agent.py   # 目录提取代理
├── frontend/                  # 前端界面
│   ├── index.html            # 招标文件分析页面
│   └── bid_analysis.html     # 投标文件检测页面
├── static/                   # 静态资源
│   └── icon/                # 图标文件
├── uploads/                  # 文件上传存储
├── temp/                     # 临时文件处理目录
├── ref/                     # 参考文档和工具
├── bid_analysis.db          # SQLite数据库文件
├── analyze_doc_structure.py # 文档结构分析工具
├── examine_database.py      # 数据库检查工具
├── clear_files_table.py     # 数据库清理工具
├── 课题方案书.md              # 项目方案文档
├── start.bat               # Windows启动脚本
├── start.sh                # Linux/macOS启动脚本
├── README.md               # 项目说明
├── LICENSE                 # 许可证
└── STRUCTURE.md            # 项目结构详细说明
```

## 🔧 技术栈

### 后端技术
- **Python 3.12 (Windows 推荐)**: COM 自动化稳定；其他平台 3.8+ 兼容
- **Flask 2.3.3**: 轻量级Web框架，RESTful API设计
- **Flask-CORS**: 跨域请求支持，前后端分离
- **SQLite**: 轻量级关系数据库，85+文件记录管理
- **python-docx**: Word文档解析处理
- **pywin32 (Windows)**: Word COM 自动化（目录定位、格式复制、拆分）
- **PyPDF2**: PDF文档内容提取
- **UUID**: 安全的文件命名和标识系统
- **python-dotenv**: 环境变量和配置管理

### 前端技术
- **HTML5/CSS3/JavaScript**: 基础Web技术，ES6+语法
- **Bootstrap 5**: 响应式UI框架，现代化设计组件
- **Grid Layout**: CSS Grid实现2x1双文件预览布局
- **拖拽上传**: 用户友好的文件上传体验
- **主题系统**: 5种可切换界面主题（蓝/紫/绿/橙/深色）
- **文件预览**: 在线预览功能，支持PDF和Word文档

### AI服务
- **多模型支持**: 支持Qwen(阿里云百炼)和豆包(字节跳动)双模型路由
- **AI Agent架构**: 模块化智能代理系统设计
  - `DocumentProcessor`: 文档处理协调器，统一管理处理工作流程
  - `ProjectInfoAgent`: 项目信息提取和匹配专用代理
  - `WordSplitterAgent`: Word文档智能拆分代理
  - `WordImageSeparatorAgent`: 图像分离和重命名代理
  - `WordTOCAgent`: 目录提取和结构化代理
  - `BaseAgent`: 通用代理基础类，支持扩展
  - `AgentManager`: 代理统一管理和调度器
- **智能分析引擎**: 文档理解、合规检查、风险评估、项目匹配

## ⚠️ 注意事项

1. **API密钥安全**: 
   - 请妥善保管您的API密钥，不要提交到版本控制系统
   - 使用`.env`文件存储敏感信息
   - 定期更换API密钥

2. **文件限制**: 
   - 支持格式：.pdf和.docx
   - 文件大小：建议不超过50MB
   - 文件内容：确保文字清晰可读

3. **网络要求**: 
   - 需要稳定的网络连接访问阿里云百炼API
   - 建议在企业网络环境下使用

4. **性能与兼容性**:
   - 大文件分析时间较长，请耐心等待
   - 建议在性能较好的服务器上部署
 - Word 相关 Agent（目录定位、拆分、图片提取重命名）在 Windows + Office + Python 3.12 + pywin32 下表现最佳
 - 非 Windows 环境将自动回退到 `python-docx` 路径（不含精确位置）

## 🐛 常见问题

### 环境配置问题

**Q: 提示"DASHSCOPE_API_KEY未设置"**  
A: 检查 `backend/.env` 文件是否存在并正确填入API密钥

**Q: 依赖安装失败**  
A: 
- 确保Python版本≥3.8
- 升级pip：`pip install --upgrade pip`
- 使用国内镜像：`pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt`

**Q: 虚拟环境激活失败**  
A: 
- Windows: 确保执行策略允许脚本运行
- 使用管理员权限运行：`Set-ExecutionPolicy RemoteSigned`

### 功能使用问题

**Q: 分析结果不准确**  
A: 
- 确保上传文件格式正确且内容清晰
- 大模型分析结果仅供参考，建议人工复核
- 复杂文档可能需要多次分析

**Q: 文件上传失败**  
A: 
- 检查文件格式（仅支持.pdf和.docx）
- 检查文件大小（建议<50MB）
- 确保文件未损坏

**Q: 前端页面无法访问**  
A: 
- 确认后端服务已正常启动
- 检查防火墙设置
- 尝试使用不同浏览器

### 部署相关问题

**Q: 如何在服务器上部署？**  
A: 
1. 克隆代码到服务器
2. 配置Python环境和依赖
3. 设置API密钥
4. 配置防火墙开放5000端口
5. 使用进程管理工具（如supervisor）管理服务

**Q: 支持Docker部署吗？**  
A: 当前版本暂不提供Docker镜像，可以根据需要自行构建

## 🔒 安全说明

1. **数据隐私**: 
   - 所有上传文件存储在本地
   - 不会泄露给第三方
   - 建议定期清理临时文件

2. **API安全**: 
   - API密钥加密存储
   - 建议使用HTTPS部署生产环境
   - 定期更新依赖包

3. **访问控制**: 
   - 建议在内网环境使用
   - 可配置IP白名单
   - 添加用户认证机制

## 🤝 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. Fork本项目
2. 创建功能分支：`git checkout -b feature/your-feature`
3. 提交更改：`git commit -am 'Add some feature'`
4. 推送分支：`git push origin feature/your-feature`
5. 提交Pull Request

### 开发规范
- 遵循PEP 8代码风格
- 添加适当的注释和文档
- 编写单元测试
- 更新相关文档

## 🗺️ 版本规划

### 当前版本 (v1.3)
- ✅ 完整的AI Agent架构
- ✅ 双页面专业化设计
- ✅ 多模型支持(Qwen/豆包)
- ✅ 智能文档处理能力
- ✅ 数据库完整管理
- ✅ 项目信息智能匹配
- ✅ Word文档高级处理

### 未来版本
- 🔄 v1.4: 批量文件处理和任务队列
- 🔄 v1.5: 用户认证和权限管理
- 🔄 v1.6: Docker容器化部署
- 🔄 v1.7: 更多文件格式支持
- 🔄 v2.0: 微服务架构和云原生部署

## 📞 技术支持

### 问题反馈
- GitHub Issues: [提交问题](https://github.com/biibibi/BidAnalysisTool/issues)
- 邮箱支持: support@bidanalysis.com

### 开发交流
- 技术讨论: [Discussions](https://github.com/biibibi/BidAnalysisTool/discussions)
- 开发文档: [Wiki](https://github.com/biibibi/BidAnalysisTool/wiki)

### 商业支持
如需定制开发或商业支持，请联系我们。

## 免责声明

本工具仅供辅助分析使用，分析结果仅供参考。实际投标过程中，请务必：
- 仔细阅读原始招标文件
- 人工核验所有关键条款
- 咨询专业人员意见
- 承担相应的投标风险

使用本工具产生的任何后果，开发团队不承担责任。

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

---

## 🌟 Star History

如果这个项目对您有帮助，请给我们一个Star⭐！

[![Star History Chart](https://api.star-history.com/svg?repos=biibibi/BidAnalysisTool&type=Date)](https://star-history.com/#biibibi/BidAnalysisTool&Date)

---

<div align="center">
  <p>Made with ❤️ by BidAnalysis Team</p>
  <p>© 2025 BidAnalysisTool. All rights reserved.</p>
</div>


