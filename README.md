# 投标文件合规性检查工具 (BidAnalysisTool)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-2.3.3-green.svg)](https://flask.palletsprojects.com/)

这是一个基于Qwen大模型的智能投标文件合规性检查工具，采用AI Agent架构设计，可以帮助您：

1. **智能分析招标文件** - 自动提取废标条款、评分标准和关键要求
2. **合规性检查** - 验证投标文件是否符合招标要求，识别潜在的废标风险
3. **项目信息提取** - 智能识别项目基本信息、技术要求和商务条件
4. **风险评估** - 提供详细的合规性评估报告和改进建议

## 🌟 功能特点

- **智能文档分析**: 支持Word（.docx）和PDF文件格式的智能解析
- **AI Agent架构**: 采用模块化AI Agent设计，易于扩展和维护
- **多维度分析**: 项目信息提取、废标条款识别、合规性检查
- **友好的Web界面**: 直观的操作界面，支持拖放上传
- **详细分析报告**: 提供完整的合规性检查报告和风险评估
- **本地化部署**: 支持本地部署，数据安全可控
- **一键启动**: 提供Windows和Linux一键启动脚本

## 🏗️ 系统架构

```
Frontend (Web UI)
    ↓ HTTP API
Backend (Flask)
    ├── File Handler (文件处理)
    ├── AI Agent Manager (AI代理管理器)
    │   ├── Project Info Agent (项目信息提取)
    │   ├── Tender Analysis Agent (招标文件分析)
    │   └── Compliance Check Agent (合规性检查)
    ├── Qwen Service (通义千问服务)
    └── Database (SQLite数据库)
```

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
1. 创建Python虚拟环境
2. 安装所需依赖
3. 检查环境配置
4. 启动后端服务

### 方式二：手动安装

#### 1. 克隆项目
```bash
git clone https://github.com/biibibi/BidAnalysisTool.git
cd BidAnalysisTool
```

#### 2. 配置Python环境

推荐使用Python 3.8或更高版本：

```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
```

#### 3. 安装依赖
```bash
pip install -r backend/requirements.txt
```

#### 4. 配置API密钥

创建环境配置文件：
```bash
cp backend/.env.template backend/.env
```

编辑 `backend/.env` 文件，填入您的阿里云百炼API密钥：
```env
DASHSCOPE_API_KEY=your_actual_api_key_here
```

**获取API密钥：**
1. 访问 [阿里云百炼控制台](https://dashscope.console.aliyun.com/)
2. 注册并登录阿里云账号
3. 在API密钥管理页面创建新的API密钥
4. 复制密钥并粘贴到 `.env` 文件中

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

### 分析流程

1. **招标文件分析**
   - 上传招标文件（.pdf或.docx格式）
   - 点击"招标文件分析"按钮
   - 系统自动提取：
     - 项目基本信息
     - 废标条款
     - 评分标准
     - 技术要求
     - 商务条件

2. **投标文件检查**
   - 上传投标文件
   - 点击"投标文件分析"按钮
   - 基于招标文件要求进行合规性检查
   - 生成详细的检查报告和改进建议

### 主要界面

- `frontend/index.html` - 基础功能界面
- `frontend/bid_analysis.html` - 完整分析界面

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
│   ├── qwen_service.py        # Qwen大模型服务
│   ├── requirements.txt       # Python依赖
│   ├── .env.template         # 环境配置模板
│   ├── ai_agents/            # AI代理模块
│   │   ├── __init__.py
│   │   ├── agent_manager.py   # 代理管理器
│   │   ├── base_agent.py      # 基础代理类
│   │   └── project_info_agent.py # 项目信息提取代理
│   └── uploads/              # 上传文件存储
├── frontend/                  # 前端界面
│   ├── index.html            # 基础功能界面
│   └── bid_analysis.html     # 完整分析界面
├── test/                     # 测试文件
│   ├── Qwen.py              # Qwen API测试
│   ├── test_upload.py       # 上传功能测试
│   └── test_project_info_agent.py # 代理测试
├── static/                   # 静态资源
│   └── icon/                # 图标文件
├── ref/                     # 参考文档
├── uploads/                 # 全局上传目录
├── start.bat               # Windows启动脚本
├── start.sh                # Linux/macOS启动脚本
├── README.md               # 项目说明
├── LICENSE                 # 许可证
└── STRUCTURE.md            # 项目结构说明
```

## 🔧 技术栈

### 后端技术
- **Python 3.8+**: 主要编程语言
- **Flask 2.3.3**: Web框架
- **Flask-CORS**: 跨域请求支持
- **SQLite**: 轻量级数据库
- **python-docx**: Word文档处理
- **PyPDF2**: PDF文档处理
- **OpenAI**: AI模型接口
- **python-dotenv**: 环境变量管理

### 前端技术
- **HTML5/CSS3/JavaScript**: 基础Web技术
- **响应式设计**: 适配不同设备
- **拖放上传**: 用户友好的文件上传

### AI服务
- **阿里云百炼平台**: 提供Qwen大模型服务
- **Agent架构**: 模块化AI代理设计

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

4. **性能优化**:
   - 大文件分析时间较长，请耐心等待
   - 建议在性能较好的服务器上部署

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

### 当前版本 (v1.0)
- ✅ 基础文档分析功能
- ✅ AI Agent架构
- ✅ Web界面
- ✅ 一键启动脚本

### 未来版本
- 🔄 v1.1: 支持更多文件格式
- 🔄 v1.2: 批量文件处理
- 🔄 v1.3: 用户认证系统
- 🔄 v1.4: Docker容器化部署
- 🔄 v2.0: 微服务架构重构

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


