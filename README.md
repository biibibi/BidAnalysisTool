# 投标文件合规性检查工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-2.3.3-green.svg)](https://flask.palletsprojects.com/)

这是一个基于Qwen大模型的投标文件合规性检查工具，可以帮助您：

1. **分析招标文件** - 自动提取废标条款和关键要求
2. **检查投标文件** - 验证投标文件是否符合招标要求，识别潜在的废标风险

## 🌟 功能特点

- 支持Word（.docx）和PDF文件格式
- 基于阿里云百炼平台的Qwen大模型进行智能分析
- 直观的Web界面，支持拖放上传
- 详细的合规性检查报告
- 风险等级评估和改进建议
- 本地部署，数据安全可控

## 📸 界面预览

![主界面](static/screenshot.png)

## 🚀 快速开始

### 1. 克隆项目或下载源码

```bash
cd BidAnalysis
```

### 2. 配置Python环境

推荐使用虚拟环境：

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate
```

### 3. 安装依赖

```bash
cd backend
pip install -r requirements.txt
```

### 4. 配置API密钥

复制环境配置模板：

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

### 5. 启动后端服务

```bash
cd backend
python run.py
```

看到以下输出表示启动成功：
```
启动投标文件合规性检查工具后端服务...
服务地址: http://0.0.0.0:5000
API文档: http://0.0.0.0:5000/api/health
 * Running on all addresses (0.0.0.0)
 * Running on http://127.0.0.1:5000
```

### 6. 打开前端界面

在浏览器中打开：`frontend/index.html`

或者直接双击 `frontend/index.html` 文件。

## 使用说明

### 第一步：分析招标文件

1. 在Web界面中上传招标文件（支持.pdf、.docx格式）
2. 点击"招标文件分析"按钮
3. 等待分析完成，查看提取的废标条款和要求

### 第二步：检查投标文件

1. 上传投标文件
2. 点击"投标文件分析"按钮
3. 系统会基于之前分析的招标文件要求，检查投标文件的合规性
4. 查看详细的检查报告和改进建议

## API接口

后端提供以下API接口：

- `POST /api/upload` - 文件上传
- `POST /api/analyze/tender` - 招标文件分析
- `POST /api/analyze/bid` - 投标文件分析
- `GET /api/analysis/{id}` - 获取分析结果
- `GET /api/health` - 健康检查

## 文件结构

```
BidAnalysis/
├── backend/                 # 后端代码
│   ├── app.py              # Flask应用主文件
│   ├── qwen_service.py     # Qwen大模型服务
│   ├── file_handler.py     # 文件处理服务
│   ├── database.py         # 数据库管理
│   ├── requirements.txt    # Python依赖
│   ├── .env               # 环境配置（需要配置API密钥）
│   └── run.py             # 启动脚本
├── frontend/               # 前端代码
│   └── index.html         # Web界面
├── test/                  # 测试文件
│   └── Qwen.py           # Qwen API测试示例
└── README.md             # 使用说明
```

## 注意事项

1. **API密钥安全**: 请妥善保管您的API密钥，不要提交到版本控制系统
2. **文件大小限制**: 目前支持最大50MB的文件上传
3. **支持格式**: 仅支持.pdf和.docx格式的文件
4. **网络要求**: 需要访问阿里云百炼API服务

## 常见问题

### Q: 提示"DASHSCOPE_API_KEY未设置"怎么办？
A: 请检查 `backend/.env` 文件中是否正确填入了API密钥。

### Q: 分析结果不准确怎么办？
A: 这是基于大模型的分析，结果仅供参考。建议：
- 确保上传的文件内容清晰完整
- 人工复核关键条款
- 根据具体项目需求调整

### Q: 支持其他文件格式吗？
A: 目前仅支持PDF和Word格式。如需支持其他格式，可以先转换为支持的格式。

### Q: 能否离线使用？
A: 本工具依赖阿里云百炼API，需要网络连接。不支持完全离线使用。

## 技术支持

如有问题或建议，请检查：
1. 依赖是否正确安装
2. API密钥是否有效
3. 网络连接是否正常
4. 文件格式是否支持

## 免责声明

本工具仅供辅助分析使用，分析结果仅供参考。实际投标过程中，请务必：
- 仔细阅读原始招标文件
- 人工核验所有关键条款
- 咨询专业人员意见
- 承担相应的投标风险

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 🤝 贡献

欢迎提交问题和拉取请求。对于重大更改，请先打开一个issue来讨论您想要更改的内容。

## 📞 支持

如果您觉得这个项目有用，请给它一个 ⭐️！

如有问题或建议，请[创建一个issue](https://github.com/yourusername/BidAnalysis/issues)。
