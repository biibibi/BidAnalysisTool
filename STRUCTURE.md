# 项目结构

```
BidAnalysis/
├── README.md                   # 项目说明文档
├── LICENSE                     # MIT许可证
├── .gitignore                 # Git忽略文件配置
├── start.bat                  # Windows启动脚本
├── start.sh                   # Linux/macOS启动脚本
├── 使用指南.md                  # 详细使用指南
├── test_system.py             # 系统测试脚本
├── test_fixed_api.py          # API测试脚本
├── test_api.py                # API连接测试
├── 
├── backend/                   # 后端代码
│   ├── app.py                # Flask主应用
│   ├── qwen_service.py       # Qwen大模型服务
│   ├── file_handler.py       # 文件处理服务
│   ├── database.py           # 数据库管理
│   ├── run.py               # 启动脚本
│   ├── requirements.txt     # Python依赖
│   ├── .env.template       # 环境变量模板
│   └── .env               # 环境变量配置（需要配置）
├── 
├── frontend/                 # 前端代码
│   └── index.html          # Web界面
├── 
├── static/                  # 静态资源
│   └── icon/              # 图标文件
│       ├── Unicom.jpeg
│       ├── Unicom1.png
│       └── yuanjing.png
├── 
├── test/                   # 测试文件
│   ├── Qwen.py           # Qwen API示例
│   └── test_tender.txt   # 测试文档
└── 
└── ref/                   # 参考文档
    └── ReferenceList.xlsx # 参考资料
```

## 主要组件说明

### 后端 (backend/)
- **app.py**: Flask web服务器，提供API接口
- **qwen_service.py**: 封装Qwen大模型调用逻辑
- **file_handler.py**: 处理文件上传和内容提取
- **database.py**: SQLite数据库操作
- **run.py**: 应用启动入口

### 前端 (frontend/)
- **index.html**: 单页面应用，包含所有UI和交互逻辑

### 配置文件
- **.env.template**: 环境变量模板，包含所有必要的配置项
- **requirements.txt**: Python依赖包列表

### 启动脚本
- **start.bat**: Windows一键启动脚本
- **start.sh**: Linux/macOS一键启动脚本
