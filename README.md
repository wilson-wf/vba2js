# VBA到WPS JS宏转换器

一个完整的项目，支持将带有VBA宏的Excel文件（.xlsm）自动转换为带有WPS JS宏的版本（.xlsx）。

## 功能特性

- 📁 **文件上传**: 支持上传带有VBA宏的Excel文件
- 🔄 **自动转换**: 自动将VBA代码转换为WPS JS宏
- 🤖 **AI增强**: 支持配置大模型API增强转换能力
- 📊 **进度展示**: 实时显示转换进度和阶段
- ✅ **代码验证**: 自动验证转换后的JS代码语法
- 📱 **Web界面**: 提供简易的Web用户界面
- 🚀 **RESTful API**: 提供完整的API接口

## 技术架构

```
┌─────────────────────────────────────────────────────────┐
│                    前端界面 (Vue.js)                    │
├─────────────────────────────────────────────────────────┤
│                    API网关 (Flask)                     │
├─────────────────────────────────────────────────────────┤
│  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐   │
│  │  文件服务    │ │  转换服务    │ │  测试服务    │   │
│  └──────────────┘ └──────────────┘ └──────────────┘   │
├─────────────────────────────────────────────────────────┤
│  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐   │
│  │  VBA解析器   │ │  语法映射器  │ │  LLM代理     │   │
│  └──────────────┘ └──────────────┘ └──────────────┘   │
├─────────────────────────────────────────────────────────┤
│                  数据存储 (SQLite)                     │
└─────────────────────────────────────────────────────────┘
```

## 快速开始

### 环境要求

- Python 3.10+
- Node.js 18+
- Linux (Ubuntu 20.04+)

### 安装依赖

```bash
# 安装Python依赖
pip install -r requirements.txt

# 安装前端依赖
cd frontend
npm install
npm run build
```

### 运行开发服务器

```bash
# 设置环境变量
export FLASK_APP=run.py
export FLASK_ENV=development

# 运行服务
python run.py
```

服务将在 http://localhost:5000 启动。

### 使用部署脚本

```bash
chmod +x deploy.sh
./deploy.sh
```

## API接口

### 文件上传

```
POST /api/upload
Content-Type: multipart/form-data

参数:
- file: xlsm文件

返回:
{ "file_id": "uuid", "status": "pending" }
```

### 触发转换

```
POST /api/convert/{file_id}
Content-Type: application/json

参数:
{ "use_llm": true }

返回:
{ "file_id": "uuid", "status": "completed", "download_url": "/api/download/{file_id}" }
```

### 查询状态

```
GET /api/convert/status/{file_id}

返回:
{ 
  "file_id": "uuid", 
  "status": "processing", 
  "progress": 50, 
  "stage": "converting" 
}
```

### 下载结果

```
GET /api/download/{file_id}

返回: 文件流 (xlsx)
```

### LLM配置

```
POST /api/config/llm
Content-Type: application/json

参数:
{ 
  "api_key": "your-api-key", 
  "endpoint": "https://api.openai.com/v1/chat/completions", 
  "model": "gpt-4" 
}

返回:
{ "success": true, "message": "LLM configuration updated" }
```

## VBA到JS宏语法映射示例

| VBA语法 | WPS JS宏语法 |
|---------|-------------|
| `Sub Hello()` | `function Hello() {}` |
| `Function Sum(a, b)` | `function Sum(a, b) {}` |
| `Dim x As Integer` | `let x = 0` |
| `MsgBox "Hello"` | `Application.MsgBox("Hello")` |
| `Range("A1").Value` | `Range("A1").Value` |
| `If x > 0 Then...End If` | `if (x > 0) {...}` |
| `For i = 1 To 10...Next` | `for (let i = 1; i <= 10; i++) {...}` |

## 项目结构

```
vba-converter/
├── app/                    # 后端应用代码
│   ├── __init__.py         # Flask应用初始化
│   ├── api/                # API路由
│   │   ├── __init__.py
│   │   └── routes.py       # REST API定义
│   ├── engine/             # 转换引擎
│   │   ├── __init__.py
│   │   ├── vba_parser.py   # VBA代码解析器
│   │   ├── syntax_mapper.py # VBA到JS语法映射
│   │   ├── llm_agent.py    # 大模型代理
│   │   └── excel_generator.py # Excel文件生成
│   ├── services/           # 业务服务
│   │   ├── __init__.py
│   │   ├── file_service.py # 文件处理服务
│   │   ├── convert_service.py # 转换服务
│   │   └── test_service.py # 测试服务
│   └── utils/              # 工具函数
│       ├── __init__.py
│       ├── helpers.py      # 辅助函数
│       └── database.py     # 数据库管理
├── frontend/               # 前端代码
│   ├── src/
│   │   └── main.js         # Vue应用入口
│   ├── index.html          # HTML模板
│   ├── package.json        # 前端依赖
│   └── vite.config.js      # Vite配置
├── data/                   # 数据存储目录
│   ├── uploads/            # 上传文件
│   ├── converted/          # 转换结果
│   └── logs/               # 日志文件
├── config.py               # 配置文件
├── requirements.txt        # Python依赖
├── run.py                  # 启动脚本
├── deploy.sh               # 部署脚本
└── tech_doc.md             # 技术方案文档
```

## 配置说明

### 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `FLASK_APP` | Flask应用入口 | `run.py` |
| `FLASK_ENV` | 运行环境 | `development` |
| `SECRET_KEY` | 密钥 | `vba-converter-secret-key` |
| `LLM_API_KEY` | LLM API密钥 | 空 |
| `LLM_ENDPOINT` | LLM API端点 | `https://api.openai.com/v1/chat/completions` |
| `LLM_MODEL` | LLM模型名称 | `gpt-4` |

### 配置文件

```python
# config.py
class Config:
    UPLOAD_FOLDER = '/opt/vba-converter/data/uploads'
    CONVERTED_FOLDER = '/opt/vba-converter/data/converted'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
```

## 安全注意事项

1. **文件上传安全**: 仅允许上传 .xlsm 文件，限制文件大小
2. **API安全**: 建议配置API密钥认证（可选扩展）
3. **数据安全**: 文件存储在服务器本地，建议定期清理
4. **LLM密钥**: API密钥存储在数据库中，建议加密存储

## License

MIT License