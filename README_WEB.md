# 📄 公文格式调整工具 - Web版

[![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-2.3%2B-green)](https://flask.palletsprojects.com/)

按照 **GB/T 9704-2012《党政机关公文格式》** 标准，自动调整 Word 文档格式的 Web 应用。

---

## 🚀 快速开始

### 1️⃣ 安装依赖

```bash
pip3 install -r requirements_web.txt
```

### 2️⃣ 启动服务

#### 方式一：双击运行（推荐）

**macOS:**
```bash
# 首次使用赋予权限
chmod +x 启动Web服务.command

# 双击运行
双击 "启动Web服务.command" 文件
```

**Windows:**
```bash
# 双击运行
双击 "启动Web服务.bat" 文件
```

#### 方式二：命令行运行

```bash
python3 app.py
```

### 3️⃣ 使用

1. 启动后会自动提示访问地址（默认：http://localhost:5000）
2. 在浏览器中打开该地址
3. 点击或拖拽 Word 文档到上传区域
4. 点击"开始处理"按钮
5. 等待处理完成后点击"下载"按钮

---

## ✨ 界面特点

### 🎨 现代化设计
- 渐变色背景，视觉效果优雅
- 响应式布局，支持移动端
- 流畅的动画效果
- 拖拽上传，操作便捷

### 💡 用户友好
- 实时文件信息显示
- 进度条反馈
- 清晰的成功/错误提示
- 一键下载处理结果

### 🔒 安全可靠
- 文件大小限制（50MB）
- 文件类型验证
- 临时文件自动清理
- 错误处理完善

---

## 📁 文件结构

```
codebuddy_gongwen/
├── app.py                    # Flask后端主程序
├── gongwen_formatter_cli.py  # 核心格式化逻辑
├── templates/
│   └── index.html           # Web前端界面
├── requirements_web.txt      # Web版依赖
├── 启动Web服务.command        # macOS启动脚本
├── 启动Web服务.bat            # Windows启动脚本
└── README_WEB.md            # Web版文档（本文件）
```

---

## 🔧 技术栈

### 后端
- **Flask**: 轻量级Web框架
- **python-docx**: Word文档处理

### 前端
- **HTML5**: 语义化结构
- **CSS3**: 现代化样式、动画
- **JavaScript**: 交互逻辑、文件上传

---

## 📋 功能特性

### ✅ 自动识别与格式化
- 主标题（方正小标宋22号，居中）
- 主送机关（仿宋16号，顶格）
- 多级标题（一、二、三、/ （一）（二）/ 1. 2. 3. / (1) (2) (3)）
- 正文、署名、日期
- 表格、图片（自动识别并跳过/居中）

### 🔧 智能修正
- 自动修正标题层级跳跃
- 自动修正编号不连续
- 清除多余标点符号
- 删除标题上方多余空行
- 统一页边距和缩进

### 📎 附件处理
- 附件独立编号
- 附件列表格式规范化
- 附件前自动分页

---

## 🌐 API 接口

### POST /upload
上传并处理Word文档

**请求:**
- Content-Type: `multipart/form-data`
- Body: `file` (Word文档)

**响应:**
```json
{
    "success": true,
    "download_url": "/download/xxx.docx",
    "filename": "done_xxx.docx"
}
```

### GET /download/<filename>
下载处理后的文档

**响应:**
- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- Body: Word文档二进制数据

---

## ⚙️ 配置说明

### app.py 配置项

```python
# 最大文件大小（字节）
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# 临时文件目录
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'docx'}
```

### 修改端口

在 `app.py` 最后一行修改：

```python
app.run(debug=True, host='0.0.0.0', port=5000)  # 改为其他端口
```

---

## 🔍 常见问题

### 1. 端口被占用

**问题:** 启动时提示 `Address already in use`

**解决:**
- 方法1: 关闭占用5000端口的程序
- 方法2: 修改 `app.py` 中的端口号

### 2. 依赖安装失败

**问题:** `pip install` 报错

**解决:**
```bash
# 使用国内镜像源
pip3 install -r requirements_web.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 3. 文件上传失败

**问题:** 上传时提示错误

**解决:**
- 检查文件格式是否为 `.docx`
- 检查文件大小是否超过50MB
- 检查文件是否已损坏

### 4. 处理后样式不对

**问题:** 格式调整不符合预期

**解决:**
- 确保系统已安装所需字体（方正小标宋简体、楷体_GB2312、仿宋_GB2312）
- 检查原文档是否使用了表格或特殊格式
- 查看终端输出的详细日志

---

## 🎯 使用建议

1. **网络访问**: 如需局域网内其他设备访问，确保防火墙允许5000端口
2. **生产部署**: 建议使用 `gunicorn` 或 `uwsgi` 部署到生产环境
3. **安全性**: 如部署到公网，建议添加用户认证和文件扫描
4. **性能**: 大文件处理可能需要较长时间，建议添加后台任务队列

---

## 🚀 生产部署

### 使用 Gunicorn

```bash
# 安装
pip install gunicorn

# 启动
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### 使用 Nginx 反向代理

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

## 📝 更新日志

### v2.0.0 (2025-01-16)
- ✅ 新增 Web 界面
- ✅ 支持拖拽上传
- ✅ 实时进度显示
- ✅ 响应式设计
- ✅ 自动清理临时文件

---

## 📞 技术支持

如有问题或建议，欢迎反馈。

---

**🎉 感谢使用公文格式调整工具 - Web版！**
