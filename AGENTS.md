# 30x-slides - PPTX 转交互式讲解页
Flask + LibreOffice + PyMuPDF + python-pptx + Anthropic + Fish Audio

<directory>
services/ - provider 配置、deck 状态、对话与 TTS 服务层（5 文件: config.py, decks.py, chat.py, tts.py, __init__.py）
templates/ - 页面模板与前端交互脚本（2 文件: index.html, viewer.html）
output/ - 转换产物与 deck 状态缓存（每个 deck 一个目录）
</directory>

<config>
app.py - HTTP 入口，负责上传、转换、页面渲染与服务层装配
requirements.txt - Python 运行依赖
nixpacks.toml - Railway 构建与启动配置
</config>

法则: app.py 只做入口装配；provider 配置单一真相源；deck 上下文不能只依赖容器本地磁盘
