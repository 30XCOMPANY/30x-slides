# 30x-slides - PPTX 转交互式讲解页
Flask + LibreOffice + PyMuPDF + python-pptx + Anthropic + Fish Audio

<directory>
templates/ - 页面模板与前端交互脚本（2 文件: index.html, viewer.html）
output/ - 转换产物与 deck 状态缓存（每个 deck 一个目录）
</directory>

<config>
app.py - 单体入口，负责上传、转换、deck 内容缓存、对话与 TTS
requirements.txt - Python 运行依赖
nixpacks.toml - Railway 构建与启动配置
</config>

法则: deck 上下文不能只依赖容器本地磁盘；对话链路必须能从前端恢复内容；旧 provider 配置及时清理
