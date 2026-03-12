"""
[INPUT]: 依赖同目录下的 config、decks、chat、tts 模块
[OUTPUT]: 对外提供 services 包边界
[POS]: 30x-slides 的服务层命名空间，承接 app.py 拆出的 provider 与状态逻辑
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""
