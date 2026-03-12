# services/
> L2 | 父级: /Users/nora/Desktop/Interactive Slide/30x-slides/AGENTS.md

成员清单
config.py: 环境变量清洗、必填校验、provider 配置常量，[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
decks.py: deck 内容读写、viewer payload 恢复、标题与 prompt 上下文构建，[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
chat.py: OpenRouter 对话、session history、导航解析、回复裁剪，[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
tts.py: Fish Audio TTS 调用与 base64 音频封装，[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
__init__.py: services 包声明，[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md

法则: provider 配置单一真相源；deck 上下文优先稳定恢复；回复长度在服务层硬约束
