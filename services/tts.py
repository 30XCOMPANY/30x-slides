"""
[INPUT]: 依赖 base64、json、urllib 与 services.config 的 Fish Audio 配置
[OUTPUT]: 对外提供 Fish Audio 文本转语音的 base64 音频
[POS]: services 的语音层，被 app.py 的 /api/tts 路由消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import base64
import json
import urllib.request

from services.config import get_fish_audio_key, get_fish_audio_reference_id


def synthesize_fish_audio(text):
    text = (text or "").strip()
    if not text:
        return ""

    print(f"[TTS] Fish Audio request: {len(text)} chars", flush=True)

    url = "https://api.fish.audio/v1/tts"
    payload = json.dumps({
        "text": text,
        "reference_id": get_fish_audio_reference_id(),
        "format": "mp3",
    })
    request = urllib.request.Request(
        url,
        data=payload.encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {get_fish_audio_key()}",
            "model": "s2-pro",
        },
    )

    with urllib.request.urlopen(request, timeout=15) as response:
        audio_bytes = response.read()

    print(f"[TTS] Fish Audio OK: {len(audio_bytes)} bytes", flush=True)
    return base64.b64encode(audio_bytes).decode()
