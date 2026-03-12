"""
[INPUT]: 依赖 base64、json、urllib 与 services.config 的 Fish Audio 配置
[OUTPUT]: 对外提供 Fish Audio 文本转语音的 base64 音频与流式音频生成器
[POS]: services 的语音层，被 app.py 的 /api/tts 路由消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import base64
import json
import urllib.request
import uuid

import msgpack
from websocket import create_connection

from services.config import get_fish_audio_key, get_fish_audio_reference_id

FISH_HTTP_URL = "https://api.fish.audio/v1/tts"
FISH_WS_URL = "wss://api.fish.audio/v1/tts/live"


def synthesize_fish_audio(text):
    text = (text or "").strip()
    if not text:
        return ""

    print(f"[TTS] Fish Audio request: {len(text)} chars", flush=True)

    payload = json.dumps({
        "text": text,
        "reference_id": get_fish_audio_reference_id(),
        "format": "mp3",
        "prosody": {"speed": 1.1},
    })
    request = urllib.request.Request(
        FISH_HTTP_URL,
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


def _pack_message(event, **payload):
    data = {"event": event}
    data.update(payload)
    return msgpack.packb(data, use_bin_type=True)


def stream_fish_audio(text):
    text = (text or "").strip()
    if not text:
        return

    print(f"[TTS STREAM] Fish Audio request: {len(text)} chars", flush=True)
    ws = create_connection(
        FISH_WS_URL,
        header=[
            f"Authorization: Bearer {get_fish_audio_key()}",
            "model: s2-pro",
        ],
        timeout=15,
    )

    try:
        ws.send_binary(
            _pack_message(
                "start",
                request={
                    "request_id": str(uuid.uuid4()),
                    "text": "",
                    "format": "mp3",
                    "reference_id": get_fish_audio_reference_id(),
                    "chunk_length": 150,
                    "latency": "normal",
                    "prosody": {"speed": 1.1},
                },
            )
        )
        ws.send_binary(_pack_message("text", text=text))
        ws.send_binary(_pack_message("flush"))
        ws.send_binary(_pack_message("stop"))

        while True:
            frame = ws.recv()
            if not frame:
                break

            message = msgpack.unpackb(frame, raw=False)
            event = message.get("event")

            if event == "audio":
                audio_chunk = message.get("audio")
                if isinstance(audio_chunk, bytes) and audio_chunk:
                    yield audio_chunk
            elif event == "error":
                raise RuntimeError(message.get("message") or "Fish Audio stream error")
            elif event in {"done", "finish"}:
                break
    finally:
        ws.close()
