"""
[INPUT]: 依赖 os、re、pathlib 获取运行环境与 provider secrets
[OUTPUT]: 对外提供环境变量清洗、启动诊断、必填校验、provider getter
[POS]: services 的配置真相源，被 app.py、chat.py、tts.py 共同消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import os
import re
from pathlib import Path

LLM_MODEL = "MiniMax-M2.5-highspeed"
LLM_FALLBACK_MODEL = "MiniMax-M2.5"
MINIMAX_BASE_URL = "https://api.minimax.io/anthropic"
REQUIRED_ENV_VARS = ("MINIMAX_API_KEY", "FISH_AUDIO_API_KEY")


def clean_env_value(value):
    return re.sub(r"\s+", "", value or "")


def env_flag(name):
    return "SET" if clean_env_value(os.environ.get(name, "")) else "MISSING"


def log_boot_env():
    print(
        f"[BOOT] cwd={os.getcwd()} file_dir={Path(__file__).resolve().parent.parent} "
        f"MINIMAX_API_KEY={env_flag('MINIMAX_API_KEY')} "
        f"FISH_AUDIO_API_KEY={env_flag('FISH_AUDIO_API_KEY')} "
        f"FISH_AUDIO_REFERENCE_ID={env_flag('FISH_AUDIO_REFERENCE_ID')}",
        flush=True,
    )


def validate_required_env():
    missing = [name for name in REQUIRED_ENV_VARS if not clean_env_value(os.environ.get(name, ""))]
    if missing:
        raise RuntimeError(
            "Missing required environment variables: "
            + ", ".join(missing)
            + ". Set them in Railway Variables or local .env."
        )


def get_minimax_key():
    return clean_env_value(os.environ.get("MINIMAX_API_KEY", ""))


def get_minimax_base_url():
    return (os.environ.get("MINIMAX_BASE_URL", MINIMAX_BASE_URL) or MINIMAX_BASE_URL).strip()


def get_minimax_base_url_candidates():
    primary = get_minimax_base_url()
    candidates = [primary]
    if ".io/" in primary:
        candidates.append(primary.replace(".io/", ".com/"))
    elif ".com/" in primary:
        candidates.append(primary.replace(".com/", ".io/"))
    return candidates


def get_fish_audio_key():
    return clean_env_value(os.environ.get("FISH_AUDIO_API_KEY", ""))


def get_fish_audio_reference_id():
    value = os.environ.get("FISH_AUDIO_REFERENCE_ID", os.environ.get("FISH_AUDIO_VOICE_ID", ""))
    return clean_env_value(value)
