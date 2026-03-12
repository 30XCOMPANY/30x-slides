"""
[INPUT]: 依赖 os、re、pathlib 获取运行环境与 provider secrets
[OUTPUT]: 对外提供环境变量清洗、启动诊断、必填校验、provider getter
[POS]: services 的配置真相源，被 app.py、chat.py、tts.py 共同消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import os
import re
from pathlib import Path

LLM_MODEL = "arcee-ai/prime"
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
REQUIRED_ENV_VARS = ("OPENROUTER_API_KEY", "FISH_AUDIO_API_KEY")


def clean_env_value(value):
    return re.sub(r"\s+", "", value or "")


def env_flag(name):
    return "SET" if clean_env_value(os.environ.get(name, "")) else "MISSING"


def log_boot_env():
    print(
        f"[BOOT] cwd={os.getcwd()} file_dir={Path(__file__).resolve().parent.parent} "
        f"OPENROUTER_API_KEY={env_flag('OPENROUTER_API_KEY')} "
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


def get_openrouter_key():
    return clean_env_value(os.environ.get("OPENROUTER_API_KEY", ""))


def get_openrouter_base_url():
    return (os.environ.get("OPENROUTER_BASE_URL", OPENROUTER_BASE_URL) or OPENROUTER_BASE_URL).strip()


def get_fish_audio_key():
    return clean_env_value(os.environ.get("FISH_AUDIO_API_KEY", ""))


def get_fish_audio_reference_id():
    value = os.environ.get("FISH_AUDIO_REFERENCE_ID", os.environ.get("FISH_AUDIO_VOICE_ID", ""))
    return clean_env_value(value)
