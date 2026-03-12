"""
[INPUT]: 依赖 anthropic SDK、services.config 的 provider 配置、services.decks 的 deck 上下文函数
[OUTPUT]: 对外提供 talk 请求处理、session history 管理、回复裁剪与导航解析
[POS]: services 的对话层，被 app.py 的 /api/talk 路由消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import re

from anthropic import Anthropic

from services.config import ANTHROPIC_MODEL, get_anthropic_key
from services.decks import (
    build_greeting,
    build_system_prompt,
    extract_slide_content,
    load_deck_content,
    restore_deck_content,
    save_deck_content,
)

_chat_histories = {}


def chat_with_llm(system_prompt, messages):
    client = Anthropic(api_key=get_anthropic_key())
    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=48,
        system=system_prompt,
        messages=messages,
    )

    text_blocks = [
        block.text
        for block in response.content
        if getattr(block, "type", "") == "text" and getattr(block, "text", "")
    ]
    content = "".join(text_blocks).strip()
    if not content:
        content = "I'm not sure how to respond to that."

    print(
        f"[LLM] model={ANTHROPIC_MODEL} content_present={bool(content)} len={len(content)}",
        flush=True,
    )
    return content


def trim_spoken_reply(reply, max_chars=110, max_words=16):
    reply = re.sub(r"\s+", " ", (reply or "")).strip()
    if not reply:
        return ""

    first_sentence = re.split(r"(?<=[.!?])\s+", reply, maxsplit=1)[0].strip()
    if first_sentence and len(first_sentence) <= max_chars:
        return first_sentence

    words = reply.split()
    if len(words) > max_words:
        clipped = " ".join(words[:max_words]).strip(" ,;:-")
        if clipped and clipped[-1] not in ".!?":
            clipped += "."
        return clipped

    clipped = reply[:max_chars].rsplit(" ", 1)[0].strip(" ,;:-")
    if clipped and clipped[-1] not in ".!?":
        clipped += "."
    return clipped or reply[:max_chars]


def parse_nav_command(reply):
    nav_match = re.search(r"\[GO:(\w+)\]", reply or "")
    if not nav_match:
        return reply.strip(), None

    nav_command = nav_match.group(1)
    spoken = re.sub(r"\s*\[GO:\w+\]\s*", " ", reply).strip()
    return spoken, nav_command


def ensure_deck_content(output_folder, deck_id, payload):
    all_content = load_deck_content(output_folder, deck_id)
    if all_content:
        return all_content

    pptx_path = f"{output_folder}/{deck_id}/deck.pptx"
    try:
        all_content = extract_slide_content(pptx_path)
    except Exception:
        all_content = []

    if all_content:
        save_deck_content(output_folder, deck_id, all_content)
        print(f"[TALK] extracted {len(all_content)} slides content", flush=True)
        return all_content

    return restore_deck_content(output_folder, deck_id, payload.get("all_content"))


def build_talk_response(deck_id, payload, output_folder):
    text = payload.get("text", "").strip()
    slide_idx = payload.get("slide_index", 0)
    voice = payload.get("voice", "af_heart")
    session_id = payload.get("session_id", "default")

    print(f"[TALK] text={text}, slide={slide_idx}, voice={voice}", flush=True)

    if not text:
        return {"text": "", "audio": "", "nav": None}

    all_content = ensure_deck_content(output_folder, deck_id, payload)
    system_prompt = build_system_prompt(all_content, slide_idx)

    history_key = f"{deck_id}:{session_id}"
    history = _chat_histories.setdefault(history_key, [])
    is_internal = text.startswith("[")

    if len(history) > 40:
        history = history[-40:]
        _chat_histories[history_key] = history

    if not is_internal:
        history.append({"role": "user", "content": text})

    messages = list(history)
    if is_internal:
        messages.append({"role": "user", "content": text})

    if text == "[GREET]":
        try:
            raw_reply = chat_with_llm(system_prompt, messages)
            print(f"[GREET LLM] {raw_reply[:120]}", flush=True)
        except Exception as exc:
            print(f"[GREET FALLBACK] {exc}", flush=True)
            raw_reply = build_greeting(all_content)
            print(f"[GREET LOCAL] {raw_reply}", flush=True)
    else:
        try:
            raw_reply = chat_with_llm(system_prompt, messages)
            print(f"[LLM] {raw_reply[:120]}", flush=True)
        except Exception as exc:
            print(f"[LLM ERROR] {exc}", flush=True)
            raw_reply = "Sorry, let me try again."

    spoken_reply, nav_command = parse_nav_command(raw_reply)
    spoken_reply = trim_spoken_reply(spoken_reply)
    print(f"[LLM TRIMMED] len={len(spoken_reply)} text={spoken_reply[:120]}", flush=True)

    history.append({"role": "assistant", "content": raw_reply})
    return {"text": spoken_reply, "nav": nav_command}
