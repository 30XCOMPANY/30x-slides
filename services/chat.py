"""
[INPUT]: 依赖 httpx 的 OpenRouter 请求、services.config 的 provider 配置、services.decks 的 deck 上下文函数
[OUTPUT]: 对外提供 talk 请求处理、流式事件输出、session history 管理、回复裁剪与导航解析
[POS]: services 的对话层，被 app.py 的 /api/talk 路由消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import json
import re

import httpx

from services.config import (
    LLM_MODEL,
    get_openrouter_base_url,
    get_openrouter_key,
)
from services.decks import (
    build_greeting,
    build_greeting_prompt,
    interactive_follow_up,
    build_local_slide_reply,
    build_system_prompt,
    extract_slide_content,
    infer_navigation_target,
    load_deck_content,
    restore_deck_content,
    save_deck_content,
)

_chat_histories = {}
_pregenerated_greetings = {}


def pregenerate_greeting(deck_id, all_content):
    """上传时调用, 预生成 LLM greeting, 用户点 mic 直接用"""
    try:
        system_prompt = build_greeting_prompt(all_content)
        messages = [{"role": "user", "content": "[GREET]"}]
        greeting = chat_with_llm(system_prompt, messages)
        _pregenerated_greetings[deck_id] = greeting
        print(f"[PREGEN] greeting ready for {deck_id}: {len(greeting)} chars", flush=True)
    except Exception as exc:
        print(f"[PREGEN ERROR] {exc}", flush=True)
        _pregenerated_greetings[deck_id] = build_greeting(all_content)


def get_pregenerated_greeting(deck_id):
    return _pregenerated_greetings.pop(deck_id, None)
LOCAL_FAST_PATTERNS = (
    "what is this slide about",
    "what's this slide about",
    "summarize this",
    "summary",
    "quick summary",
    "what is happening here",
    "what's happening here",
    "explain this slide",
)


def build_openrouter_headers():
    return {
        "Authorization": f"Bearer {get_openrouter_key()}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://30x-slides-production.up.railway.app",
        "X-Title": "30x Slides",
    }


def build_openrouter_messages(system_prompt, messages):
    chat_messages = [{"role": "system", "content": system_prompt}]
    for message in messages:
        role = message.get("role", "user")
        content = (message.get("content") or "").strip()
        if content:
            chat_messages.append({"role": role, "content": content})
    return chat_messages


def extract_openrouter_text(payload):
    choices = payload.get("choices") or []
    if not choices:
        return ""
    message = choices[0].get("message") or {}
    return (message.get("content") or "").strip()


def chat_with_llm(system_prompt, messages):
    payload = {
        "model": LLM_MODEL,
        "messages": build_openrouter_messages(system_prompt, messages),
        "max_tokens": 512,
        "temperature": 0.6,
    }
    url = f"{get_openrouter_base_url()}/chat/completions"

    with httpx.Client(timeout=10.0) as client:
        response = client.post(url, headers=build_openrouter_headers(), json=payload)
        response.raise_for_status()
        content = extract_openrouter_text(response.json()) or "I'm not sure how to respond to that."

    print(
        f"[LLM] provider=openrouter model={LLM_MODEL} content_present={bool(content)} len={len(content)}",
        flush=True,
    )
    return content


def stream_chat_with_llm(system_prompt, messages):
    payload = {
        "model": LLM_MODEL,
        "messages": build_openrouter_messages(system_prompt, messages),
        "max_tokens": 512,
        "temperature": 0.6,
        "stream": True,
    }
    url = f"{get_openrouter_base_url()}/chat/completions"

    with httpx.Client(timeout=20.0) as client:
        with client.stream("POST", url, headers=build_openrouter_headers(), json=payload) as response:
            response.raise_for_status()
            for line in response.iter_lines():
                if not line or not line.startswith("data: "):
                    continue
                chunk = line[6:].strip()
                if chunk == "[DONE]":
                    break
                event = json.loads(chunk)
                choices = event.get("choices") or []
                if not choices:
                    continue
                delta = choices[0].get("delta") or {}
                text = delta.get("content") or ""
                if text:
                    yield text

    print(f"[LLM STREAM] provider=openrouter model={LLM_MODEL} completed", flush=True)


def ensure_interactive_ending(reply):
    reply = (reply or "").strip()
    if not reply:
        return "What would you like to know?"

    # 已经有问句或邀请性结尾的，原样返回
    if reply.rstrip()[-1] == "?":
        return reply

    interactive_markers = (
        "do you want", "want the", "should i", "which part", "where do you want",
        "how do you want", "want me to", "what do you want", "what would you like",
        "which section", "want to", "move on", "next one",
    )
    if any(marker in reply.lower() for marker in interactive_markers):
        return reply

    if reply[-1] not in ".!?":
        reply += "."
    return reply


def trim_spoken_reply(reply):
    """只做空白清理, 不截断内容 — 让 LLM 把问题回答完整"""
    reply = re.sub(r"\s+", " ", (reply or "")).strip()
    if not reply:
        return ""
    # 去掉 [GO:...] 残留
    reply = re.sub(r"\s*\[GO:\w+\]\s*", " ", reply).strip()
    return reply


def parse_nav_command(reply):
    nav_match = re.search(r"\[GO:(\w+)\]", reply or "")
    if not nav_match:
        return reply.strip(), None

    nav_command = nav_match.group(1)
    spoken = re.sub(r"\s*\[GO:\w+\]\s*", " ", reply).strip()
    return spoken, nav_command


def should_use_local_explainer(text):
    normalized_text = (text or "").lower()
    return any(pattern in normalized_text for pattern in LOCAL_FAST_PATTERNS)


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


def build_response_context(deck_id, payload, output_folder):
    text = payload.get("text", "").strip()
    slide_idx = payload.get("slide_index", 0)
    voice = payload.get("voice", "af_heart")
    session_id = payload.get("session_id", "default")

    print(f"[TALK] text={text}, slide={slide_idx}, voice={voice}", flush=True)

    all_content = ensure_deck_content(output_folder, deck_id, payload)
    normalized_text = text.lower()
    inferred_nav = infer_navigation_target(text, all_content, slide_idx)
    history_key = f"{deck_id}:{session_id}"
    history = _chat_histories.setdefault(history_key, [])
    is_internal = text.startswith("[")

    if len(history) > 40:
        history = history[-40:]
        _chat_histories[history_key] = history

    if not is_internal and text:
        history.append({"role": "user", "content": text})

    messages = list(history)
    if is_internal and text:
        messages.append({"role": "user", "content": text})

    return {
        "text": text,
        "slide_idx": slide_idx,
        "all_content": all_content,
        "normalized_text": normalized_text,
        "inferred_nav": inferred_nav,
        "history_key": history_key,
        "history": history,
        "messages": messages,
        "is_internal": is_internal,
    }


def split_completed_sentences(buffer):
    parts = re.split(r"(?<=[.!?])\s+", buffer)
    completed = [part.strip() for part in parts[:-1] if part.strip()]
    remainder = parts[-1] if parts else ""
    return completed, remainder


def emit_event(event_type, **payload):
    data = {"type": event_type}
    data.update(payload)
    return (json.dumps(data, ensure_ascii=False) + "\n").encode("utf-8")


def stream_scripted_reply(reply, nav_command=None):
    cleaned = trim_spoken_reply(reply)
    sentences = [sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+", cleaned) if sentence.strip()]
    visible = ""
    for sentence in sentences:
        visible = (visible + " " + sentence).strip()
        yield emit_event("text", text=visible)
        yield emit_event("sentence", text=sentence)
    yield emit_event("done", text=cleaned, nav=nav_command)


def stream_talk_response(deck_id, payload, output_folder):
    ctx = build_response_context(deck_id, payload, output_folder)
    text = ctx["text"]
    if not text:
        yield emit_event("done", text="", nav=None)
        return

    all_content = ctx["all_content"]
    slide_idx = ctx["slide_idx"]
    normalized_text = ctx["normalized_text"]
    inferred_nav = ctx["inferred_nav"]
    messages = ctx["messages"]
    history = ctx["history"]

    if text != "[GREET]" and should_use_local_explainer(text):
        raw_reply = build_local_slide_reply(all_content, slide_idx, text)
        history.append({"role": "assistant", "content": raw_reply})
        yield from stream_scripted_reply(raw_reply, inferred_nav)
        return

    # [GREET]: 优先用预生成的, 没有才调 LLM
    if text == "[GREET]":
        if history:
            yield emit_event("done", text="", nav=None)
            return
        cached = get_pregenerated_greeting(deck_id)
        if cached:
            history.append({"role": "assistant", "content": cached})
            yield from stream_scripted_reply(cached, None)
            return
        system_prompt = build_greeting_prompt(all_content)
        fallback_fn = lambda: build_greeting(all_content)
    else:
        system_prompt = build_system_prompt(all_content, slide_idx, allow_control_tags=True)
        fallback_fn = lambda: build_local_slide_reply(all_content, slide_idx, text)

    visible = ""
    sentence_buffer = ""
    completed_sentences = []
    fallback_reply = None
    llm_nav = None  # LLM 自己决定的 [GO:N] 导航

    def strip_nav_tag(s):
        """从文本中提取并移除 [GO:...] 标签"""
        nonlocal llm_nav
        match = re.search(r"\[GO:(\w+)\]", s)
        if match:
            llm_nav = match.group(1)
            s = re.sub(r"\s*\[GO:\w+\]\s*", " ", s).strip()
        return s

    try:
        for delta in stream_chat_with_llm(system_prompt, messages):
            sentence_buffer += delta
            visible += delta
            yield emit_event("text", text=strip_nav_tag(visible).strip())
            complete, sentence_buffer = split_completed_sentences(sentence_buffer)
            for sentence in complete:
                cleaned = strip_nav_tag(sentence)
                if cleaned:
                    completed_sentences.append(cleaned)
                    yield emit_event("sentence", text=cleaned)
    except Exception as exc:
        print(f"[LLM STREAM ERROR] {exc}", flush=True)
        fallback_reply = fallback_fn()

    if fallback_reply:
        history.append({"role": "assistant", "content": fallback_reply})
        yield from stream_scripted_reply(fallback_reply, inferred_nav)
        return

    if sentence_buffer.strip():
        cleaned = strip_nav_tag(sentence_buffer.strip())
        if cleaned:
            completed_sentences.append(cleaned)
            yield emit_event("sentence", text=cleaned)

    raw_reply = " ".join(completed_sentences).strip()
    raw_reply = trim_spoken_reply(raw_reply)
    history.append({"role": "assistant", "content": raw_reply})
    # LLM 的 [GO:N] 优先, 否则用关键词推断
    final_nav = llm_nav or (inferred_nav if text != "[GREET]" else None)
    yield emit_event("done", text=raw_reply, nav=final_nav)


def build_talk_response(deck_id, payload, output_folder):
    text = payload.get("text", "").strip()
    if not text:
        return {"text": "", "audio": "", "nav": None}

    ctx = build_response_context(deck_id, payload, output_folder)
    all_content = ctx["all_content"]
    slide_idx = ctx["slide_idx"]
    normalized_text = ctx["normalized_text"]
    inferred_nav = ctx["inferred_nav"]
    history = ctx["history"]
    messages = ctx["messages"]
    system_prompt = build_system_prompt(all_content, slide_idx)

    if text != "[GREET]" and should_use_local_explainer(text):
        raw_reply = build_local_slide_reply(all_content, slide_idx, text)
        print(f"[LOCAL FAST] {raw_reply[:120]}", flush=True)
    else:
        greet_prompt = build_greeting_prompt(all_content) if text == "[GREET]" else None
        try:
            raw_reply = chat_with_llm(greet_prompt or system_prompt, messages)
            print(f"[LLM] {raw_reply[:120]}", flush=True)
        except Exception as exc:
            print(f"[LLM ERROR] {exc}", flush=True)
            raw_reply = build_greeting(all_content) if text == "[GREET]" else build_local_slide_reply(all_content, slide_idx, text)

    spoken_reply, nav_command = parse_nav_command(raw_reply)
    if inferred_nav and not nav_command:
        nav_command = inferred_nav
    spoken_reply = trim_spoken_reply(spoken_reply)
    print(f"[LLM TRIMMED] len={len(spoken_reply)} text={spoken_reply[:120]}", flush=True)

    history.append({"role": "assistant", "content": raw_reply})
    return {"text": spoken_reply, "nav": nav_command}
