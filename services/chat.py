"""
[INPUT]: 依赖 anthropic SDK、services.config 的 provider 配置、services.decks 的 deck 上下文函数
[OUTPUT]: 对外提供 talk 请求处理、流式事件输出、session history 管理、回复裁剪与导航解析
[POS]: services 的对话层，被 app.py 的 /api/talk 路由消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import json
import re

from anthropic import Anthropic

from services.config import ANTHROPIC_MODEL, get_anthropic_key
from services.decks import (
    build_greeting,
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


def chat_with_llm(system_prompt, messages):
    client = Anthropic(api_key=get_anthropic_key(), timeout=6.0)
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


def stream_chat_with_llm(system_prompt, messages):
    client = Anthropic(api_key=get_anthropic_key(), timeout=20.0)
    with client.messages.stream(
        model=ANTHROPIC_MODEL,
        max_tokens=96,
        system=system_prompt,
        messages=messages,
    ) as stream:
        for text in stream.text_stream:
            if text:
                yield text


def ensure_interactive_ending(reply):
    reply = (reply or "").strip()
    if not reply:
        return interactive_follow_up("empty", 0)

    interactive_markers = ("do you want", "want the", "should i", "which part", "where do you want", "how do you want")
    if any(marker in reply.lower() for marker in interactive_markers):
        return reply

    if reply[-1] not in ".!?":
        reply += "."
    return reply + " " + interactive_follow_up(reply, 0)


def trim_spoken_reply(reply, max_chars=320, max_words=55):
    reply = re.sub(r"\s+", " ", (reply or "")).strip()
    if not reply:
        return ""

    sentences = re.split(r"(?<=[.!?])\s+", reply)
    first_five = " ".join(sentences[:5]).strip()
    if first_five and len(first_five) <= max_chars:
        return ensure_interactive_ending(first_five)

    words = reply.split()
    if len(words) > max_words:
        clipped = " ".join(words[:max_words]).strip(" ,;:-")
        if clipped and clipped[-1] not in ".!?":
            clipped += "."
        return ensure_interactive_ending(clipped)

    clipped = reply[:max_chars].rsplit(" ", 1)[0].strip(" ,;:-")
    if clipped and clipped[-1] not in ".!?":
        clipped += "."
    return ensure_interactive_ending(clipped or reply[:max_chars])


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

    if text == "[GREET]":
        raw_reply = build_greeting(all_content)
        history.append({"role": "assistant", "content": raw_reply})
        yield from stream_scripted_reply(raw_reply, inferred_nav)
        return

    if normalized_text in {"hello", "hi", "hey", "hey there", "yo"} or should_use_local_explainer(text):
        raw_reply = build_greeting(all_content) if normalized_text in {"hello", "hi", "hey", "hey there", "yo"} else build_local_slide_reply(all_content, slide_idx, text)
        history.append({"role": "assistant", "content": raw_reply})
        yield from stream_scripted_reply(raw_reply, inferred_nav)
        return

    system_prompt = build_system_prompt(all_content, slide_idx, allow_control_tags=False)
    visible = ""
    sentence_buffer = ""
    completed_sentences = []
    fallback_reply = None

    try:
        for delta in stream_chat_with_llm(system_prompt, messages):
            sentence_buffer += delta
            visible += delta
            yield emit_event("text", text=visible.strip())
            complete, sentence_buffer = split_completed_sentences(sentence_buffer)
            for sentence in complete:
                completed_sentences.append(sentence)
                if len(completed_sentences) <= 5:
                    yield emit_event("sentence", text=sentence)
    except Exception as exc:
        print(f"[LLM STREAM ERROR] {exc}", flush=True)
        fallback_reply = build_local_slide_reply(all_content, slide_idx, text)

    if fallback_reply:
        history.append({"role": "assistant", "content": fallback_reply})
        yield from stream_scripted_reply(fallback_reply, inferred_nav)
        return

    if sentence_buffer.strip():
        completed_sentences.append(sentence_buffer.strip())
        if len(completed_sentences) <= 5:
            yield emit_event("sentence", text=sentence_buffer.strip())

    raw_reply = " ".join(completed_sentences).strip()
    raw_reply = trim_spoken_reply(raw_reply)
    history.append({"role": "assistant", "content": raw_reply})
    yield emit_event("done", text=raw_reply, nav=inferred_nav)


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

    if text == "[GREET]":
        raw_reply = build_greeting(all_content)
        print(f"[GREET LOCAL] {raw_reply}", flush=True)
    elif normalized_text in {"hello", "hi", "hey", "hey there", "yo"}:
        raw_reply = build_greeting(all_content)
        print(f"[SMALLTALK LOCAL] {raw_reply}", flush=True)
    elif should_use_local_explainer(text):
        raw_reply = build_local_slide_reply(all_content, slide_idx, text)
        print(f"[LOCAL FAST] {raw_reply[:120]}", flush=True)
    else:
        try:
            raw_reply = chat_with_llm(system_prompt, messages)
            print(f"[LLM] {raw_reply[:120]}", flush=True)
        except Exception as exc:
            print(f"[LLM ERROR] {exc}", flush=True)
            raw_reply = build_local_slide_reply(all_content, slide_idx, text)

    spoken_reply, nav_command = parse_nav_command(raw_reply)
    if inferred_nav and not nav_command:
        nav_command = inferred_nav
    spoken_reply = trim_spoken_reply(spoken_reply)
    print(f"[LLM TRIMMED] len={len(spoken_reply)} text={spoken_reply[:120]}", flush=True)

    history.append({"role": "assistant", "content": raw_reply})
    return {"text": spoken_reply, "nav": nav_command}
