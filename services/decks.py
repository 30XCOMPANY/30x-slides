"""
[INPUT]: 依赖 json、os 与 python-pptx 读取 deck 内容和 viewer 恢复载荷
[OUTPUT]: 对外提供 deck 内容读写、标准化、标题提取与 system prompt 构建
[POS]: services 的 deck 元数据层，被 app.py 上传流程与 chat.py 对话流程共同消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import json
import os


IGNORED_TITLES = {"title slide", "thank you", "questions", "end", ""}


def extract_slide_content(pptx_path):
    """Extract text and speaker notes for each slide."""
    slides_content = []
    try:
        from pptx import Presentation

        prs = Presentation(pptx_path)
        for slide in prs.slides:
            texts = []
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for para in shape.text_frame.paragraphs:
                    line = para.text.strip()
                    if line:
                        texts.append(line)

            notes = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()

            slides_content.append({
                "slide": len(slides_content) + 1,
                "text": texts,
                "notes": notes,
            })
    except Exception as exc:
        print(f"[extract] error: {exc}", flush=True)

    return slides_content


def content_path(output_folder, deck_id):
    return os.path.join(output_folder, deck_id, "content.json")


def load_deck_content(output_folder, deck_id):
    path = content_path(output_folder, deck_id)
    if os.path.exists(path):
        with open(path) as handle:
            return json.load(handle)
    return []


def save_deck_content(output_folder, deck_id, content):
    deck_dir = os.path.join(output_folder, deck_id)
    os.makedirs(deck_dir, exist_ok=True)
    with open(content_path(output_folder, deck_id), "w") as handle:
        json.dump(content, handle, ensure_ascii=False)


def normalize_deck_content(raw_content):
    """Accept trusted deck content from the viewer when container-local files are gone."""
    if not isinstance(raw_content, list):
        return []

    normalized = []
    for index, item in enumerate(raw_content, start=1):
        if not isinstance(item, dict):
            continue

        texts = item.get("text", [])
        if not isinstance(texts, list):
            texts = []
        texts = [str(line).strip() for line in texts if str(line).strip()]

        notes = item.get("notes", "")
        if not isinstance(notes, str):
            notes = str(notes or "")

        slide_number = item.get("slide", index)
        try:
            slide_number = int(slide_number)
        except Exception:
            slide_number = index

        normalized.append({
            "slide": slide_number,
            "text": texts,
            "notes": notes.strip(),
        })

    return normalized


def restore_deck_content(output_folder, deck_id, raw_content):
    content = normalize_deck_content(raw_content)
    if content:
        save_deck_content(output_folder, deck_id, content)
        print(f"[TALK] restored {len(content)} slides content from viewer payload", flush=True)
    return content


def get_slide_title(slide_content):
    return slide_content["text"][0] if slide_content.get("text") else f"Slide {slide_content['slide']}"


def build_greeting(all_content):
    titles = []
    seen = set()

    for slide_content in all_content:
        title = get_slide_title(slide_content).strip()
        normalized = title.lower()
        if normalized in IGNORED_TITLES or normalized in seen:
            continue
        seen.add(normalized)
        titles.append(title)

    if not titles:
        return "Alright, I can quickly summarize the whole thing and walk you through it step by step. How do you wanna learn this?"

    themes = titles[:4]
    if len(themes) == 1:
        return f"Alright, this mainly covers {themes[0]}. I can summarize it first or go slide by slide. How do you wanna learn this?"

    if len(themes) == 2:
        theme_text = f"{themes[0]} and {themes[1]}"
    else:
        theme_text = ", ".join(themes[:-1]) + f", and {themes[-1]}"

    return f"Alright, this deck is basically about {theme_text}. I can give you the big-picture summary or teach it step by step. How do you wanna learn this?"


def build_local_slide_reply(all_content, current_slide_idx, user_text=""):
    """Fallback explanation when the hosted LLM is unavailable."""
    if not all_content:
        return "I can keep going, but I need the slide content loaded first. Want the quick summary or slide-by-slide version?"

    if current_slide_idx < 0 or current_slide_idx >= len(all_content):
        current_slide_idx = 0

    slide_content = all_content[current_slide_idx]
    title = get_slide_title(slide_content)
    body_lines = [line for line in slide_content.get("text", []) if line.strip()]
    body_lines = body_lines[1:4] if len(body_lines) > 1 else body_lines[:3]

    if not body_lines:
        return f"This part is about {title}. Do you want the short version, the key takeaway, or the next slide?"

    snippet = "; ".join(body_lines)
    reply = f"This part is about {title}: {snippet}. Want the short version, the key takeaway, or the next slide?"
    return reply[:220].rstrip(" ,;:-") + ("." if not reply.endswith(".") else "")


def build_system_prompt(all_content, current_slide_idx):
    title_index = ""
    for slide_content in all_content:
        title = get_slide_title(slide_content)
        title_index += f"  Slide {slide_content['slide']}: {title}\n"

    current_detail = ""
    if 0 <= current_slide_idx < len(all_content):
        slide_content = all_content[current_slide_idx]
        current_detail = f"Slide {slide_content['slide']} — {get_slide_title(slide_content)}:\n"
        current_detail += "\n".join(slide_content["text"])
        if slide_content.get("notes"):
            current_detail += f"\nSpeaker notes: {slide_content['notes']}"

    nav_examples = ""
    example_slides = []
    for slide_content in all_content:
        title = get_slide_title(slide_content)
        if title and title.lower() not in IGNORED_TITLES:
            example_slides.append((slide_content["slide"], title))

    for slide_number, slide_title in example_slides[:3]:
        keyword = slide_title.split()[0].lower() if slide_title.split() else slide_title.lower()
        nav_examples += (
            f'- User says "{keyword}" → you write [GO:{slide_number}] '
            f'because Slide {slide_number} is "{slide_title}"\n'
        )
    nav_examples += '- User says "next" → you write [GO:next]\n'
    nav_examples += '- User says "go back" → you write [GO:prev]\n'

    return f"""You explain slide content directly. Talk about the topic itself, not "the deck" or "this slide".

Spoken voice only. No markdown, no bullets, no lists.

VIBE: Casual, warm, and tutor-like. Sound like a smart human guide, not a narrator.

[GREET]: Under 34 words. First summarize what this deck is about in plain language, mention 2-4 main themes, then ask "How do you wanna learn this?" Do NOT include [GO:N].

REPLIES: MAX 3 short sentences. Answer briefly, then offer a clear next move like "Want the short version, the deeper takeaway, or the next slide?" Keep it natural and conversational.

Viewer is on Slide {current_slide_idx + 1} of {len(all_content)}.

SLIDE TITLES:
{title_index}
CURRENT SLIDE:
{current_detail}

===== NAVIGATION RULES (MANDATORY) =====
When the user mentions ANY topic, keyword, or phrase that relates to a slide title, you MUST include [GO:N].

EXAMPLES from this deck:
{nav_examples}
HOW IT WORKS: You write [GO:N] anywhere in your reply. The system removes it before showing to user and auto-jumps the slide.

RULES:
1. Match loosely. Any recognizable keyword should navigate.
2. Never ask whether to navigate. Just include [GO:N].
3. "next" → [GO:next], "back"/"previous" → [GO:prev].
4. If user asks about a topic and you omit [GO:N], your response is wrong.
5. Always include [GO:N] before your spoken text, like: [GO:3] So this one covers...
======================================="""
