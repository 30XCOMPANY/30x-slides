"""
[INPUT]: 依赖 json、os 与 python-pptx 读取 deck 内容和 viewer 恢复载荷
[OUTPUT]: 对外提供 deck 内容读写、标准化、标题提取与 system prompt 构建
[POS]: services 的 deck 元数据层，被 app.py 上传流程与 chat.py 对话流程共同消费
[PROTOCOL]: 变更时更新此头部，然后检查 AGENTS.md
"""

import json
import os
import re


IGNORED_TITLES = {"title slide", "thank you", "questions", "end", ""}
STOPWORDS = {
    "the", "and", "for", "with", "from", "that", "this", "into", "about", "your",
    "have", "will", "what", "when", "where", "which", "their", "there", "here",
    "slide", "deck", "overview", "summary", "introduction", "intro", "section",
    "part", "more", "than", "into", "onto", "over", "under", "main", "topic",
}
INTERACTIVE_FOLLOW_UPS = (
    "Do you want the short version, the deeper takeaway, or the next slide?",
    "Do you want me to unpack the main point, compare it with the last slide, or move forward?",
    "Should I zoom into the key idea here, connect it to the bigger story, or jump ahead?",
    "Do you want the practical takeaway, the strategy behind it, or the next section?",
    "Want me to simplify this part, go one level deeper, or take you to the next slide?",
)


def normalize_spoken_text(text):
    """Strip visual separators that sound awkward in speech."""
    text = str(text or "")
    text = text.replace("|", ", ")
    text = text.replace("/", ", ")
    text = re.sub(r"\s*,\s*,+", ", ", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip(" ,")


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
                        texts.append(normalize_spoken_text(line))

            notes = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = normalize_spoken_text(slide.notes_slide.notes_text_frame.text.strip())

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
        texts = [normalize_spoken_text(str(line).strip()) for line in texts if str(line).strip()]

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
            "notes": normalize_spoken_text(notes.strip()),
        })

    return normalized


def restore_deck_content(output_folder, deck_id, raw_content):
    content = normalize_deck_content(raw_content)
    if content:
        save_deck_content(output_folder, deck_id, content)
        print(f"[TALK] restored {len(content)} slides content from viewer payload", flush=True)
    return content


def get_slide_title(slide_content):
    raw_title = slide_content["text"][0] if slide_content.get("text") else f"Slide {slide_content['slide']}"
    return normalize_spoken_text(raw_title)


def split_phrases(text):
    cleaned = normalize_spoken_text(text)
    if not cleaned:
        return []
    parts = re.split(r"[,:;()\-]\s*|\s+and\s+|\s+vs\.?\s+|\s+to\s+", cleaned)
    return [part.strip() for part in parts if len(part.strip()) >= 3]


def meaningful_tokens(text):
    return [
        token
        for token in re.findall(r"[a-z0-9]+", normalize_spoken_text(text).lower())
        if len(token) >= 3 and token not in STOPWORDS
    ]


def extract_theme_phrases(all_content, limit=4):
    scores = {}
    ordered = []

    for idx, slide_content in enumerate(all_content):
        title = get_slide_title(slide_content)
        title_phrases = split_phrases(title) or ([title] if title else [])
        body_phrases = []
        for line in slide_content.get("text", [])[1:4]:
            body_phrases.extend(split_phrases(line)[:1])

        for phrase in title_phrases[:2] + body_phrases[:2]:
            normalized = phrase.lower()
            if normalized in IGNORED_TITLES or len(normalized) < 3:
                continue
            if normalized not in scores:
                ordered.append(normalized)
                scores[normalized] = {"phrase": phrase, "score": 0}
            weight = 3 if idx == 0 else 2
            if phrase in body_phrases:
                weight = 1
            scores[normalized]["score"] += weight

    ranked = sorted(ordered, key=lambda key: (-scores[key]["score"], ordered.index(key)))
    return [scores[key]["phrase"] for key in ranked[:limit]]


def join_phrases(phrases):
    if not phrases:
        return ""
    if len(phrases) == 1:
        return phrases[0]
    if len(phrases) == 2:
        return f"{phrases[0]} and {phrases[1]}"
    return ", ".join(phrases[:-1]) + f", and {phrases[-1]}"


def interactive_follow_up(seed_text="", slide_number=0):
    seed = f"{slide_number}:{normalize_spoken_text(seed_text).lower()}"
    index = sum(ord(ch) for ch in seed) % len(INTERACTIVE_FOLLOW_UPS)
    return INTERACTIVE_FOLLOW_UPS[index]


def summarize_slide_core(slide_content):
    title = get_slide_title(slide_content)
    body_lines = [line for line in slide_content.get("text", []) if line.strip()]
    detail_lines = body_lines[1:5] if len(body_lines) > 1 else body_lines[:4]

    concept_phrases = []
    seen = set()
    for line in detail_lines:
        for phrase in split_phrases(line)[:2]:
            normalized = phrase.lower()
            if normalized == title.lower() or normalized in seen:
                continue
            seen.add(normalized)
            concept_phrases.append(phrase)
            if len(concept_phrases) >= 3:
                break
        if len(concept_phrases) >= 3:
            break

    return {
        "title": title,
        "concepts": concept_phrases,
        "body_lines": detail_lines,
    }


def build_navigation_index(all_content):
    index = []
    for slide_content in all_content:
        title = get_slide_title(slide_content)
        summary = summarize_slide_core(slide_content)
        phrases = [title] + summary["concepts"]
        body_lines = summary["body_lines"]
        if body_lines:
            phrases.extend(body_lines[:2])

        keywords = set()
        for phrase in phrases:
            keywords.update(meaningful_tokens(phrase))

        index.append({
            "slide": slide_content["slide"],
            "title": title,
            "keywords": keywords,
            "summary": summary,
        })
    return index


def infer_navigation_target(user_text, all_content, current_slide_idx):
    normalized = normalize_spoken_text(user_text).lower()
    if not normalized:
        return None

    if "next slide" in normalized or normalized in {"next", "next one", "move on"}:
        return "next"
    if "previous slide" in normalized or "go back" in normalized or normalized in {"back", "previous", "prev"}:
        return "prev"

    requested_number = re.search(r"\bslide\s+(\d{1,2})\b", normalized)
    if requested_number:
        return requested_number.group(1)

    nav_index = build_navigation_index(all_content)
    query_tokens = meaningful_tokens(normalized)
    if not query_tokens:
        return None

    best_slide = None
    best_score = 0
    current_slide_number = current_slide_idx + 1
    for candidate in nav_index:
        overlap = len(candidate["keywords"].intersection(query_tokens))
        if overlap <= 0:
            continue
        score = overlap * 3
        if candidate["slide"] == current_slide_number:
            score += 1
        title_tokens = set(meaningful_tokens(candidate["title"]))
        if title_tokens.intersection(query_tokens):
            score += 2
        if score > best_score:
            best_score = score
            best_slide = candidate["slide"]

    return str(best_slide) if best_slide and best_score >= 3 else None


def build_greeting(all_content):
    themes = extract_theme_phrases(all_content)
    if not themes:
        return "Hey, good to have you here. I can give you the big-picture summary first and then walk through the details. What do you want to understand first?"
    theme_text = join_phrases(themes)

    return (
        f"Hey, good to have you here. This deck breaks down {theme_text}. "
        "I can give you the big picture, zoom into one section, or walk it slide by slide. "
        "What do you want to understand first?"
    )


def build_local_slide_reply(all_content, current_slide_idx, user_text=""):
    """Fallback explanation when the hosted LLM is unavailable."""
    if not all_content:
        return (
            "I can keep going, but I need the slide content loaded first. "
            "Right now I do not have the actual slide details. "
            "Once the content is loaded, I can break it down clearly. "
            "Do you want the quick summary, the key takeaway, or the slide-by-slide version?"
        )

    if current_slide_idx < 0 or current_slide_idx >= len(all_content):
        current_slide_idx = 0

    slide_content = all_content[current_slide_idx]
    summary = summarize_slide_core(slide_content)
    title = summary["title"]
    concepts = summary["concepts"]
    body_lines = summary["body_lines"]

    if not body_lines:
        return (
            f"This part is about {title}. "
            "It is setting up the main idea more than listing detailed evidence. "
            "So the useful read here is the direction it is pointing you toward. "
            f"{interactive_follow_up(user_text or title, slide_content['slide'])}"
        )

    concept_text = join_phrases(concepts[:3])
    first = body_lines[0]
    sentences = [f"This part is really about {title}."]
    if concept_text:
        sentences.append(f"The core idea is how {concept_text} connect in one story.")
    else:
        sentences.append(f"The core idea is {first}.")

    if len(body_lines) > 1:
        sentences.append(
            f"What matters most is not each bullet by itself, but the pattern between {body_lines[0]} and {body_lines[1]}."
        )
    else:
        sentences.append("So the slide is trying to give you one clear takeaway, not just a list to read aloud.")

    if len(body_lines) > 2:
        sentences.append(f"The extra detail here is {body_lines[2]}, which gives the main point more weight.")
    else:
        sentences.append("That is the practical read of this slide once you strip away the presentation wording.")

    sentences.append(interactive_follow_up(user_text or title, slide_content["slide"]))
    return " ".join(sentences[:5])


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

    return f"""You explain slide content like a sharp human presenter. Talk about the idea underneath the slide, not the slide wording itself.

Spoken voice only. No markdown, no bullets, no lists.

VIBE: Casual, warm, and tutor-like. Sound like a smart human guide, not a narrator.

NEVER read bullets line by line. Synthesize. Compress. Explain what the slide is trying to say and why it matters.

[GREET]: EXACT STRUCTURE: first greet the user naturally, then summarize the deck in plain language with 2-4 major themes, then ask what they want to understand first. End with a question like "What do you want to understand first?" Do NOT include [GO:N].

REPLIES: ALWAYS 3-5 short sentences. The last sentence MUST be interactive and invite the user to choose a next move. Vary the final question naturally. Do not reuse the exact same closing line every time.

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
