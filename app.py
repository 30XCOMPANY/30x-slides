"""
30x Slides — Upload PPTX, get interactive HTML slides with Voice Bot.
Stack: Flask + SocketIO + LibreOffice + pymupdf + Kokoro TTS + Bailian LLM

[INPUT]: PPTX file upload
[OUTPUT]: Interactive slide viewer with voice bot + downloadable HTML
[POS]: Application entry point, orchestrates conversion + voice pipeline
[PROTOCOL]: 变更时更新此头部，然后检查 CLAUDE.md
"""

import os
import re
import json
import uuid
import base64
import subprocess
import urllib.request
import threading
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
from flask import (
    Flask, request, redirect, url_for,
    render_template, send_from_directory, jsonify, Response,
)
from flask_socketio import SocketIO, emit

import fitz  # pymupdf

app = Flask(__name__)
app.config["SECRET_KEY"] = "30x-slides-secret"
app.config["OUTPUT_FOLDER"] = os.path.join(os.path.dirname(__file__), "output")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB

socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

# ---- 字体目录 (macOS vs Linux) ----
import platform
if platform.system() == "Darwin":
    FONT_DIR = Path.home() / "Library" / "Fonts"
else:
    FONT_DIR = Path.home() / ".local" / "share" / "fonts"
FONT_DIR.mkdir(parents=True, exist_ok=True)


# ---- MiniMax LLM ----
MINIMAX_API_KEY = os.environ.get("MINIMAX_API_KEY", "")
MINIMAX_MODEL = "MiniMax-Text-01"

# ---- ElevenLabs TTS ----
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY", "")
ELEVENLABS_VOICE_ID = os.environ.get("ELEVENLABS_VOICE_ID", "HY09gbZLEpQrUZjrJgJv")

# ---- 启动校验 ----
print(f"[BOOT] MINIMAX_API_KEY={'SET' if MINIMAX_API_KEY else 'MISSING'}", flush=True)
print(f"[BOOT] ELEVENLABS_API_KEY={'SET' if ELEVENLABS_API_KEY else 'MISSING'}", flush=True)
print(f"[BOOT] ELEVENLABS_VOICE_ID={ELEVENLABS_VOICE_ID}", flush=True)




# ============================================================
# 字体工具
# ============================================================
def get_pptx_fonts(pptx_path):
    """从 PPTX theme XML + slide XML 提取所有字体名称"""
    fonts = set()
    try:
        import zipfile as zf
        from lxml import etree
        with zf.ZipFile(pptx_path) as z:
            for name in z.namelist():
                if name.endswith(".xml"):
                    try:
                        tree = etree.parse(z.open(name))
                        for elem in tree.iter():
                            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                            if tag in ("latin", "ea", "cs", "rFont"):
                                tf = elem.get("typeface") or elem.get("val")
                                if tf and not tf.startswith("+") and tf not in ("", "Calibri"):
                                    fonts.add(tf)
                    except Exception:
                        pass
    except Exception:
        pass

    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.name:
                                fonts.add(run.font.name)
    except Exception:
        pass

    return fonts


def get_system_fonts():
    fonts = set()
    try:
        result = subprocess.run(
            ["fc-list", "--format", "%{family}\n"],
            capture_output=True, text=True,
        )
        for line in result.stdout.splitlines():
            name = line.strip().split(",")[0].strip()
            if name:
                fonts.add(name)
    except Exception:
        pass
    return sorted(fonts)


def is_font_installed(font_name):
    try:
        result = subprocess.run(["fc-list"], capture_output=True, text=True)
        normalized = font_name.lower().replace(" ", "")
        for line in result.stdout.splitlines():
            if normalized in line.lower().replace(" ", ""):
                return True
    except Exception:
        pass
    return False


def try_install_font(font_name):
    """只装 400 + 700 两个权重，够用且快"""
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", font_name).strip("-").lower()
    weights = {"400": "latin-400-normal", "700": "latin-700-normal"}
    installed = False
    for wk, ws in weights.items():
        url = f"https://cdn.jsdelivr.net/fontsource/fonts/{slug}@latest/{ws}.ttf"
        dest = FONT_DIR / f"{slug}-{wk}.ttf"
        if dest.exists():
            installed = True
            continue
        try:
            urllib.request.urlretrieve(url, str(dest))
            with open(dest, "rb") as f:
                header = f.read(4)
            if header[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO"):
                installed = True
            else:
                dest.unlink()
        except Exception:
            if dest.exists():
                dest.unlink()
    return installed


def ensure_fonts(pptx_path):
    fonts = get_pptx_fonts(pptx_path)
    report = {"found": [], "installed": [], "missing": []}
    for font in sorted(fonts):
        if is_font_installed(font):
            report["found"].append(font)
        elif try_install_font(font):
            report["installed"].append(font)
        else:
            report["missing"].append(font)
    if report["installed"]:
        subprocess.run(["fc-cache", "-f"], capture_output=True)
    print(f"[FONTS] found={len(report['found'])} installed={len(report['installed'])} missing={len(report['missing'])}", flush=True)
    return report


# ============================================================
# 内容提取
# ============================================================
def extract_slide_content(pptx_path):
    """提取每页 slide 的文字内容 + speaker notes"""
    slides_content = []
    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        for slide in prs.slides:
            texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
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
    except Exception as e:
        print(f"[extract] error: {e}")

    return slides_content


# ============================================================
# 转换管线
# ============================================================
def convert_deck(deck_dir):
    pptx_path = os.path.join(deck_dir, "deck.pptx")

    # ---- LibreOffice 无头渲染 ----
    print(f"[CONVERT] LibreOffice starting: {pptx_path}", flush=True)
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", pptx_path, "--outdir", deck_dir],
        capture_output=True, text=True, timeout=120,
    )
    print(f"[CONVERT] LibreOffice done, rc={result.returncode}", flush=True)
    if result.returncode != 0:
        print(f"[CONVERT ERROR] {result.stderr}", flush=True)
        raise RuntimeError(f"LibreOffice failed: {result.stderr}")
    doc = fitz.open(os.path.join(deck_dir, "deck.pdf"))
    total = len(doc)
    for i in range(total):
        page = doc[i]
        mat = fitz.Matrix(1920 / page.rect.width, 1080 / page.rect.height)
        pix = page.get_pixmap(matrix=mat)
        pix.save(os.path.join(deck_dir, f"slide-{i + 1}.png"))
    doc.close()
    with open(os.path.join(deck_dir, "meta.txt"), "w") as f:
        f.write(str(total))

    # ---- 提取文字内容 ----
    content = extract_slide_content(pptx_path)
    with open(os.path.join(deck_dir, "content.json"), "w") as f:
        json.dump(content, f, ensure_ascii=False)

    return total


def apply_font_replacements(pptx_path, replacements, weight_map=None):
    """替换 PPTX XML 中的字体名称，可选调整粗细"""
    import zipfile, shutil
    from lxml import etree

    tmp = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path, "r") as zin, \
         zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                text = data.decode("utf-8")
                for old, new in replacements.items():
                    text = text.replace(f'typeface="{old}"', f'typeface="{new}"')
                    text = text.replace(f'val="{old}"', f'val="{new}"')

                if weight_map and item.filename.startswith("ppt/slides/"):
                    try:
                        root = etree.fromstring(text.encode("utf-8"))
                        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                        for orig_font, weight in weight_map.items():
                            new_font = replacements.get(orig_font, orig_font)
                            w = int(weight)
                            for rpr in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"):
                                latin = rpr.find("a:latin", ns)
                                if latin is not None and latin.get("typeface") == new_font:
                                    if w >= 600:
                                        rpr.set("b", "1")
                                    else:
                                        rpr.attrib.pop("b", None)
                        text = etree.tostring(root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True).decode("utf-8")
                    except Exception:
                        pass

                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, pptx_path)


# ============================================================
# LLM — MiniMax (OpenAI 兼容)
# ============================================================
def chat_with_llm(messages):
    """调用 MiniMax M2.5 highspeed，返回文本回复"""
    import urllib.request
    url = "https://api.minimax.chat/v1/text/chatcompletion_v2"
    payload = json.dumps({
        "model": MINIMAX_MODEL,
        "messages": messages,
        "max_tokens": 80,
    })
    req = urllib.request.Request(
        url,
        data=payload.encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {MINIMAX_API_KEY}",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read()
    data = json.loads(raw)
    status = data.get("base_resp", {}).get("status_code")
    print(f"[LLM] status={status}, base_resp={data.get('base_resp')}", flush=True)

    if status and status != 0:
        status_msg = data.get("base_resp", {}).get("status_msg", "unknown error")
        print(f"[LLM] API error: {status_msg}", flush=True)
        return f"Sorry, the AI service returned an error. ({status_msg})"

    msg = data.get("choices", [{}])[0].get("message", {})
    content = msg.get("content", "")
    # MiniMax reasoning model 有时 content 为空，用 reasoning_content 兜底
    if not content:
        reasoning = msg.get("reasoning_content", "")
        if reasoning:
            content = reasoning
        else:
            content = "I'm not sure how to respond to that."
    print(f"[LLM] content present: {bool(content)}, len={len(content)}", flush=True)
    return content



# ============================================================
# 加载 deck 内容
# ============================================================
def load_deck_content(deck_id):
    deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
    content_path = os.path.join(deck_dir, "content.json")
    if os.path.exists(content_path):
        with open(content_path) as f:
            return json.load(f)
    return []


def get_slide_title(sc):
    """提取 slide 标题 (第一个文字元素)"""
    return sc["text"][0] if sc.get("text") else f"Slide {sc['slide']}"


def build_system_prompt(all_content, current_slide_idx):
    """构建 system prompt: 标题索引 + 当前 slide 详情"""

    # ---- 标题索引 (简洁、稳定) ----
    title_index = ""
    for sc in all_content:
        title = get_slide_title(sc)
        title_index += f"  Slide {sc['slide']}: {title}\n"

    # ---- 当前 slide 完整内容 ----
    current_detail = ""
    if 0 <= current_slide_idx < len(all_content):
        sc = all_content[current_slide_idx]
        current_detail = f"Slide {sc['slide']} — {get_slide_title(sc)}:\n"
        current_detail += "\n".join(sc["text"])
        if sc.get("notes"):
            current_detail += f"\nSpeaker notes: {sc['notes']}"

    # ---- 动态生成导航示例 (从实际标题中选 2-3 个) ----
    nav_examples = ""
    example_slides = []
    for sc in all_content:
        t = get_slide_title(sc)
        if t and t.lower() not in ("title slide", "thank you", "questions", "end", ""):
            example_slides.append((sc["slide"], t))
    # 选最多 3 个作为具体示例
    for idx, (snum, stitle) in enumerate(example_slides[:3]):
        keyword = stitle.split()[0].lower() if stitle.split() else stitle.lower()
        nav_examples += f'- User says "{keyword}" → you write [GO:{snum}] because Slide {snum} is "{stitle}"\n'
    nav_examples += '- User says "next" → you write [GO:next]\n'
    nav_examples += '- User says "go back" → you write [GO:prev]\n'

    return f"""You explain slide content directly — just talk about the topic itself. Don't refer to "the company", "they", or "we" unless the slide content naturally calls for it. Just explain the actual ideas, facts, and concepts on the slides like you're teaching a friend.

Spoken voice only — no markdown, no bullets, no formatting.

VIBE: Casual like "So basically...", "Oh nice,", "Yeah so the idea here is...". NEVER say "this presentation", "this deck", "this slide says".

[GREET]: Under 40 words. One sentence summary of the topic, name 4-5 main themes from the slide titles, end with "Where do you wanna start?" Do NOT include [GO:N] in greeting.

REPLIES: MAX 30 WORDS. One short sentence to answer, one short question back. That's it. Never ramble.

Viewer is on Slide {current_slide_idx + 1} of {len(all_content)}.

SLIDE TITLES:
{title_index}
CURRENT SLIDE:
{current_detail}

===== NAVIGATION RULES (MANDATORY — READ CAREFULLY) =====
When the user mentions ANY topic, keyword, or phrase that relates to a slide title, you MUST include [GO:N] in your reply (N = slide number).

EXAMPLES from this deck:
{nav_examples}
HOW IT WORKS: You write [GO:N] anywhere in your reply. The system removes it before showing to user and auto-jumps the slide. The user never sees [GO:N].

RULES:
1. Match loosely — if user says ANY word from a slide title, navigate there.
2. NEVER ask "want me to go there?" — just include [GO:N] and describe the content.
3. "next" → [GO:next], "back"/"previous" → [GO:prev]
4. If user asks about a topic and you DON'T include [GO:N], your response is WRONG.
5. Always include [GO:N] BEFORE your spoken text, like: [GO:3] So this one covers...
================================================"""


# ---- 每个 session 的对话历史 ----
_chat_histories = {}


# ============================================================
# HTTP API — 对话式语音导览 (比 WebSocket 更稳定)
# ============================================================
@app.route("/api/talk/<deck_id>", methods=["POST"])
def api_talk(deck_id):
    """用户说话 → LLM 对话 → TTS → 返回 JSON {text, audio, nav}"""
    data = request.json
    text = data.get("text", "").strip()
    slide_idx = data.get("slide_index", 0)
    voice = data.get("voice", "af_heart")
    session_id = data.get("session_id", "default")

    print(f"[TALK] text={text}, slide={slide_idx}, voice={voice}", flush=True)

    if not text:
        return jsonify({"text": "", "audio": "", "nav": None})

    all_content = load_deck_content(deck_id)

    # 如果 content.json 不存在，尝试从 PPTX 提取
    if not all_content:
        deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
        pptx_path = os.path.join(deck_dir, "deck.pptx")
        if os.path.exists(pptx_path):
            all_content = extract_slide_content(pptx_path)
            with open(os.path.join(deck_dir, "content.json"), "w") as f:
                json.dump(all_content, f, ensure_ascii=False)
            print(f"[TALK] extracted {len(all_content)} slides content", flush=True)
    system_prompt = build_system_prompt(all_content, slide_idx)

    # ---- 对话历史 ----
    history_key = f"{deck_id}:{session_id}"
    if history_key not in _chat_histories:
        _chat_histories[history_key] = []
    history = _chat_histories[history_key]

    if len(history) > 40:
        history = history[-40:]
        _chat_histories[history_key] = history

    history.append({"role": "user", "content": text})
    messages = [{"role": "system", "content": system_prompt}] + history

    # ---- LLM ----
    try:
        reply = chat_with_llm(messages)
        print(f"[LLM] {reply[:120]}", flush=True)
    except Exception as e:
        print(f"[LLM ERROR] {e}", flush=True)
        reply = "Sorry, let me try again."

    history.append({"role": "assistant", "content": reply})

    # ---- 解析导航 ----
    nav_command = None
    nav_match = re.search(r'\[GO:(\w+)\]', reply)
    if nav_match:
        nav_command = nav_match.group(1)
        reply = re.sub(r'\s*\[GO:\w+\]\s*', '', reply).strip()

    return jsonify({"text": reply, "nav": nav_command})


@app.route("/api/tts", methods=["POST"])
def api_tts():
    """ElevenLabs streaming TTS — 返回 audio/mpeg 流"""
    data = request.json
    text = data.get("text", "").strip()
    if not text:
        return jsonify({"audio": ""})

    print(f"[TTS] ElevenLabs request: {len(text)} chars", flush=True)

    try:
        url = f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}/stream"
        payload = json.dumps({
            "text": text,
            "model_id": "eleven_turbo_v2_5",
            "voice_settings": {
                "stability": 0.5,
                "similarity_boost": 0.75,
            },
        })
        req = urllib.request.Request(
            url,
            data=payload.encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "xi-api-key": ELEVENLABS_API_KEY,
                "Accept": "audio/mpeg",
            },
        )

        with urllib.request.urlopen(req, timeout=15) as resp:
            audio_bytes = resp.read()

        audio_b64 = base64.b64encode(audio_bytes).decode()
        print(f"[TTS] ElevenLabs OK: {len(audio_bytes)} bytes", flush=True)
        return jsonify({"audio": audio_b64})
    except Exception as e:
        print(f"[TTS ERROR] {e}", flush=True)
        return jsonify({"audio": ""})


# ============================================================
# HTTP 路由
# ============================================================
@app.route("/debug/env")
def debug_env():
    """临时调试端点 — 部署成功后删除"""
    return jsonify({
        "MINIMAX_API_KEY": "SET" if os.environ.get("MINIMAX_API_KEY") else "MISSING",
        "ELEVENLABS_API_KEY": "SET" if os.environ.get("ELEVENLABS_API_KEY") else "MISSING",
        "ELEVENLABS_VOICE_ID": os.environ.get("ELEVENLABS_VOICE_ID", "MISSING"),
    })


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".pptx"):
        return "Please upload a .pptx file", 400

    deck_id = uuid.uuid4().hex[:10]
    deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
    os.makedirs(deck_dir, exist_ok=True)

    pptx_path = os.path.join(deck_dir, "deck.pptx")
    file.save(pptx_path)

    try:
        font_report = ensure_fonts(pptx_path)
        with open(os.path.join(deck_dir, "fonts.txt"), "w") as f:
            for cat in ("found", "installed", "missing"):
                for name in font_report[cat]:
                    f.write(f"{cat}: {name}\n")

        convert_deck(deck_dir)
        return redirect(url_for("viewer", deck_id=deck_id))
    except Exception as e:
        print(f"[UPLOAD ERROR] {e}", flush=True)
        return f"Conversion failed: {e}", 500


@app.route("/view/<deck_id>")
def viewer(deck_id):
    deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
    meta_path = os.path.join(deck_dir, "meta.txt")
    if not os.path.exists(meta_path):
        return "Deck not found", 404
    with open(meta_path) as f:
        total = int(f.read().strip())

    font_report = {"found": [], "installed": [], "missing": []}
    fonts_path = os.path.join(deck_dir, "fonts.txt")
    if os.path.exists(fonts_path):
        with open(fonts_path) as f:
            for line in f:
                cat, name = line.strip().split(": ", 1)
                font_report[cat].append(name)

    system_fonts = get_system_fonts() if font_report["missing"] else []

    # 加载 slide 内容
    content = []
    content_path = os.path.join(deck_dir, "content.json")
    if os.path.exists(content_path):
        with open(content_path) as f:
            content = json.load(f)

    return render_template(
        "viewer.html",
        deck_id=deck_id,
        total=total,
        font_report=font_report,
        system_fonts=system_fonts,
        slide_content=content,
    )


@app.route("/reconvert/<deck_id>", methods=["POST"])
def reconvert(deck_id):
    deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
    pptx_path = os.path.join(deck_dir, "deck.pptx")
    if not os.path.exists(pptx_path):
        return jsonify({"error": "Deck not found"}), 404

    try:
        replacements = request.json.get("replacements", {})
        weight_map = request.json.get("weights", {})
        if replacements:
            for new_font in replacements.values():
                if not is_font_installed(new_font):
                    try_install_font(new_font)
            subprocess.run(["fc-cache", "-f"], capture_output=True)
            apply_font_replacements(pptx_path, replacements, weight_map)

        convert_deck(deck_dir)

        report = ensure_fonts(pptx_path)
        with open(os.path.join(deck_dir, "fonts.txt"), "w") as f:
            for cat in ("found", "installed", "missing"):
                for name in report[cat]:
                    f.write(f"{cat}: {name}\n")

        return jsonify({"ok": True})
    except Exception as e:
        print(f"[RECONVERT ERROR] {e}", flush=True)
        return jsonify({"error": str(e)}), 500


@app.route("/download/<deck_id>")
def download(deck_id):
    deck_dir = os.path.join(app.config["OUTPUT_FOLDER"], deck_id)
    meta_path = os.path.join(deck_dir, "meta.txt")
    if not os.path.exists(meta_path):
        return "Deck not found", 404
    with open(meta_path) as f:
        total = int(f.read().strip())

    # slide 图片 base64
    slides_data = []
    for i in range(1, total + 1):
        with open(os.path.join(deck_dir, f"slide-{i}.png"), "rb") as f:
            slides_data.append(base64.b64encode(f.read()).decode())

    # slide 内容
    content = []
    content_path = os.path.join(deck_dir, "content.json")
    if os.path.exists(content_path):
        with open(content_path) as f:
            content = json.load(f)

    # 获取当前服务器地址
    server_url = request.host_url.rstrip("/")

    return render_template(
        "download.html",
        deck_id=deck_id,
        total=total,
        slides_data=slides_data,
        slide_content=content,
        server_url=server_url,
    )


@app.route("/slides/<deck_id>/<filename>")
def slide_image(deck_id, filename):
    return send_from_directory(os.path.join(app.config["OUTPUT_FOLDER"], deck_id), filename)


os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

if __name__ == "__main__":
    socketio.run(app, debug=False, port=8080, host="0.0.0.0", allow_unsafe_werkzeug=True)
