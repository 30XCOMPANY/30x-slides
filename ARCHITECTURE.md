# 30X Slides — Architecture & Roadmap

## What It Does

Upload a PPTX, get a self-contained HTML file with embedded Voice Bot.
The bot guides viewers through slides — narrating, answering questions, controlling navigation.

## Pipeline

```
PPTX Upload
    │
    ├── LibreOffice (--headless)  →  PDF
    │                                  │
    │                           pymupdf (fitz)  →  PNG slides (2881×1620)
    │
    ├── python-pptx + lxml  →  Extract per-slide text & speaker notes
    │
    ├── fontsource CDN  →  Auto-install missing font weights (100-900)
    │
    └── Flask assembles everything into:
            │
            ▼
    ┌─────────────────────────────────────┐
    │  Single HTML File (downloadable)    │
    │                                     │
    │  ┌───────────────────────────────┐  │
    │  │  Slides (base64 PNG)          │  │
    │  └───────────────────────────────┘  │
    │  ┌───────────────────────────────┐  │
    │  │  Slide Content (JSON)         │  │
    │  │  - per-slide text             │  │
    │  │  - speaker notes              │  │
    │  └───────────────────────────────┘  │
    │  ┌───────────────────────────────┐  │
    │  │  Voice Bot (JS)               │  │
    │  │  - narration (TTS)            │  │
    │  │  - slide navigation           │  │
    │  │  - Q&A about slide content    │  │
    │  └───────────────────────────────┘  │
    └─────────────────────────────────────┘
```

## Tech Stack

| Layer      | Tool              | Why                                    |
|------------|-------------------|----------------------------------------|
| PPTX → PDF | LibreOffice       | Free, cross-platform, headless, stable |
| PDF → PNG  | pymupdf           | Fast C library, exact pixel control    |
| Text extraction | python-pptx + lxml | Reads shapes, notes, theme fonts  |
| Font install | fontsource CDN   | 18 weight variants per font family     |
| Web server | Flask             | Simple, Python-native                  |
| Voice TTS  | TBD               | See options below                      |
| Voice LLM  | TBD               | See options below                      |

## Voice Bot Options

### Option A: Fully Self-Contained (Offline)
- **TTS**: Web Speech API (browser built-in)
- **Navigation**: JS controls, no LLM needed
- **Content**: Pre-extracted slide text baked into HTML
- **Pros**: Zero cost, works offline, no API keys in HTML
- **Cons**: Robotic voice, no conversational Q&A

### Option B: Server-Connected (Online)
- **TTS**: ElevenLabs / OpenAI TTS / other
- **LLM**: Claude / Qwen (Bailian) / other for Q&A
- **Flow**: HTML calls back to 30X Slides server for voice + AI
- **Pros**: Natural voice, can answer questions about slides
- **Cons**: Needs internet, API costs, server must stay running

### Option C: Hybrid (Recommended)
- Narration script pre-generated at upload time (LLM writes script from slide content)
- TTS audio pre-rendered and embedded as base64 in HTML
- Result: natural voice, fully offline, zero runtime API calls
- Q&A: optional, connects to server only if user asks questions

## File Structure

```
30x-slides/
├── app.py                 # Flask app, orchestrates pipeline
├── templates/
│   ├── index.html         # Upload page (30X glass design)
│   └── viewer.html        # Slide viewer + font panel
├── static/
│   └── hero-bg.mp4        # Upload page background video
└── output/
    └── {deck_id}/
        ├── deck.pptx      # Uploaded file
        ├── deck.pdf        # LibreOffice output
        ├── slide-1.png     # Per-slide renders
        ├── slide-2.png
        ├── meta.txt        # Slide count
        ├── fonts.txt       # Font detection report
        └── content.json    # Per-slide text + notes (TODO)
```

## Implementation Order

### Phase 1: Stable Core (DONE)
- [x] PPTX upload + conversion pipeline
- [x] Font auto-detection + auto-install (18 weight variants)
- [x] Font replacement UI (glass modal + weight dropdown)
- [x] Slide viewer with 30X design system
- [x] Self-contained HTML download (base64 PNG)

### Phase 2: Content Extraction (NEXT)
- [ ] Extract per-slide text content (shapes + paragraphs)
- [ ] Extract speaker notes
- [ ] Save as content.json per deck
- [ ] Embed content.json in downloadable HTML

### Phase 3: Voice Bot
- [ ] Decide TTS approach (A/B/C above)
- [ ] Build voice bot UI component (mic button, transcript)
- [ ] Implement narration (bot reads slide content aloud)
- [ ] Implement navigation (bot controls slide transitions)
- [ ] Embed voice bot JS in downloadable HTML

### Phase 4: Conversational Q&A (Optional)
- [ ] LLM integration for answering viewer questions
- [ ] Context: slide content + speaker notes
- [ ] "What does this slide mean?" / "Go back to the chart"
