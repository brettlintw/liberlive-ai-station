# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a single-file Streamlit web app (`app.py`) called **Liberlive AI Station** — a live-performance guitar chord/lyrics chart tool. It lets a performer pull up a song's lyrics+chords, auto-transpose them into a target key, render them with colored chord tags above the lyrics, and save/reload charts to local disk so they can be recalled instantly during a gig.

There is no separate backend/frontend split — Streamlit handles both UI and server logic in this one script, driven entirely through `st.session_state`.

## Commands

```bash
pip install -r requirements.txt   # streamlit, requests, beautifulsoup4, docx2txt, python-docx, Pillow
streamlit run app.py              # start the app (default: http://localhost:8501)
```

There is no lint config, no test suite, and no build step in this repo. Verify changes by running the app and exercising the flow manually (load a song from the menu → edit chords → transpose → check the rendered chart in "演出模式").

## Architecture

`app.py` is organized into numbered comment blocks (`# --- 1. ... ---` through `# --- 10. ... ---`), each a distinct concern (weather/date header, song menu DB, session init, transpose engine, CSS, sidebar, metadata panel, the three tabs). **When discussing or making changes, refer to the relevant numbered section** rather than the whole file — this mirrors how the project owner (Brett) has directed prior edits in `prompt-liberlive專案.docx`, and keeps diffs scoped to one concern.

Key mechanics:

- **Chord transposition** (`transpose_engine`, section 6): parses `[Chord]` tokens embedded inline in lyric text, normalizes flats to sharps, shifts by semitone steps through the fixed `KEYS` list, and rejoins. Slash chords (e.g. `G/B`) are transposed on both sides.
- **Chord coloring** (`COLOR_MAP`, section 1, applied in section 10's rendering loop): color is keyed off the **root letter only** — for a slash chord like `G/B`, only `G` (the part before the slash) determines the color, per the original spec in `prompt-liberlive專案.docx`.
- **Rendering model** (tab_play in section 10): each lyric line is split on `\[[^\]]+\]` into chord/text tokens, then re-emitted **character-by-character** as `<div class="char-unit">` pairs (chord tag stacked above one character) so that chord labels stay pinned above the exact syllable they modify, independent of font size or line wrapping. This char-by-char DOM approach is intentional — collapsing it into per-word spans breaks vertical chord alignment.
- **Persistence** (tab_cloud, section... "雲端/實體曲庫"): charts are saved as `.txt` files under `liberlive_saved_tracks/` (gitignored) using a simple `KEY:value` metadata header, a `---` separator, then raw chord/lyric text. Loading parses that same format back into `st.session_state`.
- **Lyrics fetch** (`fetch_lyrics_v18_core`, section 6): scrapes a given URL with BeautifulSoup looking for `div.chord-content`, `<pre>`, or `.post-content`. Per the spec doc, lyrics must come from real third-party sources, not be AI-generated — this function has no LLM involved by design.
- **Session state contract**: nearly all UI state lives in `st.session_state` (`buffer` = current chord/lyric text, `meta` = song metadata dict, `editor_main` = the text area's bound key). Widgets read/write these directly; most action buttons end with `st.rerun()` to force the render loop to pick up the new buffer/meta.

## Important context / gaps to be aware of

- `Liberlive_AI_Station_v71.0_Features.docx` describes an aspirational "v71.0" feature set (AI-powered web search/grounding for chord lookup, paid/free API key fallback, PDF export, history log, auto section-labeling like intro/chorus with strum/drum cues). The actual `app.py` is versioned **v25.0** and does not implement most of this — image "OCR" input is a hardcoded mock string, there is no LLM/API-key integration, and there is no PDF export or history log yet. Don't assume features described in that docx exist in the code; verify against `app.py` directly.
- `prompt-liberlive專案.docx` is the original requirements doc from the project owner and is the best source of *intent* behind design choices (chord color-by-scale-degree mapping, slash-chord root-color rule, no-AI-generated-lyrics rule, section-index-based communication convention above).
- `Brett_Score_C.txt` is just a sample chord/lyric chart (七里香) usable for manual testing of parsing/transposition/rendering.
- `C2.png` is a UI layout reference image mentioned in the prompt doc's visual-layout section.
