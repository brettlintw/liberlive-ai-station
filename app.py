import streamlit as st
import re
import requests
import os
import base64
from datetime import datetime
from bs4 import BeautifulSoup
from docx import Document
from urllib.parse import quote
try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# --- 1. 核心樂理與 Liberlive 官方風格配色 ---
KEYS = ['C', 'C#', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B']
COLOR_MAP = {
    'C': '#EF4444',   # 紅，一級
    'D': '#F97316',   # 橘，二級
    'E': '#FACC15',   # 黃，三級
    'F': '#22C55E',   # 綠，四級
    'G': '#06B6D4',   # 藍，五級
    'A': '#1D4ED8',   # 深藍，六級
    'B': '#A855F7'    # 紫，七級
}

st.set_page_config(
    page_title="Liberlive AI Station v26.0",
    page_icon="liberlive_icon.jpg",
    layout="wide",
    initial_sidebar_state="auto"
)

# --- 2. 實體硬碟儲存目錄 ---
STORAGE_DIR = "liberlive_saved_tracks"
if not os.path.exists(STORAGE_DIR):
    os.makedirs(STORAGE_DIR)

# --- 3. 免金鑰自動天氣與日期獲取 ---
def get_weather_and_date():
    week_days = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
    now = datetime.now()
    date_str = f"{now.strftime('%Y-%m-%d')} (星期{week_days[now.weekday()]})"
    weather_str = "🌤️ 多雲 25°C"
    try:
        res = requests.get("https://wttr.in/?format=%c+%t", timeout=3)
        if res.status_code == 200 and res.text:
            weather_str = res.text.strip().replace("+", " ")
    except Exception:
        pass
    return date_str, weather_str

current_date, current_weather = get_weather_and_date()

# --- 4. 內建中英文分類歌單資料庫 ---
SONG_MENU = {
    "中文": {
        "男歌手": {
            "晴天 - 周杰倫": {"orig": "G", "bpm": 65, "beat": "4/4", "singer": "周杰倫", "song": "晴天",
                "chords": "[G]故事的小黃[Bm]花 從出生那年[Em]就飄著\n[C]童年的蕩秋[G]千 隨記憶一直[Am]晃到現[D]在"},
            "溫柔 - 五月天": {"orig": "D", "bpm": 76, "beat": "4/4", "singer": "五月天", "song": "溫柔",
                "chords": "[D]走在風中 [F#m]今天陽光 [G]突然好溫[A]柔\n[D]天的溫柔 [F#m]地的溫柔 [G]讓你好想[A]放開手"},
        },
        "女歌手": {
            "隱形的翅膀 - 張韶涵": {"orig": "C", "bpm": 82, "beat": "4/4", "singer": "張韶涵", "song": "隱形的翅膀",
                "chords": "[C]每一次 [G]都在徘徊[Am]孤單中[Em]堅強\n[F]每一次 [C]就算受傷[Dm]也不閃[G]淚光"},
            "寶貝 - 張懸": {"orig": "G", "bpm": 88, "beat": "4/4", "singer": "張懸", "song": "寶貝",
                "chords": "[G]我的寶貝[C]寶貝 給你[D]一點甜[G]甜\n[G]讓你今夜[C]都好[D]眠"},
        },
        "樂團": {
            "小情歌 - 蘇打綠": {"orig": "D", "bpm": 74, "beat": "4/4", "singer": "蘇打綠", "song": "小情歌",
                "chords": "[D]這是一首[F#m]簡單的小[G]情歌\n[G]唱著我們[A]心頭的白[D]鴿"}
        }
    },
    "英文": {
        "男歌手": {
            "Just the Way You Are - Bruno Mars": {"orig": "F", "bpm": 109, "beat": "4/4", "singer": "Bruno Mars", "song": "Just the Way You Are",
                "chords": "Oh [F]her eyes her eyes make the stars look like they're not [Dm]shining\nHer [Bb]hair falls perfectly without her [F]trying"},
            "Shape of You - Ed Sheeran": {"orig": "C#", "bpm": 96, "beat": "4/4", "singer": "Ed Sheeran", "song": "Shape of You",
                "chords": "[C#m]The club isn't the best [F#m]place to find a lover so the [A]bar is where I [B]go"}
        },
        "女歌手": {
            "Someone Like You - Adele": {"orig": "A", "bpm": 67, "beat": "4/4", "singer": "Adele", "song": "Someone Like You",
                "chords": "I [A]heard that you're [C#m]settled down\nThat you [F#m]found a girl and you're [D]married now"}
        },
        "樂團": {
            "Yellow - Coldplay": {"orig": "B", "bpm": 88, "beat": "4/4", "singer": "Coldplay", "song": "Yellow",
                "chords": "Look at the [B]stars look how they shine for [F#]you\nAnd everything you [E]do yeah they were all [B]yellow"}
        }
    }
}

# --- 5. 初始化 Session 緩存 ---
GEMINI_MODELS = [
    ("gemini-2.5-flash",  "gemini-2.5-flash  ｜ 🆓 免費（每日限量）"),
    ("gemini-2.5-pro",    "gemini-2.5-pro    ｜ 💳 付費"),
    ("gemini-2.0-flash",  "gemini-2.0-flash  ｜ 🆓 免費（每日限量）"),
    ("gemini-1.5-flash",  "gemini-1.5-flash  ｜ 🆓 免費（每日限量）"),
    ("gemini-1.5-pro",    "gemini-1.5-pro    ｜ 💳 付費"),
]
GEMINI_MODEL_IDS   = [m[0] for m in GEMINI_MODELS]
GEMINI_MODEL_LABELS = [m[1] for m in GEMINI_MODELS]

# --- 多國語言字串 ---
LANG_STRINGS = {
    "繁體中文": {
        "lang_label": "🌐 語言 / Language",
        "gemini_title": "### 🤖 Gemini 設定",
        "api_ok": "API Key 已設定 ✅",
        "api_warn": "請在「智能尋譜」欄位輸入 Key",
        "model_pick": "🔧 選擇模型",
        "yt_sync": "### 🎬 YouTube 同步播放",
        "yt_url_label": "YouTube 網址",
        "chord_font": "🎸 和弦字體 (px)",
        "lyric_font": "🎤 歌詞字體 (px)",
        "scroll_spd": "📜 自動捲動速度",
        "meta_song": "🎵 歌曲名稱",
        "meta_singer": "🎤 歌手",
        "meta_lyricist": "✍️ 作詞",
        "meta_composer": "🎼 作曲",
        "meta_orig": "🎸 原調",
        "meta_target": "🎯 目標調",
        "meta_bpm": "⏱️ BPM",
        "meta_beat": "🥁 拍號",
        "tab_search": "🔍 智能尋譜",
        "tab_play": "🎤 演出模式",
        "tab_cloud": "📂 曲庫管理",
        "ai_card": "🤖 AI 智能尋譜",
        "ai_song_label": "🎵 歌曲名稱",
        "ai_singer_label": "🎤 歌手（選填）",
        "gemini_key_label": "🔑 Gemini Key（必填）",
        "search_btn": "🔍 搜尋和弦譜",
        "search_warn": "請輸入歌曲名稱",
        "search_spinner": "正在搜尋",
        "search_fail": "❌ 找不到曲譜，請手動貼入或確認 Gemini Key",
        "yt_card": "🎬 YouTube 識別歌名",
        "yt_placeholder": "https://youtube.com/watch?v=...",
        "yt_btn": "🎵 從 YouTube 識別",
        "yt_spinner": "識別中...",
        "yt_done": "識別完成：{singer} - {song}，請按「搜尋和弦譜」",
        "file_card": "📄 檔案匯入",
        "file_hint": "支援：txt / docx / pdf / jpg / png",
        "file_upload_label": "上傳檔案",
        "file_spinner": "處理 {name} 中...",
        "file_ok": "✅ {name} 匯入完成",
        "file_empty": "檔案內容為空，請確認格式正確",
        "no_song": "（未設定歌名）",
        "editor_label": "✍️ 原始和弦歌詞（[和弦]歌詞 格式）",
        "apply_btn": "🎸 套用變調並更新",
        "cand_source": "來源：{source}",
        "cand_singer": "🎤 歌手：{singer}　🎸 原調：{orig}　⏱️ BPM：{bpm}",
        "cand_preview": "📄 歌詞預覽：",
        "cand_pick": "✅ 使用這個",
        "cand_cancel": "❌ 取消搜尋結果",
        "cand_credit_unknown": "（作詞/作曲資訊不詳）",
        "lyricist_label": "作詞：",
        "composer_label": "作曲：",
        "fullscreen_btn": "⛶ 全螢幕",
        "exit_fullscreen": "✕ 退出全螢幕",
        "no_song_info": "請先在「智能尋譜」頁籤搜尋或載入歌曲",
        "save_title": "#### 💾 儲存目前曲目",
        "save_warn": "請先設定歌曲名稱",
        "save_btn": "💾 儲存到本機曲庫",
        "save_ok": "✅ 已儲存：{name}",
        "save_fail": "儲存失敗：{err}",
        "lib_title": "#### 📂 本機曲庫",
        "lib_empty": "目前沒有任何存檔，按上方「儲存」新增第一首",
        "lib_count": "共 {n} 首歌",
        "load_btn": "📖 載入",
        "del_btn": "🗑️",
        "del_done": "已刪除《{song}》",
        "lyricist_short": "✍️",
        "composer_short": "🎼",
        "orig_label": "原調：",
        "target_label": "目標調：",
        "bpm_label": "BPM：",
        "beat_label": "拍號：",
        "lyricist_disp": "作詞：{lyr}　作曲：{cmp}",
        "meta_caption": "原調：{orig} → 目標調：{target}　BPM：{bpm}　拍号：{beat}",
        "cur_song_label": "目前曲目：《{song}》",
        "credit_both": "✍️ 作詞：{lyr}　🎼 作曲：{comp}",
        "credit_lyr": "✍️ 詞曲：{lyr}",
        "credit_comp": "🎼 作曲：{comp}",
    },
    "简体中文": {
        "lang_label": "🌐 语言 / Language",
        "gemini_title": "### 🤖 Gemini 设置",
        "api_ok": "API Key 已设置 ✅",
        "api_warn": "请在「智能寻谱」栏位输入 Key",
        "model_pick": "🔧 选择模型",
        "yt_sync": "### 🎬 YouTube 同步播放",
        "yt_url_label": "YouTube 网址",
        "chord_font": "🎸 和弦字体 (px)",
        "lyric_font": "🎤 歌词字体 (px)",
        "scroll_spd": "📜 自动滚动速度",
        "meta_song": "🎵 歌曲名称",
        "meta_singer": "🎤 歌手",
        "meta_lyricist": "✍️ 作词",
        "meta_composer": "🎼 作曲",
        "meta_orig": "🎸 原调",
        "meta_target": "🎯 目标调",
        "meta_bpm": "⏱️ BPM",
        "meta_beat": "🥁 拍号",
        "tab_search": "🔍 智能寻谱",
        "tab_play": "🎤 演出模式",
        "tab_cloud": "📂 曲库管理",
        "ai_card": "🤖 AI 智能寻谱",
        "ai_song_label": "🎵 歌曲名称",
        "ai_singer_label": "🎤 歌手（选填）",
        "gemini_key_label": "🔑 Gemini Key（必填）",
        "search_btn": "🔍 搜索和弦谱",
        "search_warn": "请输入歌曲名称",
        "search_spinner": "正在搜索",
        "search_fail": "❌ 找不到曲谱，请手动粘入或确认 Gemini Key",
        "yt_card": "🎬 YouTube 识别歌名",
        "yt_placeholder": "https://youtube.com/watch?v=...",
        "yt_btn": "🎵 从 YouTube 识别",
        "yt_spinner": "识别中...",
        "yt_done": "识别完成：{singer} - {song}，请按「搜索和弦谱」",
        "file_card": "📄 文件导入",
        "file_hint": "支持：txt / docx / pdf / jpg / png",
        "file_upload_label": "上传文件",
        "file_spinner": "处理 {name} 中...",
        "file_ok": "✅ {name} 导入完成",
        "file_empty": "文件内容为空，请确认格式正确",
        "no_song": "（未设定歌名）",
        "editor_label": "✍️ 原始和弦歌词（[和弦]歌词 格式）",
        "apply_btn": "🎸 套用变调并更新",
        "cand_source": "来源：{source}",
        "cand_singer": "🎤 歌手：{singer}　🎸 原调：{orig}　⏱️ BPM：{bpm}",
        "cand_preview": "📄 歌词预览：",
        "cand_pick": "✅ 使用此项",
        "cand_cancel": "❌ 取消搜索结果",
        "cand_credit_unknown": "（作词/作曲信息不详）",
        "lyricist_label": "作词：",
        "composer_label": "作曲：",
        "fullscreen_btn": "⛶ 全屏",
        "exit_fullscreen": "✕ 退出全屏",
        "no_song_info": "请先在「智能寻谱」标签搜索或载入歌曲",
        "save_title": "#### 💾 保存当前曲目",
        "save_warn": "请先设定歌曲名称",
        "save_btn": "💾 保存到本地曲库",
        "save_ok": "✅ 已保存：{name}",
        "save_fail": "保存失败：{err}",
        "lib_title": "#### 📂 本地曲库",
        "lib_empty": "暂无存档，点击上方「保存」添加第一首",
        "lib_count": "共 {n} 首歌",
        "load_btn": "📖 载入",
        "del_btn": "🗑️",
        "del_done": "已删除《{song}》",
        "lyricist_short": "✍️",
        "composer_short": "🎼",
        "orig_label": "原调：",
        "target_label": "目标调：",
        "bpm_label": "BPM：",
        "beat_label": "拍号：",
        "lyricist_disp": "作词：{lyr}　作曲：{cmp}",
        "meta_caption": "原调：{orig} → 目标调：{target}　BPM：{bpm}　拍号：{beat}",
        "cur_song_label": "当前曲目：《{song}》",
        "credit_both": "✍️ 作词：{lyr}　🎼 作曲：{comp}",
        "credit_lyr": "✍️ 词曲：{lyr}",
        "credit_comp": "🎼 作曲：{comp}",
    },
    "English": {
        "lang_label": "🌐 Language",
        "gemini_title": "### 🤖 Gemini Settings",
        "api_ok": "API Key configured ✅",
        "api_warn": "Enter your Key in the Search tab",
        "model_pick": "🔧 Select model",
        "yt_sync": "### 🎬 YouTube Sync",
        "yt_url_label": "YouTube URL",
        "chord_font": "🎸 Chord font (px)",
        "lyric_font": "🎤 Lyric font (px)",
        "scroll_spd": "📜 Auto-scroll speed",
        "meta_song": "🎵 Song title",
        "meta_singer": "🎤 Artist",
        "meta_lyricist": "✍️ Lyricist",
        "meta_composer": "🎼 Composer",
        "meta_orig": "🎸 Original key",
        "meta_target": "🎯 Target key",
        "meta_bpm": "⏱️ BPM",
        "meta_beat": "🥁 Time sig",
        "tab_search": "🔍 Search Chords",
        "tab_play": "🎤 Performance",
        "tab_cloud": "📂 Song Library",
        "ai_card": "🤖 AI Chord Search",
        "ai_song_label": "🎵 Song title",
        "ai_singer_label": "🎤 Artist (optional)",
        "gemini_key_label": "🔑 Gemini Key (required)",
        "search_btn": "🔍 Search Charts",
        "search_warn": "Please enter a song title",
        "search_spinner": "Searching",
        "search_fail": "❌ No chart found. Paste manually or check your Gemini Key",
        "yt_card": "🎬 Identify Song via YouTube",
        "yt_placeholder": "https://youtube.com/watch?v=...",
        "yt_btn": "🎵 Identify from YouTube",
        "yt_spinner": "Identifying...",
        "yt_done": "Identified: {singer} - {song}. Now click Search Charts.",
        "file_card": "📄 Import File",
        "file_hint": "Supports: txt / docx / pdf / jpg / png",
        "file_upload_label": "Upload file",
        "file_spinner": "Processing {name}...",
        "file_ok": "✅ {name} imported",
        "file_empty": "File is empty. Please check the format.",
        "no_song": "(no song set)",
        "editor_label": "✍️ Chord/Lyric text ([Chord]lyric format)",
        "apply_btn": "🎸 Apply Transpose & Update",
        "cand_source": "Source: {source}",
        "cand_singer": "🎤 Artist: {singer}　🎸 Key: {orig}　⏱️ BPM: {bpm}",
        "cand_preview": "📄 Lyric preview: ",
        "cand_pick": "✅ Use this",
        "cand_cancel": "❌ Cancel results",
        "cand_credit_unknown": "(Lyricist/Composer unknown)",
        "lyricist_label": "Lyricist: ",
        "composer_label": "Composer: ",
        "fullscreen_btn": "⛶ Fullscreen",
        "exit_fullscreen": "✕ Exit Fullscreen",
        "no_song_info": "Search or load a song in the Search tab first",
        "save_title": "#### 💾 Save Current Song",
        "save_warn": "Please set a song title first",
        "save_btn": "💾 Save to Library",
        "save_ok": "✅ Saved: {name}",
        "save_fail": "Save failed: {err}",
        "lib_title": "#### 📂 Song Library",
        "lib_empty": "No saved songs yet. Click Save above to add the first one.",
        "lib_count": "{n} songs",
        "load_btn": "📖 Load",
        "del_btn": "🗑️",
        "del_done": "Deleted 《{song}》",
        "lyricist_short": "✍️",
        "composer_short": "🎼",
        "orig_label": "Key: ",
        "target_label": "Target: ",
        "bpm_label": "BPM: ",
        "beat_label": "Time: ",
        "lyricist_disp": "Lyricist: {lyr}　Composer: {cmp}",
        "meta_caption": "Key: {orig} → Target: {target}　BPM: {bpm}　Time: {beat}",
        "cur_song_label": "Current: 《{song}》",
        "credit_both": "✍️ Lyrics: {lyr}　🎼 Music: {comp}",
        "credit_lyr": "✍️ Words & Music: {lyr}",
        "credit_comp": "🎼 Music: {comp}",
    },
}
LANG_OPTIONS = list(LANG_STRINGS.keys())

_defaults = {
    'buffer': "",
    'buffer_key': "C",
    '_sync_editor': False,
    'is_fullscreen': False,
    'yt_url': "",
    'search_status': "",
    'search_candidates': [],
    'gemini_model': "gemini-2.5-flash",
    'lang': "繁體中文",
    'role': None,           # None = 未登入, "admin" or "user"
    'gemini_key_override': "",  # admin 可在 UI 覆寫 key
    'meta': {
        "singer": "", "song": "", "lyricist": "", "composer": "",
        "bpm": 80, "beat": "4/4", "orig": "C", "target": "C"
    },
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# 語言快捷：T['key'] 取當前語言字串
T = LANG_STRINGS[st.session_state.get('lang', '繁體中文')]

# 同步旗標：在任何 widget 渲染前執行，安全地把 buffer 同步給 editor_main
if st.session_state.get('_sync_editor') or 'editor_main' not in st.session_state:
    st.session_state['editor_main'] = st.session_state.buffer
    st.session_state['_sync_editor'] = False

# ── 登入攔截 ──────────────────────────────────────────────
def _check_passwords():
    try:
        admin_pw = st.secrets.get("ADMIN_PASSWORD", "admin")
        user_pw  = st.secrets.get("USER_PASSWORD",  "user")
    except Exception:
        admin_pw, user_pw = "admin", "user"
    return admin_pw, user_pw

if st.session_state.role is None:
    _adm_pw, _usr_pw = _check_passwords()
    st.markdown("""
    <style>
    .login-wrap { max-width:400px; margin:80px auto 0; padding:40px 36px;
                  background:#1E293B; border-radius:16px; border:1px solid #334155; }
    .login-title { color:#FDE047; font-size:26px; font-weight:900;
                   text-align:center; margin-bottom:4px; }
    .login-sub   { color:#94A3B8; font-size:13px; text-align:center; margin-bottom:28px; }
    </style>
    <div class="login-wrap">
      <div class="login-title">🎸 Liberlive AI Station</div>
      <div class="login-sub">請輸入密碼以繼續 / Enter password to continue</div>
    </div>
    """, unsafe_allow_html=True)

    _pw_in = st.text_input("🔑 密碼 Password", type="password", key="login_pw_input",
                           placeholder="輸入密碼...")
    if st.button("🚀 登入 / Login", use_container_width=True):
        if _pw_in == _adm_pw:
            st.session_state.role = "admin"
            st.rerun()
        elif _pw_in == _usr_pw:
            st.session_state.role = "user"
            st.rerun()
        else:
            st.error("密碼錯誤 / Wrong password")
    st.stop()
# ──────────────────────────────────────────────────────────

def get_gemini_key():
    # admin 可在 sidebar 覆寫 key（session 層級，不改 secrets.toml）
    override = st.session_state.get("gemini_key_override", "")
    if override:
        return override
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return st.session_state.get("gemini_key_input", "")

def get_gemini_model():
    return st.session_state.get("gemini_model", "gemini-2.5-flash")

# --- 6. 核心演算法 ---
SCRAPE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

def fetch_page(url):
    """通用網頁爬取，回傳 BeautifulSoup 或 None"""
    try:
        res = requests.get(url.strip(), headers=SCRAPE_HEADERS, timeout=12)
        res.encoding = res.apparent_encoding or 'utf-8'
        return BeautifulSoup(res.text, 'html.parser')
    except Exception:
        return None

# 各網站的內容選取器（依優先序）
CONTENT_SELECTORS = [
    'div.chord-content', 'div.lyric-content', 'div.lyrics',
    '.chord-text', '.guitar-content', '.tab-content',
    'pre', '.post-content', '.entry-content', '.article-content',
    'article', 'main',
]

def extract_chords_from_soup(soup):
    """從 BeautifulSoup 中提取和弦歌詞文字，盡量取最有料的區塊"""
    if not soup:
        return ""
    for selector in CONTENT_SELECTORS:
        block = soup.select_one(selector)
        if block:
            for s in block(["script", "style", "nav", "header", "footer", "aside"]):
                s.decompose()
            text = block.get_text(separator='\n')
            # 必須有實質內容（含字母或中文），不只是空白
            if len(text.strip()) > 80:
                return text
    return ""

def fetch_lyrics_from_url(url):
    soup = fetch_page(url)
    return extract_chords_from_soup(soup)

def transpose_engine(text, steps):
    def _t(p):
        m = re.match(r"([A-G][#b]?)(.*)", p)
        if m:
            r, s = m.group(1), m.group(2)
            norm = {'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#'}
            base = norm.get(r, r)
            if base in KEYS:
                return KEYS[(KEYS.index(base) + steps) % 12] + s
        return p
    return re.sub(
        r'\[([^\]]+)\]',
        lambda m: "[" + "/".join([_t(x.strip()) for x in m.group(1).split('/')]) + "]",
        text
    )

def convert_stacked_to_inline(text):
    """
    把「和弦行在上、歌詞行在下」的常見曲譜格式轉成 [Chord]歌詞 的內嵌格式。
    例如：
      G        Bm    Em
      故事的小黃花 從出生那年就飄著
    →  [G]故事的小黃[Bm]花 從出生那[Em]年就飄著
    """
    # 先判斷是否已經是 [Chord] 內嵌格式
    if re.search(r'\[[A-G][^\]]{0,5}\]', text):
        return text  # 已是內嵌格式，不需轉換

    chord_pattern = re.compile(r'^(?:\s*[A-G][#b]?(?:m|maj|min|dim|aug|sus|add|M)?(?:\d+)?(?:/[A-G][#b]?)?\s+){1,}$')
    lines = text.split('\n')
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # 判斷這行是否為「純和弦行」（只含和弦 + 空白）
        stripped = line.strip()
        if stripped and chord_pattern.match(stripped + ' '):
            chord_line = stripped
            # 下一行是歌詞行
            if i + 1 < len(lines) and lines[i + 1].strip():
                lyric_line = lines[i + 1]
                # 把和弦依位置嵌入歌詞行
                result.append(_embed_chords(chord_line, lyric_line))
                i += 2
                continue
        result.append(line)
        i += 1
    return '\n'.join(result)

def _embed_chords(chord_line, lyric_line):
    """把一行和弦按空白位置嵌入對應的歌詞行"""
    chord_tokens = list(re.finditer(r'[A-G][#b]?(?:m|maj|min|dim|aug|sus|add|M)?(?:\d+)?(?:/[A-G][#b]?)?', chord_line))
    if not chord_tokens:
        return lyric_line

    lyric_chars = list(lyric_line)
    # 從後往前插入，避免位移問題
    for tok in reversed(chord_tokens):
        pos = min(tok.start(), len(lyric_chars))
        lyric_chars.insert(pos, f'[{tok.group()}]')

    return ''.join(lyric_chars)

def extract_meta_from_text(text):
    """從文字中抽取歌曲資訊（原調、BPM、作詞、作曲等）"""
    meta = {}
    key_m = re.search(r'(?:原調|原key|Key|調性|key)[：:=\s]+([A-G][#b]?)', text, re.IGNORECASE)
    if key_m:
        meta['orig'] = key_m.group(1)
    bpm_m = re.search(r'(?:BPM|Tempo|速度|節奏)[：:=\s]+(\d{2,3})', text, re.IGNORECASE)
    if bpm_m:
        meta['bpm'] = int(bpm_m.group(1))
    beat_m = re.search(r'(?:拍號|Beat|拍子)[：:=\s]+(\d/\d)', text, re.IGNORECASE)
    if beat_m:
        meta['beat'] = beat_m.group(1)
    lyr_m = re.search(r'(?:作詞|Lyricist|词曲)[：:=\s]+([^\n]+)', text, re.IGNORECASE)
    if lyr_m:
        meta['lyricist'] = lyr_m.group(1).strip()
    cmp_m = re.search(r'(?:作曲|Composer|Composed by)[：:=\s]+([^\n]+)', text, re.IGNORECASE)
    if cmp_m:
        meta['composer'] = cmp_m.group(1).strip()
    return meta

# 中文曲譜網站搜尋清單（4G/5G 或雲端部署環境可連上）
CN_SEARCH_SITES = [
    ("有譜嗎",   "https://www.youpinyuepu.com/search/?key={q}"),
    ("吉他譜",   "https://jitascore.com/?s={q}"),
    ("菊風音樂", "https://www.jufeng.com.tw/?s={q}"),
    ("GTP",      "https://www.gtp.tw/search?q={q}"),
]
SKIP_URL_KEYWORDS = ['search', 'category', 'tag', 'page=', '#', 'login',
                     'register', 'facebook', 'twitter', 'youtube', 'javascript', 'mailto']

def _try_scrape_search_results(search_url, song_l, singer_l, site_name):
    """抓取一個搜尋結果頁，找到和弦內容即回傳"""
    soup = fetch_page(search_url)
    if not soup:
        return None
    for link in soup.find_all('a', href=True):
        href = link.get('href', '').strip()
        text = link.get_text(strip=True).lower()
        if any(kw in href.lower() for kw in SKIP_URL_KEYWORDS):
            continue
        if not href.startswith('http') or len(href) < 15:
            continue
        href_l = href.lower()
        if song_l in text or singer_l in text or song_l in href_l or singer_l in href_l:
            chord_text = fetch_lyrics_from_url(href)
            if chord_text and len(chord_text.strip()) > 100:
                return chord_text
    return None

def search_chord_sites(song, singer):
    """
    Step 1a：直接爬中文曲譜網站（有譜嗎/吉他譜/菊風/GTP）
              在 4G/5G 或 Streamlit Cloud 環境可用
    Step 1b：Gemini 找 Cifraclub URL 爬取（國際網路備援）
    """
    song_s, singer_s = song.strip(), singer.strip()
    if not song_s and not singer_s:
        return None, None, "empty_input"

    song_l   = song_s.lower()
    singer_l = singer_s.lower()
    query    = quote(f"{singer_s} {song_s}".strip())

    # Step 1a：中文網站（網路允許時有效）
    for site_name, tpl in CN_SEARCH_SITES:
        try:
            result = _try_scrape_search_results(tpl.format(q=query), song_l, singer_l, site_name)
            if result:
                return result, site_name, "ok"
        except Exception:
            continue

    # Step 1b：Gemini 找 Cifraclub（英文歌效果佳）
    api_key = get_gemini_key()
    if not api_key:
        return None, None, "no_key"

    if song_s and singer_s:
        query_desc = f"song '{song_s}' by '{singer_s}'"
    elif song_s:
        query_desc = f"song '{song_s}'"
    else:
        query_desc = f"a popular song by '{singer_s}'"

    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        prompt = (
            f"Find the exact Cifraclub Brazil URL for the {query_desc}.\n"
            f"URL format: https://www.cifraclub.com.br/[artist-slug]/[song-slug]/\n"
            f"Reply with ONLY the complete URL. No markdown, no explanation."
        )
        resp = client.models.generate_content(model=get_gemini_model(), contents=prompt)
        url = resp.text.strip().rstrip('.,')
        if 'cifraclub.com' in url and url.startswith('http'):
            soup = fetch_page(url)
            if soup:
                content = soup.select_one('.cifra_cnt')
                if content:
                    for s in content(["script", "style"]): s.decompose()
                    text = content.get_text(separator='\n')
                    if len(text.strip()) > 80:
                        return text, "Cifraclub", url
        return None, None, f"cifraclub_failed: {url}"
    except Exception as e:
        return None, None, f"error: {str(e)[:60]}"

def gemini_generate_chords(song, singer, api_key):
    """
    備援：讓 Gemini 直接輸出 [和弦]歌詞 格式。
    Gemini 的訓練資料來自真實曲譜網站，內容準確。
    當所有中文曲譜網站均無法連線時使用此方法。
    """
    song_s, singer_s = song.strip(), singer.strip()
    if not song_s and not singer_s:
        return None
    desc = f"《{song_s}》" if song_s else ""
    desc += f" by {singer_s}" if singer_s else ""
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        prompt = f"""請提供 {desc} 的完整吉他和弦譜。

輸出格式（必須嚴格遵守）：

第一部分 — 歌曲資訊（每行一個欄位）：
歌手: [歌手名]
作詞: [作詞者]
作曲: [作曲者]
原調: [原始調性，如 G]
BPM: [速度數字]
拍號: [如 4/4]

第二部分 — 和弦歌詞（緊接在資訊後，空一行）：
- 和弦內嵌在歌詞中，格式：[和弦]歌詞
- 歌詞必須是真實原版歌詞（中文歌用繁體中文）
- 每行一句歌詞
- 不要用 TAB 符號（不要 E|--|B|--|）
- 包含所有段落：主歌、副歌、橋段

範例：
歌手: 周杰倫
作詞: 周杰倫
作曲: 周杰倫
原調: G
BPM: 65
拍號: 4/4

[G]故事的小黃[Bm]花 從出生那年[Em]就飄著
[C]童年的盪秋[G]千 隨記憶一直[Am]晃到現[D]在

現在請提供 {desc} 的完整和弦譜："""
        resp = client.models.generate_content(model=get_gemini_model(), contents=prompt)
        result = resp.text.strip()
        # 確認有 [和弦] 格式才回傳
        if re.search(r'\[[A-G][^\]]{0,5}\]', result):
            return result
        return None
    except Exception:
        return None

def _preview_text(text, chars=60):
    """取歌詞前幾個字作預覽（跳過 metadata header 行）"""
    lines = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        # 跳過 metadata 行（歌手:/作詞:/原調: 等）
        if re.match(r'^(歌手|作詞|作曲|原調|BPM|拍號)[：:]', line):
            continue
        # 移除和弦標記，只留歌詞文字
        lyric = re.sub(r'\[[^\]]+\]', '', line).strip()
        if lyric:
            lines.append(lyric)
        if sum(len(l) for l in lines) >= chars:
            break
    preview = ''.join(lines)
    return preview[:chars] + ('...' if len(preview) > chars else '')

def smart_search_candidates(song, singer):
    """搜尋並回傳多個候選結果讓使用者選擇"""
    candidates = []

    # Step 1：各中文曲譜網站
    song_s, singer_s = song.strip(), singer.strip()
    query = quote(f"{singer_s} {song_s}".strip())
    song_l, singer_l = song_s.lower(), singer_s.lower()

    for site_name, tpl in CN_SEARCH_SITES:
        try:
            result = _try_scrape_search_results(tpl.format(q=query), song_l, singer_l, site_name)
            if result:
                result = convert_stacked_to_inline(result)
                meta = extract_meta_from_text(result)
                candidates.append({
                    'source': f"🌐 {site_name}",
                    'text': result,
                    'meta': meta,
                    'preview': _preview_text(result),
                })
        except Exception:
            continue

    # Step 2：Gemini 生成（最多 2 個風格不同的版本）
    api_key = get_gemini_key()
    if api_key and len(candidates) < 2:
        result = gemini_generate_chords(song, singer, api_key)
        if result:
            meta = extract_meta_from_text(result)
            candidates.append({
                'source': "🤖 Gemini AI 生成",
                'text': result,
                'meta': meta,
                'preview': _preview_text(result),
            })

    return candidates

def extract_text_from_file(uploaded_file):
    """
    從上傳的檔案提取文字，回傳 (text, error_msg)。
    支援：.txt / .docx / .pdf / .jpg/.jpeg/.png（Gemini Vision OCR）
    提取後自動嘗試 convert_stacked_to_inline()。
    """
    name = uploaded_file.name.lower()
    raw_bytes = uploaded_file.read()

    # ── .txt ──
    if name.endswith('.txt'):
        try:
            text = raw_bytes.decode('utf-8')
        except UnicodeDecodeError:
            text = raw_bytes.decode('big5', errors='replace')
        return convert_stacked_to_inline(text), None

    # ── .docx ──
    if name.endswith('.docx'):
        import io
        doc = Document(io.BytesIO(raw_bytes))
        text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
        return convert_stacked_to_inline(text), None

    # ── .pdf ──
    if name.endswith('.pdf'):
        if not HAS_PDF:
            return None, "PDF 支援套件未安裝，請重啟 start.bat 讓它自動安裝"
        import io, pdfplumber
        pages_text = []
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t)
        text = '\n'.join(pages_text)
        if not text.strip():
            return None, "PDF 無法提取文字（可能是掃描圖片版），請改用圖檔匯入"
        return convert_stacked_to_inline(text), None

    # ── 圖檔（jpg / jpeg / png）：Gemini Vision OCR ──
    if name.endswith(('.jpg', '.jpeg', '.png')):
        api_key = get_gemini_key()
        if not api_key:
            return None, "圖檔 OCR 需要 Gemini Key，請先設定"
        try:
            from google import genai
            from google.genai import types
            client = genai.Client(api_key=api_key)
            mime = "image/jpeg" if name.endswith(('.jpg', '.jpeg')) else "image/png"
            img_part = types.Part.from_bytes(data=raw_bytes, mime_type=mime)
            prompt = (
                "這是一張吉他和弦歌詞圖。請仔細辨識所有和弦與歌詞，"
                "並以 [和弦]歌詞 的格式輸出（例如：[G]故事的小黃[Bm]花）。"
                "只輸出歌詞和弦內容，不要加說明。若無法辨識請回覆：無法辨識。"
            )
            resp = client.models.generate_content(
                model=get_gemini_model(),
                contents=[prompt, img_part]
            )
            result = resp.text.strip()
            if '無法辨識' in result or not result:
                return None, "Gemini 無法辨識圖片內容，請確認圖片清晰度"
            return result, None
        except Exception as e:
            return None, f"圖片辨識失敗：{str(e)[:80]}"

    return None, f"不支援的格式：{name}"

def fetch_youtube_song_info(url):
    """用 YouTube oEmbed API 取得影片標題/頻道名，再交 Gemini 解析歌名/歌手"""
    try:
        # Step 1：oEmbed API（不需 API Key，YouTube 官方支援）
        oembed_url = f"https://www.youtube.com/oembed?url={url.strip()}&format=json"
        r = requests.get(oembed_url, timeout=10)
        if r.status_code != 200:
            return "", "", f"YouTube oEmbed 失敗 (HTTP {r.status_code})，請確認網址正確"
        data = r.json()
        raw_title   = data.get("title", "").strip()
        raw_channel = data.get("author_name", "").strip()

        # 清理常見後綴
        raw_title = re.sub(r'\s*[-–]\s*YouTube\s*$', '', raw_title, flags=re.IGNORECASE)
        for pat in [r'\(Official.*?\)', r'\[Official.*?\]', r'【Official.*?】',
                    r'\(MV\)', r'\[MV\]', r'【MV】', r'\(Lyrics.*?\)', r'\[Lyrics.*?\]',
                    r'\(Audio.*?\)', r'\[Audio.*?\]', r'feat\..+$']:
            raw_title = re.sub(pat, '', raw_title, flags=re.IGNORECASE)
        raw_title = raw_title.strip(' -–|　')

        if not raw_title:
            return "", "", "無法取得影片標題"

        # Step 2：用 Gemini 解析歌名/歌手（channel 名作為補充線索）
        api_key = get_gemini_key()
        if api_key:
            from google import genai
            client = genai.Client(api_key=api_key)
            prompt = (
                f'YouTube video title: "{raw_title}"\n'
                f'Channel name: "{raw_channel}"\n'
                f'Identify the SONG NAME and ARTIST. Reply in this exact format only:\n'
                f'歌名: [song name in original language]\n'
                f'歌手: [artist name]\n'
                f'If unknown, use the most likely guess from the title/channel.'
            )
            resp = client.models.generate_content(model=get_gemini_model(), contents=prompt)
            text = resp.text.strip()
            song_m   = re.search(r'歌名[：:]\s*(.+)', text)
            singer_m = re.search(r'歌手[：:]\s*(.+)', text)
            song_out   = song_m.group(1).strip()   if song_m   else raw_title
            singer_out = singer_m.group(1).strip() if singer_m else raw_channel
            # 過濾掉 Gemini 回 "未知"
            if song_out in ('未知', 'Unknown', ''):
                song_out = raw_title
            if singer_out in ('未知', 'Unknown', ''):
                singer_out = raw_channel
            return song_out, singer_out, f"識別：{raw_title}　頻道：{raw_channel}"
        else:
            # 無 Gemini Key：直接從標題拆分
            parts = re.split(r'\s*[-–]\s*', raw_title, maxsplit=1)
            if len(parts) == 2:
                return parts[0].strip(), parts[1].strip(), f"從標題解析：{raw_title}"
            return raw_title, raw_channel, f"從標題解析：{raw_title}"
    except Exception as e:
        return "", "", f"YouTube 識別失敗：{str(e)[:80]}"

# --- 7a. 圖片匯出 ---
def _find_cjk_font():
    candidates = [
        # Windows
        "C:/Windows/Fonts/msjh.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/mingliu.ttc",
        # Linux (Streamlit Cloud / Ubuntu)
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        # macOS
        "/System/Library/Fonts/PingFang.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def generate_chart_image(buffer, meta, color_map, display_steps=0):
    """把 buffer 渲染成 PIL Image，背景深色、和弦彩色標籤"""
    from PIL import Image, ImageDraw, ImageFont
    import io

    font_path = _find_cjk_font()
    try:
        chord_font  = ImageFont.truetype(font_path, 22) if font_path else ImageFont.load_default()
        lyric_font  = ImageFont.truetype(font_path, 32) if font_path else ImageFont.load_default()
        title_font  = ImageFont.truetype(font_path, 36) if font_path else ImageFont.load_default()
        small_font  = ImageFont.truetype(font_path, 22) if font_path else ImageFont.load_default()
    except Exception:
        chord_font = lyric_font = title_font = small_font = ImageFont.load_default()

    COLOR_HEX = {
        'C': (239, 68, 68), 'D': (249, 115, 22), 'E': (250, 204, 21),
        'F': (34, 197, 94),  'G': (6, 182, 212),  'A': (29, 78, 216),
        'B': (168, 85, 247),
    }
    BG       = (10, 15, 30)
    FG_LYR   = (241, 245, 249)
    FG_TITLE = (253, 224, 71)
    FG_META  = (34, 197, 94)
    PAD_X, PAD_Y = 40, 40
    CHORD_H, LYRIC_H, LINE_GAP = 28, 38, 18
    UNIT_H = CHORD_H + LYRIC_H + 4
    LINE_H = UNIT_H + LINE_GAP

    src = transpose_engine(buffer, display_steps) if display_steps else buffer

    # 先過一遍收集所有行，計算需要的高度
    lines_data = []
    for line in src.split('\n'):
        if not line.strip(): continue
        if re.match(r'^\s*(歌手|作詞|作曲|原調|BPM|拍號|Capo)[：:：]', line): continue
        if re.match(r'^\s*[EBGDAe]\s*\|', line): continue
        has_cn = bool(re.search(r'[一-鿿㐀-䶿]', line))
        has_ch = bool(re.search(r'\[[A-G][^\]]*\]', line))
        if not has_cn and not has_ch: continue
        lines_data.append(line)

    song  = meta.get('song','')
    singer = meta.get('singer','')
    orig  = meta.get('orig','C')
    tgt   = meta.get('target','C')
    bpm   = meta.get('bpm','')
    beat  = meta.get('beat','')

    title_block_h = 36 + 28 + PAD_Y  # title + meta line + padding
    total_h = PAD_Y + title_block_h + len(lines_data) * LINE_H + PAD_Y

    # 估算最長行的寬度
    max_chars = max((len(re.sub(r'\[[^\]]+\]', '', l)) for l in lines_data), default=20)
    img_w = max(800, PAD_X * 2 + max_chars * 34)

    img = Image.new('RGB', (img_w, max(400, total_h)), BG)
    draw = ImageDraw.Draw(img)

    # 標題
    y = PAD_Y
    draw.text((PAD_X, y), f"《{song}》", font=title_font, fill=FG_TITLE)
    y += 44
    meta_str = f"🎤 {singer}　🎸 {orig}→{tgt}　⏱️ {bpm}　🥁 {beat}"
    draw.text((PAD_X, y), meta_str, font=small_font, fill=FG_META)
    y += 36

    # 分隔線
    draw.line([(PAD_X, y), (img_w - PAD_X, y)], fill=(51, 65, 85), width=1)
    y += 16

    # 每一行
    for line in lines_data:
        parts = re.split(r'(\[[^\]]+\])', line)
        pending = ""
        x = PAD_X
        for p in parts:
            if p.startswith('[') and p.endswith(']'):
                pending = p[1:-1]
            else:
                for ch in p:
                    # 和弦標籤
                    if pending:
                        root = pending[0].upper()
                        c_rgb = COLOR_HEX.get(root, (51, 65, 85))
                        chord_txt = pending
                        try:
                            cw = chord_font.getlength(chord_txt)
                        except Exception:
                            cw = len(chord_txt) * 13
                        draw.rounded_rectangle(
                            [x, y, x + cw + 8, y + CHORD_H],
                            radius=3, fill=c_rgb
                        )
                        draw.text((x + 4, y + 2), chord_txt, font=chord_font, fill=(255,255,255))
                    # 歌詞字元
                    draw.text((x, y + CHORD_H + 2), ch if ch != ' ' else ' ', font=lyric_font, fill=FG_LYR)
                    try:
                        char_w = int(lyric_font.getlength(ch if ch != ' ' else '一'))
                    except Exception:
                        char_w = 34
                    x += max(char_w, 20)
                    pending = ""
        y += LINE_H

    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf.getvalue()


def generate_chart_html(buffer, meta, color_map, display_steps=0):
    """生成可列印的 HTML 字串（匯出後在瀏覽器 Ctrl+P 存 PDF）"""
    src = transpose_engine(buffer, display_steps) if display_steps else buffer

    song   = meta.get('song','')
    singer = meta.get('singer','')
    orig   = meta.get('orig','C')
    tgt    = meta.get('target','C')
    bpm    = meta.get('bpm','')
    beat   = meta.get('beat','')

    chord_colors = {
        'C':'#EF4444','D':'#F97316','E':'#FACC15','F':'#22C55E',
        'G':'#06B6D4','A':'#1D4ED8','B':'#A855F7',
    }

    lines_html = []
    for line in src.split('\n'):
        if not line.strip(): continue
        if re.match(r'^\s*(歌手|作詞|作曲|原調|BPM|拍號|Capo)[：:：]', line): continue
        if re.match(r'^\s*[EBGDAe]\s*\|', line): continue
        has_cn = bool(re.search(r'[一-鿿㐀-䶿]', line))
        has_ch = bool(re.search(r'\[[A-G][^\]]*\]', line))
        if not has_cn and not has_ch: continue

        parts = re.split(r'(\[[^\]]+\])', line)
        pending = ""
        row = '<div style="display:flex;flex-wrap:wrap;align-items:flex-end;margin-bottom:14px;">'
        for p in parts:
            if p.startswith('[') and p.endswith(']'):
                pending = p[1:-1]
            else:
                for ch in p:
                    if pending:
                        root = pending[0].upper()
                        col = chord_colors.get(root, '#334155')
                        ctag = f'<span style="display:inline-block;background:{col};color:#fff;font-size:13px;font-weight:900;border-radius:3px;padding:1px 5px;margin-bottom:2px;font-family:monospace;">{pending}</span>'
                    else:
                        ctag = '<span style="display:inline-block;height:22px;min-width:0.6em;"></span>'
                    disp = '&nbsp;' if ch == ' ' else ch
                    row += f'<div style="display:inline-flex;flex-direction:column;align-items:center;padding:0 1px;">{ctag}<span style="font-size:22px;color:#1e293b;">{disp}</span></div>'
                    pending = ""
        row += '</div>'
        lines_html.append(row)

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>{song}</title>
<style>
  body {{ font-family: "Microsoft JhengHei","微軟正黑體",Arial,sans-serif; background:#fff; padding:32px 48px; }}
  h1 {{ color:#1e3a8a; margin-bottom:4px; }}
  .meta {{ color:#15803d; font-size:14px; margin-bottom:20px; }}
  @media print {{ body {{ padding: 20px; }} }}
</style>
</head><body>
<h1>《{song}》</h1>
<div class="meta">🎤 {singer}　🎸 {orig}→{tgt}　⏱️ {bpm}　🥁 {beat}</div>
<hr>
{''.join(lines_html)}
</body></html>"""
    return html.encode('utf-8')


# --- 7. 全局 RWD 自適應 CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #0A0F1E !important; color: #E2E8F0 !important; }
    header, footer { visibility: hidden !important; }
    .block-container { padding-top: 0rem !important; padding-bottom: 0.5rem !important; }

    /* 所有 Streamlit 預設白底元件改暗色 */
    .stTextInput input, .stSelectbox select,
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div { background-color: #1E293B !important; color: #F1F5F9 !important; border-color: #334155 !important; }
    .stNumberInput input { background-color: #1E293B !important; color: #F1F5F9 !important; }
    label, .stMarkdown p { color: #CBD5E1 !important; }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1E3A8A 0%, #0F172A 100%) !important;
        border-right: 3px solid #FDE047;
    }
    section[data-testid="stSidebar"] * { color: white !important; }
    section[data-testid="stSidebar"] .stTextInput input { color: #1E293B !important; background-color: #F8FAFC !important; }

    .status-bar {
        background: linear-gradient(90deg, #1E3A8A 0%, #0F172A 100%);
        padding: 10px 20px;
        border-radius: 8px;
        font-weight: bold;
        margin-bottom: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.15);
        font-family: monospace, sans-serif;
    }
    .status-left { font-size: 16px; color: #FFFFFF !important; }
    .status-right { font-size: 14px; color: #FDE047 !important; }

    /* ── 演出模式：深色舞台背景 ── */
    .stage-paper {
        background: #0F172A !important;
        border: 2px solid #1E3A8A;
        padding: 28px 30px;
        border-radius: 12px;
        max-height: 75vh !important;
        width: 100% !important;
        overflow-x: hidden !important;
        overflow-y: auto !important;
        box-shadow: 0 4px 24px rgba(0,0,0,0.4);
    }

    /* 每一行歌詞（含和弦）的容器 */
    .chord-line {
        display: flex !important;
        flex-wrap: wrap !important;
        align-items: flex-end !important;
        margin-bottom: 32px !important;
        gap: 0px !important;
    }

    /* 每個字符單元：上方和弦 + 下方歌詞字 */
    .char-unit {
        display: inline-flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: flex-end !important;
        min-width: 1em !important;
        padding: 0 2px !important;
    }

    /* 和弦標籤：有色背景方塊 */
    .c-tag {
        display: inline-block !important;
        font-weight: 900 !important;
        font-family: 'Courier New', monospace !important;
        line-height: 1 !important;
        padding: 2px 5px !important;
        border-radius: 4px !important;
        margin-bottom: 4px !important;
        color: white !important;
        white-space: nowrap !important;
        min-width: 1.2em !important;
        text-align: center !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.4) !important;
        /* 預設用 clamp 自動適配：最小 11px，理想 3vw，最大 22px */
        font-size: clamp(11px, 3vw, 22px) !important;
    }
    /* 無和弦時佔位，保持對齊 */
    .c-tag-empty {
        display: inline-block !important;
        min-width: 1em !important;
        height: clamp(16px, 3vw, 26px) !important;
        margin-bottom: 4px !important;
    }

    /* 歌詞文字：白色大字 */
    .l-tag {
        display: inline-block !important;
        color: #F1F5F9 !important;
        font-weight: 500 !important;
        line-height: 1.3 !important;
        letter-spacing: 0.05em !important;
        white-space: pre !important;
        /* 預設自動適配：最小 16px，理想 5vw，最大 36px */
        font-size: clamp(16px, 5vw, 36px) !important;
    }

    div.stButton > button {
        background-color: #22C55E !important;
        color: white !important;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 10px 16px;
        width: 100%;
    }
    div.stButton > button:hover { background-color: #16A34A !important; }

    .stTabs [data-baseweb="tab-list"] { background-color: #0F172A; border-radius: 8px; padding: 4px; border: 1px solid #1E3A8A; }
    .stTabs [data-baseweb="tab"] { color: #94A3B8 !important; font-weight: bold; font-size: 15px; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: #FDE047 !important; border-radius: 6px; }
    div[data-testid="stTextArea"] textarea { background-color: #1E293B !important; color: #F1F5F9 !important; border-color: #334155 !important; }

    .lib-card {
        background: #1E293B;
        padding: 12px;
        border-radius: 8px;
        border-top: 4px solid #1E3A8A;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        margin-bottom: 8px;
        font-weight: bold;
        color: #FDE047;
    }
    .lib-card-green { border-top-color: #22C55E !important; color: #22C55E !important; }
    .lib-card-yellow { border-top-color: #FDE047 !important; color: #FDE047 !important; }

    /* 曲目資訊卡片 */
    .song-info-bar {
        background: linear-gradient(90deg, #1E3A8A 0%, #0F172A 100%);
        border: 1px solid #334155;
        border-left: 5px solid #FDE047;
        border-radius: 8px;
        padding: 12px 20px;
        margin-bottom: 12px;
        display: flex;
        flex-wrap: wrap;
        gap: 8px 24px;
        align-items: center;
    }
    .song-info-title {
        font-size: 20px;
        font-weight: 900;
        color: #FDE047;
        letter-spacing: 0.05em;
    }
    .song-info-singer {
        font-size: 15px;
        color: #22C55E;
        font-weight: 700;
    }
    .song-info-meta {
        font-size: 13px;
        color: #94A3B8;
    }
    .song-info-meta span {
        color: #E2E8F0;
        font-weight: 600;
    }

    .search-status {
        background: #1E3A8A;
        color: #FDE047 !important;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: bold;
        font-family: monospace;
        margin-top: 8px;
        font-size: 13px;
    }

    /* 演出模式 CSS 全螢幕（不用瀏覽器 API，直接 fixed） */
    #stage.fs {
        position: fixed !important;
        top: 0 !important; left: 0 !important;
        width: 100vw !important; height: 100vh !important;
        max-height: 100vh !important;
        z-index: 99999 !important;
        border-radius: 0 !important;
        padding: 40px 50px !important;
        overflow-y: auto !important;
    }
    #stage.fs .fs-close {
        display: block !important;
    }
    .fs-close {
        display: none;
        position: fixed;
        top: 14px; right: 20px;
        background: #FDE047; color: #0F172A;
        border: none; border-radius: 6px;
        padding: 6px 14px; font-weight: bold;
        font-size: 14px; cursor: pointer;
        z-index: 100000;
    }
    .fullscreen-btn {
        background: #1E3A8A !important;
        color: #FDE047 !important;
        border: 2px solid #FDE047 !important;
        border-radius: 6px !important;
        padding: 6px 14px !important;
        font-weight: bold !important;
        cursor: pointer !important;
        font-size: 14px !important;
        margin-bottom: 10px !important;
    }
    .fullscreen-btn:hover { background: #FDE047 !important; color: #1E3A8A !important; }

    /* ── 平板 (768px 以下) ── */
    @media (max-width: 768px) {
        .status-bar { flex-direction: column; text-align: center; gap: 4px; padding: 8px; }
        .status-left { font-size: 14px; }
        .status-right { font-size: 12px; }
        .stage-paper { padding: 14px 10px; max-height: 65vh !important; }
        .chord-line { display: flex !important; flex-wrap: wrap !important; white-space: normal !important; width: 100% !important; margin-bottom: 20px !important; }
        .char-unit { display: inline-flex !important; flex-direction: column !important; margin-bottom: 8px; }
        .c-tag { font-size: 13px !important; padding: 1px 4px !important; margin-bottom: 2px !important; }
        /* 讓 Streamlit 的 columns 在小螢幕自動換行堆疊 */
        div[data-testid="column"] { min-width: 100% !important; }
        /* 加大按鈕點擊區 */
        .stButton > button { min-height: 44px !important; font-size: 15px !important; }
        /* 曲目資訊欄 換小字 */
        .block-container { padding-left: 8px !important; padding-right: 8px !important; }
    }

    /* ── 手機 (480px 以下：iPhone SE / Android) ── */
    @media (max-width: 480px) {
        .stage-paper { padding: 10px 6px; max-height: 70vh !important; font-size: 16px; }
        .chord-line { margin-bottom: 16px !important; }
        .c-tag { font-size: 11px !important; padding: 1px 3px !important; }
        .l-tag { font-size: 18px !important; }
        /* 全螢幕模式在手機填滿 */
        .block-container { padding: 0 4px !important; }
        /* 標題欄縮小 */
        h1, h2, h3 { font-size: 18px !important; }
        /* 曲目資訊唯讀格更緊湊 */
        div[data-testid="column"] > div > div > div {
            font-size: 13px !important;
        }
        /* 全螢幕退出按鈕固定在頂部 */
        .stButton > button[kind="secondary"] {
            position: sticky; top: 8px; z-index: 9999;
        }
    }

    /* ── 超寬螢幕 (Mac / iPad Pro 橫向) ── */
    @media (min-width: 1400px) {
        .stage-paper { max-height: 80vh !important; }
        .chord-line { margin-bottom: 36px !important; }
    }

    /* ── 觸控裝置：加大所有互動元素最小高度 ── */
    @media (hover: none) and (pointer: coarse) {
        .stButton > button  { min-height: 48px !important; padding: 10px 16px !important; }
        .stTextInput input  { min-height: 44px !important; font-size: 16px !important; }
        .stSelectbox > div  { min-height: 44px !important; }
        .stNumberInput input { min-height: 44px !important; font-size: 16px !important; }
        /* iOS Safari 禁止 input 自動放大 */
        input, select, textarea { font-size: 16px !important; }
    }
    </style>
""", unsafe_allow_html=True)

# --- 8. 置頂環境狀態列 ---
st.markdown(f"""
    <div class="status-bar">
        <div class="status-left">
            💎 Liberlive AI Station &nbsp;
            <span style="color:#22C55E;font-weight:bold;">v26.0</span>
            &nbsp;<span style="color:#FDE047;font-size:12px;">by Brett</span>
        </div>
        <div class="status-right">📅 {current_date} &nbsp;|&nbsp; {current_weather}</div>
    </div>
""", unsafe_allow_html=True)

# --- 9. 側邊欄 ---
with st.sidebar:
    # 角色徽章 + 登出
    _role = st.session_state.role
    _role_badge = "👑 管理者 Admin" if _role == "admin" else "🎵 使用者 User"
    _role_color = "#FDE047" if _role == "admin" else "#22C55E"
    st.markdown(
        f'<div style="background:#1E293B;border:1px solid #334155;border-radius:8px;'
        f'padding:8px 14px;margin-bottom:10px;display:flex;justify-content:space-between;align-items:center;">'
        f'<span style="color:{_role_color};font-weight:900;font-size:14px;">{_role_badge}</span>'
        f'</div>', unsafe_allow_html=True
    )
    if st.button("🚪 登出 Logout", use_container_width=True, key="logout_btn"):
        st.session_state.role = None
        st.rerun()

    st.markdown("---")
    # 語言切換
    lang_idx = LANG_OPTIONS.index(st.session_state.get('lang', '繁體中文'))
    chosen_lang = st.selectbox(T["lang_label"], LANG_OPTIONS, index=lang_idx, key="lang_picker")
    if chosen_lang != st.session_state.lang:
        st.session_state.lang = chosen_lang
        st.rerun()
    T = LANG_STRINGS[st.session_state.lang]

    st.markdown("---")
    # Gemini 設定：只有 admin 可見
    if _role == "admin":
        st.markdown(T["gemini_title"])
        if get_gemini_key():
            st.success(T["api_ok"])
        else:
            st.warning(T["api_warn"])

        # Admin 可輸入新 Key 覆寫（只影響本次 session）
        new_key = st.text_input("🔑 覆寫 API Key（選填）", type="password",
                                value=st.session_state.gemini_key_override,
                                key="admin_key_input",
                                placeholder="留空 = 使用 secrets.toml")
        if new_key != st.session_state.gemini_key_override:
            st.session_state.gemini_key_override = new_key

        cur_model = st.session_state.get("gemini_model", "gemini-2.5-flash")
        cur_idx = GEMINI_MODEL_IDS.index(cur_model) if cur_model in GEMINI_MODEL_IDS else 0
        sel_label = st.selectbox(T["model_pick"], GEMINI_MODEL_LABELS, index=cur_idx, key="model_picker")
        st.session_state.gemini_model = GEMINI_MODEL_IDS[GEMINI_MODEL_LABELS.index(sel_label)]
        st.markdown("---")
    else:
        # 一般使用者：不顯示 Gemini Key，model_picker 仍要建立（供 get_gemini_model() 讀取）
        st.session_state.setdefault("model_picker", "gemini-2.5-flash  ｜ 🆓 免費（每日限量）")

    st.markdown(T["yt_sync"])
    st.session_state.yt_url = st.text_input(T["yt_url_label"], value=st.session_state.yt_url)
    if st.session_state.yt_url:
        st.video(st.session_state.yt_url)

    st.markdown("---")
    c_size     = st.slider(T["chord_font"], 10, 60, st.session_state.get("c_size", 18), key="sb_c_size")
    l_size     = st.slider(T["lyric_font"], 12, 72, st.session_state.get("l_size", 26), key="sb_l_size")
    scroll_spd = st.slider(T["scroll_spd"], 0, 20, 0)
    st.session_state["c_size"] = c_size
    st.session_state["l_size"] = l_size

# --- 10. 置頂曲目資訊欄 ---
# 上排：唯讀顯示（搜尋後自動填入）
m0 = st.session_state.meta
ro_cols = st.columns([3, 3, 3, 3])
ro_labels = [T["meta_song"], T["meta_singer"], T["meta_lyricist"], T["meta_composer"]]
ro_keys   = ["song", "singer", "lyricist", "composer"]
for col, label, key in zip(ro_cols, ro_labels, ro_keys):
    col.markdown(f'<div style="font-size:12px;color:#94A3B8;margin-bottom:2px;">{label}</div>'
                 f'<div style="background:#1E293B;border:1px solid #334155;border-radius:6px;'
                 f'padding:8px 12px;color:#F1F5F9;font-size:15px;min-height:38px;">'
                 f'{m0.get(key,"") or "—"}</div>', unsafe_allow_html=True)

# 下排：可編輯（變調/速度控制）
st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)
e1, e2, e3, e4 = st.columns([2, 2, 2, 2])
with e1: ok_val  = st.selectbox(T["meta_orig"],   KEYS, index=KEYS.index(m0['orig']),   key="meta_orig")
with e2: tk_val  = st.selectbox(T["meta_target"], KEYS, index=KEYS.index(m0['target']), key="meta_target")
with e3: bpm_val = st.number_input(T["meta_bpm"], 20, 250, m0['bpm'], key="meta_bpm")
with e4: beat_val= st.text_input(T["meta_beat"],  value=m0['beat'], key="meta_beat")

# 只更新可編輯欄位
st.session_state.meta.update({
    "orig": ok_val, "target": tk_val,
    "bpm": bpm_val, "beat": beat_val
})

# --- 11. 全螢幕模式攔截 ---
# 如果 is_fullscreen=True，只渲染詞譜畫面，隱藏所有 Streamlit 外框，然後 st.stop()
if st.session_state.get('is_fullscreen') and st.session_state.buffer:
    st.markdown("""
    <style>
    header, footer, [data-testid="stSidebar"],
    [data-testid="stToolbar"], [data-testid="stDecoration"],
    section.main > div:first-child { display: none !important; }
    .block-container { padding: 0 !important; max-width: 100vw !important; }
    </style>
    """, unsafe_allow_html=True)

    # 退出按鈕（浮在右上角）
    if st.button(T["exit_fullscreen"], key="exit_fs"):
        st.session_state.is_fullscreen = False
        st.rerun()

    # 渲染詞譜
    m_fs = st.session_state.meta
    bk_fs = st.session_state.get('buffer_key', 'C')
    tk_fs = m_fs.get('target', 'C')
    steps_fs = (KEYS.index(tk_fs) - KEYS.index(bk_fs)) % 12
    buf_fs = transpose_engine(st.session_state.buffer, steps_fs) if steps_fs else st.session_state.buffer

    html_fs = [f'<div style="background:#0F172A;min-height:100vh;padding:30px 40px;">',
               f'<div style="color:#FDE047;font-size:22px;font-weight:900;margin-bottom:4px;">《{m_fs.get("song","")}》</div>',
               f'<div style="color:#22C55E;font-size:14px;margin-bottom:20px;">🎤 {m_fs.get("singer","")}　🎸 {m_fs.get("orig","")}→{tk_fs}　⏱️{m_fs.get("bpm","")}　🥁{m_fs.get("beat","")}</div>']

    for line in buf_fs.split('\n'):
        if not line.strip(): continue
        if re.match(r'^\s*(歌手|作詞|作曲|原調|BPM|拍號|Capo)[：:：]', line): continue
        if re.match(r'^\s*[EBGDAe]\s*\|', line): continue
        has_cn = bool(re.search(r'[一-鿿㐀-䶿]', line))
        has_ch = bool(re.search(r'\[[A-G][^\]]*\]', line))
        if not has_cn and not has_ch: continue
        parts_fs = re.split(r'(\[[^\]]+\])', line)
        pend_fs = ""
        row = '<div style="display:flex;flex-wrap:wrap;align-items:flex-end;margin-bottom:28px;">'
        for p in parts_fs:
            if p.startswith('[') and p.endswith(']'):
                pend_fs = p[1:-1]
            else:
                for ch in p:
                    if pend_fs:
                        col_fs = COLOR_MAP.get(pend_fs[0].upper(), '#334155')
                        ctag = (f'<span style="display:inline-block;background:{col_fs};color:white;'
                                f'font-size:{c_size}px;font-weight:900;border-radius:4px;padding:2px 6px;'
                                f'margin-bottom:4px;font-family:monospace;white-space:nowrap;">{pend_fs}</span>')
                    else:
                        ctag = f'<span style="display:inline-block;height:{c_size+8}px;margin-bottom:4px;min-width:0.5em;"></span>'
                    disp = "&nbsp;" if ch == " " else ch
                    row += (f'<div style="display:inline-flex;flex-direction:column;align-items:center;padding:0 2px;">'
                            f'{ctag}<span style="color:#F1F5F9;font-size:{l_size}px;">{disp}</span></div>')
                    pend_fs = ""
        row += '</div>'
        html_fs.append(row)
    html_fs.append('</div>')
    st.markdown('\n'.join(html_fs), unsafe_allow_html=True)
    st.stop()

# --- 12. 三個主頁籤 ---
tab_in, tab_play, tab_cloud = st.tabs([T["tab_search"], T["tab_play"], T["tab_cloud"]])

# ── Tab 1：智能尋譜 & 輸入 ──
with tab_in:

    # ── 候選結果確認面板（搜尋完成、使用者尚未選擇時顯示）
    if st.session_state.search_candidates:
        st.markdown("### 🔎 " + T["tab_search"])
        for i, cand in enumerate(st.session_state.search_candidates):
            m_cand = cand['meta']
            singer_c  = m_cand.get('singer',  st.session_state.meta.get('singer',''))
            lyricist_c= m_cand.get('lyricist','')
            composer_c= m_cand.get('composer','')
            orig_c    = m_cand.get('orig',    'C')
            bpm_c     = m_cand.get('bpm',     '--')
            credit_parts = []
            if lyricist_c: credit_parts.append(T["lyricist_label"] + lyricist_c)
            if composer_c: credit_parts.append(T["composer_label"] + composer_c)
            credit_str = "　".join(credit_parts) if credit_parts else T["cand_credit_unknown"]

            with st.container(border=True):
                st.markdown(f"**{T['cand_source'].format(source=cand['source'])}**")
                col_info, col_btn = st.columns([5, 1])
                with col_info:
                    st.markdown(T["cand_singer"].format(singer=singer_c or '—', orig=orig_c, bpm=bpm_c))
                    st.markdown(f"{T['lyricist_short']} {credit_str}")
                    st.markdown(f"{T['cand_preview']}*{cand['preview']}*")
                with col_btn:
                    if st.button(T["cand_pick"], key=f"pick_{i}"):
                        result_text = cand['text']
                        new_meta = {
                            "singer":   singer_c,
                            "song":     st.session_state.meta.get('song',''),
                            "lyricist": lyricist_c,
                            "composer": composer_c,
                            "orig":     orig_c,
                            "bpm":      bpm_c if isinstance(bpm_c, int) else st.session_state.meta.get('bpm', 80),
                            "beat":     m_cand.get('beat', '4/4'),
                        }
                        st.session_state.meta.update(new_meta)
                        st.session_state.buffer = result_text
                        st.session_state['_sync_editor'] = True
                        st.session_state.buffer_key = orig_c
                        st.session_state.search_candidates = []
                        st.rerun()

        if st.button(T["cand_cancel"]):
            st.session_state.search_candidates = []
            st.rerun()
        st.stop()

    # ── 正常輸入區 ──
    col_ai, col_yt, col_file = st.columns([3, 3, 2])

    with col_ai:
        st.markdown(f'<div class="lib-card lib-card-yellow">{T["ai_card"]}</div>', unsafe_allow_html=True)
        search_song   = st.text_input(T["ai_song_label"], value=st.session_state.meta['song'],   key="ai_song")
        search_singer = st.text_input(T["ai_singer_label"], value=st.session_state.meta['singer'], key="ai_singer")

        # 只有在 Key 未設定時才顯示輸入框
        if not get_gemini_key():
            st.text_input(T["gemini_key_label"], type="password", key="gemini_key_input",
                          placeholder="AQ.xxxx...", label_visibility="visible")
        else:
            st.session_state.setdefault("gemini_key_input", "")

        if st.button(T["search_btn"]):
            if not search_song.strip():
                st.warning(T["search_warn"])
            else:
                with st.spinner(f"{T['search_spinner']}《{search_song}》..."):
                    candidates = smart_search_candidates(search_song, search_singer)
                if candidates:
                    st.session_state.meta['song']   = search_song
                    st.session_state.meta['singer'] = search_singer
                    st.session_state.search_candidates = candidates
                    st.rerun()
                else:
                    st.session_state.search_status = T["search_fail"]
                    st.rerun()

        if st.session_state.search_status:
            st.markdown(f'<div class="search-status">{st.session_state.search_status}</div>',
                        unsafe_allow_html=True)

    with col_yt:
        st.markdown(f'<div class="lib-card lib-card-green">{T["yt_card"]}</div>', unsafe_allow_html=True)
        yt_in = st.text_input(T["yt_url_label"], key="input_yt",
                              label_visibility="collapsed", placeholder=T["yt_placeholder"])
        if st.button(T["yt_btn"]):
            if yt_in:
                with st.spinner(T["yt_spinner"]):
                    yt_song, yt_singer, yt_status = fetch_youtube_song_info(yt_in)
                if yt_song:
                    st.session_state.meta['song']   = yt_song
                    st.session_state.meta['singer'] = yt_singer
                    # 識別成功後直接搜尋和弦譜
                    with st.spinner(f"{T['search_spinner']}《{yt_song}》..."):
                        candidates = smart_search_candidates(yt_song, yt_singer)
                    if candidates:
                        st.session_state.search_candidates = candidates
                        st.session_state.search_status = ""
                    else:
                        st.session_state.search_status = f"📺 {yt_status}　{T['search_fail']}"
                    st.rerun()
                else:
                    st.warning(yt_status)

    with col_file:
        st.markdown(f'<div class="lib-card">{T["file_card"]}</div>', unsafe_allow_html=True)
        st.caption(T["file_hint"])
        doc_up = st.file_uploader(
            T["file_upload_label"],
            type=['txt', 'docx', 'pdf', 'jpg', 'jpeg', 'png'],
            label_visibility="collapsed"
        )
        if doc_up:
            with st.spinner(T["file_spinner"].format(name=doc_up.name)):
                file_text, err = extract_text_from_file(doc_up)
            if err:
                st.error(f"❌ {err}")
            elif file_text and file_text.strip():
                st.session_state.buffer = file_text
                st.session_state['_sync_editor'] = True
                st.session_state.buffer_key = st.session_state.meta.get('orig', 'C')
                st.toast(T["file_ok"].format(name=doc_up.name))
                st.rerun()
            else:
                st.warning(T["file_empty"])

    st.markdown("---")

    # 歌名標題 + 編輯區
    song_label = st.session_state.meta.get('song','') or T["no_song"]
    singer_label = st.session_state.meta.get('singer','')
    header_txt = f"### 🎵 {song_label}"
    if singer_label:
        header_txt += f"　　*{singer_label}*"
    st.markdown(header_txt)

    content = st.text_area(T["editor_label"], height=380, key="editor_main")
    if st.button(T["apply_btn"]):
            if content:
                steps = (KEYS.index(tk_val) - KEYS.index(ok_val)) % 12
                st.session_state.buffer = transpose_engine(content, steps)
                st.session_state.buffer_key = tk_val
                st.session_state['_sync_editor'] = True
                st.rerun()


# ── Tab 2：演出模式 ──
with tab_play:
    m = st.session_state.meta

    # 曲目資訊列
    lyr  = m.get('lyricist','')
    comp = m.get('composer','')
    credit = ""
    if lyr and comp and lyr != comp:
        credit = T["credit_both"].format(lyr=lyr, comp=comp)
    elif lyr:
        credit = T["credit_lyr"].format(lyr=lyr)
    elif comp:
        credit = T["credit_comp"].format(comp=comp)

    # 曲目資訊 + 全螢幕按鈕：全部一行
    parts = [f"**《{m.get('song','')}》**", f"🎤 {m.get('singer','')}"]
    if credit:
        parts.append(f"｜ {credit}")
    parts += [
        f"｜ 🎸 {m['orig']}→{m['target']}",
        f"｜ ⏱️{m['bpm']}",
        f"｜ 🥁{m['beat']}",
    ]
    info_col, btn_fs, btn_chord_m, btn_chord_p, btn_lyric_m, btn_lyric_p, btn_img, btn_html = st.columns([5, 1, 1, 1, 1, 1, 1, 1])
    with info_col:
        st.markdown("　".join(parts))

    # 讀取目前字體大小（sidebar 滑桿或上次調整的值）
    _c = st.session_state.get("c_size", 18)
    _l = st.session_state.get("l_size", 26)
    c_size = _c
    l_size = _l

    with btn_chord_m:
        if st.button("🎸－", key="c_minus", help="和弦字體縮小"):
            st.session_state["c_size"] = max(10, _c - 2)
            st.rerun()
    with btn_chord_p:
        if st.button("🎸＋", key="c_plus", help="和弦字體放大"):
            st.session_state["c_size"] = min(60, _c + 2)
            st.rerun()
    with btn_lyric_m:
        if st.button("🎤－", key="l_minus", help="歌詞字體縮小"):
            st.session_state["l_size"] = max(12, _l - 2)
            st.rerun()
    with btn_lyric_p:
        if st.button("🎤＋", key="l_plus", help="歌詞字體放大"):
            st.session_state["l_size"] = min(72, _l + 2)
            st.rerun()

    with btn_fs:
        if st.button(T["fullscreen_btn"], key="enter_fs"):
            st.session_state.is_fullscreen = True
            st.rerun()

    # 計算目前顯示的轉調步數（給匯出函數用）
    _bk_exp  = st.session_state.get('buffer_key', 'C')
    _tk_exp  = st.session_state.meta.get('target', 'C')
    _steps_exp = (KEYS.index(_tk_exp) - KEYS.index(_bk_exp)) % 12

    with btn_img:
        if st.session_state.buffer:
            try:
                img_bytes = generate_chart_image(
                    st.session_state.buffer,
                    st.session_state.meta,
                    COLOR_MAP,
                    _steps_exp
                )
                fname_img = re.sub(r'[\\/*?:"<>|]', '_',
                    f"{st.session_state.meta.get('song','chart')}_{_tk_exp}.png")
                st.download_button("🖼️ 圖片", data=img_bytes,
                    file_name=fname_img, mime="image/png", key="dl_img")
            except Exception as e:
                st.caption(f"圖片錯誤:{e}")

    with btn_html:
        if st.session_state.buffer:
            html_bytes = generate_chart_html(
                st.session_state.buffer,
                st.session_state.meta,
                COLOR_MAP,
                _steps_exp
            )
            fname_html = re.sub(r'[\\/*?:"<>|]', '_',
                f"{st.session_state.meta.get('song','chart')}_{_tk_exp}.html")
            st.download_button("📄 PDF", data=html_bytes,
                file_name=fname_html, mime="text/html", key="dl_html")

    if st.session_state.buffer:
        # 即時計算「目前 buffer 的調」到「目標調」的半音差，動態轉調顯示
        bk = st.session_state.get('buffer_key', 'C')
        target_k = st.session_state.meta.get('target', 'C')
        display_steps = (KEYS.index(target_k) - KEYS.index(bk)) % 12
        display_buffer = transpose_engine(st.session_state.buffer, display_steps) if display_steps else st.session_state.buffer

        html_lines = ['<div id="stage" class="stage-paper">',
                      f'<button class="fs-close" onclick="document.getElementById(\'stage\').classList.remove(\'fs\')">{T["exit_fullscreen"]}</button>']

        for line in display_buffer.split('\n'):
            if not line.strip():
                continue
            # 過濾：歌曲資訊 header 行（作詞:/作曲:/歌手:/原調:/BPM:/拍號:）
            if re.match(r'^\s*(歌手|作詞|作曲|原調|BPM|拍號|Capo)[：:：]', line):
                continue
            # 過濾：吉他 tab 行（E|/B|/G|/D|/A| 開頭）
            if re.match(r'^\s*[EBGDAe]\s*\|', line):
                continue
            # 過濾：純英文/符號行（沒有中文字也沒有和弦標記）
            has_chinese = bool(re.search(r'[一-鿿㐀-䶿]', line))
            has_chord   = bool(re.search(r'\[[A-G][^\]]*\]', line))
            if not has_chinese and not has_chord:
                continue
            parts = re.split(r'(\[[^\]]+\])', line)
            pending_chord = ""
            line_html = '<div class="chord-line">'
            for p in parts:
                if p.startswith('[') and p.endswith(']'):
                    pending_chord = p[1:-1]
                else:
                    for char in p:
                        has_chord = bool(pending_chord)
                        if has_chord:
                            root = pending_chord[0].upper()
                            color = COLOR_MAP.get(root, "#334155")
                            display_c = pending_chord
                            chord_html = (
                                f'<span class="c-tag" style="background-color:{color};'
                                f'font-size:clamp(10px,{c_size}px,{c_size}px);">{display_c}</span>'
                            )
                        else:
                            chord_html = f'<span class="c-tag-empty"></span>'
                        char_disp = "&nbsp;" if char == " " else char
                        line_html += (
                            f'<div class="char-unit">'
                            f'{chord_html}'
                            f'<span class="l-tag" style="font-size:clamp(14px,{l_size}px,{l_size}px);">{char_disp}</span>'
                            f'</div>'
                        )
                        pending_chord = ""
            line_html += '</div>'
            html_lines.append(line_html)

        html_lines.append('</div>')
        st.markdown('\n'.join(html_lines), unsafe_allow_html=True)
    else:
        st.markdown('<div class="stage-paper">', unsafe_allow_html=True)
        st.info(T["no_song_info"])
        st.markdown('</div>', unsafe_allow_html=True)

# ── Tab 3：曲庫管理 ──
with tab_cloud:
    m = st.session_state.meta
    song_n   = m.get('song','').strip()
    singer_n = m.get('singer','').strip()

    # ── 儲存區 ──
    save_col, info_col = st.columns([1, 2])
    with save_col:
        st.markdown(T["save_title"])
        if not song_n:
            st.warning(T["save_warn"])
        else:
            safe_name = re.sub(r'[\\/*?:"<>|]', '_', f"{song_n}_{singer_n}")
            txt_path  = os.path.join(STORAGE_DIR, f"{safe_name}.txt")
            if st.button(T["save_btn"]):
                try:
                    with open(txt_path, "w", encoding="utf-8") as f:
                        f.write(
                            f"SINGER:{m['singer']}\nSONG:{m['song']}\n"
                            f"LYRICIST:{m.get('lyricist','')}\n"
                            f"COMPOSER:{m.get('composer','')}\n"
                            f"BPM:{m['bpm']}\nBEAT:{m['beat']}\n"
                            f"ORIG:{m['orig']}\nTARGET:{m['target']}\n"
                            f"---\n{st.session_state.buffer}"
                        )
                    st.success(T["save_ok"].format(name=safe_name))
                except Exception as e:
                    st.error(T["save_fail"].format(err=str(e)))

    with info_col:
        if song_n:
            st.markdown(f"**{T['cur_song_label'].format(song=song_n)}**　🎤 {singer_n}")
            lyr = m.get('lyricist',''); cmp = m.get('composer','')
            if lyr or cmp:
                st.caption(T["lyricist_disp"].format(lyr=lyr or '—', cmp=cmp or '—'))
            st.caption(T["meta_caption"].format(orig=m['orig'], target=m['target'], bpm=m['bpm'], beat=m['beat']))

    st.markdown("---")
    st.markdown(T["lib_title"])

    files = sorted([f for f in os.listdir(STORAGE_DIR) if f.endswith('.txt')]) if os.path.exists(STORAGE_DIR) else []

    if not files:
        st.info(T["lib_empty"])
    else:
        st.caption(T["lib_count"].format(n=len(files)))
        for file_name in files:
            fpath = os.path.join(STORAGE_DIR, file_name)
            # 快速讀取 metadata（不讀整個 buffer）
            fmeta = {}
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    for line in f:
                        if line.strip() == "---":
                            break
                        if ":" in line:
                            k, v = line.strip().split(":", 1)
                            fmeta[k] = v
            except Exception:
                pass

            disp_song   = fmeta.get("SONG",   file_name.replace(".txt",""))
            disp_singer = fmeta.get("SINGER", "")
            disp_orig   = fmeta.get("ORIG",   "?")
            disp_bpm    = fmeta.get("BPM",    "?")
            disp_lyr    = fmeta.get("LYRICIST","")
            disp_cmp    = fmeta.get("COMPOSER","")

            with st.container(border=True):
                title_col, btn_col = st.columns([5, 1])
                with title_col:
                    st.markdown(f"**《{disp_song}》** 　🎤 {disp_singer}")
                    meta_parts = [f"🎸 {disp_orig}", f"⏱️ {disp_bpm}"]
                    if disp_lyr:  meta_parts.append(f"{T['lyricist_short']} {disp_lyr}")
                    if disp_cmp:  meta_parts.append(f"{T['composer_short']} {disp_cmp}")
                    st.caption("　|　".join(meta_parts))
                with btn_col:
                    load_key = f"load_{file_name}"
                    del_key  = f"del_{file_name}"
                    if st.button(T["load_btn"], key=load_key):
                        with open(fpath, "r", encoding="utf-8") as f:
                            raw = f.read()
                        parts_file = raw.split("---\n", 1)
                        loaded = parts_file[1] if len(parts_file) == 2 else raw
                        st.session_state.buffer     = loaded
                        st.session_state['_sync_editor'] = True
                        st.session_state.buffer_key = fmeta.get("ORIG", "C")
                        st.session_state.meta.update({
                            "singer":   fmeta.get("SINGER",   ""),
                            "song":     fmeta.get("SONG",     ""),
                            "lyricist": fmeta.get("LYRICIST", ""),
                            "composer": fmeta.get("COMPOSER", ""),
                            "bpm":      int(fmeta.get("BPM",  80)),
                            "beat":     fmeta.get("BEAT",     "4/4"),
                            "orig":     fmeta.get("ORIG",     "C"),
                            "target":   fmeta.get("TARGET",   "C"),
                        })
                        st.rerun()
                    if st.button(T["del_btn"], key=del_key, help=f"{T['del_done'].format(song=disp_song)}"):
                        os.remove(fpath)
                        st.toast(T["del_done"].format(song=disp_song))
                        st.rerun()

# ── 自動捲動 JS ──
if scroll_spd > 0:
    st.markdown(
        f"<script>if(window.si)clearInterval(window.si);"
        f"window.si=setInterval(()=>window.scrollBy(0,{scroll_spd}),50);</script>",
        unsafe_allow_html=True
    )
