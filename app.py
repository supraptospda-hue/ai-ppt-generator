# ============================================================
# AI POWERPOINT GENERATOR - PRO EDITION
# Kategori + Font Size + Progress Bar + Animasi Per-Kalimat
# ============================================================
import io
import json
import os
import re
import time
from typing import Optional

import requests
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from groq import Groq
from lxml import etree

# ============================================================
# 1. KONFIGURASI & SECRETS
# ============================================================
st.set_page_config(
    page_title="AI PowerPoint Generator Pro",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def _get_secret(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.getenv(name, default)


GROQ_API_KEY = _get_secret("GROQ_API_KEY", "")
UNSPLASH_ACCESS_KEY = _get_secret("UNSPLASH_ACCESS_KEY", "")
PEXELS_API_KEY = _get_secret("PEXELS_API_KEY", "")
PIXABAY_API_KEY = _get_secret("PIXABAY_API_KEY", "")

# ============================================================
# 2. KATEGORI & FONT SIZE PRESETS
# ============================================================
CATEGORIES = {
    "Umum": "",
    "Pendidikan": "Fokus pada aspek pendidikan, pembelajaran, siswa, guru, kurikulum, dan pengembangan akademik.",
    "Pengetahuan": "Fokus pada wawasan umum, sains, teknologi, fakta menarik, dan pengetahuan luas.",
    "Sejarah": "Fokus pada peristiwa bersejarah, tokoh penting, kronologi, dan pelajaran dari masa lalu.",
    "Dongeng": "Fokus pada cerita rakyat, legenda, fabel, dan narasi imajinatif yang mendidik.",
    "Berita": "Fokus pada peristiwa terkini, analisis berita, data faktual, dan perkembangan terbaru.",
    "Custom (Manual)": "Ikuti instruksi bahan materi pengguna secara ketat.",
}

# Skala font (dalam Pt) untuk judul slide & body
FONT_SCALES = {
    "Kecil":        {"title": 24, "body": 14, "quote": 20, "stat_value": 36, "cover": 36,
                     "subtitle": 14, "author": 14, "timeline_year": 18, "timeline_text": 11,
                     "stat_label": 12, "closing": 44, "closing_sub": 18},
    "Sedang":       {"title": 28, "body": 16, "quote": 24, "stat_value": 44, "cover": 42,
                     "subtitle": 16, "author": 16, "timeline_year": 20, "timeline_text": 12,
                     "stat_label": 13, "closing": 54, "closing_sub": 20},
    "Besar":        {"title": 32, "body": 19, "quote": 28, "stat_value": 52, "cover": 48,
                     "subtitle": 18, "author": 18, "timeline_year": 23, "timeline_text": 14,
                     "stat_label": 15, "closing": 62, "closing_sub": 24},
    "Sangat Besar": {"title": 38, "body": 22, "quote": 32, "stat_value": 60, "cover": 54,
                     "subtitle": 20, "author": 20, "timeline_year": 26, "timeline_text": 16,
                     "stat_label": 17, "closing": 72, "closing_sub": 28},
}

# ============================================================
# 3. TEMA WARNA
# ============================================================
THEMES = {
    "Professional Blue": {
        "primary": "#0F172A", "accent": "#2563EB",
        "text": "#334155", "bg": "#FFFFFF", "soft": "#F1F5F9",
    },
    "Modern Purple": {
        "primary": "#1E1B4B", "accent": "#7C3AED",
        "text": "#3F3F46", "bg": "#FFFFFF", "soft": "#F5F3FF",
    },
    "Elegant Emerald": {
        "primary": "#064E3B", "accent": "#10B981",
        "text": "#334155", "bg": "#FFFFFF", "soft": "#ECFDF5",
    },
    "Bold Sunset": {
        "primary": "#7C2D12", "accent": "#F97316",
        "text": "#3F3F46", "bg": "#FFFBEB", "soft": "#FEF3C7",
    },
    "Minimal Charcoal": {
        "primary": "#111827", "accent": "#6366F1",
        "text": "#4B5563", "bg": "#FFFFFF", "soft": "#F3F4F6",
    },
}

# ============================================================
# 4. CSS - UI FULL WEB (HIJAU + BORDER MERAH MENYALA)
# ============================================================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        /* ===== BACKGROUND HIJAU ===== */
        .stApp {
            background:
                radial-gradient(circle at 15% 10%, rgba(16,185,129,0.30), transparent 45%),
                radial-gradient(circle at 85% 0%, rgba(34,197,94,0.28), transparent 45%),
                radial-gradient(circle at 50% 100%, rgba(20,83,45,0.55), transparent 55%),
                linear-gradient(135deg, #022c22 0%, #064e3b 45%, #065f46 100%);
            background-attachment: fixed;
            min-height: 100vh;
        }

        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 3rem !important;
            max-width: 1400px !important;
        }

        /* ===== HERO ===== */
        .hero { text-align: center; padding: 3rem 1rem 2rem; }
        .hero-badge {
            display: inline-block; padding: 0.4rem 1rem;
            background: rgba(16,185,129,0.20);
            border: 1px solid rgba(34,197,94,0.55);
            color: #bbf7d0; border-radius: 999px;
            font-size: 0.8rem; font-weight: 600;
            letter-spacing: 0.05em; margin-bottom: 1.2rem;
            backdrop-filter: blur(10px);
        }
        .hero h1 {
            font-size: 3.2rem; font-weight: 800;
            background: linear-gradient(135deg, #ffffff 0%, #a7f3d0 50%, #6ee7b7 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin: 0 0 0.8rem 0; line-height: 1.15;
        }
        .hero p { color: #d1fae5; font-size: 1.1rem; max-width: 700px; margin: 0 auto; }

        /* ===== FEATURE CARDS ===== */
        .feature-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1rem; margin: 1.5rem 0 2.5rem;
        }
        .feature-card {
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(34,197,94,0.30);
            border-radius: 16px; padding: 1.2rem 1.4rem;
            backdrop-filter: blur(12px); transition: all 0.25s ease;
        }
        .feature-card:hover {
            transform: translateY(-3px);
            border-color: rgba(34,197,94,0.75);
            background: rgba(16,185,129,0.12);
        }
        .feature-card .icon { font-size: 1.6rem; margin-bottom: 0.4rem; }
        .feature-card .title { color: #ffffff; font-weight: 600; font-size: 0.95rem; }
        .feature-card .desc { color: #d1fae5; font-size: 0.82rem; margin-top: 0.2rem; }

        /* ===== FORM BOX ===== */
        [data-testid="stForm"] {
            background: rgba(255,255,255,0.97) !important;
            padding: 2rem !important;
            border-radius: 20px !important;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.6) !important;
            border: 2px solid #16a34a !important;
        }

        [data-testid="stForm"] label {
            color: #0f172a !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
        }

        /* ===== INPUT: BORDER MERAH MENYALA + FONT HITAM ===== */
        [data-testid="stForm"] input,
        [data-testid="stForm"] textarea,
        [data-testid="stForm"] select,
        [data-testid="stForm"] .stTextInput input,
        [data-testid="stForm"] .stNumberInput input,
        [data-testid="stForm"] .stTextArea textarea,
        [data-testid="stForm"] div[data-baseweb="select"] > div,
        [data-testid="stForm"] div[data-baseweb="input"] input {
            background-color: #ffffff !important;
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
            caret-color: #dc2626 !important;
            border: 2px solid #ff0000 !important;
            border-radius: 10px !important;
            font-size: 0.95rem !important;
            font-weight: 500 !important;
            box-shadow: 0 0 6px rgba(255, 0, 0, 0.55), 
                        0 0 12px rgba(255, 0, 0, 0.25) !important;
            transition: all 0.2s ease-in-out !important;
        }

        [data-testid="stForm"] input:focus,
        [data-testid="stForm"] textarea:focus,
        [data-testid="stForm"] div[data-baseweb="select"] > div:focus-within {
            border: 2px solid #ff1744 !important;
            box-shadow: 0 0 10px rgba(255, 0, 0, 0.85), 
                        0 0 20px rgba(255, 0, 0, 0.45) !important;
            outline: none !important;
        }

        /* Placeholder text lebih gelap */
        [data-testid="stForm"] input::placeholder,
        [data-testid="stForm"] textarea::placeholder {
            color: #6b7280 !important;
            opacity: 1 !important;
        }

        /* Dropdown value text */
        [data-testid="stForm"] div[data-baseweb="select"] span {
            color: #000000 !important;
        }

        /* Number input +/- buttons */
        [data-testid="stForm"] button[data-testid="stNumberInput-StepUp"],
        [data-testid="stForm"] button[data-testid="stNumberInput-StepDown"] {
            background-color: #ffffff !important;
            color: #000000 !important;
            border: 2px solid #ff0000 !important;
        }

        /* ===== SUBMIT BUTTON ===== */
        [data-testid="stForm"] button[type="submit"] {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
            color: #ffffff !important;
            border-radius: 10px !important;
            border: none !important;
            font-weight: 700 !important;
            padding: 0.75rem 1.5rem !important;
            width: 100% !important;
            font-size: 1rem !important;
            transition: all 0.2s ease-in-out !important;
            box-shadow: 0 8px 20px -4px rgba(220, 38, 38, 0.6) !important;
        }
        [data-testid="stForm"] button[type="submit"]:hover {
            transform: translateY(-1px);
            box-shadow: 0 12px 28px -4px rgba(220, 38, 38, 0.85) !important;
        }

        /* ===== DOWNLOAD BUTTON ===== */
        .stDownloadButton button {
            background: linear-gradient(135deg, #16a34a 0%, #15803d 100%) !important;
            color: #ffffff !important;
            border-radius: 10px !important;
            border: none !important;
            font-weight: 700 !important;
            width: 100% !important;
            padding: 0.75rem 1.5rem !important;
            font-size: 1rem !important;
            box-shadow: 0 8px 20px -4px rgba(22, 163, 74, 0.6) !important;
        }
        .stDownloadButton button:hover {
            transform: translateY(-1px);
            box-shadow: 0 12px 28px -4px rgba(22, 163, 74, 0.85) !important;
        }

        .stAlert { border-radius: 12px !important; }

        .section-title {
            color: #ecfdf5; font-size: 1.1rem; font-weight: 700;
            margin: 1.5rem 0 0.8rem; display: flex; align-items: center; gap: 0.5rem;
        }
        .section-title::before {
            content: ""; width: 4px; height: 22px;
            background: linear-gradient(180deg, #ef4444, #dc2626); border-radius: 4px;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        .preview-label {
            color: #d1fae5; font-size: 0.8rem; font-weight: 600;
            text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 0.5rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 5. AI CONTENT GENERATION
# ============================================================
def generate_presentation_content(topic: str, num_slides: int, optional_material: str,
                                   language: str = "Indonesia", category: str = "Umum"):
    client = Groq(api_key=GROQ_API_KEY)

    category_context = CATEGORIES.get(category, "")
    if category.startswith("Custom: "):
        category_context = f"Ikuti instruksi bahan materi pengguna secara ketat: {category[8:]}"

    layout_guide = """
    - "title_bullets": judul + 4-6 bullet points
    - "two_column": judul + 2 kolom masing-masing 3 poin
    - "quote": judul + 1 kutipan inspiratif + atribusi
    - "stats": judul + 3 angka/statistik penting
    - "timeline": judul + 4 milestone berurutan
    - "image_left" / "image_right": judul + 3 poin + 1 gambar besar
    - "auto": AI bebas memilih
    """

    prompt = f"""
Kamu adalah desainer presentasi profesional tingkat dunia.
Buat konten presentasi dalam bahasa {language}.

KATEGORI: {category}
KONTEKS KATEGORI: {category_context if category_context else "Umum / bebas."}

DETAIL:
- Judul: {topic}
- Jumlah slide isi: {num_slides}
- Bahan materi: {optional_material if optional_material else "Tidak ada. Kembangkan materi relevan, mendalam, dan faktual sesuai kategori."}

PILIHAN LAYOUT:
{layout_guide}

FORMAT OUTPUT — JSON SAJA:
{{
  "slides": [
    {{
      "title": "Judul Slide",
      "layout": "title_bullets",
      "points": ["poin 1", "poin 2", "poin 3", "poin 4"],
      "columns": {{"left": ["a","b"], "right": ["c","d"]}},
      "quote": "kutipan inspiratif",
      "author": "Nama Tokoh",
      "stats": [
        {{"value": "85%", "label": "Deskripsi singkat"}},
        {{"value": "2.5x", "label": "Deskripsi singkat"}},
        {{"value": "10jt", "label": "Deskripsi singkat"}}
      ],
      "timeline": [
        {{"time": "2020", "text": "Kejadian penting"}},
        {{"time": "2022", "text": "Kejadian penting"}},
        {{"time": "2024", "text": "Kejadian penting"}},
        {{"time": "2026", "text": "Kejadian penting"}}
      ],
      "image_keyword": "english keyword for stock photo",
      "speaker_notes": "Catatan pembicara 2-3 kalimat."
    }}
  ]
}}

ATURAN:
1. Tepat {num_slides} slide di array "slides".
2. Isi field sesuai layout yang dipilih.
3. "image_keyword" WAJIB bahasa Inggris, spesifik & visual.
4. "speaker_notes" untuk presenter.
5. Output HARUS JSON valid.
"""

    priority_models = [
        "llama-3.3-70b-versatile",
        "llama-3.1-8b-instant",
        "mixtral-8x7b-32768",
    ]

    try:
        remote_models = [m.id for m in client.models.list().data]
        valid_remote = [
            m for m in remote_models
            if not any(x in m.lower() for x in ["guard", "whisper", "safeguard", "vision", "tool"])
        ]
        candidate_models = priority_models + [m for m in valid_remote if m not in priority_models]
    except Exception:
        candidate_models = priority_models

    last_error = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model=model_name,
                    response_format={"type": "json_object"},
                    temperature=0.7,
                )
                raw = response.choices[0].message.content
                result_json = json.loads(raw)

                if "slides" in result_json:
                    slides = result_json["slides"]
                elif isinstance(result_json, list):
                    slides = result_json
                else:
                    slides = list(result_json.values())[0]

                normalized = []
                for s in slides:
                    normalized.append({
                        "title": s.get("title", "Untitled"),
                        "layout": s.get("layout", "auto"),
                        "points": s.get("points") or [],
                        "columns": s.get("columns") or {"left": [], "right": []},
                        "quote": s.get("quote", ""),
                        "author": s.get("author", ""),
                        "stats": s.get("stats") or [],
                        "timeline": s.get("timeline") or [],
                        "image_keyword": s.get("image_keyword", topic),
                        "speaker_notes": s.get("speaker_notes", ""),
                    })
                return normalized
            except Exception as e:
                last_error = e
                time.sleep(0.5)
                continue

    raise Exception(f"Gagal memproses dengan Groq. Error terakhir: {last_error}")


# ============================================================
# 6. IMAGE FETCHER
# ============================================================
@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image_bytes(query: str) -> Optional[bytes]:
    query_clean = re.sub(r"[^a-zA-Z0-9\s]", "", query).strip() or "abstract background"

    # Unsplash
    try:
        url = f"https://api.unsplash.com/photos/random?query={query_clean}&client_id={UNSPLASH_ACCESS_KEY}&orientation=landscape"
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            img_url = r.json()["urls"]["regular"]
            img_data = requests.get(img_url, timeout=15).content
            if img_data and len(img_data) > 5000:
                return img_data
    except Exception:
        pass

    # Pexels
    if PEXELS_API_KEY:
        try:
            headers = {"Authorization": PEXELS_API_KEY}
            url = f"https://api.pexels.com/v1/search?query={query_clean}&per_page=1&orientation=landscape"
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code == 200:
                photos = r.json().get("photos", [])
                if photos:
                    img_url = photos[0]["src"]["large"]
                    img_data = requests.get(img_url, timeout=15).content
                    if img_data and len(img_data) > 5000:
                        return img_data
        except Exception:
            pass

    # Pixabay
    if PIXABAY_API_KEY:
        try:
            url = f"https://pixabay.com/api/?key={PIXABAY_API_KEY}&q={query_clean}&image_type=photo&orientation=horizontal&per_page=3"
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                hits = r.json().get("hits", [])
                if hits:
                    img_url = hits[0]["largeImageURL"]
                    img_data = requests.get(img_url, timeout=15).content
                    if img_data and len(img_data) > 5000:
                        return img_data
        except Exception:
            pass

    # Pollinations AI
    try:
        encoded = requests.utils.quote(query_clean)
        url = f"https://image.pollinations.ai/prompt/{encoded}?width=1280&height=720&nologo=true"
        r = requests.get(url, timeout=25)
        if r.status_code == 200 and len(r.content) > 5000:
            return r.content
    except Exception:
        pass

    return None


def make_placeholder_image(width=1280, height=720, text="Image", color="#6366F1"):
    img = Image.new("RGB", (width, height), color)
    draw = ImageDraw.Draw(img)
    for y in range(height):
        alpha = int(40 * (y / height))
        draw.line([(0, y), (width, y)], fill=(max(0, 60 - alpha), max(0, 70 - alpha), 150))
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", 48)
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((width - tw) / 2, (height - th) / 2), text, fill="white", font=font)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()


# ============================================================
# 7. ANIMASI & TRANSISI (SHAPE-LEVEL, PER-KALIMAT)
# ============================================================
NSMAP_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSMAP_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NSMAP_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NSMAP_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _p(tag): return f"{{{NSMAP_P}}}{tag}"
def _a(tag): return f"{{{NSMAP_A}}}{tag}"


def _add_transition(slide, effect: str = "fade", duration_ms: int = 700):
    """Tambah efek transisi antar slide."""
    sld = slide._element

    old = sld.find(_p("transition"))
    if old is not None:
        sld.remove(old)
    for ac in sld.findall(f"{{{NSMAP_MC}}}AlternateContent"):
        choice = ac.find(f"{{{NSMAP_MC}}}Choice")
        if choice is not None and choice.find(_p("transition")) is not None:
            sld.remove(ac)

    xml = f'''
    <mc:AlternateContent
        xmlns:mc="{NSMAP_MC}"
        xmlns:p="{NSMAP_P}"
        xmlns:p14="{NSMAP_P14}">
      <mc:Choice Requires="p14">
        <p:transition spd="med" p14:dur="{duration_ms}">
          <p:{effect}/>
        </p:transition>
      </mc:Choice>
      <mc:Fallback>
        <p:transition spd="med">
          <p:{effect}/>
        </p:transition>
      </mc:Fallback>
    </mc:AlternateContent>
    '''
    sld.append(etree.fromstring(xml))


def _build_shape_fade_effect(shape_id, anim_base_id, dur_ms=500):
    """Animasi fade-in untuk SATU shape (yang berisi satu kalimat)."""
    b = anim_base_id
    return f'''
    <p:par>
      <p:cTn id="{b}" fill="hold">
        <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
        <p:childTnLst>
          <p:par>
            <p:cTn id="{b+1}" fill="hold">
              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="{b+2}" presetID="10" presetClass="entr"
                         presetSubtype="0" fill="hold" nodeType="clickEffect">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:set>
                        <p:cBhvr>
                          <p:cTn id="{b+3}" dur="1" fill="hold">
                            <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          </p:cTn>
                          <p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>
                          <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                        </p:cBhvr>
                        <p:to><p:strVal val="visible"/></p:to>
                      </p:set>
                      <p:animEffect transition="in" filter="fade">
                        <p:cBhvr>
                          <p:cTn id="{b+4}" dur="{dur_ms}"/>
                          <p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>
                        </p:cBhvr>
                      </p:animEffect>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:childTnLst>
      </p:cTn>
    </p:par>
    '''


def add_animations_to_slide(slide, shape_order=None):
    """Tambahkan animasi fade-in untuk SEMUA shape di slide."""
    sld = slide._element

    old_timing = sld.find(_p("timing"))
    if old_timing is not None:
        sld.remove(old_timing)

    if shape_order:
        shape_map = {}
        for shp in slide.shapes:
            try:
                shape_map[shp.shape_id] = shp
            except Exception:
                continue
        targets = [shape_map[sid] for sid in shape_order if sid in shape_map]
    else:
        targets = []
        for shp in slide.shapes:
            try:
                shp.shape_id
            except Exception:
                continue
            try:
                if shp.has_text_frame:
                    if any(p.text.strip() for p in shp.text_frame.paragraphs):
                        targets.append(shp)
            except Exception:
                continue

    if not targets:
        return

    effects_xml = []
    anim_counter = 10

    for shp in targets:
        try:
            sp_id = shp.shape_id
        except Exception:
            continue
        effects_xml.append(_build_shape_fade_effect(sp_id, anim_counter))
        anim_counter += 10

    main_seq_children = "\n".join(effects_xml)

    timing_xml = f'''
    <p:timing xmlns:p="{NSMAP_P}" xmlns:a="{NSMAP_A}">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {main_seq_children}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>
                  <p:cond evt="onPrev" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:prevCondLst>
                <p:nextCondLst>
                  <p:cond evt="onNext" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>
    '''

    sld.append(etree.fromstring(timing_xml))


def _fix_slide_xml_order(slide):
    """Paksa urutan: cSld → clrMapOvr → transition → timing."""
    sld = slide._element

    csld = sld.find(_p("cSld"))
    clr_map = sld.find(_p("clrMapOvr"))
    transition = sld.find(_p("transition"))
    timing = sld.find(_p("timing"))

    alt_transition = None
    for ac in sld.findall(f"{{{NSMAP_MC}}}AlternateContent"):
        choice = ac.find(f"{{{NSMAP_MC}}}Choice")
        if choice is not None and choice.find(_p("transition")) is not None:
            alt_transition = ac
            break

    transition_element = alt_transition if alt_transition is not None else transition

    if transition_element is None and timing is None:
        return

    if csld is not None:
        sld.remove(csld)
    if clr_map is not None:
        sld.remove(clr_map)
    if transition_element is not None:
        sld.remove(transition_element)
    if timing is not None:
        sld.remove(timing)

    other_children = []
    for child in list(sld):
        if child is csld or child is clr_map or child is transition_element or child is timing:
            continue
        other_children.append(child)
        sld.remove(child)

    if csld is not None:
        sld.append(csld)
    if clr_map is not None:
        sld.append(clr_map)
    if transition_element is not None:
        sld.append(transition_element)
    if timing is not None:
        sld.append(timing)
    for oc in other_children:
        sld.append(oc)


# ============================================================
# 8. BUILDER PPTX (SHAPE-PER-KALIMAT UNTUK ANIMASI)
# ============================================================
def hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tb, tf


def add_rect(slide, left, top, width, height, fill_hex, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if not line:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_rounded(slide, left, top, width, height, fill_hex):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_single_line_textbox(slide, left, top, width, height, text,
                            font_size, font_color, bold=False,
                            bullet_char="•", align=PP_ALIGN.LEFT,
                            anchor=MSO_ANCHOR.TOP):
    """Buat textbox dengan SATU baris teks (1 paragraf). Untuk animasi per-kalimat."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    if bullet_char:
        p.text = f"{bullet_char}  {text}"
    else:
        p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = "Calibri"
    p.alignment = align
    return tb


def build_presentation(topic, slides_data, theme: dict, animate: bool = True,
                       transition: str = "fade", font_scale: dict = None):
    if font_scale is None:
        font_scale = FONT_SCALES["Sedang"]

    FS_TITLE = font_scale["title"]
    FS_BODY = font_scale["body"]
    FS_QUOTE = font_scale["quote"]
    FS_STAT = font_scale["stat_value"]
    FS_COVER = font_scale["cover"]
    FS_SUBTITLE = font_scale["subtitle"]
    FS_AUTHOR = font_scale["author"]
    FS_TL_YEAR = font_scale["timeline_year"]
    FS_TL_TEXT = font_scale["timeline_text"]
    FS_STAT_LABEL = font_scale["stat_label"]
    FS_CLOSING = font_scale["closing"]
    FS_CLOSING_SUB = font_scale["closing_sub"]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    COLOR_PRIMARY = hex_to_rgb(theme["primary"])
    COLOR_ACCENT = hex_to_rgb(theme["accent"])
    COLOR_TEXT = hex_to_rgb(theme["text"])
    COLOR_SOFT = hex_to_rgb(theme["soft"])

    blank_layout = prs.slide_layouts[6]
    SW, SH = prs.slide_width, prs.slide_height

    # ==================== COVER ====================
    cover = prs.slides.add_slide(blank_layout)
    add_rect(cover, 0, 0, SW, SH, theme["bg"])
    add_rect(cover, Inches(0), Inches(0), Inches(0.35), SH, theme["accent"])

    add_single_line_textbox(
        cover, Inches(1.1), Inches(2.1), Inches(5.8), Inches(1.5),
        topic, FS_COVER, COLOR_PRIMARY, bold=True, bullet_char="",
        align=PP_ALIGN.LEFT
    )
    add_single_line_textbox(
        cover, Inches(1.1), Inches(3.8), Inches(5.8), Inches(0.5),
        "AI Generated Presentation", FS_SUBTITLE, COLOR_ACCENT, bold=True, bullet_char=""
    )
    add_single_line_textbox(
        cover, Inches(1.1), Inches(4.4), Inches(5.8), Inches(0.4),
        f"Theme: {theme.get('name', 'Custom')}", max(10, FS_SUBTITLE - 4), COLOR_TEXT, bullet_char=""
    )

    add_rect(cover, Inches(1.1), Inches(2.0), Inches(0.9), Inches(0.06), theme["accent"])

    cover_bytes = fetch_image_bytes(topic)
    if not cover_bytes:
        cover_bytes = make_placeholder_image(1280, 720, topic[:40], theme["accent"])
    cover_img = io.BytesIO(cover_bytes)
    cover.shapes.add_picture(cover_img, Inches(7.0), Inches(0), Inches(6.333), SH)

    overlay = add_rect(cover, Inches(7.0), Inches(0), Inches(6.333), SH, theme["primary"])
    overlay.fill.fore_color.rgb = COLOR_PRIMARY
    overlay.fill.transparency = 0.55

    if animate:
        add_animations_to_slide(cover)
    _add_transition(cover, transition, 700)
    _fix_slide_xml_order(cover)

    # ==================== CONTENT SLIDES ====================
    for item in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        add_rect(slide, 0, 0, SW, SH, theme["bg"])
        add_rect(slide, 0, 0, SW, Inches(0.15), theme["accent"])

        layout = (item.get("layout") or "auto").strip()
        if layout == "auto":
            if item.get("stats"):
                layout = "stats"
            elif item.get("timeline"):
                layout = "timeline"
            elif item.get("quote"):
                layout = "quote"
            elif item.get("columns", {}).get("left"):
                layout = "two_column"
            else:
                layout = "title_bullets"

        title_text = item.get("title", "")
        img_bytes = fetch_image_bytes(item.get("image_keyword", topic))
        if not img_bytes:
            img_bytes = make_placeholder_image(1280, 720, item.get("image_keyword", "image")[:30], theme["accent"])

        # TITLE (1 shape)
        add_single_line_textbox(
            slide, Inches(0.7), Inches(0.5), Inches(11.9), Inches(0.9),
            title_text, FS_TITLE, COLOR_PRIMARY, bold=True, bullet_char=""
        )
        add_rect(slide, Inches(0.7), Inches(1.42), Inches(0.7), Inches(0.06), theme["accent"])

        # ---- LAYOUT LOGIC (SETIAP KALIMAT = 1 TEXTBOX) ----

        if layout == "title_bullets":
            points = item["points"][:6]
            y_start = Inches(2.0)
            line_height = Inches(0.7)
            for idx, point in enumerate(points):
                add_single_line_textbox(
                    slide,
                    Inches(0.7), y_start + line_height * idx,
                    Inches(6.5), Inches(0.65),
                    point, FS_BODY, COLOR_TEXT, bullet_char="•"
                )
            img = io.BytesIO(img_bytes)
            slide.shapes.add_picture(img, Inches(7.6), Inches(1.9), Inches(5.0), Inches(4.8))

        elif layout == "two_column":
            cols = item.get("columns", {})
            left_points = cols.get("left") or item.get("points", [])[: len(item.get("points", [])) // 2]
            right_points = cols.get("right") or item.get("points", [])[len(item.get("points", [])) // 2:]

            add_rounded(slide, Inches(0.7), Inches(1.9), Inches(5.9), Inches(4.8), theme["soft"])
            add_rounded(slide, Inches(6.9), Inches(1.9), Inches(5.7), Inches(4.8), theme["soft"])

            for idx, point in enumerate(left_points[:5]):
                add_single_line_textbox(
                    slide,
                    Inches(1.0), Inches(2.2) + Inches(0.7) * idx,
                    Inches(5.3), Inches(0.65),
                    point, max(12, FS_BODY - 1), COLOR_TEXT, bullet_char="●"
                )
            for idx, point in enumerate(right_points[:5]):
                add_single_line_textbox(
                    slide,
                    Inches(7.2), Inches(2.2) + Inches(0.7) * idx,
                    Inches(5.1), Inches(0.65),
                    point, max(12, FS_BODY - 1), COLOR_TEXT, bullet_char="●"
                )

        elif layout == "quote":
            add_rounded(slide, Inches(1.2), Inches(2.2), Inches(10.9), Inches(3.8), theme["soft"])
            quote_text = item.get("quote") or (item["points"][0] if item["points"] else title_text)
            add_single_line_textbox(
                slide, Inches(1.8), Inches(2.8), Inches(9.7), Inches(1.8),
                f'"{quote_text}"', FS_QUOTE, COLOR_PRIMARY, bullet_char="",
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE
            )
            add_single_line_textbox(
                slide, Inches(1.8), Inches(4.8), Inches(9.7), Inches(0.6),
                f"— {item.get('author', 'Anonim')}", FS_AUTHOR, COLOR_ACCENT, bold=True,
                bullet_char="", align=PP_ALIGN.CENTER
            )

        elif layout == "stats":
            stats = item.get("stats") or []
            if not stats:
                stats = [{"value": "—", "label": "No data"}]
            n = min(len(stats), 3)
            card_w = Inches(3.9)
            gap = Inches(0.35)
            total_w = card_w * n + gap * (n - 1)
            start_x = (SW - total_w) / 2

            for idx, stat in enumerate(stats[:3]):
                x = start_x + (card_w + gap) * idx
                add_rounded(slide, x, Inches(2.4), card_w, Inches(3.6), theme["soft"])
                add_single_line_textbox(
                    slide, x, Inches(2.8), card_w, Inches(1.6),
                    str(stat.get("value", "—")), FS_STAT, COLOR_ACCENT, bold=True,
                    bullet_char="", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE
                )
                add_single_line_textbox(
                    slide, x + Inches(0.3), Inches(4.8), card_w - Inches(0.6), Inches(1.0),
                    stat.get("label", ""), FS_STAT_LABEL, COLOR_TEXT,
                    bullet_char="", align=PP_ALIGN.CENTER
                )

        elif layout == "timeline":
            tl = item.get("timeline") or []
            if not tl:
                tl = [{"time": "—", "text": "No data"}]
            n = min(len(tl), 4)
            add_rect(slide, Inches(1.0), Inches(4.1), Inches(11.3), Inches(0.05), theme["accent"])

            step = Inches(11.3) / max(n - 1, 1)
            for idx, t in enumerate(tl[:4]):
                x = Inches(1.0) + step * idx
                dot = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL, x - Inches(0.12), Inches(3.95), Inches(0.35), Inches(0.35)
                )
                dot.fill.solid()
                dot.fill.fore_color.rgb = COLOR_ACCENT
                dot.line.fill.background()

                add_single_line_textbox(
                    slide, x - Inches(0.7), Inches(2.5), Inches(1.8), Inches(1.0),
                    str(t.get("time", "")), FS_TL_YEAR, COLOR_PRIMARY, bold=True,
                    bullet_char="", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE
                )
                add_single_line_textbox(
                    slide, x - Inches(0.9), Inches(4.5), Inches(2.2), Inches(1.8),
                    t.get("text", ""), FS_TL_TEXT, COLOR_TEXT,
                    bullet_char="", align=PP_ALIGN.CENTER
                )

        elif layout in ("image_left", "image_right"):
            img = io.BytesIO(img_bytes)
            points = item["points"][:5]
            if layout == "image_right":
                slide.shapes.add_picture(img, Inches(7.6), Inches(1.9), Inches(5.0), Inches(4.8))
                x_text = Inches(0.7)
                w_text = Inches(6.5)
            else:
                slide.shapes.add_picture(img, Inches(0.7), Inches(1.9), Inches(5.0), Inches(4.8))
                x_text = Inches(6.3)
                w_text = Inches(6.3)

            for idx, point in enumerate(points):
                add_single_line_textbox(
                    slide,
                    x_text, Inches(2.0) + Inches(0.7) * idx,
                    w_text, Inches(0.65),
                    point, FS_BODY, COLOR_TEXT, bullet_char="•"
                )

        # SPEAKER NOTES
        if item.get("speaker_notes"):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = item["speaker_notes"]

        # ANIMASI + TRANSISI
        if animate:
            add_animations_to_slide(slide)
        _add_transition(slide, transition, 700)
        _fix_slide_xml_order(slide)

    # ==================== CLOSING ====================
    end = prs.slides.add_slide(blank_layout)
    add_rect(end, 0, 0, SW, SH, theme["primary"])
    add_rect(end, 0, 0, Inches(0.35), SH, theme["accent"])

    add_single_line_textbox(
        end, Inches(1.5), Inches(2.8), Inches(10.3), Inches(1.2),
        "Terima Kasih", FS_CLOSING, hex_to_rgb("#FFFFFF"), bold=True,
        bullet_char="", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE
    )
    add_single_line_textbox(
        end, Inches(1.5), Inches(4.2), Inches(10.3), Inches(0.8),
        "Questions & Discussion", FS_CLOSING_SUB, COLOR_ACCENT,
        bullet_char="", align=PP_ALIGN.CENTER
    )

    if animate:
        add_animations_to_slide(end)
    _add_transition(end, transition, 700)
    _fix_slide_xml_order(end)

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer


# ============================================================
# 9. PREVIEW THUMBNAIL
# ============================================================
def render_preview(slides_data, topic, theme):
    previews = []
    W, H = 960, 540

    def hex2rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

    primary = hex2rgb(theme["primary"])
    accent = hex2rgb(theme["accent"])
    text_c = hex2rgb(theme["text"])
    bg = hex2rgb(theme["bg"])
    soft = hex2rgb(theme["soft"])

    try:
        f_title = ImageFont.truetype("DejaVuSans-Bold.ttf", 42)
        f_big = ImageFont.truetype("DejaVuSans-Bold.ttf", 28)
        f_body = ImageFont.truetype("DejaVuSans.ttf", 20)
        f_small = ImageFont.truetype("DejaVuSans.ttf", 15)
    except Exception:
        f_title = f_big = f_body = f_small = ImageFont.load_default()

    def wrap(text, font, max_w, draw):
        words = text.split()
        lines, line = [], ""
        for w in words:
            test = (line + " " + w).strip()
            bbox = draw.textbbox((0, 0), test, font=font)
            if bbox[2] - bbox[0] > max_w and line:
                lines.append(line)
                line = w
            else:
                line = test
        if line:
            lines.append(line)
        return lines

    # Cover
    img = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 20, H], fill=accent)
    d.rectangle([80, 140, 100, 170], fill=accent)
    lines = wrap(topic, f_title, 500, d)
    y = 200
    for ln in lines[:3]:
        d.text((80, y), ln, fill=primary, font=f_title)
        y += 52
    d.text((80, y + 10), "AI Generated Presentation", fill=accent, font=f_body)
    d.rectangle([W - 380, 0, W, H], fill=soft)
    d.text((W - 330, H // 2), "[ Cover Image ]", fill=text_c, font=f_body)
    previews.append(img)

    for s in slides_data:
        img = Image.new("RGB", (W, H), bg)
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, W, 8], fill=accent)

        t_lines = wrap(s["title"], f_big, 850, d)
        y = 40
        for ln in t_lines[:2]:
            d.text((50, y), ln, fill=primary, font=f_big)
            y += 38
        d.rectangle([50, y + 6, 90, y + 11], fill=accent)

        layout = s.get("layout", "title_bullets")
        body_y = y + 40

        if layout in ("title_bullets", "auto"):
            pts = s.get("points", [])[:6]
            for pt in pts:
                for ln in wrap("• " + pt, f_body, 500, d)[:2]:
                    d.text((50, body_y), ln, fill=text_c, font=f_body)
                    body_y += 28
                body_y += 8
            d.rectangle([W - 420, 160, W - 50, H - 60], fill=soft)
            d.text((W - 320, H // 2), "[ Image ]", fill=text_c, font=f_body)

        elif layout == "two_column":
            d.rectangle([50, body_y, 470, H - 60], fill=soft)
            d.rectangle([490, body_y, W - 50, H - 60], fill=soft)
            d.text((70, body_y + 15), "Column A", fill=primary, font=f_body)
            d.text((510, body_y + 15), "Column B", fill=primary, font=f_body)

        elif layout == "quote":
            d.rectangle([80, body_y, W - 80, H - 80], fill=soft)
            q = s.get("quote") or (s["points"][0] if s.get("points") else "")
            for i, ln in enumerate(wrap('"' + q + '"', f_body, W - 200, d)[:4]):
                d.text((W // 2 - 350, body_y + 40 + i * 30), ln, fill=primary, font=f_body)
            d.text((W // 2 - 100, H - 130), f"— {s.get('author', 'Anonim')}",
                   fill=accent, font=f_small)

        elif layout == "stats":
            stats = s.get("stats", [])[:3]
            cw = (W - 100 - 40) // 3
            for i, st in enumerate(stats):
                x = 50 + i * (cw + 20)
                d.rectangle([x, body_y, x + cw, H - 80], fill=soft)
                d.text((x + cw // 2 - 30, body_y + 40), str(st.get("value", "—")),
                       fill=accent, font=f_title)
                for j, ln in enumerate(wrap(st.get("label", ""), f_small, cw - 20, d)[:3]):
                    d.text((x + 10, body_y + 120 + j * 20), ln, fill=text_c, font=f_small)

        elif layout == "timeline":
            d.rectangle([50, body_y + 80, W - 50, body_y + 85], fill=accent)
            tl = s.get("timeline", [])[:4]
            for i, t in enumerate(tl):
                x = 50 + i * ((W - 100) // max(len(tl) - 1, 1))
                d.ellipse([x - 10, body_y + 72, x + 10, body_y + 92], fill=accent)
                d.text((x - 25, body_y + 20), str(t.get("time", "")), fill=primary, font=f_small)
                for j, ln in enumerate(wrap(t.get("text", ""), f_small, 140, d)[:3]):
                    d.text((x - 60, body_y + 110 + j * 18), ln, fill=text_c, font=f_small)

        elif layout in ("image_left", "image_right"):
            if layout == "image_right":
                d.rectangle([W - 400, body_y, W - 50, H - 60], fill=soft)
                d.text((W - 320, H // 2), "[ Image ]", fill=text_c, font=f_body)
                pts = s.get("points", [])[:5]
                yy = body_y
                for pt in pts:
                    d.text((50, yy), "• " + pt[:60], fill=text_c, font=f_body)
                    yy += 32
            else:
                d.rectangle([50, body_y, 380, H - 60], fill=soft)
                d.text((150, H // 2), "[ Image ]", fill=text_c, font=f_body)
                pts = s.get("points", [])[:5]
                yy = body_y
                for pt in pts:
                    d.text((420, yy), "• " + pt[:60], fill=text_c, font=f_body)
                    yy += 32

        previews.append(img)

    # Closing
    img = Image.new("RGB", (W, H), primary)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 20, H], fill=accent)
    d.text((W // 2 - 180, H // 2 - 40), "Terima Kasih", fill=(255, 255, 255), font=f_title)
    d.text((W // 2 - 150, H // 2 + 30), "Questions & Discussion", fill=accent, font=f_body)
    previews.append(img)

    return previews


# ============================================================
# 10. UI STREAMLIT
# ============================================================
inject_css()

st.markdown(
    """
    <div class="hero">
        <div class="hero-badge">✨ POWERED BY GROQ AI + UNSPLASH</div>
        <h1>AI PowerPoint Generator</h1>
        <p>Ubah ide menjadi presentasi profesional dalam hitungan detik — lengkap dengan gambar relevan, animasi teks, dan efek transisi.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="feature-grid">
        <div class="feature-card"><div class="icon">🎨</div><div class="title">5 Tema Warna</div><div class="desc">Profesional, Modern, Elegant & lainnya</div></div>
        <div class="feature-card"><div class="icon">📂</div><div class="title">6 Kategori Konten</div><div class="desc">Pendidikan, Sejarah, Dongeng, Berita, Custom</div></div>
        <div class="feature-card"><div class="icon">🔠</div><div class="title">Ukuran Font</div><div class="desc">Kecil, Sedang, Besar, Sangat Besar</div></div>
        <div class="feature-card"><div class="icon">🎬</div><div class="title">Animasi & Transisi</div><div class="desc">Teks muncul per kalimat di SEMUA slide</div></div>
        <div class="feature-card"><div class="icon">🖼️</div><div class="title">Gambar Relevan</div><div class="desc">Auto-fetch dari 4 sumber</div></div>
        <div class="feature-card"><div class="icon">👁️</div><div class="title">Preview Live</div><div class="desc">Lihat slide sebelum download</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.form("ppt_form", clear_on_submit=False):
    col1, col2 = st.columns([2, 1])
    with col1:
        judul = st.text_input(
            "📝 Judul Presentasi",
            placeholder="Contoh: Manfaat AI dalam Pendidikan Modern",
        )
    with col2:
        bahasa = st.selectbox("🌐 Bahasa", ["Indonesia", "English", "Bilingual (ID+EN)"])

    col_cat1, col_cat2 = st.columns([1, 1])
    with col_cat1:
        kategori_pilihan = st.selectbox(
            "📂 Kategori Konten",
            list(CATEGORIES.keys()),
            index=0,
            help="Pilih kategori untuk memfokuskan gaya & isi konten"
        )
    with col_cat2:
        kategori_custom = st.text_input(
            "✏️ Kategori Custom (opsional)",
            placeholder="Contoh: Teknologi AI, Kuliner Nusantara...",
            disabled=(kategori_pilihan != "Custom (Manual)"),
        )

    col3, col4, col5 = st.columns([1, 1, 1])
    with col3:
        jumlah_slide = st.number_input(
            "🎞️ Jumlah Slide Isi", min_value=1, max_value=15, value=5
        )
    with col4:
        tema_pilihan = st.selectbox("🎨 Tema Warna", list(THEMES.keys()))
    with col5:
        transisi = st.selectbox("✨ Efek Transisi", ["fade", "push", "wipe", "split", "cover", "dissolve"])

    col_fs1, col_fs2, col_fs3 = st.columns([1, 1, 1])
    with col_fs1:
        font_scale_pilihan = st.selectbox(
            "🔠 Ukuran Font Slide",
            list(FONT_SCALES.keys()),
            index=1,
            help="Perbesar/Perkecil font judul & body di semua slide"
        )
    with col_fs2:
        use_anim = st.checkbox("🎬 Animasi teks (muncul per kalimat)", value=True)
    with col_fs3:
        use_preview = st.checkbox("👁️ Preview thumbnail", value=True)

    materi = st.text_area(
        "📚 Bahan Materi (Opsional)",
        placeholder="Tempelkan poin-poin khusus, outline, atau referensi yang ingin dibahas...",
        height=120,
    )

    submitted = st.form_submit_button("✨ Generate Presentation", use_container_width=True)


# ============================================================
# 11. HANDLE SUBMIT (DENGAN PROGRESS BAR)
# ============================================================
if submitted:
    if not judul.strip():
        st.warning("⚠️ Judul presentasi tidak boleh kosong.")
    elif not GROQ_API_KEY:
        st.error("❌ GROQ_API_KEY belum diset.")
    else:
        try:
            # ====== PROGRESS BAR 0-100% ======
            progress_bar = st.progress(0, text="🚀 Memulai proses...")
            status_text = st.empty()
            time.sleep(0.2)

            # Tentukan kategori final
            if kategori_pilihan == "Custom (Manual)" and kategori_custom.strip():
                kategori_final = f"Custom: {kategori_custom.strip()}"
            else:
                kategori_final = kategori_pilihan

            # --- Tahap 1: 0-15% ---
            progress_bar.progress(5, text="📋 Menyiapkan parameter...")
            time.sleep(0.15)
            progress_bar.progress(15, text=f"🤖 AI menyusun materi ({kategori_final})...")
            time.sleep(0.15)

            # --- Tahap 2: 15-55% ---
            progress_bar.progress(25, text="🧠 AI sedang berpikir & menulis konten...")
            slides_json = generate_presentation_content(
                judul, int(jumlah_slide), materi, bahasa, kategori_final
            )
            progress_bar.progress(55, text=f"✅ Konten selesai! {len(slides_json)} slide dibuat.")
            time.sleep(0.2)

            # --- Tahap 3: 55-70% ---
            progress_bar.progress(60, text="🎨 Menyiapkan tema & font...")
            theme = dict(THEMES[tema_pilihan])
            theme["name"] = tema_pilihan
            font_scale = FONT_SCALES[font_scale_pilihan]
            time.sleep(0.15)

            # --- Tahap 4: 70-95% ---
            progress_bar.progress(70, text="🖼️ Mengunduh gambar dari sumber...")
            time.sleep(0.15)
            progress_bar.progress(80, text="📐 Membangun slide & animasi...")

            ppt_file = build_presentation(
                judul, slides_json, theme,
                animate=use_anim,
                transition=transisi,
                font_scale=font_scale,
            )

            # --- Tahap 5: 100% ---
            progress_bar.progress(95, text="📦 Menyusun file PPTX...")
            time.sleep(0.2)
            progress_bar.progress(100, text="🎉 Presentasi siap diunduh!")

            st.session_state["last_ppt"] = ppt_file
            st.session_state["last_slides"] = slides_json
            st.session_state["last_topic"] = judul
            st.session_state["last_theme"] = theme
            st.session_state["last_preview"] = use_preview

            st.success(
                f"🎉 Presentasi berhasil dibuat! "
                f"Total {len(slides_json) + 2} slide · "
                f"Kategori: **{kategori_final}** · "
                f"Font: **{font_scale_pilihan}**"
            )

            # Auto-hide progress bar
            time.sleep(1.5)
            progress_bar.empty()
            status_text.empty()

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}")


# ============================================================
# 12. HASIL & PREVIEW
# ============================================================
if "last_ppt" in st.session_state:
    slides_json = st.session_state["last_slides"]
    judul = st.session_state["last_topic"]
    theme = st.session_state["last_theme"]

    st.markdown(
        '<div class="section-title">📥 Download Presentasi</div>',
        unsafe_allow_html=True,
    )

    file_name = re.sub(r"[^a-zA-Z0-9_-]", "_", judul)[:50] + ".pptx"
    st.download_button(
        label="📥 Download File PPTX",
        data=st.session_state["last_ppt"],
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

    if st.session_state.get("last_preview", True):
        st.markdown(
            '<div class="section-title">👁️ Preview Slide</div>',
            unsafe_allow_html=True,
        )
        with st.spinner("Rendering preview..."):
            try:
                previews = render_preview(slides_json, judul, theme)
                cols = st.columns(3)
                labels = ["Cover"] + [f"Slide {i+1}" for i in range(len(slides_json))] + ["Closing"]
                for i, img in enumerate(previews):
                    with cols[i % 3]:
                        st.markdown(
                            f'<div class="preview-label">{labels[i] if i < len(labels) else f"Slide {i+1}"}</div>',
                            unsafe_allow_html=True,
                        )
                        st.image(img, use_container_width=True)
            except Exception as e:
                st.warning(f"Preview gagal dirender: {e}")

    with st.expander("🔍 Lihat Detail Konten Slide"):
        for i, s in enumerate(slides_json, 1):
            st.markdown(f"**Slide {i}: {s['title']}** — _layout: `{s.get('layout', 'auto')}`_")
            if s.get("points"):
                for p in s["points"]:
                    st.markdown(f"- {p}")
            if s.get("quote"):
                st.markdown(f"> *\"{s['quote']}\"* — {s.get('author', '')}")
            if s.get("stats"):
                for stt in s["stats"]:
                    st.markdown(f"- **{stt.get('value')}** — {stt.get('label')}")
            if s.get("timeline"):
                for t in s["timeline"]:
                    st.markdown(f"- **{t.get('time')}**: {t.get('text')}")
            if s.get("speaker_notes"):
                st.caption(f"🎤 Speaker notes: {s['speaker_notes']}")
            st.divider()


# ============================================================
# 13. FOOTER
# ============================================================
st.markdown(
    """
    <div style="text-align:center; color:#a7f3d0; font-size:0.8rem; margin-top:3rem;">
        Built with ❤️ using Streamlit · Groq · python-pptx · Pillow
    </div>
    """,
    unsafe_allow_html=True,
)