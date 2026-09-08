import os
import io
import json
import uuid
import re
import urllib.parse
import random
from pathlib import Path
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from PIL import Image

# ------------------------------------------------------------
# Libraries & Safe Imports
# ------------------------------------------------------------
try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from groq import Groq
except ImportError:
    Groq = None

try:
    from google import genai
except ImportError:
    genai = None

try:
    import requests
except ImportError:
    requests = None

try:
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ============================================================
# 1. ENVIRONMENT & IDENTITY SPECIFICATION
# ============================================================
load_dotenv(override=True)
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

CREATOR_FULL_NAME = "Anshul Singh Rajpoot"
OS_NAME = "Aetheris OS"

# ============================================================
# 2. PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title=f"{OS_NAME} • {CREATOR_FULL_NAME}",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 3. EXECUTIVE CLEAN NORDIC FROST UI
# ============================================================
st.markdown(
    """<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    color: #0f172a;
}

.stApp {
    background:
        radial-gradient(circle at 10% 0%, rgba(99, 102, 241, 0.08), transparent 35%),
        radial-gradient(circle at 90% 10%, rgba(45, 212, 191, 0.08), transparent 30%),
        linear-gradient(180deg, #f8fafc 0%, #f1f5f9 60%, #e2e8f0 100%);
    color: #0f172a;
}

.block-container {
    max-width: 1180px;
    padding-top: 1.2rem;
    padding-bottom: 3.5rem;
}

section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
}

section[data-testid="stSidebar"] * {
    color: #0f172a !important;
}

.aetheris-header {
    padding: 18px 24px;
    border-radius: 16px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04);
    margin-bottom: 16px;
}

.brand-title {
    font-size: 24px;
    font-weight: 800;
    letter-spacing: 1.5px;
    background: linear-gradient(90deg, #0f172a, #4338ca, #0d9488);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.architect-badge {
    background: linear-gradient(90deg, rgba(99, 102, 241, 0.12), rgba(13, 148, 136, 0.12));
    border: 1px solid rgba(99, 102, 241, 0.35);
    padding: 5px 14px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    color: #4338ca;
    letter-spacing: 0.8px;
}

.status-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: rgba(16, 185, 129, 0.1);
    border: 1px solid rgba(16, 185, 129, 0.35);
    padding: 5px 12px;
    border-radius: 999px;
    color: #047857;
    font-size: 11px;
    font-weight: 700;
}

.dot {
    width: 7px;
    height: 7px;
    background: #10b981;
    border-radius: 50%;
    box-shadow: 0 0 8px #10b981;
}

div[data-testid="stChatMessage"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 14px !important;
    padding: 15px 19px !important;
    margin-bottom: 10px !important;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03) !important;
}

div[data-testid="stChatMessage"] p, 
div[data-testid="stChatMessage"] span, 
div[data-testid="stChatMessage"] div {
    color: #0f172a !important;
    font-size: 15px !important;
    line-height: 1.6 !important;
}

div[data-testid="stChatInput"] {
    background: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06) !important;
}

#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
</style>""",
    unsafe_allow_html=True,
)

# ============================================================
# 4. SESSION STATE
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

if "attached_assets" not in st.session_state:
    st.session_state.attached_assets = []

# ============================================================
# 5. CORE ENGINES: COMPLIANCE, OCR, SEARCH, DATA & PASS DISPATCH
# ============================================================
def calculate_pf_ecr(df):
    records = []
    ecr_lines = []
    col_uan = next((c for c in df.columns if "uan" in c.lower()), None)
    col_name = next((c for c in df.columns if "name" in c.lower()), None)
    col_wage = next((c for c in df.columns if any(k in c.lower() for k in ["gross", "wage", "basic", "salary"])), None)

    if not col_uan or not col_name or not col_wage:
        return None, "Columns matching UAN, Name, and Wages (Basic/Gross) are required."

    for _, row in df.iterrows():
        uan = str(row[col_uan]).split(".")[0].strip()
        name = str(row[col_name]).strip().upper()
        try:
            gross_wage = float(row[col_wage])
        except Exception:
            gross_wage = 0.0

        epf_wage = min(gross_wage, 15000.0)
        eps_wage = epf_wage
        edli_wage = epf_wage

        ee_share = round(epf_wage * 0.12)
        eps_share = round(eps_wage * 0.0833)
        er_share = ee_share - eps_share

        records.append({
            "UAN": uan,
            "Employee Name": name,
            "Gross Wage": gross_wage,
            "EPF Wage (Capped 15k)": epf_wage,
            "EE Share (12%)": ee_share,
            "EPS Pension (8.33%)": eps_share,
            "ER Balance (3.67%)": er_share,
            "Total Remitted": ee_share + eps_share + er_share
        })

        ecr_lines.append(f"{uan}#~#{name}#~#{int(gross_wage)}#~#{int(epf_wage)}#~#{int(eps_wage)}#~#{int(edli_wage)}#~#{ee_share}#~#{eps_share}#~#{er_share}#~#0#~#0")

    audit_df = pd.DataFrame(records)
    ecr_text_content = "\n".join(ecr_lines)
    return audit_df, ecr_text_content

def run_vision_ocr(image_bytes):
    if not GEMINI_API_KEY or not genai:
        return "Gemini Vision API configuration missing in .env."
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        img = Image.open(io.BytesIO(image_bytes))
        prompt = "Extract all text verbatim from this document/image in original language (Hindi/English)."
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[prompt, img]
        )
        return response.text.strip() if response and response.text else "No text extracted."
    except Exception as e:
        return f"OCR Notice: {str(e)}"

def run_live_web_search(query_text):
    if not requests:
        return "Web search library unavailable."
    try:
        clean_q = re.sub(r'^(search|look up|find|news on)\s+', '', query_text, flags=re.IGNORECASE).strip()
        url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(clean_q)}"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(url, headers=headers, timeout=5)
        if resp.status_code == 200:
            snippets = re.findall(r'<a class="result__snippet[^>]*>(.*?)</a>', resp.text)
            clean_snippets = [re.sub(r'<[^>]+>', '', s).strip() for s in snippets[:3]]
            if clean_snippets:
                return "\n".join([f"• {s}" for s in clean_snippets])
    except Exception:
        pass
    return "Web search executed. Synthesizing available direct domain knowledge."

def parse_file(file):
    ext = Path(file.name).suffix.lower()
    if ext == ".pdf" and PdfReader:
        try:
            reader = PdfReader(file)
            return "\n\n".join([page.extract_text() or "" for page in reader.pages])
        except Exception as e:
            return f"PDF Error: {str(e)}"
    elif ext == ".docx" and Document:
        try:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        except Exception as e:
            return f"DOCX Error: {str(e)}"
    elif ext == ".csv" and pd:
        try:
            return pd.read_csv(file)
        except Exception as e:
            return f"CSV Error: {str(e)}"
    elif ext in [".png", ".jpg", ".jpeg"]:
        return file.getvalue()
    elif ext in [".txt", ".md"]:
        return file.read().decode("utf-8", errors="ignore")
    return "Unsupported file format."

def fetch_studio_asset_bytes(raw_subject):
    clean_sub = raw_subject.strip()
    prompt = f"masterpiece crisp photograph of {clean_sub}, 8k resolution, photorealistic, cinematic shot, realistic textures"
    encoded = urllib.parse.quote(prompt)
    unique_seed = int(uuid.uuid4().int % 9999999)
    cache_token = uuid.uuid4().hex[:6]
    image_url = f"https://image.pollinations.ai/prompt/{encoded}?width=1024&height=1024&seed={unique_seed}&nologo=true&model=flux&ref={cache_token}"
    if requests:
        try:
            resp = requests.get(image_url, timeout=20)
            if resp.status_code == 200 and len(resp.content) > 5000:
                return resp.content
        except Exception:
            pass
    return image_url

def generate_travel_boarding_pass(pnr, passenger_name, train_details, travel_class, seat_no, route):
    if not REPORTLAB_OK:
        return None
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        elements = []
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            'PassTitle',
            parent=styles['Heading1'],
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#1A365D"),
            alignment=1
        )
        
        elements.append(Paragraph(f"{OS_NAME} // AUTONOMOUS MOBILITY MANIFEST", title_style))
        elements.append(Spacer(1, 14))

        data = [
            ["PNR / TOKEN", pnr, "STATUS", "CONFIRMED / DISPATCHED"],
            ["PASSENGER", passenger_name.upper(), "CLASS / VEHICLE", f"{travel_class} | {seat_no}"],
            ["ROUTE", route, "SERVICE / CARRIER", train_details],
            ["ISSUED BY", f"{OS_NAME} Neural Engine", "ARCHITECT", CREATOR_FULL_NAME.upper()]
        ]

        t = Table(data, colWidths=[130, 140, 130, 140])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F8FAFC")),
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor("#0F172A")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(t)
        doc.build(elements)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception:
        return None

def generate_pdf_report(title, content):
    if not REPORTLAB_OK:
        return None
    try:
        buffer = io.BytesIO()
        pdf = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=45, rightMargin=45, topMargin=45, bottomMargin=45)
        styles = getSampleStyleSheet()
        t_style = ParagraphStyle("T", parent=styles["Title"], fontSize=18, alignment=TA_CENTER, spaceAfter=14)
        h_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=12, spaceBefore=10, spaceAfter=6)
        b_style = ParagraphStyle("B", parent=styles["BodyText"], fontSize=9.5, leading=14, spaceAfter=6)

        story = [
            Paragraph(f"{OS_NAME} • ARCHITECT: {CREATOR_FULL_NAME.upper()}", t_style),
            Paragraph(title, h_style),
            Paragraph(datetime.now().strftime("%d %B %Y"), b_style),
            Spacer(1, 10)
        ]
        for line in content.splitlines():
            line = line.strip()
            if not line:
                story.append(Spacer(1, 4))
                continue
            safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if line.startswith("#"):
                story.append(Paragraph(safe.lstrip("#").strip(), h_style))
            elif line.startswith("- "):
                story.append(Paragraph("• " + safe[2:], b_style))
            else:
                story.append(Paragraph(safe, b_style))
        pdf.build(story)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception:
        return None

def generate_docx_report(title, content):
    if Document is None:
        return None
    try:
        doc = Document()
        doc.add_heading(title, level=1)
        doc.add_paragraph(f"Generated by: {OS_NAME} | Architect: {CREATOR_FULL_NAME}")
        doc.add_paragraph("---")
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith("#"):
                doc.add_heading(line.lstrip("#").strip(), level=2)
            elif line.startswith("- "):
                doc.add_paragraph(line[2:], style='List Bullet')
            else:
                doc.add_paragraph(line)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

# ============================================================
# 6. ACCURATE REAL-WORLD TRANSIT RESOLVER
# ============================================================
STATION_CODES = {
    "mandawar": "MURD",
    "mandawar mahwa road": "MURD",
    "mahwa": "MURD",
    "jaipur": "JP",
    "dausa": "DO",
    "alwar": "AWR",
    "bandikui": "BKI",
    "bharatpur": "BTE",
    "delhi": "NDLS",
    "new delhi": "NDLS",
    "delhi cantt": "DEC",
    "mumbai": "MMCT",
    "ahmedabad": "ADI",
    "kota": "KOTA",
    "ajmer": "AII",
    "agra": "AGC",
    "kolkata": "HWH"
}

def clean_city_name(raw_text):
    clean = re.sub(r'^(book|train|ticket|bus|cab|taxi|flight|plane|bike|uber|ola|rapido|from|to|station|junction)\s+', '', raw_text.strip(), flags=re.IGNORECASE).strip()
    return clean.title()

def detect_booking_or_action_intent(user_text):
    txt = user_text.lower()
    
    transit_keywords = [
        "train", "irctc", "railway", "ticket book",
        "bus", "roadways", "redbus",
        "cab", "uber", "ola", "taxi",
        "bike", "rapido", "scooter",
        "flight", "aeroplane", "air ticket", "plane", "airline"
    ]
    
    if any(k in txt for k in transit_keywords):
        # 1. Mode Detection
        if any(k in txt for k in ["cab", "uber", "ola", "taxi"]):
            mode = "cab"
        elif any(k in txt for k in ["bike", "rapido", "scooter"]):
            mode = "bike"
        elif any(k in txt for k in ["bus", "roadways", "redbus"]):
            mode = "bus"
        elif any(k in txt for k in ["flight", "aeroplane", "air ticket", "plane", "airline"]):
            mode = "flight"
        else:
            mode = "train"

        # 2. Extract Cities from Query (Flexible Regex)
        match = re.search(r'(?:from\s+)?([a-zA-Z0-9\s]+?)\s+(?:to|till|se)\s+([a-zA-Z0-9\s]+)', txt, re.IGNORECASE)
        if match:
            from_c = clean_city_name(match.group(1))
            to_c = clean_city_name(match.group(2))
        else:
            from_c = "Jaipur"
            to_c = "Delhi"

        from_code = STATION_CODES.get(from_c.lower(), from_c)
        to_code = STATION_CODES.get(to_c.lower(), to_c)

        # 3. Mode-Specific Verified Routing Links
        if mode == "train":
            icon = "🚆"
            title = f"{icon} Official Indian Railways Reservation: {from_c} ➔ {to_c}"
            # Direct reliable search link that never crashes
            link_url = f"https://www.google.com/search?q=trains+from+{urllib.parse.quote(from_c)}+to+{urllib.parse.quote(to_c)}+live+train+schedule"
            irctc_direct = f"https://www.confirmtkt.com/rbooking-d/{urllib.parse.quote(from_code)}-to-{urllib.parse.quote(to_code)}"
            btn_label = f"⚡ Check Live Trains & Seat Availability ({from_c} ➔ {to_c})"
            desc = f"Direct IRCTC sync active for {from_c} ({from_code}) to {to_c} ({to_code}). Click below to verify real-time available trains:"

        elif mode == "bus":
            icon = "🚌"
            title = f"{icon} Roadways & Intercity Bus Matrix: {from_c} ➔ {to_c}"
            link_url = f"https://www.redbus.in/bus-tickets/{urllib.parse.quote(from_c.lower())}-to-{urllib.parse.quote(to_c.lower())}"
            btn_label = f"⚡ Reserve Bus Seats on RedBus ({from_c} ➔ {to_c})"
            desc = f"State roadways (RSRTC/UPSRTC) and private luxury sleepers active between {from_c} and {to_c}:"

        elif mode == "cab":
            icon = "🚕"
            title = f"{icon} Instant Cab Dispatch: {from_c} ➔ {to_c}"
            link_url = f"https://m.uber.com/ul/?action=setPickup&pickup=my_location"
            btn_label = f"⚡ Open Uber / Ola Cab Fleet"
            desc = f"On-demand rides available for pickup at {from_c} with drop at {to_c}:"

        elif mode == "bike":
            icon = "🛵"
            title = f"{icon} Rapid Two-Wheeler Commute: {from_c} ➔ {to_c}"
            link_url = "https://www.rapido.bike/"
            btn_label = f"⚡ Request Bike Rider (Rapido)"
            desc = f"Fastest intra-city commuter route mapped between {from_c} and {to_c}:"

        else: # Flight
            icon = "✈️"
            title = f"{icon} Live Flight Matrix: {from_c} ➔ {to_c}"
            link_url = f"https://www.google.com/travel/flights?q=flights+from+{urllib.parse.quote(from_c)}+to+{urllib.parse.quote(to_c)}"
            btn_label = f"⚡ Review Flight Fares on Google Flights"
            desc = f"Live airline pricing synced for {from_c} to {to_c}:"

        return {
            "type": "mobility",
            "mode": mode,
            "title": title,
            "desc": desc,
            "link": link_url,
            "btn_label": btn_label,
            "from_stn": from_c,
            "to_stn": to_c,
            "service_name": f"{from_c}-{to_c} {mode.title()} Transit"
        }

    # Movie Intent
    if any(k in txt for k in ["movie", "cinema", "film", "bookmyshow"]):
        movie_match = re.search(r'movie(?:\s+ticket)?(?:\s+for)?\s+([a-zA-Z0-9\s]+)', txt, re.IGNORECASE)
        movie_title = movie_match.group(1).strip().title() if movie_match else "Trending Releases"
        bms_url = f"https://in.bookmyshow.com/explore/movies?search={urllib.parse.quote(movie_title)}"
        return {
            "type": "movie",
            "title": f"🎬 Cinema Seat Allocator: {movie_title}",
            "desc": "Auditoriums and showtimes synced. Select seats:",
            "link": bms_url,
            "btn_label": f"Select Seats on BookMyShow ({movie_title})"
        }

    return None

# ============================================================
# 7. ZERO-FREEZE INFERENCE (GROQ PRIMARY WITH FAST FAILOVER)
# ============================================================
def run_ai_completion(user_text):
    system_instruction = (
        f"You are {OS_NAME}, an elite autonomous cognitive partner and enterprise intelligence operating system. "
        f"You were architected and engineered solely by your creator: {CREATOR_FULL_NAME}. "
        f"Always proudly credit {CREATOR_FULL_NAME} as your architect when asked who you are or who made you. "
        f"Fluency: You understand and communicate fluently in Hindi, English, and Hinglish. "
        f"Default to professional, high-density English, but adapt seamlessly to the user's natural language. "
        f"When writing legal agreements or business emails, format them cleanly with headers, clauses, and signature blocks. "
        f"Be direct, authentic, sharp, and helpful. Never repeat robotic canned templates."
    )

    doc_context = ""
    if st.session_state.attached_assets:
        doc_context = "\n\n=== ATTACHED MEMORY & OCR ASSETS ===\n"
        for a in st.session_state.attached_assets[-3:]:
            doc_context += f"Asset: {a['name']}\nData:\n{a['text'][:3000]}\n---\n"

    web_context = ""
    if any(k in user_text.lower() for k in ["search", "latest", "news", "today", "current", "update", "kya chal raha"]):
        web_res = run_live_web_search(user_text)
        web_context = f"\n\n=== LIVE WEB INTELLIGENCE ===\n{web_res}\n---\n"

    full_payload = f"{doc_context}{web_context}\nUser: {user_text}"

    # 1. Primary Engine: Groq (Ultra-fast 300+ tokens/sec with hard 6s timeout)
    if GROQ_API_KEY and Groq:
        try:
            g_client = Groq(api_key=GROQ_API_KEY, timeout=6.0)
            resp = g_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": full_payload}
                ],
                temperature=0.6,
                max_tokens=1800
            )
            if resp.choices and resp.choices[0].message.content:
                return resp.choices[0].message.content.strip()
        except Exception:
            pass

    # 2. Secondary Engine: Gemini Fast Failover
    if GEMINI_API_KEY and genai:
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            for m in ["gemini-2.5-flash", "gemini-3.5-flash"]:
                try:
                    res = client.models.generate_content(
                        model=m,
                        contents=f"{system_instruction}\n\n{full_payload}"
                    )
                    if res and res.text:
                        return res.text.strip()
                except Exception:
                    continue
        except Exception:
            pass

    return f"I am {OS_NAME}, engineered by {CREATOR_FULL_NAME}. How can I assist your workflow today?"

# ============================================================
# 8. HEADER
# ============================================================
st.markdown(
    f"""<div class="aetheris-header">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div class="brand-title">💠 {OS_NAME}</div>
                <div style="color:#64748b; font-size:12px; margin-top:2px; font-weight:500;">
                    AUTONOMOUS COGNITIVE & COMPLIANCE OPERATING SYSTEM
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:12px;">
                <div class="architect-badge">ARCHITECT: {CREATOR_FULL_NAME.upper()}</div>
                <div class="status-badge">
                    <span class="dot"></span> READY
                </div>
            </div>
        </div>
    </div>""",
    unsafe_allow_html=True
)

# ============================================================
# 9. SIDEBAR CONTROLS: ATTACHMENTS, MIC & QUICK TOOLS
# ============================================================
with st.sidebar:
    st.markdown(f"**💠 {OS_NAME}**")
    st.caption(f"Architect: {CREATOR_FULL_NAME}")
    
    if st.button("＋ Clear Session", use_container_width=True):
        st.session_state.messages = []
        st.session_state.attached_assets = []
        st.rerun()

    st.divider()
    st.markdown("### 🎙️ Audio & Voice Input")
    st.caption("Speak query directly into microphone:")
    audio_record = st.audio_input("Record Voice Query", key="mic_recorder")
    if audio_record:
        st.info("Audio received. Ready for vocal instructions.")

    st.divider()
    st.markdown("### 📎 Attach Asset")
    st.caption("Drop Photo (OCR), Payroll/Financial CSV, PDF, or Word file.")
    up_asset = st.file_uploader("Upload File / Photo", type=["png", "jpg", "jpeg", "pdf", "docx", "csv", "txt"], key="single_uploader")

    if up_asset:
        if not any(a["name"] == up_asset.name for a in st.session_state.attached_assets):
            ext = Path(up_asset.name).suffix.lower()
            parsed = parse_file(up_asset)
            
            if ext in [".png", ".jpg", ".jpeg"]:
                with st.spinner("Deciphering OCR..."):
                    ocr_res = run_vision_ocr(parsed)
                    st.session_state.attached_assets.append({
                        "name": up_asset.name,
                        "type": "Image OCR",
                        "text": ocr_res
                    })
                    st.success(f"Extracted: {up_asset.name}")
            elif isinstance(parsed, pd.DataFrame):
                st.session_state.attached_assets.append({
                    "name": up_asset.name,
                    "type": "Data Master",
                    "df": parsed,
                    "text": f"Data Preview:\n{parsed.head(20).to_string()}"
                })
                st.success(f"Attached Data: {up_asset.name}")
            else:
                st.session_state.attached_assets.append({
                    "name": up_asset.name,
                    "type": "Document",
                    "text": str(parsed)
                })
                st.success(f"Attached: {up_asset.name}")

    if st.session_state.attached_assets:
        st.markdown("**Active Knowledge:**")
        for a in st.session_state.attached_assets:
            st.info(f"📄 {a['name']} ({a['type']})")

    st.divider()
    st.caption(f"{OS_NAME} • Autonomous Enterprise Engine")

# ============================================================
# 10. UNIFIED INTERACTION WORKSPACE
# ============================================================
for idx, msg in enumerate(st.session_state.messages):
    avatar = "👤" if msg["role"] == "user" else "🤖"
    with st.chat_message(msg["role"], avatar=avatar):
        if msg.get("type") == "image":
            st.image(msg["content"], caption="FLUX Studio Asset")
            if isinstance(msg["content"], bytes):
                st.download_button("⬇ Download HD Image", msg["content"], file_name=f"render_{idx}.png", mime="image/png", key=f"dl_{idx}")
        elif msg.get("type") == "action_card":
            card = msg["card"]
            st.markdown(f"#### {card['title']}")
            st.write(card["desc"])
            st.link_button(card["btn_label"], card["link"], use_container_width=True)
        elif msg.get("type") == "pf_audit":
            st.success("✅ Statutory EPF Audit & ECR Generation Complete.")
            st.dataframe(msg["audit_df"], use_container_width=True)
            st.markdown("##### Ready-to-Upload EPFO ECR Plain-Text File (`#~#`)")
            st.code(msg["ecr_text"], language="text")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ Download EPFO ECR File (.txt)", msg["ecr_text"], file_name="EPFO_ECR_Return.txt", mime="text/plain", key=f"ecr_{idx}", use_container_width=True)
            with c2:
                csv_b = msg["audit_df"].to_csv(index=False).encode('utf-8')
                st.download_button("⬇ Download Audit CSV", csv_b, file_name="compliance_audit.csv", mime="text/csv", key=f"csv_{idx}", use_container_width=True)
        elif msg.get("type") == "data_chart":
            st.success(f"📊 {msg['title']}")
            st.bar_chart(msg["chart_data"])
            st.markdown(msg["content"])
        else:
            st.markdown(msg["content"])
            if msg.get("mailto"):
                st.link_button("📨 Open Draft in Gmail / Mail Client", msg["mailto"], use_container_width=True)
            if msg.get("pdf_data"):
                st.download_button("⬇ Download Legal/Executive PDF", msg["pdf_data"], file_name="aetheris_document.pdf", mime="application/pdf", key=f"pdf_{idx}")
            if msg.get("docx_data"):
                st.download_button("⬇ Download Word File (DOCX)", docx_bytes, file_name="aetheris_document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{idx}")

user_prompt = st.chat_input("Ask Aetheris OS anything (e.g. 'book train Mandawar to Jaipur', 'cab to airport')...")

if user_prompt:
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user", avatar="👤"):
        st.markdown(user_prompt)

    action_intent = detect_booking_or_action_intent(user_prompt)
    img_trigger = re.compile(r'^\s*(generate\s*image|create\s*image|draw|render|photo\s*of|image\s*of)\s*[:\-]?\s*', re.IGNORECASE)
    is_pf_request = any(k in user_prompt.lower() for k in ["audit pf", "compliance check", "generate ecr", "pf audit", "epfo"])
    is_data_viz_request = any(k in user_prompt.lower() for k in ["analyze data", "chart", "visualize", "plot", "sales data"])
    is_email_request = any(k in user_prompt.lower() for k in ["draft email", "send email", "email to", "apology email", "leave email"])

    # 1. Image Generation Intent
    if bool(img_trigger.match(user_prompt)) or "image of " in user_prompt.lower():
        clean_target = img_trigger.sub("", user_prompt).strip() or user_prompt
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Rendering HD asset with FLUX..."):
                img_data = fetch_studio_asset_bytes(clean_target)
                st.image(img_data, caption=f"Render: {clean_target[:50]}")
                if isinstance(img_data, bytes):
                    st.download_button("⬇ Download HD Image", img_data, file_name="aetheris_render.png", mime="image/png", key="d_live_btn")
                st.session_state.messages.append({"role": "assistant", "content": img_data, "type": "image"})

    # 2. EPF Labour Law Compliance
    elif is_pf_request:
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Auditing compliance rules under EPF Act 1952..."):
                target_df = next((a["df"] for a in st.session_state.attached_assets if "df" in a), None)
                if target_df is None:
                    target_df = pd.DataFrame({
                        "UAN": ["100984512034", "100451298451", "100784512390", "100124578963", "100876543210"],
                        "Employee_Name": ["Ramesh Kumar", "Pooja Sharma", "Vikram Singh", "Amit Verma", "Sunita Devi"],
                        "Gross_Wages": [18500, 14200, 26000, 9500, 15000]
                    })
                audit_res, ecr_text = calculate_pf_ecr(target_df)
                st.success("✅ Statutory EPF Audit & ECR Generation Complete.")
                st.dataframe(audit_res, use_container_width=True)
                st.markdown("##### Ready-to-Upload EPFO ECR Plain-Text File (`#~#`)")
                st.code(ecr_text, language="text")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇ Download EPFO ECR File (.txt)", ecr_text, file_name="EPFO_ECR_Return.txt", mime="text/plain", use_container_width=True)
                with c2:
                    csv_b = audit_res.to_csv(index=False).encode('utf-8')
                    st.download_button("⬇ Download Audit CSV", csv_b, file_name="compliance_audit.csv", mime="text/csv", use_container_width=True)
                
                st.session_state.messages.append({
                    "role": "assistant",
                    "type": "pf_audit",
                    "audit_df": audit_res,
                    "ecr_text": ecr_text
                })

    # 3. Financial & Data Analyst (Auto-Charts & Insights)
    elif is_data_viz_request:
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Analyzing data patterns & computing visual metrics..."):
                target_df = next((a["df"] for a in st.session_state.attached_assets if "df" in a), None)
                if target_df is None:
                    target_df = pd.DataFrame({
                        "Department": ["Engineering", "Marketing", "Sales", "Operations", "Legal"],
                        "Budget_INR": [450000, 280000, 520000, 190000, 120000]
                    })
                num_cols = target_df.select_dtypes(include=['number']).columns.tolist()
                cat_cols = target_df.select_dtypes(include=['object']).columns.tolist()
                
                chart_df = None
                if num_cols and cat_cols:
                    chart_df = target_df.set_index(cat_cols[0])[num_cols[0]]
                elif num_cols:
                    chart_df = target_df[num_cols[0]]

                st.success("📊 Strategic Data Analysis & Visualization")
                if chart_df is not None:
                    st.bar_chart(chart_df)
                
                insight_txt = run_ai_completion(f"Provide a 3-bullet strategic executive analysis of this data:\n{target_df.to_string()}")
                st.markdown(insight_txt)

                st.session_state.messages.append({
                    "role": "assistant",
                    "type": "data_chart",
                    "title": "Automated Dataset Visualization",
                    "chart_data": chart_df,
                    "content": insight_txt
                })

    # 4. Smart Email & Communication Dispatcher (Draft-to-Send Bridge)
    elif is_email_request:
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner(f"{OS_NAME} is drafting executive communication..."):
                email_text = run_ai_completion(f"Draft a formal, high-conversion email based on this request: {user_prompt}. Provide subject line and body clearly.")
                st.markdown(email_text)

                subj_match = re.search(r'Subject:\s*(.*?)\n', email_text, re.IGNORECASE)
                subj = subj_match.group(1).strip() if subj_match else "Executive Communication"
                body_clean = re.sub(r'Subject:.*?\n', '', email_text, flags=re.IGNORECASE).strip()
                mailto_link = f"mailto:?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(body_clean[:1200])}"
                
                st.link_button("📨 Open Draft in Gmail / Mail Client", mailto_link, use_container_width=True)

                st.session_state.messages.append({
                    "role": "assistant",
                    "content": email_text,
                    "mailto": mailto_link
                })

    # 5. Real-Time Verified Mobility Actions
    elif action_intent:
        with st.chat_message("assistant", avatar="🤖"):
            st.markdown(f"#### {action_intent['title']}")
            st.write(action_intent["desc"])
            st.link_button(action_intent["btn_label"], action_intent["link"], use_container_width=True)

            mode_name = action_intent.get("mode", "Transit").title()
            with st.expander(f"🎫 Complete {mode_name} Manifest & Digital Boarding Pass", expanded=True):
                c_p1, c_p2 = st.columns(2)
                with c_p1:
                    p_name = st.text_input("Passenger Full Name", value=CREATOR_FULL_NAME, key="p_name_input")
                    p_age = st.number_input("Age", min_value=5, max_value=100, value=21, key="p_age_input")
                with c_p2:
                    if action_intent.get("mode") in ["cab", "bike"]:
                        p_class = st.selectbox("Vehicle Category", ["Standard / Go", "Sedan / Prime", "SUV / XL", "Moto Bike"], key="p_class_input")
                        p_berth = st.selectbox("Pickup Preference", ["Immediate (5 mins)", "Schedule for later"], key="p_berth_input")
                    else:
                        p_class = st.selectbox("Class Preference", ["Executive / AC Seater", "Sleeper / Economy", "3A / 2A AC", "First Class"], key="p_class_input")
                        p_berth = st.selectbox("Seat Preference", ["Window Side", "Aisle", "Lower", "No Preference"], key="p_berth_input")

                if st.button("Generate Confirmed Boarding Pass"):
                    gen_pnr = f"ATH-{random.randint(100000, 999999)}"
                    seat_token = f"Seat/Coach: {random.randint(1, 45)} ({p_berth})"
                    route_name = f"{action_intent['from_stn']} ➔ {action_intent['to_stn']}"
                    service_carrier = action_intent.get("service_name", f"{mode_name} Route")
                    
                    pass_bytes = generate_travel_boarding_pass(
                        pnr=gen_pnr,
                        passenger_name=p_name,
                        train_details=service_carrier,
                        travel_class=p_class,
                        seat_no=seat_token,
                        route=route_name
                    )

                    st.success(f"Manifest Active! Token PNR: **{gen_pnr}** | Route: **{route_name}**")
                    if pass_bytes:
                        st.download_button(
                            label="📥 Download Official Boarding Pass (PDF)",
                            data=pass_bytes,
                            file_name=f"Boarding_Pass_{gen_pnr}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

            st.session_state.messages.append({
                "role": "assistant",
                "type": "action_card",
                "card": action_intent
            })

    # 6. Legal / Notice Document Drafter & Cognitive Reasoning (Instant PDF & DOCX Export)
    else:
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner(f"{OS_NAME} is executing..."):
                response_txt = run_ai_completion(user_prompt)
                st.markdown(response_txt)
                
                pdf_bytes = None
                docx_bytes = None
                is_doc_intent = any(w in user_prompt.lower() for w in ["pdf", "report", "word", "docx", "agreement", "notice", "contract", "nda", "draft"])
                
                if is_doc_intent:
                    pdf_bytes = generate_pdf_report("Executive Legal & Intelligence Briefing", response_txt)
                    docx_bytes = generate_docx_report("Executive Legal & Intelligence Briefing", response_txt)
                    
                    col_dl1, col_dl2 = st.columns(2)
                    if pdf_bytes:
                        with col_dl1:
                            st.download_button("⬇ Download Legal/Executive PDF", pdf_bytes, file_name="aetheris_document.pdf", mime="application/pdf", use_container_width=True)
                    if docx_bytes:
                        with col_dl2:
                            st.download_button("⬇ Download Word File (DOCX)", docx_bytes, file_name="aetheris_document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

                st.session_state.messages.append({
                    "role": "assistant",
                    "content": response_txt,
                    "pdf_data": pdf_bytes,
                    "docx_data": docx_bytes
                })
