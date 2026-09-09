import os
import io
import json
import uuid
import re
import urllib.parse
from pathlib import Path
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from PIL import Image

# ------------------------------------------------------------
# Core Safe Imports
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
    from docx.shared import Pt, RGBColor
    DOCX_OK = True
except ImportError:
    Document = None
    DOCX_OK = False

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
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
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
# 4. SESSION STATE INITIALIZATION
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

if "attached_assets" not in st.session_state:
    st.session_state.attached_assets = []

# ============================================================
# 5. ROBUST CONVERTERS: TEXT TO DOCX & TEXT TO PDF
# ============================================================
def create_docx_file(title, text_content):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        # Title
        h = doc.add_heading(title, level=1)
        h.style.font.name = 'Arial'
        
        meta = doc.add_paragraph()
        meta_run = meta.add_run(f"System: {OS_NAME} | Architect: {CREATOR_FULL_NAME} | Date: {datetime.now().strftime('%d-%b-%Y')}")
        meta_run.font.size = Pt(9)
        meta_run.font.color.rgb = RGBColor(100, 116, 139)
        
        doc.add_paragraph("―" * 45)

        for paragraph_text in text_content.split("\n\n"):
            clean_p = paragraph_text.strip()
            if not clean_p:
                continue
            
            # Check for heading lines
            if clean_p.startswith("# "):
                doc.add_heading(clean_p.replace("# ", "").strip(), level=2)
            elif clean_p.startswith("## "):
                doc.add_heading(clean_p.replace("## ", "").strip(), level=3)
            elif clean_p.startswith("- ") or clean_p.startswith("* "):
                for line in clean_p.splitlines():
                    if line.strip().startswith(("- ", "* ")):
                        doc.add_paragraph(line.strip()[2:], style='List Bullet')
                    else:
                        doc.add_paragraph(line.strip())
            else:
                p = doc.add_paragraph(clean_p)
                p.style.font.name = 'Arial'
                p.style.font.size = Pt(11)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        st.error(f"DOCX Generation Error: {str(e)}")
        return None

def create_pdf_file(title, text_content):
    if not REPORTLAB_OK:
        return None
    try:
        buffer = io.BytesIO()
        pdf = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=40,
            rightMargin=40,
            topMargin=40,
            bottomMargin=40
        )
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "DocTitle",
            parent=styles["Title"],
            fontSize=16,
            leading=20,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0f172a"),
            spaceAfter=10
        )
        meta_style = ParagraphStyle(
            "DocMeta",
            parent=styles["Normal"],
            fontSize=8.5,
            leading=12,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#64748b"),
            spaceAfter=14
        )
        body_style = ParagraphStyle(
            "DocBody",
            parent=styles["Normal"],
            fontSize=10,
            leading=14,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#1e293b"),
            spaceAfter=8
        )

        story = [
            Paragraph(f"<b>{title}</b>", title_style),
            Paragraph(f"{OS_NAME} Verified • Architect: {CREATOR_FULL_NAME} • {datetime.now().strftime('%d %B %Y')}", meta_style),
            Spacer(1, 8)
        ]

        for block in text_content.splitlines():
            clean = block.strip()
            if not clean:
                story.append(Spacer(1, 4))
                continue
            safe_text = clean.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if safe_text.startswith("#"):
                story.append(Paragraph(f"<b>{safe_text.lstrip('#').strip()}</b>", body_style))
            elif safe_text.startswith(("- ", "* ")):
                story.append(Paragraph(f"• {safe_text[2:]}", body_style))
            else:
                story.append(Paragraph(safe_text, body_style))

        pdf.build(story)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"PDF Generation Error: {str(e)}")
        return None

# ============================================================
# 6. UNIVERSAL OCR EXTRACTION (HINDI / SANSKRIT / ENGLISH)
# ============================================================
def extract_ocr_text(image_bytes):
    if not GEMINI_API_KEY or not genai:
        return "Gemini API key is required for OCR processing."
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        img = Image.open(io.BytesIO(image_bytes))
        prompt = (
            "You are an expert OCR transcription engine. Extract all text verbatim from this document/image. "
            "Maintain the exact layout, paragraphs, numbers, and punctuation. "
            "Transcribe accurately whether it is Hindi (Devanagari), Sanskrit, English, or Hinglish. "
            "Do NOT summarize, do NOT omit words, and do NOT add explanatory notes. Return only the raw extracted text."
        )
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[prompt, img]
        )
        return response.text.strip() if response and response.text else "No legible text found in image."
    except Exception as e:
        return f"OCR Extraction Error: {str(e)}"

# ============================================================
# 7. AI ENGINE (GROQ PRIMARY WITH FAST FAILOVER)
# ============================================================
def run_ai_completion(user_text):
    system_instruction = (
        f"You are {OS_NAME}, an autonomous cognitive operating system created and engineered solely by {CREATOR_FULL_NAME}. "
        f"You understand and write accurately in Hindi, English, Sanskrit, and mixed Hinglish. "
        f"Whenever writing formal letters, legal notices, research, or study roadmaps, generate complete, structured output."
    )

    doc_context = ""
    if st.session_state.attached_assets:
        doc_context = "\n\n=== RECENT DOCUMENT CONTEXT ===\n"
        for a in st.session_state.attached_assets[-2:]:
            doc_context += f"File: {a['name']}\nContent:\n{a['text'][:4000]}\n---\n"

    full_payload = f"{doc_context}\nUser Request: {user_text}"

    if GROQ_API_KEY and Groq:
        try:
            g_client = Groq(api_key=GROQ_API_KEY, timeout=8.0)
            resp = g_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": full_payload}
                ],
                temperature=0.4,
                max_tokens=2200
            )
            if resp.choices and resp.choices[0].message.content:
                return resp.choices[0].message.content.strip()
        except Exception:
            pass

    if GEMINI_API_KEY and genai:
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            res = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"{system_instruction}\n\n{full_payload}"
            )
            if res and res.text:
                return res.text.strip()
        except Exception:
            pass

    return f"{OS_NAME} is active. Please ask your query."

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
# 9. SIDEBAR: PHOTO-TO-WORD / PHOTO-TO-PDF OCR TERMINAL
# ============================================================
with st.sidebar:
    st.markdown(f"**💠 {OS_NAME} CORE**")
    st.caption(f"Architect: {CREATOR_FULL_NAME}")
    
    if st.button("＋ New Clean Session", use_container_width=True):
        st.session_state.messages = []
        st.session_state.attached_assets = []
        st.rerun()

    st.divider()
    st.markdown("### 📷 Photo to Word & PDF (OCR)")
    st.caption("Upload any photo (Hindi / English / Sanskrit) to extract verbatim & download as DOCX or PDF.")
    
    uploaded_file = st.file_uploader(
        "Upload Image / Document", 
        type=["png", "jpg", "jpeg", "pdf", "txt"], 
        key="ocr_uploader"
    )

    if uploaded_file:
        file_bytes = uploaded_file.getvalue()
        file_ext = Path(uploaded_file.name).suffix.lower()

        if file_ext in [".png", ".jpg", ".jpeg"]:
            if not any(a["name"] == uploaded_file.name for a in st.session_state.attached_assets):
                with st.spinner("Transcribing verbatim text via Gemini Vision..."):
                    extracted_text = extract_ocr_text(file_bytes)
                    
                    # Generate Downloadable Files Instantly
                    docx_bytes = create_docx_file(f"Transcribed - {uploaded_file.name}", extracted_text)
                    pdf_bytes = create_pdf_file(f"Transcribed - {uploaded_file.name}", extracted_text)

                    st.session_state.attached_assets.append({
                        "name": uploaded_file.name,
                        "type": "OCR Document",
                        "text": extracted_text,
                        "docx": docx_bytes,
                        "pdf": pdf_bytes
                    })
                    st.success("Extracted successfully!")

    # Display Download Buttons in Sidebar for any processed OCR
    if st.session_state.attached_assets:
        for idx, asset in enumerate(st.session_state.attached_assets):
            st.markdown(f"**📄 {asset['name']}**")
            with st.expander("👁️ View Extracted Text", expanded=False):
                st.text_area("Extracted Verbatim", asset["text"], height=160, key=f"preview_{idx}")
            
            c_d1, c_d2 = st.columns(2)
            if asset.get("docx"):
                with c_d1:
                    st.download_button(
                        label="⬇ Word (.docx)",
                        data=asset["docx"],
                        file_name=f"{Path(asset['name']).stem}_transcribed.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"side_docx_{idx}",
                        use_container_width=True
                    )
            if asset.get("pdf"):
                with c_d2:
                    st.download_button(
                        label="⬇ PDF (.pdf)",
                        data=asset["pdf"],
                        file_name=f"{Path(asset['name']).stem}_transcribed.pdf",
                        mime="application/pdf",
                        key=f"side_pdf_{idx}",
                        use_container_width=True
                    )
            st.divider()

# ============================================================
# 10. MAIN CHAT & TEXT-TO-DOCX / TEXT-TO-PDF ENGINE
# ============================================================
for idx, msg in enumerate(st.session_state.messages):
    avatar = "👤" if msg["role"] == "user" else "🤖"
    with st.chat_message(msg["role"], avatar=avatar):
        st.markdown(msg["content"])
        
        # Clickable downloads for text messages
        if msg.get("docx_data") or msg.get("pdf_data"):
            c1, c2 = st.columns(2)
            if msg.get("docx_data"):
                with c1:
                    st.download_button(
                        label="⬇ Download Word File (.docx)",
                        data=msg["docx_data"],
                        file_name=f"Aetheris_Document_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"chat_docx_{idx}",
                        use_container_width=True
                    )
            if msg.get("pdf_data"):
                with c2:
                    st.download_button(
                        label="⬇ Download Document (.pdf)",
                        data=msg["pdf_data"],
                        file_name=f"Aetheris_Document_{idx}.pdf",
                        mime="application/pdf",
                        key=f"chat_pdf_{idx}",
                        use_container_width=True
                    )

user_query = st.chat_input("Type anything or ask to draft a document (e.g. 'draft an RTI appeal letter in Hindi')...")

if user_query:
    st.session_state.messages.append({"role": "user", "content": user_query})
    with st.chat_message("user", avatar="👤"):
        st.markdown(user_query)

    with st.chat_message("assistant", avatar="🤖"):
        with st.spinner(f"{OS_NAME} is processing..."):
            response_text = run_ai_completion(user_query)
            st.markdown(response_text)

            # Auto-generate downloadable files whenever drafting or document intent is present
            doc_triggers = ["draft", "letter", "report", "pdf", "word", "docx", "application", "appeal", "notice", "agreement", "notes"]
            is_doc = any(k in user_query.lower() for k in doc_triggers) or len(response_text) > 450

            docx_data = None
            pdf_data = None
            if is_doc:
                docx_data = create_docx_file("Aetheris Generated Document", response_text)
                pdf_data = create_pdf_file("Aetheris Generated Document", response_text)

                c1, c2 = st.columns(2)
                if docx_data:
                    with c1:
                        st.download_button(
                            label="⬇ Download Word File (.docx)",
                            data=docx_data,
                            file_name="Aetheris_Document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                if pdf_data:
                    with c2:
                        st.download_button(
                            label="⬇ Download Document (.pdf)",
                            data=pdf_data,
                            file_name="Aetheris_Document.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

            st.session_state.messages.append({
                "role": "assistant",
                "content": response_text,
                "docx_data": docx_data,
                "pdf_data": pdf_data
            })
