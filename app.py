import os
import io
import json
import uuid
import re
import base64
from pathlib import Path
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from PIL import Image

# ------------------------------------------------------------
# Core Safe Imports
# ------------------------------------------------------------
try:
    from pypdf import PdfReader
    PDF_OK = True
except ImportError:
    PdfReader = None
    PDF_OK = False

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
    import requests
    REQUESTS_OK = True
except ImportError:
    requests = None
    REQUESTS_OK = False

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
# 1. BULLETPROOF KEY RESOLUTION (STREAMLIT SECRETS + .ENV)
# ============================================================
load_dotenv(override=True)

def get_secret(key_name):
    # 1. Check Streamlit Cloud Secrets first
    if hasattr(st, "secrets") and key_name in st.secrets:
        return str(st.secrets[key_name]).strip()
    # 2. Check os.environ / .env
    val = os.getenv(key_name, "").strip()
    if val:
        return val
    return ""

GROQ_API_KEY = get_secret("GROQ_API_KEY")
GEMINI_API_KEY = get_secret("GEMINI_API_KEY")

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
}
</style>""",
    unsafe_allow_html=True,
)

# ============================================================
# 3. SESSION STATE
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

if "processed_docs" not in st.session_state:
    st.session_state.processed_docs = []

# ============================================================
# 4. DIRECT DETERMINISTIC CONVERTERS
# ============================================================
def build_docx_bytes(title, content_text):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        head = doc.add_heading(title, level=1)
        meta = doc.add_paragraph()
        meta_run = meta.add_run(f"System: {OS_NAME} | Architect: {CREATOR_FULL_NAME} | Date: {datetime.now().strftime('%d-%b-%Y')}")
        meta_run.font.size = Pt(9)
        meta_run.font.color.rgb = RGBColor(100, 116, 139)
        doc.add_paragraph("―" * 45)

        for paragraph in content_text.split("\n\n"):
            clean_p = paragraph.strip()
            if not clean_p:
                continue
            p = doc.add_paragraph(clean_p)
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def build_pdf_bytes(title, content_text):
    if not REPORTLAB_OK:
        return None
    try:
        buf = io.BytesIO()
        pdf = SimpleDocTemplate(buf, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()

        t_style = ParagraphStyle("DocT", parent=styles["Title"], fontSize=15, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a"))
        m_style = ParagraphStyle("DocM", parent=styles["Normal"], fontSize=8.5, alignment=TA_CENTER, textColor=colors.HexColor("#64748b"), spaceAfter=14)
        b_style = ParagraphStyle("DocB", parent=styles["Normal"], fontSize=10, leading=14, alignment=TA_LEFT, textColor=colors.HexColor("#1e293b"), spaceAfter=8)

        story = [
            Paragraph(f"<b>{title}</b>", t_style),
            Paragraph(f"{OS_NAME} Converted • Architect: {CREATOR_FULL_NAME}", m_style),
            Spacer(1, 10)
        ]

        for line in content_text.splitlines():
            clean = line.strip()
            if not clean:
                story.append(Spacer(1, 4))
                continue
            safe = clean.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe, b_style))

        pdf.build(story)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

# ============================================================
# 5. ZERO-DEPENDENCY NATIVE OCR ENGINE (DIRECT REST + GROQ)
# ============================================================
def extract_text_from_pdf(file_bytes):
    if not PDF_OK:
        return "pypdf library missing."
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        extracted = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text and text.strip():
                extracted.append(f"--- Page {i+1} ---\n" + text.strip())
        return "\n\n".join(extracted) if extracted else "No selectable text found in PDF."
    except Exception as e:
        return f"PDF Extraction Error: {str(e)}"

def extract_text_from_image(image_bytes, mime_type="image/jpeg"):
    prompt = (
        "Transcribe all text from this image VERBATIM in its original script and language (Hindi, Sanskrit, English, or Hinglish). "
        "Keep line breaks, punctuation, and exact spellings. "
        "Do NOT summarize, explain, or omit anything. Return only the raw extracted text."
    )

    # 1. Primary: Direct Gemini REST API (Bypasses library/version conflicts)
    if GEMINI_API_KEY and REQUESTS_OK:
        b64_data = base64.b64encode(image_bytes).decode("utf-8")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime_type, "data": b64_data}}
                ]
            }]
        }
        try:
            resp = requests.post(url, json=payload, timeout=25)
            if resp.status_code == 200:
                res_data = resp.json()
                text = res_data["candidates"][0]["content"]["parts"][0]["text"].strip()
                if text:
                    return text
        except Exception:
            pass

    # 2. Secondary: Groq Vision Fallback
    if GROQ_API_KEY and Groq:
        try:
            g_client = Groq(api_key=GROQ_API_KEY, timeout=20.0)
            b64_data = base64.b64encode(image_bytes).decode("utf-8")
            resp = g_client.chat.completions.create(
                model="llama-3.2-11b-vision-preview",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{b64_data}"}}
                        ]
                    }
                ],
                temperature=0.1
            )
            if resp.choices and resp.choices[0].message.content:
                return resp.choices[0].message.content.strip()
        except Exception:
            pass

    return "Error: Unable to transcribe. Please ensure GEMINI_API_KEY or GROQ_API_KEY is saved in Streamlit Cloud Secrets."

# ============================================================
# 6. HEADER
# ============================================================
st.markdown(
    f"""<div class="aetheris-header">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div class="brand-title">💠 {OS_NAME}</div>
                <div style="color:#64748b; font-size:12px; margin-top:2px; font-weight:500;">
                    AUTONOMOUS DOCUMENT CONVERSION & ENTERPRISE ENGINE
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:12px;">
                <div class="architect-badge">ARCHITECT: {CREATOR_FULL_NAME.upper()}</div>
            </div>
        </div>
    </div>""",
    unsafe_allow_html=True
)

# ============================================================
# 7. SIDEBAR: DIRECT DOCUMENT TO WORD / PDF WORKSPACE
# ============================================================
with st.sidebar:
    st.markdown("### 📥 Universal File Converter")
    st.caption("Upload Photo (Hindi/Eng/Sanskrit) or PDF to convert into Word & PDF:")
    
    upload = st.file_uploader("Upload Image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="file_converter")

    if upload:
        b_data = upload.getvalue()
        fname = upload.name
        fext = Path(fname).suffix.lower()
        mime = "image/png" if fext == ".png" else "image/jpeg"

        if not any(d["name"] == fname for d in st.session_state.processed_docs):
            with st.spinner(f"Reading and transcribing {fname}..."):
                if fext in [".png", ".jpg", ".jpeg"]:
                    extracted = extract_text_from_image(b_data, mime_type=mime)
                elif fext == ".pdf":
                    extracted = extract_text_from_pdf(b_data)
                else:
                    extracted = "Unsupported file format."

                docx_out = build_docx_bytes(f"Extracted - {Path(fname).stem}", extracted)
                pdf_out = build_pdf_bytes(f"Extracted - {Path(fname).stem}", extracted)

                st.session_state.processed_docs.append({
                    "name": fname,
                    "content": extracted,
                    "docx": docx_out,
                    "pdf": pdf_out
                })
                st.success(f"Processed: {fname}")

    if st.session_state.processed_docs:
        st.divider()
        st.markdown("**📁 Converted Deliverables:**")
        for idx, item in enumerate(st.session_state.processed_docs):
            st.markdown(f"**{item['name']}**")
            with st.expander("👁️ View Extracted Content", expanded=True):
                st.text_area("Verbatim Text", item["content"], height=140, key=f"txt_{idx}")
            
            c1, c2 = st.columns(2)
            if item.get("docx"):
                with c1:
                    st.download_button(
                        "⬇ Word (.docx)",
                        item["docx"],
                        file_name=f"{Path(item['name']).stem}_converted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{idx}",
                        use_container_width=True
                    )
            if item.get("pdf"):
                with c2:
                    st.download_button(
                        "⬇ PDF (.pdf)",
                        item["pdf"],
                        file_name=f"{Path(item['name']).stem}_converted.pdf",
                        mime="application/pdf",
                        key=f"dl_pdf_{idx}",
                        use_container_width=True
                    )
            st.divider()

    if st.button("＋ Clear Workspace", use_container_width=True):
        st.session_state.processed_docs = []
        st.session_state.messages = []
        st.rerun()

# ============================================================
# 8. MAIN WORKSPACE: TEXT TO WORD / PDF & COGNITIVE CHAT
# ============================================================
for idx, msg in enumerate(st.session_state.messages):
    avatar = "👤" if msg["role"] == "user" else "🤖"
    with st.chat_message(msg["role"], avatar=avatar):
        st.markdown(msg["content"])
        if msg.get("docx") or msg.get("pdf"):
            c1, c2 = st.columns(2)
            if msg.get("docx"):
                with c1:
                    st.download_button(
                        "⬇ Download Word (.docx)",
                        msg["docx"],
                        file_name=f"Document_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"c_docx_{idx}",
                        use_container_width=True
                    )
            if msg.get("pdf"):
                with c2:
                    st.download_button(
                        "⬇ Download PDF (.pdf)",
                        msg["pdf"],
                        file_name=f"Document_{idx}.pdf",
                        mime="application/pdf",
                        key=f"c_pdf_{idx}",
                        use_container_width=True
                    )

user_prompt = st.chat_input("Enter text or ask to draft a document...")

if user_prompt:
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user", avatar="👤"):
        st.markdown(user_prompt)

    with st.chat_message("assistant", avatar="🤖"):
        is_convert_req = any(k in user_prompt.lower() for k in ["convert", "word me", "pdf me", "docx me", "badlo", "photo"])
        has_docs = len(st.session_state.processed_docs) > 0

        if is_convert_req and has_docs:
            last_doc = st.session_state.processed_docs[-1]
            out_text = f"✅ Extracted content from **{last_doc['name']}**:\n\n{last_doc['content']}"
            docx_file = last_doc["docx"]
            pdf_file = last_doc["pdf"]
            st.markdown(out_text)
        else:
            instruction = f"You are {OS_NAME}, engineered by {CREATOR_FULL_NAME}. Write exhaustive, complete, verbatim text as requested."
            out_text = ""
            if GROQ_API_KEY and Groq:
                try:
                    client = Groq(api_key=GROQ_API_KEY, timeout=8.0)
                    r = client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[{"role": "system", "content": instruction}, {"role": "user", "content": user_prompt}],
                        temperature=0.3
                    )
                    out_text = r.choices[0].message.content.strip()
                except Exception:
                    pass
            if not out_text:
                out_text = user_prompt

            st.markdown(out_text)
            docx_file = build_docx_bytes("Aetheris Generated Document", out_text)
            pdf_file = build_pdf_bytes("Aetheris Generated Document", out_text)

        c1, c2 = st.columns(2)
        if docx_file:
            with c1:
                st.download_button(
                    "⬇ Download Word File (.docx)",
                    docx_file,
                    file_name="Aetheris_Document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        if pdf_file:
            with c2:
                st.download_button(
                    "⬇ Download PDF File (.pdf)",
                    pdf_file,
                    file_name="Aetheris_Document.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

        st.session_state.messages.append({
            "role": "assistant",
            "content": out_text,
            "docx": docx_file,
            "pdf": pdf_file
        })
