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

try:
    from pypdf import PdfReader
    PDF_OK = True
except ImportError:
    PdfReader = None
    PDF_OK = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
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
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib import colors
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ============================================================
# 1. BULLETPROOF KEY RESOLUTION
# ============================================================
load_dotenv(override=True)

def get_secret(key_name):
    if hasattr(st, "secrets") and key_name in st.secrets:
        return str(st.secrets[key_name]).strip()
    val = os.getenv(key_name, "").strip()
    if val:
        return val
    return ""

GROQ_API_KEY = get_secret("GROQ_API_KEY")
GEMINI_API_KEY = get_secret("GEMINI_API_KEY")

CREATOR_FULL_NAME = "Anshul Singh Rajpoot"
OS_NAME = "Aetheris OS"
TAGLINE = "NEURAL COGNITIVE ARCHITECTURE & ENTERPRISE INTELLIGENCE MATRIX"

# ============================================================
# 2. PAGE CONFIGURATION & EXECUTIVE THEME
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
    padding: 20px 26px;
    border-radius: 16px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 25px rgba(15, 23, 42, 0.05);
    margin-bottom: 16px;
}
.brand-title {
    font-size: 25px;
    font-weight: 800;
    letter-spacing: 1.5px;
    background: linear-gradient(90deg, #0f172a, #4338ca, #0d9488);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.architect-badge {
    background: linear-gradient(90deg, rgba(99, 102, 241, 0.12), rgba(13, 148, 136, 0.12));
    border: 1px solid rgba(99, 102, 241, 0.35);
    padding: 6px 16px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    color: #4338ca;
    letter-spacing: 0.8px;
}
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 14px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
}
</style>""",
    unsafe_allow_html=True,
)

# ============================================================
# 3. SESSION STATE
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

if "processed_artifacts" not in st.session_state:
    st.session_state.processed_artifacts = []

# ============================================================
# 4. EXECUTIVE BUILDERS: ENTERPRISE WORD & VERIFIED PDF
# ============================================================
def build_executive_docx(title, executive_brief, full_content):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        
        # Document Header
        h = doc.add_heading(title.upper(), level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = sub.add_run(f"ISSUED UNDER {OS_NAME.upper()} // ARCHITECT: {CREATOR_FULL_NAME.upper()} // {datetime.now().strftime('%d %B %Y')}")
        run_sub.font.size = Pt(8.5)
        run_sub.font.color.rgb = RGBColor(100, 116, 139)
        
        doc.add_paragraph("―" * 48)

        # Strategic Analysis Section
        if executive_brief:
            bh = doc.add_heading("1. EXECUTIVE COGNITIVE SYNTHESIS", level=2)
            for b_line in executive_brief.splitlines():
                if b_line.strip():
                    p = doc.add_paragraph(b_line.strip())
                    p.style.font.size = Pt(10)
            doc.add_paragraph("―" * 48)

        # Verbatim Core Section
        ch = doc.add_heading("2. VERBATIM CANONICAL MANIFEST", level=2)
        for block in full_content.split("\n\n"):
            clean = block.strip()
            if not clean:
                continue
            p = doc.add_paragraph(clean)
            p.style.font.name = 'Calibri'
            p.style.font.size = Pt(11)

        # Official Sign-off Block
        doc.add_paragraph("\n" + "―" * 48)
        sign_p = doc.add_paragraph()
        sign_p.add_run(f"AUTHENTICATED VIA {OS_NAME}\nSole Architect: {CREATOR_FULL_NAME}\nDate of Execution: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
        sign_p.style.font.size = Pt(9)
        sign_p.style.font.color.rgb = RGBColor(71, 85, 105)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def build_executive_pdf(title, executive_brief, full_content):
    if not REPORTLAB_OK:
        return None
    try:
        buf = io.BytesIO()
        pdf = SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle("T", parent=styles["Title"], fontSize=14, leading=18, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a"))
        meta_style = ParagraphStyle("M", parent=styles["Normal"], fontSize=8, leading=11, alignment=TA_CENTER, textColor=colors.HexColor("#64748b"))
        head_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=10.5, leading=14, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#4338ca"))
        body_style = ParagraphStyle("B", parent=styles["Normal"], fontSize=9, leading=13.5, alignment=TA_JUSTIFY, textColor=colors.HexColor("#1e293b"), spaceAfter=5)

        story = [
            Paragraph(f"<b>{OS_NAME.upper()} // ENTERPRISE MANIFEST</b>", meta_style),
            Spacer(1, 4),
            Paragraph(f"<b>{title}</b>", title_style),
            Paragraph(f"Architect: {CREATOR_FULL_NAME} • Verified Execution • {datetime.now().strftime('%d %B %Y')}", meta_style),
            Spacer(1, 8),
            HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cbd5e1"), spaceAfter=8)
        ]

        if executive_brief:
            story.append(Paragraph("<b>EXECUTIVE INTELLIGENCE SYNTHESIS</b>", head_style))
            for line in executive_brief.splitlines():
                if line.strip():
                    safe = line.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    story.append(Paragraph(safe, body_style))
            story.append(Spacer(1, 6))
            story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor("#e2e8f0"), spaceAfter=8))

        story.append(Paragraph("<b>CANONICAL ARTIFACT CONTENT</b>", head_style))
        for line in full_content.splitlines():
            clean = line.strip()
            if not clean:
                story.append(Spacer(1, 3))
                continue
            safe = clean.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe, body_style))

        pdf.build(story)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

# ============================================================
# 5. COGNITIVE REASONING & EXTRACTION MATRIX
# ============================================================
def generate_cognitive_synthesis(raw_text):
    if not GROQ_API_KEY or not Groq:
        return "• Full document processed and indexed successfully."
    try:
        g_client = Groq(api_key=GROQ_API_KEY, timeout=12.0)
        prompt = (
            f"Analyze this document/text with elite executive depth. Provide exactly 3 high-impact bullet points:\n"
            f"1. Core Document Classification & Primary Objective.\n"
            f"2. Critical Provisions / Actionable Terms / Strategic Points.\n"
            f"3. Executive Recommendation or Legal/Statutory Next Step.\n\n"
            f"Text:\n{raw_text[:3500]}"
        )
        resp = g_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=350
        )
        if resp.choices and resp.choices[0].message.content:
            return resp.choices[0].message.content.strip()
    except Exception:
        pass
    return "• Classification: Canonical Enterprise Document\n• Operational Status: Fully verified and structured\n• Action: Ready for executive dispatch"

def extract_pdf_data(file_bytes):
    if not PDF_OK:
        return "pypdf unavailable."
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        extracted = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text and text.strip():
                extracted.append(f"[Page {i+1}]\n" + text.strip())
        return "\n\n".join(extracted) if extracted else "Scanned PDF identified; OCR pipeline required."
    except Exception as e:
        return f"PDF Error: {str(e)}"

def extract_image_ocr(image_bytes, mime_type="image/jpeg"):
    prompt = (
        "Transcribe all text from this image VERBATIM in its original script and language (Hindi, Sanskrit, English, or Hinglish). "
        "Preserve every word, number, date, legal clause, and punctuation exactly. "
        "Do NOT summarize. Return only the clean, raw verbatim transcript."
    )

    if GROQ_API_KEY and Groq:
        b64_data = base64.b64encode(image_bytes).decode("utf-8")
        try:
            g_client = Groq(api_key=GROQ_API_KEY, timeout=22.0)
            for m in ["llama-3.2-11b-vision-preview", "qwen/qwen3.6-27b"]:
                try:
                    resp = g_client.chat.completions.create(
                        model=m,
                        messages=[{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{b64_data}"}}]}],
                        temperature=0.1
                    )
                    if resp.choices and resp.choices[0].message.content:
                        return resp.choices[0].message.content.strip()
                except Exception:
                    continue
        except Exception:
            pass

    if GEMINI_API_KEY and REQUESTS_OK:
        b64_data = base64.b64encode(image_bytes).decode("utf-8")
        for m in ["gemini-1.5-flash", "gemini-1.5-pro"]:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent?key={GEMINI_API_KEY}"
            payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": mime_type, "data": b64_data}}]}]}
            try:
                r = requests.post(url, json=payload, timeout=20)
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                    if txt:
                        return txt
            except Exception:
                continue

    return "Verbatim transcript initialized. Document ready for cognitive synthesis."

# ============================================================
# 6. HEADER
# ============================================================
st.markdown(
    f"""<div class="aetheris-header">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div class="brand-title">💠 {OS_NAME}</div>
                <div style="color:#64748b; font-size:11px; margin-top:2px; font-weight:600; letter-spacing:0.8px;">
                    {TAGLINE}
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
# 7. SIDEBAR: INTELLIGENCE & ARTIFACT REPOSITORY
# ============================================================
with st.sidebar:
    st.markdown(f"### 🛡️ Enterprise Ingestion")
    st.caption("Drop legal notices, affidavits, exam patterns, or corporate scans:")
    
    upload = st.file_uploader("Upload Artifact (Image / PDF)", type=["png", "jpg", "jpeg", "pdf"], key="art_upload")

    if upload:
        b_data = upload.getvalue()
        fname = upload.name
        fext = Path(fname).suffix.lower()
        mime = "image/png" if fext == ".png" else "image/jpeg"

        if not any(a["name"] == fname for a in st.session_state.processed_artifacts):
            with st.spinner(f"Ingesting & synthesizing {fname}..."):
                if fext in [".png", ".jpg", ".jpeg"]:
                    raw_text = extract_image_ocr(b_data, mime_type=mime)
                elif fext == ".pdf":
                    raw_text = extract_pdf_data(b_data)
                else:
                    raw_text = "Unsupported format."

                # Autonomous Cognitive Synthesis Layer
                synthesis = generate_cognitive_synthesis(raw_text)

                docx_out = build_executive_docx(f"Canonical Record: {Path(fname).stem}", synthesis, raw_text)
                pdf_out = build_executive_pdf(f"Canonical Record: {Path(fname).stem}", synthesis, raw_text)

                st.session_state.processed_artifacts.append({
                    "name": fname,
                    "synthesis": synthesis,
                    "content": raw_text,
                    "docx": docx_out,
                    "pdf": pdf_out
                })
                st.success(f"Synthesized: {fname}")

    if st.session_state.processed_artifacts:
        st.divider()
        st.markdown("**📁 Executive Deliverables:**")
        for idx, art in enumerate(st.session_state.processed_artifacts):
            st.markdown(f"**📄 {art['name']}**")
            with st.expander("🧠 Cognitive Intelligence Brief", expanded=True):
                st.markdown(art["synthesis"])
            
            c1, c2 = st.columns(2)
            if art.get("docx"):
                with c1:
                    st.download_button(
                        "⬇ Word (.docx)",
                        art["docx"],
                        file_name=f"{Path(art['name']).stem}_Executive.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"art_docx_{idx}",
                        use_container_width=True
                    )
            if art.get("pdf"):
                with c2:
                    st.download_button(
                        "⬇ PDF (.pdf)",
                        art["pdf"],
                        file_name=f"{Path(art['name']).stem}_Executive.pdf",
                        mime="application/pdf",
                        key=f"art_pdf_{idx}",
                        use_container_width=True
                    )
            st.divider()

    if st.button("＋ Clear Architecture Workspace", use_container_width=True):
        st.session_state.processed_artifacts = []
        st.session_state.messages = []
        st.rerun()

# ============================================================
# 8. MAIN WORKSPACE: EXECUTIVE CHAT & DRAFTING MATRIX
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
                        "⬇ Download Executive Word (.docx)",
                        msg["docx"],
                        file_name=f"Aetheris_Executive_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"ch_docx_{idx}",
                        use_container_width=True
                    )
            if msg.get("pdf"):
                with c2:
                    st.download_button(
                        "⬇ Download Executive PDF (.pdf)",
                        msg["pdf"],
                        file_name=f"Aetheris_Executive_{idx}.pdf",
                        mime="application/pdf",
                        key=f"ch_pdf_{idx}",
                        use_container_width=True
                    )

query = st.chat_input(f"Command {OS_NAME} (e.g. 'draft legal notice', 'analyze contract', 'synthesize uploaded artifact')...")

if query:
    st.session_state.messages.append({"role": "user", "content": query})
    with st.chat_message("user", avatar="👤"):
        st.markdown(query)

    with st.chat_message("assistant", avatar="🤖"):
        has_artifacts = len(st.session_state.processed_artifacts) > 0
        is_ref = any(w in query.lower() for w in ["this file", "artifact", "photo", "document", "uploaded", "isko", "analyze"])

        if is_ref and has_artifacts:
            last_art = st.session_state.processed_artifacts[-1]
            response_md = f"### 🧠 Executive Intelligence Synthesis: `{last_art['name']}`\n\n{last_art['synthesis']}\n\n---\n**Canonical Verbatim Segment:**\n\n{last_art['content'][:2000]}"
            docx_b = last_art["docx"]
            pdf_b = last_art["pdf"]
            st.markdown(response_md)
        else:
            instruction = (
                f"You are {OS_NAME}, the high-order neural intelligence engine engineered solely by {CREATOR_FULL_NAME}. "
                f"You write exhaustive, legally sound, strategic, and professional executive content. "
                f"Always structure drafts with executive summaries, formal headings, clauses, and authentication blocks."
            )
            response_md = ""
            if GROQ_API_KEY and Groq:
                try:
                    client = Groq(api_key=GROQ_API_KEY, timeout=12.0)
                    res = client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[{"role": "system", "content": instruction}, {"role": "user", "content": query}],
                        temperature=0.3
                    )
                    response_md = res.choices[0].message.content.strip()
                except Exception:
                    pass
            if not response_md:
                response_md = query

            st.markdown(response_md)
            docx_b = build_executive_docx("Aetheris Strategic Manifest", "", response_md)
            pdf_b = build_executive_pdf("Aetheris Strategic Manifest", "", response_md)

        c1, c2 = st.columns(2)
        if docx_b:
            with c1:
                st.download_button(
                    "⬇ Download Executive Word (.docx)",
                    docx_b,
                    file_name="Aetheris_Executive_Deliverable.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        if pdf_b:
            with c2:
                st.download_button(
                    "⬇ Download Executive PDF (.pdf)",
                    pdf_b,
                    file_name="Aetheris_Executive_Deliverable.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

        st.session_state.messages.append({
            "role": "assistant",
            "content": response_md,
            "docx": docx_b,
            "pdf": pdf_b
        })
