import os
import io
import json
import uuid
import re
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
    from google import genai
except ImportError:
    genai = None

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
# 1. IDENTITY & ENVIRONMENT
# ============================================================
load_dotenv(override=True)
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

CREATOR_FULL_NAME = "Anshul Singh Rajpoot"
OS_NAME = "Aetheris OS"

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
# 4. DIRECT DETERMINISTIC CONVERTERS (ZERO HALLUCINATION)
# ============================================================
def build_docx_bytes(title, content_text):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        
        # Header Metadata
        head = doc.add_heading(title, level=1)
        meta = doc.add_paragraph()
        meta_run = meta.add_run(f"System: {OS_NAME} | Architect: {CREATOR_FULL_NAME} | Date: {datetime.now().strftime('%d-%b-%Y')}")
        meta_run.font.size = Pt(9)
        meta_run.font.color.rgb = RGBColor(100, 116, 139)
        doc.add_paragraph("―" * 45)

        # Body Paragraphs verbatim
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
    except Exception as e:
        st.error(f"DOCX Build Error: {str(e)}")
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
    except Exception as e:
        st.error(f"PDF Build Error: {str(e)}")
        return None

# ============================================================
# 5. VERBATIM EXTRACTION ENGINES (IMAGE OCR & PDF PARSER)
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
        return "\n\n".join(extracted) if extracted else "No selectable text found in PDF (scanned PDF requires OCR image upload)."
    except Exception as e:
        return f"PDF Extraction Error: {str(e)}"

def extract_text_from_image(image_bytes):
    if not GEMINI_API_KEY or not genai:
        return "Gemini API key missing for OCR processing."
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        img = Image.open(io.BytesIO(image_bytes))
        prompt = (
            "Transcribe all text from this image VERBATIM. "
            "Preserve every single word, sentence, number, Hindi, Sanskrit, or English character exactly as written. "
            "Do NOT summarize. Do NOT omit anything. Do NOT add conversational greetings or explanations. "
            "Only output the transcribed raw text."
        )
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[prompt, img]
        )
        return response.text.strip() if response and response.text else "No legible text found."
    except Exception as e:
        return f"Vision OCR Error: {str(e)}"

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

        # Process only if not already processed in this state
        if not any(d["name"] == fname for d in st.session_state.processed_docs):
            with st.spinner(f"Converting {fname} verbatim..."):
                if fext in [".png", ".jpg", ".jpeg"]:
                    extracted = extract_text_from_image(b_data)
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
            with st.expander("👁️ View Extracted Content"):
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

user_prompt = st.chat_input("Enter text to convert to Word/PDF or ask to draft a document...")

if user_prompt:
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user", avatar="👤"):
        st.markdown(user_prompt)

    with st.chat_message("assistant", avatar="🤖"):
        # Check if user specifically asks to convert the uploaded document
        is_convert_req = any(k in user_prompt.lower() for k in ["convert", "word me", "pdf me", "docx me", "badlo"])
        has_docs = len(st.session_state.processed_docs) > 0

        if is_convert_req and has_docs:
            last_doc = st.session_state.processed_docs[-1]
            out_text = f"✅ Successfully converted **{last_doc['name']}**! Here is the full extracted content:\n\n{last_doc['content']}"
            docx_file = last_doc["docx"]
            pdf_file = last_doc["pdf"]
            st.markdown(out_text)
        else:
            # Generate AI text via Groq/Gemini
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
            if not out_text and GEMINI_API_KEY and genai:
                try:
                    gclient = genai.Client(api_key=GEMINI_API_KEY)
                    gr = gclient.models.generate_content(model="gemini-2.5-flash", contents=f"{instruction}\n\n{user_prompt}")
                    out_text = gr.text.strip()
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
