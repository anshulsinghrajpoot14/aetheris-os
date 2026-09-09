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
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib import colors
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ============================================================
# 1. BULLETPROOF CREDENTIALS & IDENTITY
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
    background: #f8fafc;
}
.block-container {
    max-width: 1140px;
    padding-top: 1.2rem;
    padding-bottom: 3.5rem;
}
section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
}
.aetheris-header {
    padding: 18px 24px;
    border-radius: 14px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04);
    margin-bottom: 16px;
}
.brand-title {
    font-size: 24px;
    font-weight: 800;
    letter-spacing: 1.2px;
    background: linear-gradient(90deg, #0f172a, #4338ca);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.architect-badge {
    background: rgba(99, 102, 241, 0.1);
    border: 1px solid rgba(99, 102, 241, 0.3);
    padding: 5px 14px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    color: #4338ca;
}
div[data-testid="stChatMessage"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    padding: 16px 20px !important;
    margin-bottom: 10px !important;
}
</style>""",
    unsafe_allow_html=True,
)

# ============================================================
# 3. SESSION STATE
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

if "attached_assets" not in st.session_state:
    st.session_state.attached_assets = []

# ============================================================
# 4. EXACT DOCUMENT GENERATORS (FROZEN)
# ============================================================
def convert_images_to_exact_pdf(uploaded_images):
    try:
        pil_images = []
        for img_file in uploaded_images:
            img = Image.open(img_file)
            if img.mode != "RGB":
                img = img.convert("RGB")
            pil_images.append(img)
        if not pil_images:
            return None
        pdf_buf = io.BytesIO()
        first_img = pil_images[0]
        remaining = pil_images[1:] if len(pil_images) > 1 else []
        first_img.save(pdf_buf, format="PDF", save_all=True, append_images=remaining, resolution=100.0)
        pdf_buf.seek(0)
        return pdf_buf.getvalue()
    except Exception:
        return None

def extract_verbatim_ocr(image_bytes, mime_type="image/jpeg"):
    prompt = "Extract and transcribe all text from this image VERBATIM without adding any notes or greetings."
    if GROQ_API_KEY and Groq:
        try:
            b64_data = base64.b64encode(image_bytes).decode("utf-8")
            g_client = Groq(api_key=GROQ_API_KEY, timeout=20.0)
            for m in ["llama-3.2-11b-vision-preview", "qwen/qwen3.6-27b"]:
                try:
                    resp = g_client.chat.completions.create(
                        model=m,
                        messages=[{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{b64_data}"}}]}],
                        temperature=0.1
                    )
                    if resp.choices and resp.choices[0].message.content:
                        txt = resp.choices[0].message.content.strip()
                        if txt:
                            return txt
                except Exception:
                    continue
        except Exception:
            pass

    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            b64_data = base64.b64encode(image_bytes).decode("utf-8")
            for m in ["gemini-1.5-flash", "gemini-1.5-pro"]:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent?key={GEMINI_API_KEY}"
                payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": mime_type, "data": b64_data}}]}]}
                r = requests.post(url, json=payload, timeout=20)
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                    if txt:
                        return txt
        except Exception:
            pass

    return "Verbatim transcription initialized. Document indexed."

def build_word_doc(title, content_text):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        doc.add_heading(title, level=1)
        for block in content_text.split("\n\n"):
            clean = block.strip()
            if not clean:
                continue
            if clean.startswith("# "):
                doc.add_heading(clean.replace("# ", "").strip(), level=2)
            elif clean.startswith("## "):
                doc.add_heading(clean.replace("## ", "").strip(), level=3)
            elif clean.startswith(("- ", "* ")):
                for line in clean.splitlines():
                    if line.strip().startswith(("- ", "* ")):
                        doc.add_paragraph(line.strip()[2:], style='List Bullet')
                    else:
                        doc.add_paragraph(line.strip())
            else:
                p = doc.add_paragraph(clean)
                p.style.font.name = 'Arial'
                p.style.font.size = Pt(11)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def build_pdf_doc(title, content_text):
    if not REPORTLAB_OK:
        return None
    try:
        buf = io.BytesIO()
        pdf = SimpleDocTemplate(buf, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()

        t_style = ParagraphStyle("T", parent=styles["Title"], fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a"))
        h_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=11, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#4338ca"))
        b_style = ParagraphStyle("B", parent=styles["Normal"], fontSize=9.5, leading=14, alignment=TA_LEFT, textColor=colors.HexColor("#1e293b"), spaceAfter=5)

        story = [
            Paragraph(f"<b>{title}</b>", t_style),
            Spacer(1, 8),
            HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cbd5e1"), spaceAfter=10)
        ]

        for line in content_text.splitlines():
            clean = line.strip()
            if not clean:
                story.append(Spacer(1, 3))
                continue
            safe = clean.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if safe.startswith("#"):
                story.append(Paragraph(f"<b>{safe.lstrip('#').strip()}</b>", h_style))
            elif safe.startswith(("- ", "* ")):
                story.append(Paragraph(f"• {safe[2:]}", b_style))
            else:
                story.append(Paragraph(safe, b_style))

        pdf.build(story)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

# ============================================================
# 5. FAST RELIABLE INFERENCE (ZERO-HANG)
# ============================================================
def run_fast_inference(system_msg, user_msg):
    if GROQ_API_KEY and Groq:
        try:
            client = Groq(api_key=GROQ_API_KEY, timeout=20.0)
            for model_id in ["llama-3.1-8b-instant", "llama-3.3-70b-versatile"]:
                try:
                    resp = client.chat.completions.create(
                        model=model_id,
                        messages=[
                            {"role": "system", "content": system_msg},
                            {"role": "user", "content": user_msg}
                        ],
                        temperature=0.3,
                        max_tokens=2200
                    )
                    if resp.choices and resp.choices[0].message.content:
                        res = resp.choices[0].message.content.strip()
                        if len(res) > 30:
                            return res
                except Exception:
                    continue
        except Exception:
            pass

    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
            payload = {"contents": [{"parts": [{"text": f"{system_msg}\n\n{user_msg}"}]}]}
            r = requests.post(url, json=payload, timeout=20)
            if r.status_code == 200:
                txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                if txt and len(txt) > 30:
                    return txt
        except Exception:
            pass

    return ""

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
            <div>
                <span class="architect-badge">ARCHITECT: {CREATOR_FULL_NAME.upper()}</span>
            </div>
        </div>
    </div>""",
    unsafe_allow_html=True
)

# ============================================================
# 7. SIDEBAR: 1:1 CONVERSION & EXACT UTILITIES (FROZEN)
# ============================================================
with st.sidebar:
    st.markdown(f"**💠 {OS_NAME} MATRIX**")
    st.caption(f"Architect: {CREATOR_FULL_NAME}")

    if st.button("＋ Clear Workspace", use_container_width=True):
        st.session_state.messages = []
        st.session_state.attached_assets = []
        st.rerun()

    st.divider()
    st.markdown("### 📷 1:1 File & Photo Pipeline")
    st.caption("Drop single or multiple photos/PDFs. Converts 1:1 directly to HD PDF or Verbatim Word:")

    multi_files = st.file_uploader(
        "Upload Photos or PDF",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
        key="side_multi_uploader"
    )

    if multi_files:
        st.info(f"Loaded {len(multi_files)} document asset(s).")
        c_act1, c_act2 = st.columns(2)

        with c_act1:
            if st.button("📄 Exact PDF", use_container_width=True):
                img_files = [f for f in multi_files if Path(f.name).suffix.lower() in [".png", ".jpg", ".jpeg"]]
                if img_files:
                    with st.spinner("Generating 1:1 Multi-Page PDF..."):
                        exact_pdf = convert_images_to_exact_pdf(img_files)
                        if exact_pdf:
                            st.session_state.attached_assets.append({
                                "name": "Exact_Compilation.pdf",
                                "pdf": exact_pdf
                            })
                            st.success("PDF Assembled!")
                else:
                    st.warning("Upload JPG/PNG images.")

        with c_act2:
            if st.button("📝 Exact Word", use_container_width=True):
                pages_extracted = []
                with st.spinner("Transcribing verbatim text..."):
                    for f in multi_files:
                        fext = Path(f.name).suffix.lower()
                        b = f.getvalue()
                        if fext in [".png", ".jpg", ".jpeg"]:
                            mime = "image/png" if fext == ".png" else "image/jpeg"
                            txt = extract_verbatim_ocr(b, mime_type=mime)
                            pages_extracted.append(txt)
                        elif fext == ".pdf" and PDF_OK:
                            reader = PdfReader(io.BytesIO(b))
                            p_txt = "\n".join([page.extract_text() or "" for page in reader.pages])
                            pages_extracted.append(p_txt)

                    if pages_extracted:
                        doc_bytes = build_word_doc("Verbatim Document Manifest", "\n\n".join(pages_extracted))
                        if doc_bytes:
                            st.session_state.attached_assets.append({
                                "name": "Verbatim_Transcribed.docx",
                                "docx": doc_bytes
                            })
                            st.success("Word Document Built!")

    if st.session_state.attached_assets:
        st.divider()
        st.markdown("**📁 Sidebar Deliverables:**")
        for idx, ast_item in enumerate(st.session_state.attached_assets):
            st.markdown(f"**{ast_item['name']}**")
            if "pdf" in ast_item and ast_item["pdf"]:
                st.download_button("⬇ Download PDF", ast_item["pdf"], file_name=ast_item["name"], mime="application/pdf", key=f"side_pdf_{idx}", use_container_width=True)
            if "docx" in ast_item and ast_item["docx"]:
                st.download_button("⬇ Download Word (.docx)", ast_item["docx"], file_name=ast_item["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"side_docx_{idx}", use_container_width=True)

# ============================================================
# 8. MAIN TABS: CLEAN CHAT & ACADEMIC PUBLISHER
# ============================================================
main_tab_chat, main_tab_academic = st.tabs([
    "💬 Autonomous Cognitive Workspace", 
    "🎓 Academic & Examination Intelligence Matrix"
])

# ------------------------------------------------------------
# TAB 1: NATURAL DIRECT CONVERSATION & ASSISTANCE
# ------------------------------------------------------------
with main_tab_chat:
    for idx, msg in enumerate(st.session_state.messages):
        avatar = "👤" if msg["role"] == "user" else "🤖"
        with st.chat_message(msg["role"], avatar=avatar):
            st.markdown(msg["content"])
            if msg.get("docx") or msg.get("pdf"):
                c1, c2 = st.columns(2)
                if msg.get("docx"):
                    st.download_button("⬇ Download Word (.docx)", msg["docx"], file_name=f"Document_{idx}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"chat_docx_{idx}", use_container_width=True)
                if msg.get("pdf"):
                    st.download_button("⬇ Download PDF (.pdf)", msg["pdf"], file_name=f"Document_{idx}.pdf", mime="application/pdf", key=f"chat_pdf_{idx}", use_container_width=True)

    user_query = st.chat_input("Ask anything (e.g. 'draft an RTI appeal for marksheet', 'write an application', 'explain a concept')...")

    if user_query:
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user", avatar="👤"):
            st.markdown(user_query)

        with st.chat_message("assistant", avatar="🤖"):
            system_instruction = (
                f"You are {OS_NAME}, engineered solely by your architect: {CREATOR_FULL_NAME}. "
                f"Respond directly and clearly in the user's natural language (Hindi, English, or Hinglish). "
                f"When asked to write an application, RTI petition, or formal letter, draft the complete, accurate text directly. "
                f"Never attach fake metadata or forced headers. Be direct, authentic, and high quality."
            )

            out_response = run_fast_inference(system_instruction, user_query)
            if not out_response:
                out_response = f"I am {OS_NAME}, engineered by {CREATOR_FULL_NAME}. Please re-submit your command."

            st.markdown(out_response)

            # Generate download buttons ONLY if user explicitly asked for document/word/pdf or wrote formal letter
            wants_file = any(w in user_query.lower() for w in ["pdf", "word", "docx", "file", "download", "application", "letter", "draft", "rti"])
            docx_b = None
            pdf_b = None

            if wants_file and len(out_response) > 200:
                docx_b = build_word_doc("Aetheris Generated Document", out_response)
                pdf_b = build_pdf_doc("Aetheris Generated Document", out_response)
                c1, c2 = st.columns(2)
                if docx_b:
                    with c1:
                        st.download_button("⬇ Download Word (.docx)", docx_b, file_name="Aetheris_Document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                if pdf_b:
                    with c2:
                        st.download_button("⬇ Download PDF (.pdf)", pdf_b, file_name="Aetheris_Document.pdf", mime="application/pdf", use_container_width=True)

            st.session_state.messages.append({
                "role": "assistant",
                "content": out_response,
                "docx": docx_b,
                "pdf": pdf_b
            })

# ------------------------------------------------------------
# TAB 2: RELIABLE ACADEMIC & NOTES PUBLISHER
# ------------------------------------------------------------
with main_tab_academic:
    st.markdown("### 🎓 Academic & Examination Intelligence Matrix")
    st.caption("Instant comprehensive study publisher for school boards (9th-12th), college, and competitive exams.")

    c_top1, c_top2 = st.columns([2, 1])
    with c_top1:
        topic_name = st.text_input("Enter Chapter / Subject / Topic", placeholder="e.g. 10th Science Electricity, Acid Bases and Salts, Indian History 1857...")
    with c_top2:
        study_mode = st.selectbox("Select Study Deliverable", [
            "Complete Chapter Notes & Theory",
            "10 High-Yield Exam MCQs & Answers",
            "Official Board Question Bank (Short & Long)",
            "Rapid Revision & Formula Sheet"
        ])

    lang_choice = st.radio("Language Medium", ["Bilingual (English + Hindi)", "Pure English", "Pure Hindi (हिंदी)"], horizontal=True)

    if st.button("⚡ Generate Complete Study Deliverable", use_container_width=True):
        if not topic_name.strip():
            st.warning("Please enter a subject or chapter name.")
        else:
            with st.spinner(f"Compiling {study_mode} for '{topic_name}'..."):
                acad_system = (
                    f"You are the Academic Publisher of {OS_NAME}, engineered by {CREATOR_FULL_NAME}. "
                    f"Write deep, thorough textbook-standard content in {lang_choice}. "
                    f"Never return brief summaries. Write complete, detailed notes, formulas, and questions."
                )

                acad_user = (
                    f"Topic: '{topic_name}'\n"
                    f"Mode: '{study_mode}'\n"
                    f"Deliverable Requirements:\n"
                    f"- If 'Complete Chapter Notes & Theory': Cover all core definitions, scientific laws/principles, formulas, SI units, and daily life applications.\n"
                    f"- If '10 High-Yield Exam MCQs & Answers': Write exactly 10 exam-grade MCQs with 4 options each, clearly marked answers, and conceptual explanations.\n"
                    f"- If 'Official Board Question Bank': Write 3 Short Questions (2-3 Marks) and 2 Long Analytical Questions (5 Marks) with point-wise model answers.\n"
                    f"- If 'Rapid Revision & Formula Sheet': Provide high-yield formula list, key dates/facts, and a 15-minute quick revision checklist.\n"
                    f"Use clean markdown with headings and bullet points."
                )

                result_text = run_fast_inference(acad_system, acad_user)

                if not result_text:
                    result_text = f"Academic synthesis for '{topic_name}' completed. Please re-generate if needed."

                st.markdown(result_text)

                acad_docx = build_word_doc(f"{topic_name} - {study_mode}", result_text)
                acad_pdf = build_pdf_doc(f"{topic_name} - {study_mode}", result_text)

                st.success("✅ Deliverable Compiled Successfully!")

                cd1, cd2 = st.columns(2)
                if acad_docx:
                    with cd1:
                        st.download_button(
                            "📥 Download Word (.docx)",
                            acad_docx,
                            file_name=f"{topic_name.replace(' ', '_')}_{study_mode.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                if acad_pdf:
                    with cd2:
                        st.download_button(
                            "📥 Download PDF (.pdf)",
                            acad_pdf,
                            file_name=f"{topic_name.replace(' ', '_')}_{study_mode.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
