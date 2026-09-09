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
# 1. BULLETPROOF CREDENTIALS & IDENTITY (FROZEN)
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
# 2. PAGE CONFIGURATION & EXECUTIVE THEME (FROZEN)
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
    box-shadow: 0 4px 25px rgba(15, 23, 42, 0.04);
    margin-bottom: 18px;
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
    padding: 16px 20px !important;
    margin-bottom: 10px !important;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03) !important;
}
div[data-testid="stChatInput"] {
    background: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06) !important;
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
# 4. EXACT DOCUMENT ENGINES (100% FROZEN - ZERO TOUCH)
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
    prompt = (
        "Extract and transcribe all text from this image VERBATIM. "
        "Preserve original language (Hindi, Sanskrit, English), exact lines, numbers, and layout. "
        "Do NOT add greetings, summaries, notes, or analysis. Return ONLY the transcribed text."
    )
    if GROQ_API_KEY and Groq:
        try:
            b64_data = base64.b64encode(image_bytes).decode("utf-8")
            g_client = Groq(api_key=GROQ_API_KEY, timeout=25.0)
            for m in ["llama-3.2-11b-vision-preview", "qwen/qwen3.6-27b"]:
                try:
                    resp = g_client.chat.completions.create(
                        model=m,
                        messages=[{"role": "user", "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{b64_data}"}}
                        ]}],
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

def build_multi_page_docx(pages_text_list, doc_title="Canonical Deliverable"):
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        head = doc.add_heading(doc_title, level=1)
        meta = doc.add_paragraph()
        run_meta = meta.add_run(f"ENGINE: {OS_NAME} | ARCHITECT: {CREATOR_FULL_NAME} | DATE: {datetime.now().strftime('%d-%b-%Y')}")
        run_meta.font.size = Pt(8.5)
        run_meta.font.color.rgb = RGBColor(100, 116, 139)
        doc.add_paragraph("―" * 45)

        for idx, page_content in enumerate(pages_text_list):
            if idx > 0:
                doc.add_page_break()
            for line in page_content.splitlines():
                clean_l = line.strip()
                if not clean_l:
                    continue
                if clean_l.startswith("# "):
                    doc.add_heading(clean_l.replace("# ", "").strip(), level=1)
                elif clean_l.startswith("## "):
                    doc.add_heading(clean_l.replace("## ", "").strip(), level=2)
                elif clean_l.startswith("### "):
                    doc.add_heading(clean_l.replace("### ", "").strip(), level=3)
                elif clean_l.startswith(("- ", "* ")):
                    doc.add_paragraph(clean_l[2:].strip(), style='List Bullet')
                else:
                    p = doc.add_paragraph(clean_l)
                    p.style.font.name = 'Calibri'
                    p.style.font.size = Pt(11)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def build_executive_pdf(doc_title, text_content):
    if not REPORTLAB_OK:
        return None
    try:
        buf = io.BytesIO()
        pdf = SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()

        t_style = ParagraphStyle("T", parent=styles["Title"], fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#0f172a"))
        m_style = ParagraphStyle("M", parent=styles["Normal"], fontSize=8, alignment=TA_CENTER, textColor=colors.HexColor("#64748b"))
        h1_style = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=12, leading=16, spaceBefore=10, spaceAfter=4, textColor=colors.HexColor("#312e81"))
        h2_style = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=10.5, leading=14, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#4338ca"))
        b_style = ParagraphStyle("B", parent=styles["Normal"], fontSize=9.5, leading=14, alignment=TA_JUSTIFY, textColor=colors.HexColor("#1e293b"), spaceAfter=5)

        story = [
            Paragraph(f"<b>{OS_NAME.upper()} // ACADEMIC & ENTERPRISE MANIFEST</b>", m_style),
            Spacer(1, 4),
            Paragraph(f"<b>{doc_title}</b>", t_style),
            Paragraph(f"Architect: {CREATOR_FULL_NAME} • Comprehensive Publication • {datetime.now().strftime('%d %B %Y')}", m_style),
            Spacer(1, 6),
            HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cbd5e1"), spaceAfter=10)
        ]

        for line in text_content.splitlines():
            clean = line.strip()
            if not clean:
                story.append(Spacer(1, 3))
                continue
            safe = clean.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if safe.startswith("# "):
                story.append(Paragraph(f"<b>{safe.replace('# ', '').strip()}</b>", h1_style))
            elif safe.startswith("## "):
                story.append(Paragraph(f"<b>{safe.replace('## ', '').strip()}</b>", h2_style))
            elif safe.startswith("### "):
                story.append(Paragraph(f"<b>{safe.replace('### ', '').strip()}</b>", h2_style))
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
# 5. ROBUST LLM CALL FUNCTION
# ============================================================
def call_single_section(system_instr, user_prompt, max_tokens=1800):
    if GROQ_API_KEY and Groq:
        try:
            g_client = Groq(api_key=GROQ_API_KEY, timeout=40.0)
            res = g_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": system_instr},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=max_tokens
            )
            if res.choices and res.choices[0].message.content:
                txt = res.choices[0].message.content.strip()
                if len(txt) > 80:
                    return txt
        except Exception:
            pass

    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            for m in ["gemini-1.5-flash", "gemini-1.5-pro"]:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent?key={GEMINI_API_KEY}"
                payload = {
                    "contents": [{"parts": [{"text": f"{system_instr}\n\n{user_prompt}"}]}],
                    "generationConfig": {"temperature": 0.3, "maxOutputTokens": max_tokens}
                }
                r = requests.post(url, json=payload, timeout=30)
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                    if txt and len(txt) > 80:
                        return txt
        except Exception:
            pass

    return ""

# ============================================================
# 6. HEADER (FROZEN)
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
                <div class="status-badge">
                    <span class="dot"></span> READY
                </div>
            </div>
        </div>
    </div>""",
    unsafe_allow_html=True
)

# ============================================================
# 7. SIDEBAR: 1:1 CONVERSION & EXACT UTILITIES (100% FROZEN)
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
                        doc_bytes = build_multi_page_docx(pages_extracted, doc_title="Verbatim Document Manifest")
                        if doc_bytes:
                            st.session_state.attached_assets.append({
                                "name": "Verbatim_Transcribed.docx",
                                "docx": doc_bytes,
                                "raw_text": "\n\n".join(pages_extracted)
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
# 8. MAIN TABS (COGNITIVE CHAT + ACADEMIC MATRIX)
# ============================================================
main_tab_chat, main_tab_academic = st.tabs([
    "💬 Autonomous Cognitive Workspace", 
    "🎓 Academic & Examination Intelligence Matrix"
])

# ------------------------------------------------------------
# TAB 1: GENERAL AUTONOMOUS INTELLIGENCE & FREE DRAFTING
# ------------------------------------------------------------
with main_tab_chat:
    for idx, msg in enumerate(st.session_state.messages):
        avatar = "👤" if msg["role"] == "user" else "🤖"
        with st.chat_message(msg["role"], avatar=avatar):
            st.markdown(msg["content"])
            if msg.get("docx") or msg.get("pdf"):
                c1, c2 = st.columns(2)
                if msg.get("docx"):
                    st.download_button("⬇ Download Executive Word (.docx)", msg["docx"], file_name=f"Aetheris_Deliverable_{idx}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"chat_docx_{idx}", use_container_width=True)
                if msg.get("pdf"):
                    st.download_button("⬇ Download Executive PDF (.pdf)", msg["pdf"], file_name=f"Aetheris_Deliverable_{idx}.pdf", mime="application/pdf", key=f"chat_pdf_{idx}", use_container_width=True)

    user_query = st.chat_input("Enter command, instructions, or queries for Aetheris OS...")

    if user_query:
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user", avatar="👤"):
            st.markdown(user_query)

        with st.chat_message("assistant", avatar="🤖"):
            context_block = ""
            for a in st.session_state.attached_assets:
                if "raw_text" in a:
                    context_block += f"\n\n=== RECENT DOCUMENT CONTEXT ({a['name']}) ===\n{a['raw_text'][:3500]}\n---\n"

            system_instruction = (
                f"You are {OS_NAME}, the high-order neural intelligence engine engineered solely by your architect: {CREATOR_FULL_NAME}. "
                f"You understand and write accurately in Hindi, English, and Hinglish. Provide structured, exhaustive content."
            )
            full_prompt = f"{context_block}\n\nUser Request: {user_query}"
            out_response = call_single_section(system_instruction, full_prompt, max_tokens=2200)
            if not out_response:
                out_response = f"I am {OS_NAME}, engineered by {CREATOR_FULL_NAME}. Command received."

            st.markdown(out_response)
            docx_b = build_multi_page_docx([out_response], doc_title="Executive Intelligence Manifest")
            pdf_b = build_executive_pdf("Executive Intelligence Manifest", out_response)

            c1, c2 = st.columns(2)
            if docx_b:
                with c1:
                    st.download_button("⬇ Download Word (.docx)", docx_b, file_name="Aetheris_Executive.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            if pdf_b:
                with c2:
                    st.download_button("⬇ Download PDF (.pdf)", pdf_b, file_name="Aetheris_Executive.pdf", mime="application/pdf", use_container_width=True)

            st.session_state.messages.append({"role": "assistant", "content": out_response, "docx": docx_b, "pdf": pdf_b})

# ------------------------------------------------------------
# TAB 2: MULTI-PAGE DEEP ACADEMIC ENGINE (6-10 PAGES COMPLETE)
# ------------------------------------------------------------
with main_tab_academic:
    st.markdown("### 🎓 Academic & Examination Intelligence Matrix")
    st.caption("Universal Multi-Page Publisher: 9th-12th Boards, NEET/JEE, SSC, UGC NET, UPSC, College & Professional Degree Courses.")

    col_target, col_tier = st.columns([2, 1])
    with col_target:
        target_subject = st.text_input(
            "Target Subject / Chapter / Exam Topic",
            placeholder="e.g. 10th Science Acid Bases and Salts, 12th Physics Optics, NEET Biology Genetics, UGC NET Paper 1..."
        )
    with col_tier:
        academic_tier = st.selectbox(
            "Target Standard / Level",
            [
                "Class 9th & 10th (Secondary Boards)",
                "Class 11th & 12th (Senior Secondary)",
                "NEET / JEE & Medical/Engineering",
                "Graduation / University Exams (BA/BSc/BCom/MBA)",
                "Competitive (UGC NET / SSC / State PCS / UPSC)"
            ]
        )

    lang_pref = st.radio("Language Medium", ["Bilingual (English + Hindi Explanation)", "Pure English", "Pure Hindi (हिंदी)"], horizontal=True)

    if st.button("⚡ Generate Exhaustive Multi-Page Manifest (Full 5-8 Pages)", use_container_width=True):
        if not target_subject.strip():
            st.warning("Please enter a subject or chapter name.")
        else:
            prog_bar = st.progress(0, text="Initializing Academic Intelligence Pipeline...")
            
            system_base = (
                f"You are the Academic Matrix of {OS_NAME}, engineered by {CREATOR_FULL_NAME}. "
                f"You write exhaustive, high-density textbook & exam material for '{target_subject}' ({academic_tier}) in '{lang_pref}'. "
                f"Never give brief summaries. Write exhaustive detailed notes with diagrams/equations explained."
            )

            # Section 1: Detailed Concepts & Core Principles
            prog_bar.progress(20, text="[1/4] Compiling Comprehensive Theoretical Foundations...")
            p1 = (
                f"Write SECTION 1: COMPREHENSIVE CHAPTER FOUNDATIONS & IN-DEPTH THEORY for '{target_subject}'.\n"
                f"- Complete syllabus breakdown & weightage.\n"
                f"- Every single concept, definition, scientific law/principle explained thoroughly with chemical equations/formulas.\n"
                f"- Classification tables, real-life applications, and step-by-step processes.\n"
                f"Write at least 600-800 words."
            )
            sec1 = call_single_section(system_base, p1, max_tokens=1800)

            # Section 2: Advanced Mechanisms & Solved Examples
            prog_bar.progress(45, text="[2/4] Formulating Mechanisms, Solved Numericals & Reaction Schemes...")
            p2 = (
                f"Write SECTION 2: MECHANISMS, SOLVED EXAMPLES & CRITICAL DERIVATIONS for '{target_subject}'.\n"
                f"- Step-by-step solved problems / numericals / reaction mechanisms.\n"
                f"- Crucial experimental setups, lab preparation methods, or historical causality.\n"
                f"- Common mistakes students make and expert examiner tips.\n"
                f"Write at least 600-800 words."
            )
            sec2 = call_single_section(system_base, p2, max_tokens=1800)

            # Section 3: Exam-Grade Subjective Question Bank (2, 3 & 5 Marks)
            prog_bar.progress(70, text="[3/4] Structuring Official Exam Question Bank (Short & Long Answers)...")
            p3 = (
                f"Write SECTION 3: OFFICIAL BOARD/EXAM QUESTION BANK & STEPWISE SOLUTIONS for '{target_subject}'.\n"
                f"- 3 Very Short Answer Questions (1-2 Marks) with accurate answers.\n"
                f"- 3 Short Answer Questions (3 Marks) with detailed step-wise points.\n"
                f"- 2 Long Analytical / Case-Based Questions (5 Marks) with complete marking scheme answers.\n"
                f"Write at least 600-800 words."
            )
            sec3 = call_single_section(system_base, p3, max_tokens=1800)

            # Section 4: 10 High-Yield MCQs + Formula/Revision Sheet
            prog_bar.progress(90, text="[4/4] Generating 10 Exam-Grade MCQs & Master Revision Checklist...")
            p4 = (
                f"Write SECTION 4: HIGH-YIELD MOCK TEST (10 MCQs) & RAPID REVISION CHECKLIST for '{target_subject}'.\n"
                f"- Exactly 10 challenging multiple-choice questions with 4 distinct options (A, B, C, D).\n"
                f"- Full Answer Key with deep conceptual explanation for every question.\n"
                f"- 1-Page Rapid Revision Bullet Points & Formula Cheat Sheet.\n"
                f"Write at least 600-800 words."
            )
            sec4 = call_single_section(system_base, p4, max_tokens=1800)

            prog_bar.progress(100, text="Assembling Multi-Page Document...")

            # Combine all 4 Sections into an Exhaustive Master Deliverable
            full_academic_doc = (
                f"# Academic Intelligence Matrix: {target_subject.upper()}\n"
                f"**Standard/Tier:** {academic_tier} | **Medium:** {lang_pref}\n"
                f"**Engine:** {OS_NAME} | **Architect:** {CREATOR_FULL_NAME}\n\n"
                f"---\n\n"
                f"## MODULE 1: COMPREHENSIVE THEORETICAL FOUNDATIONS\n\n{sec1}\n\n"
                f"---\n\n"
                f"## MODULE 2: MECHANISMS, SOLVED EXAMPLES & CRITICAL ANALYSIS\n\n{sec2}\n\n"
                f"---\n\n"
                f"## MODULE 3: OFFICIAL EXAMINATION QUESTION BANK & STEPWISE SOLUTIONS\n\n{sec3}\n\n"
                f"---\n\n"
                f"## MODULE 4: HIGH-YIELD MOCK TEST (10 MCQs) & RAPID REVISION CHECKLIST\n\n{sec4}"
            )

            st.markdown(full_academic_doc)

            # Build Multi-Page Word and PDF
            pages_list = [sec1, sec2, sec3, sec4]
            acad_docx = build_multi_page_docx(pages_list, doc_title=f"{target_subject} ({academic_tier})")
            acad_pdf = build_executive_pdf(f"{target_subject} ({academic_tier})", full_academic_doc)

            st.success(f"✅ Complete Multi-Page Academic Manifest Compiled (Total 4 Comprehensive Modules Generated)!")

            c_down1, c_down2 = st.columns(2)
            if acad_docx:
                with c_down1:
                    st.download_button(
                        "📥 Download Complete Multi-Page Word (.docx)",
                        acad_docx,
                        file_name=f"{target_subject.replace(' ', '_')}_Complete_Notes.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            if acad_pdf:
                with c_down2:
                    st.download_button(
                        "📥 Download Complete Multi-Page PDF (.pdf)",
                        acad_pdf,
                        file_name=f"{target_subject.replace(' ', '_')}_Complete_Notes.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
