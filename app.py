import os
import io
import json
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
    from docx.shared import Pt, Inches
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

# ============================================================
# 1. ENVIRONMENT & SECRETS
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
    background: #f8fafc;
}
.block-container {
    max-width: 1120px;
    padding-top: 1.5rem;
    padding-bottom: 3.5rem;
}
.aetheris-header {
    padding: 18px 24px;
    border-radius: 14px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04);
    margin-bottom: 20px;
}
.brand-title {
    font-size: 22px;
    font-weight: 800;
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
</style>""",
    unsafe_allow_html=True,
)

# ============================================================
# 3. DIRECT 1:1 EXACT CONVERTERS (ZERO LOSS / NO FAKE AI TEXT)
# ============================================================

def convert_images_to_exact_pdf(uploaded_images):
    """Takes original raw images and compiles them directly into a 1:1 exact multi-page PDF."""
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
        
        first_img.save(
            pdf_buf,
            format="PDF",
            save_all=True,
            append_images=remaining,
            resolution=100.0
        )
        pdf_buf.seek(0)
        return pdf_buf.getvalue()
    except Exception as e:
        st.error(f"Image to PDF Error: {str(e)}")
        return None

def extract_verbatim_ocr(image_bytes, mime_type="image/jpeg"):
    """Extracts verbatim text from image without adding any extra commentary."""
    prompt = (
        "Extract and transcribe all text from this image VERBATIM. "
        "Preserve the original language (Hindi, Sanskrit, English), exact lines, numbers, and layout. "
        "Do NOT add greetings, summaries, notes, or analysis. Return ONLY the transcribed text."
    )
    
    # 1. Groq Vision
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

    # 2. Gemini REST API
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

    return "No text detected or API connection issue."

def build_multi_page_word(pages_text_list):
    """Creates a clean Word (.docx) document with each page separated by real Page Breaks."""
    if not DOCX_OK:
        return None
    try:
        doc = Document()
        for idx, page_content in enumerate(pages_text_list):
            if idx > 0:
                doc.add_page_break()
            
            p_head = doc.add_paragraph()
            r = p_head.add_run(f"--- Page {idx + 1} ---")
            r.font.size = Pt(9.5)
            r.font.bold = True

            for line in page_content.splitlines():
                clean_line = line.strip()
                if clean_line:
                    p = doc.add_paragraph(clean_line)
                    p.style.font.name = 'Arial'
                    p.style.font.size = Pt(11)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        st.error(f"Word Generation Error: {str(e)}")
        return None

# ============================================================
# 4. HEADER
# ============================================================
st.markdown(
    f"""<div class="aetheris-header">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div class="brand-title">💠 {OS_NAME}</div>
                <div style="color:#64748b; font-size:12px; font-weight:500;">
                    1:1 EXACT DOCUMENT & IMAGE CONVERSION ENGINE
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
# 5. CORE INTERFACE: 2 DEDICATED CLEAN TABS
# ============================================================
tab_photo, tab_text = st.tabs(["📷 Photo & File Converter (Exact 1:1)", "✍️ Text to Word & PDF Generator"])

# ------------------------------------------------------------
# TAB 1: EXACT PHOTO / PDF CONVERTER
# ------------------------------------------------------------
with tab_photo:
    st.markdown("### 1:1 Photo / Scan to Exact PDF & Word")
    st.caption("Select one or multiple photos. It directly combines your photos into an exact high-res PDF or transcribes to Word without any extra AI commentary.")

    files = st.file_uploader(
        "Drop Photo(s) or PDF here (Multiple files allowed)",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
        key="multi_file_uploader"
    )

    if files:
        st.info(f"Loaded {len(files)} file(s). Choose conversion below:")
        
        col1, col2 = st.columns(2)
        
        # Action 1: Exact Photo to PDF (CamScanner style)
        with col1:
            st.markdown("#### 📄 Exact Photocopy to PDF")
            st.caption("Direct 1:1 conversion. Embeds your exact photos as full pages in a clean PDF.")
            if st.button("Generate Exact PDF (All Pages)", use_container_width=True):
                image_files = [f for f in files if Path(f.name).suffix.lower() in [".png", ".jpg", ".jpeg"]]
                if image_files:
                    with st.spinner("Stitching photos into exact PDF pages..."):
                        pdf_bytes = convert_images_to_exact_pdf(image_files)
                        if pdf_bytes:
                            st.success("✅ Exact Multi-Page PDF Created!")
                            st.download_button(
                                label="📥 Download Exact PDF",
                                data=pdf_bytes,
                                file_name="Converted_Document.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                else:
                    st.warning("Please upload image files (JPG/PNG) to compile into a photo-identical PDF.")

        # Action 2: Multi-Page Text to Word (.docx)
        with col2:
            st.markdown("#### 📝 Verbatim Text to Word (.docx)")
            st.caption("Transcribes each page verbatim into Word with proper page breaks.")
            if st.button("Extract Verbatim to Word (.docx)", use_container_width=True):
                pages_extracted = []
                with st.spinner(f"Extracting all {len(files)} page(s) verbatim..."):
                    for idx, f in enumerate(files):
                        fext = Path(f.name).suffix.lower()
                        b_data = f.getvalue()
                        
                        if fext in [".png", ".jpg", ".jpeg"]:
                            mime = "image/png" if fext == ".png" else "image/jpeg"
                            txt = extract_verbatim_ocr(b_data, mime_type=mime)
                            pages_extracted.append(txt)
                        elif fext == ".pdf" and PDF_OK:
                            reader = PdfReader(io.BytesIO(b_data))
                            for p in reader.pages:
                                pages_extracted.append(p.extract_text() or "")

                    if pages_extracted:
                        docx_bytes = build_multi_page_word(pages_extracted)
                        if docx_bytes:
                            st.success(f"✅ Converted {len(pages_extracted)} page(s) to Word!")
                            st.download_button(
                                label="📥 Download Multi-Page Word (.docx)",
                                data=docx_bytes,
                                file_name="Verbatim_Document.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )

# ------------------------------------------------------------
# TAB 2: TEXT TO WORD / PDF GENERATOR
# ------------------------------------------------------------
with tab_text:
    st.markdown("### ✍️ Pure Text to Word & PDF")
    st.caption("Paste any text, notes, or Hindi/English content. Download it instantly as clean Word or PDF without any AI summaries.")

    user_text = st.text_area("Paste or Type Content Here:", height=240, placeholder="Paste your complete Hindi / English / Sanskrit text here...")

    if user_text.strip():
        c_w1, c_w2 = st.columns(2)
        with c_w1:
            doc_single = build_multi_page_word([user_text])
            if doc_single:
                st.download_button(
                    label="📥 Download as Word (.docx)",
                    data=doc_single,
                    file_name="Typed_Document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        with c_w2:
            st.info("Tip: If you have photos, use the first tab for exact 1:1 photocopy PDF output.")
