from pathlib import Path
import streamlit as st

from config import SYSTEM_NAME, CREATOR_NAME, TAGLINE
from core_engine import ask_aetheris
from document_engine import (
    convert_images_to_exact_pdf,
    read_file_content,
    export_to_docx,
    export_to_pdf
)

st.set_page_config(
    page_title=f"{SYSTEM_NAME} • {CREATOR_NAME}",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; color: #0f172a; }
.stApp { background: #f8fafc; }
.block-container { max-width: 1140px; padding-top: 1.2rem; padding-bottom: 3.5rem; }
.aetheris-header { padding: 18px 24px; border-radius: 14px; background: #ffffff; border: 1px solid #e2e8f0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04); margin-bottom: 16px; }
.brand-title { font-size: 24px; font-weight: 800; background: linear-gradient(90deg, #0f172a, #4338ca); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
.architect-badge { background: rgba(99, 102, 241, 0.1); border: 1px solid rgba(99, 102, 241, 0.3); padding: 5px 14px; border-radius: 999px; font-size: 11px; font-weight: 700; color: #4338ca; }
div[data-testid="stChatMessage"] { background: #ffffff !important; border: 1px solid #e2e8f0 !important; border-radius: 12px !important; margin-bottom: 10px !important; }
</style>""", unsafe_allow_html=True)

if "messages" not in st.session_state:
    st.session_state.messages = []
if "side_docs" not in st.session_state:
    st.session_state.side_docs = []

# Header
st.markdown(f"""<div class="aetheris-header">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <div class="brand-title">💠 {SYSTEM_NAME}</div>
            <div style="color:#64748b; font-size:11px; font-weight:600;">{TAGLINE}</div>
        </div>
        <span class="architect-badge">ARCHITECT: {CREATOR_NAME.upper()}</span>
    </div>
</div>""", unsafe_allow_html=True)

# Sidebar: 1:1 Image to PDF & OCR (100% Isolated)
with st.sidebar:
    st.markdown(f"**💠 {SYSTEM_NAME} MATRIX**")
    if st.button("＋ Clear Workspace", use_container_width=True):
        st.session_state.messages = []
        st.session_state.side_docs = []
        st.rerun()

    st.divider()
    st.markdown("### 📷 1:1 File & Photo Pipeline")
    files = st.file_uploader("Upload Photos or PDF", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    
    if files:
        st.info(f"Loaded {len(files)} asset(s).")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📄 Exact PDF", use_container_width=True):
                imgs = [f for f in files if Path(f.name).suffix.lower() in [".png", ".jpg", ".jpeg"]]
                if imgs:
                    pdf_bytes = convert_images_to_exact_pdf(imgs)
                    if pdf_bytes:
                        st.session_state.side_docs.append({"name": "Exact_Compilation.pdf", "pdf": pdf_bytes})
                        st.success("PDF Assembled!")
                else:
                    st.warning("Upload JPG/PNG images.")
        with c2:
            if st.button("📝 Exact Word", use_container_width=True):
                pages = [read_file_content(f) for f in files]
                valid_pages = [p for p in pages if p]
                if valid_pages:
                    docx_bytes = export_to_docx("Verbatim Manifest", "\n\n".join(valid_pages))
                    if docx_bytes:
                        st.session_state.side_docs.append({"name": "Verbatim_Transcribed.docx", "docx": docx_bytes})
                        st.success("Word Document Built!")

    if st.session_state.side_docs:
        st.divider()
        st.markdown("**📁 Sidebar Deliverables:**")
        for idx, item in enumerate(st.session_state.side_docs):
            st.markdown(f"**{item['name']}**")
            if "pdf" in item:
                st.download_button("⬇ Download PDF", item["pdf"], file_name=item["name"], mime="application/pdf", key=f"s_pdf_{idx}", use_container_width=True)
            if "docx" in item:
                st.download_button("⬇ Download Word", item["docx"], file_name=item["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"s_docx_{idx}", use_container_width=True)

# Main Navigation
tab_chat, tab_academic = st.tabs(["💬 Autonomous Workspace", "🎓 Academic & Examination Matrix"])

# TAB 1: PURE DIRECT CHAT (Clean, Authentic, No forced download buttons)
with tab_chat:
    for m in st.session_state.messages:
        with st.chat_message(m["role"], avatar="👤" if m["role"] == "user" else "🤖"):
            st.markdown(m["content"])

    user_query = st.chat_input("Ask anything (RTI drafting, conceptual doubts, competitive exam queries)...")
    if user_query:
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user", avatar="👤"):
            st.markdown(user_query)

        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Generating response..."):
                sys_instruction = (
                    f"You are {SYSTEM_NAME}, engineered by {CREATOR_NAME}. "
                    f"Answer authentically and thoroughly in Hindi, English, or Hinglish as appropriate. "
                    f"If asked to write an application, RTI, or legal notice, provide the complete, professional draft without skipping details."
                )
                reply = ask_aetheris(sys_instruction, user_query)
                st.markdown(reply)

        st.session_state.messages.append({"role": "assistant", "content": reply})

# TAB 2: ACADEMIC & COPY EVALUATION
with tab_academic:
    st.markdown("### 🎓 Academic & Examination Intelligence Matrix")
    mode = st.radio("Select Intelligence Mode:", ["📖 In-Depth Study Notes & MCQs", "📋 Question Paper & Answer Sheet Evaluator"], horizontal=True)

    if mode.startswith("📖"):
        c1, c2 = st.columns([2, 1])
        with c1:
            topic = st.text_input("Enter Topic / Chapter", placeholder="e.g. 10th Science Electricity, Modern Indian History 1857...")
        with c2:
            tier = st.selectbox("Standard / Exam", ["Secondary Boards (9th-10th)", "Senior Secondary (11th-12th)", "NEET / JEE", "UGC NET / SSC / State PCS"])

        if st.button("⚡ Generate Exhaustive Notes", use_container_width=True):
            if not topic.strip():
                st.warning("Please specify a topic.")
            else:
                with st.spinner(f"Compiling textbook-grade material for '{topic}'..."):
                    sys_academic = (
                        f"You are the Academic Head of {SYSTEM_NAME}. Write a high-density, multi-topic study manifest for '{topic}' ({tier}). "
                        f"Include: 1. Core Theory & Governing Laws/Formulas. 2. Solved Step-by-Step Examples. 3. 5 Exam Questions (Short & Long) with model answers. 4. 5 Exam-Grade MCQs with explanations."
                    )
                    notes = ask_aetheris(sys_academic, f"Generate complete study package for: {topic}", max_tokens=2800)
                    st.markdown(notes)

                    d_bytes = export_to_docx(f"{topic} Notes", notes)
                    p_bytes = export_to_pdf(f"{topic} Notes", notes)
                    st.success("✅ Manifest Compiled! Download below:")
                    cd1, cd2 = st.columns(2)
                    if d_bytes:
                        with cd1:
                            st.download_button("📥 Download Word (.docx)", d_bytes, file_name=f"{topic}_Notes.docx", use_container_width=True)
                    if p_bytes:
                        with cd2:
                            st.download_button("📥 Download PDF (.pdf)", p_bytes, file_name=f"{topic}_Notes.pdf", use_container_width=True)

    else:
        st.markdown("#### 📋 Examination & OMR Copy Evaluator")
        cq, ca = st.columns(2)
        with cq:
            q_file = st.file_uploader("1. Upload Question Paper (PDF / Image)", type=["pdf", "png", "jpg", "jpeg", "txt"], key="eval_q")
        with ca:
            a_file = st.file_uploader("2. Upload Student Answer Sheet (PDF / Image)", type=["pdf", "png", "jpg", "jpeg", "txt"], key="eval_a")

        if st.button("⚖️ Evaluate Answer Sheet", use_container_width=True):
            if not q_file or not a_file:
                st.warning("Please upload BOTH the Question Paper and Answer Sheet.")
            else:
                with st.spinner("Extracting text and performing question-by-question evaluation..."):
                    q_text = read_file_content(q_file)
                    a_text = read_file_content(a_file)

                    eval_sys = (
                        f"You are the Chief Examiner of {SYSTEM_NAME}. Evaluate the student's answer sheet strictly against the question paper. "
                        f"Output: 1. Scorecard (Marks obtained vs Total). 2. Question-by-question scoring and deductions. 3. Key mistakes and advice."
                    )
                    eval_report = ask_aetheris(eval_sys, f"QUESTION PAPER:\n{q_text[:3000]}\n\nSTUDENT ANSWERS:\n{a_text[:3000]}", max_tokens=2800)
                    st.markdown(eval_report)

                    ev_docx = export_to_docx("Evaluation Scorecard", eval_report)
                    ev_pdf = export_to_pdf("Evaluation Scorecard", eval_report)
                    cd1, cd2 = st.columns(2)
                    if ev_docx:
                        with cd1:
                            st.download_button("📥 Download Scorecard (.docx)", ev_docx, file_name="Evaluation_Report.docx", use_container_width=True)
                    if ev_pdf:
                        with cd2:
                            st.download_button("📥 Download Scorecard (.pdf)", ev_pdf, file_name="Evaluation_Report.pdf", use_container_width=True)
