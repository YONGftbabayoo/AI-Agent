import streamlit as st
import os
import zipfile
from datetime import datetime
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
from ai_backend import (
    extract_text_from_pdf, extract_text_from_docx, extract_text_from_image,
    detect_language, translate_contextual, correct_grammar, summarize_text,
    score_translation_quality, extract_entities, classify_topic,
    ensure_term_consistency, build_glossary, save_to_docx,
    translate_deepl, translate_literal, send_email_with_attachments
)
from gtts import gTTS

st.set_page_config(page_title="Mandarin AI Translator", layout="wide")

st.markdown("""
<style>
    .main {
        background-color: #f7f9fb;
        padding: 2rem;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

st.title("🈶 Chinese ➔ Bahasa Indonesia AI Translator")
st.caption("Powerful Chinese to Indonesia AI Agent translator.")

st.sidebar.title("📂 Choose input mode")
mode = st.sidebar.radio("Choose Mode:", ["Single File", "Batch Files"])

def zip_batch_results(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, folder_path)
                zipf.write(full_path, arcname=arcname)

if mode == "Single File":
    st.header("📄 Process Single Document")

    uploaded_file = st.file_uploader("Upload file (PDF, DOCX, Image)", type=["pdf", "docx", "png", "jpg", "jpeg"])
    manual_text = st.text_area("✍️ Or paste Chinese text here", height=150)

    text = ""
    if uploaded_file:
        file_path = f"temp_upload/{uploaded_file.name}"
        os.makedirs("temp_upload", exist_ok=True)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())

        if file_path.endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
        elif file_path.endswith(".docx"):
            text = extract_text_from_docx(file_path)
        else:
            text = extract_text_from_image(file_path)

        st.success("✅ Successfully read the file.")

    elif manual_text.strip():
        text = manual_text.strip()

    if text:
        st.divider()
        st.subheader("📑 Original text")
        st.text_area("", text, height=200)

        if st.button("🚀 Start translate"):
            with st.spinner("🔄 Detecting language and translating..."):
                language = detect_language(text)
                contextual = translate_contextual(text) or translate_deepl(text) or translate_literal(text) or ""

            if not contextual:
                st.error("❌ Failed to translate. Try again.")
                st.stop()

            contextual = ensure_term_consistency(contextual)

            with st.spinner("🛠️ Checking grammar..."):
                corrected = correct_grammar(contextual)

            with st.spinner("🧠 Summarize your document..."):
                summary = summarize_text(corrected)

            with st.spinner("📊 Evaluate the translation's quality..."):
                quality_score = score_translation_quality(text, corrected)

            with st.spinner("🔍 Analyze entity and topic..."):
                entities = extract_entities(text)
                topic = classify_topic(corrected)

            base_name = f"translation_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            os.makedirs("Result", exist_ok=True)
            glossary_path = f"Result/glossary_{base_name}.txt"
            docx_path = f"Result/{base_name}.docx"
            audio_path = f"Result/{base_name}.mp3"

            build_glossary(text, contextual, glossary_path)
            save_to_docx(text, contextual, contextual, corrected, summary, quality_score, entities, topic, docx_path)

            try:
                tts = gTTS(contextual, lang='id')
                tts.save(audio_path)
                st.session_state['audio_path'] = audio_path
            except Exception as e:
                st.warning(f"⚠️ Failed to create audio: {e}")

            # Simpan hasil ke session state
            st.session_state['docx_path'] = docx_path
            st.session_state['glossary_path'] = glossary_path

            st.success("✅ Translation done!")

            with st.expander("📘 See the result"):
                st.text_area("Translation (contextual and fixed)", corrected, height=200)

            with st.expander("🧠 Document summarization"):
                st.write(summary)

            with st.expander("📊 Score dan analysis"):
                st.write("**Translation quality score:**", quality_score)
                st.write("**Main topic:**", topic)

            with st.expander("🔍 Detected entity's name"):
                for ent in entities:
                    st.markdown(f"- **{ent['text']}** ({ent['label']})")

            col1, col2, col3 = st.columns(3)
            with col1:
                with open(docx_path, "rb") as f:
                    st.download_button("⬇️ Download result (.docx)", data=f, file_name=os.path.basename(docx_path))
            with col2:
                with open(glossary_path, "rb") as f:
                    st.download_button("⬇️ Download glossary", data=f, file_name=os.path.basename(glossary_path))
            with col3:
                if 'audio_path' in st.session_state:
                    with open(st.session_state['audio_path'], "rb") as f:
                        st.download_button("⬇️ Download Audio (.mp3)", data=f, file_name=os.path.basename(st.session_state['audio_path']))

            if 'audio_path' in st.session_state:
                st.subheader("🎵 Listen audio")
                audio_file = open(st.session_state['audio_path'], 'rb')
                audio_bytes = audio_file.read()
                st.audio(audio_bytes, format='audio/mp3')

st.subheader("✉️ Send to email")
email = st.text_input("Enter the destination email...")
if email and st.button("Send the results"):
    if 'docx_path' in st.session_state and 'glossary_path' in st.session_state:
        try:
            attachments = [st.session_state['docx_path'], st.session_state['glossary_path']]
            if 'audio_path' in st.session_state:
                attachments.append(st.session_state['audio_path'])
            send_email_with_attachments(email, "Your document translation result", "Following are the result of your translation:", attachments)
            st.success(f"📬 Email successfully send to {email}! Thanks for using this AI Agent! ✨")
        except Exception as e:
            st.error(f"❌ Failed to send email: {e}")
    else:
        st.warning("⚠️ Please translate the document first before sending email.")

elif mode == "Batch Files":
    st.header("📁 Process multiples document")

    uploaded_files = st.file_uploader("Upload many file (PDF, DOCX, Gambar)", type=["pdf", "docx", "png", "jpg", "jpeg"], accept_multiple_files=True)

    if uploaded_files and st.button("🚀 Translate all"):
        st.info("🔄 Processing all files. Please wait...")
        os.makedirs("Result/Batch", exist_ok=True)
        for uploaded_file in uploaded_files:
            file_path = f"temp_upload/{uploaded_file.name}"
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if file_path.endswith(".pdf"):
                text = extract_text_from_pdf(file_path)
            elif file_path.endswith(".docx"):
                text = extract_text_from_docx(file_path)
            else:
                text = extract_text_from_image(file_path)

            contextual = translate_contextual(text) or translate_deepl(text) or translate_literal(text) or ""
            contextual = ensure_term_consistency(contextual)
            corrected = correct_grammar(contextual)
            summary = summarize_text(corrected)
            quality_score = score_translation_quality(text, corrected)
            entities = extract_entities(text)
            topic = classify_topic(corrected)

            base_name = os.path.splitext(uploaded_file.name)[0]
            file_folder = f"Result/Batch/{base_name}"
            os.makedirs(file_folder, exist_ok=True)
            glossary_path = f"{file_folder}/glossary.txt"
            docx_path = f"{file_folder}/translation_results.docx"
            build_glossary(text, contextual, glossary_path)
            save_to_docx(text, contextual, contextual, corrected, summary, quality_score, entities, topic, docx_path)

        zip_path = "Result/Batch/batch_results.zip"
        zip_batch_results("Result/Batch", zip_path)

        st.success("✅ All files have been translated!")
        with open(zip_path, "rb") as f:
            st.download_button("⬇️ Download all files (ZIP)", data=f, file_name="batch_results.zip")

        st.subheader("✉️ Send all batch results to email")
        email_batch = st.text_input("Enter your destination email for the batch....")
        if email_batch and st.button("Send Batch Email"):
            try:
                send_email_with_attachments(email_batch, "Your Batch Translation Results", "Here are the results of all your batch documents:", [zip_path])
                st.success(f"📬 All have been successfully send to {email_batch}! Thanks for using this AI Agent! ✨")
            except Exception as e:
                st.error(f"❌ Failed to send email: {e}")
