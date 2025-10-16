import streamlit as st
from transformer.app import AcademicTextHumanizer, NLP_GLOBAL, download_nltk_resources
from transformer.document_parser import (
    parse_document_file,
    detect_file_type,
    get_plain_text, 
    formatted_document_to_html, 
    formatted_document_to_docx,
    text_to_formatted_document,
    FormattedDocument,
    FormattedParagraph,
    FormattedRun
)
from nltk.tokenize import word_tokenize
import io




def main():
    """
    The `main` function sets up a Streamlit page for transforming user-provided text into a more formal
    academic style by expanding contractions, adding academic transitions, and optionally converting
    sentences to passive voice or replacing words with synonyms.
    """
    # Download NLTK resources if needed
    download_nltk_resources()

    # Configure Streamlit page
    st.set_page_config(
        page_title="Humanize AI Generated text",
        page_icon="🤖",
        layout="wide",
        initial_sidebar_state="expanded",
        menu_items={
            'Get Help': "https://github.com/DadaNanjesha/AI-Text-Humanizer-App/issues",
            'Report a bug': "https://github.com/DadaNanjesha/AI-Text-Humanizer-App/issues",
            'About': "# This app is used to Humanize AI generated text"
        }
    )

    # --- Custom CSS for Title Centering and Additional Styling ---
    st.markdown(
        """
        <style>
        /* Center the main title */
        .title {
            text-align: center;
            font-size: 2em;
            font-weight: bold;
            margin-top: 0.5em;
        }
        /* Center the subtitle / introduction block */
        .intro {
            text-align: left;
            line-height: 1.6;
            margin-bottom: 1.2em;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # --- Title / Intro ---
    st.markdown("<div class='title'>🧔🏻‍♂️Humanize AI🤖 Generated text</div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class='intro'>
        <p><b>This app transforms your text into a more formal academic style by:<b><br>
        • Expanding contractions<br>
        • Adding academic transitions<br>
        • <em>Optionally</em> converting some sentences to passive voice<br>
        • <em>Optionally</em> replacing words with synonyms for a more formal tone.</p>
        <hr>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Checkboxes
    use_passive = st.checkbox("Enable Passive Voice Transformation", value=False)
    use_synonyms = st.checkbox("Enable Synonym Replacement", value=False)

    # Controls
    formality = st.slider("Formality level", 0.0, 1.0, 0.5, 0.1, help="Increase to apply more transformations")
    ensure_change = st.checkbox("Ensure at least one change per sentence", value=True)

    # Text input
    user_text = st.text_area("Enter your text here:")

    # File upload
    uploaded_file = st.file_uploader("Or upload a .txt, .docx, or .pdf file:", type=["txt","docx","pdf"])
    formatted_doc = None
    if uploaded_file is not None:
        try:
            # Read file once and wrap with BytesIO
            file_bytes = uploaded_file.read()
            file_stream = io.BytesIO(file_bytes)
            
            # Parse using unified document parser
            formatted_doc = parse_document_file(file_stream)
            user_text = get_plain_text(formatted_doc)
            
            # Provide feedback on detected type
            file_type = detect_file_type(file_stream)
            st.success(f"Successfully loaded .{file_type} file with {len(formatted_doc.paragraphs)} paragraphs, {len(formatted_doc.list_items)} list items, and {len(formatted_doc.tables)} tables.")
        except Exception as e:
            error_message = str(e)
            if "encrypted" in error_message.lower():
                st.error("This PDF is password-protected. Upload an unencrypted file.")
            elif "no extractable text" in error_message.lower() or "scanned" in error_message.lower():
                st.warning("This PDF has no extractable text (scanned). Please upload a text-based PDF.")
            else:
                st.error("Failed to parse file. Ensure it's a valid .txt/.docx/.pdf.")
            user_text = ""
            formatted_doc = None

    # Button
    if st.button("Transform to Academic Style"):
        if not user_text.strip():
            st.warning("Please enter or upload some text to transform.")
        else:
            with st.spinner("Transforming text..."):
                # Input stats
                input_word_count = len(word_tokenize(user_text,language='english', preserve_line=True))
                doc_input = NLP_GLOBAL(user_text)
                input_sentence_count = len(list(doc_input.sents))

                # Transform probabilities based on formality level
                p_passive = max(0.0, min(1.0, 0.1 + 0.7 * formality))
                p_syn = max(0.0, min(1.0, 0.15 + 0.7 * formality))
                p_academic = max(0.0, min(1.0, 0.2 + 0.6 * formality))

                # Transformer instance with enforcement of visible change
                humanizer = AcademicTextHumanizer(
                    p_passive=p_passive,
                    p_synonym_replacement=p_syn,
                    p_academic_transition=p_academic,
                    ensure_change=ensure_change
                )

                # Inform user if high-quality synonym selection model isn't available
                if use_synonyms and getattr(humanizer, 'model', None) is None:
                    st.info("Synonym replacement is running in fallback mode (WordNet-only). For improved results, ensure internet access or a cached Hugging Face model.")
                
                # Transform using appropriate method based on whether we have formatted document
                if formatted_doc is not None:
                    # Use humanize_document to preserve formatting
                    output_formatted_doc = humanizer.humanize_document(
                        formatted_doc, 
                        use_passive=use_passive, 
                        use_synonyms=use_synonyms
                    )
                    transformed = get_plain_text(output_formatted_doc)
                else:
                    # For manual text input, first build a simple FormattedDocument then call humanize_document
                    input_formatted_doc = text_to_formatted_document(user_text)
                    output_formatted_doc = humanizer.humanize_document(
                        input_formatted_doc,
                        use_passive=use_passive,
                        use_synonyms=use_synonyms
                    )
                    transformed = get_plain_text(output_formatted_doc)

                # Output
                st.subheader("Transformed Text:")
                
                # Display options
                display_mode = st.radio("Display format:", ["Plain Text", "HTML Preview"], horizontal=True)
                
                if display_mode == "Plain Text":
                    st.write(transformed)
                else:
                    # Render as HTML
                    html_content = formatted_document_to_html(output_formatted_doc)
                    st.markdown(html_content, unsafe_allow_html=True)

                # Add download buttons
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="Download .txt",
                        data=transformed,
                        file_name="transformed_output.txt",
                        mime="text/plain"
                    )
                
                with col2:
                    # Generate DOCX for download
                    try:
                        docx_bytes = formatted_document_to_docx(output_formatted_doc)
                        st.download_button(
                            label="Download .docx (with formatting)",
                            data=docx_bytes,
                            file_name="transformed_output.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"Error generating .docx file: {str(e)}")

                # Output stats
                output_word_count = len(word_tokenize(transformed,language='english', preserve_line=True))
                doc_output = NLP_GLOBAL(transformed)
                output_sentence_count = len(list(doc_output.sents))

                st.markdown(
                    f"**Input Word Count**: {input_word_count} "
                    f"| **Sentence Count**: {input_sentence_count}  "
                    f"| **Output Word Count**: {output_word_count} "
                    f"| **Sentence Count**: {output_sentence_count}"
                )

    st.markdown("---")


if __name__ == "__main__":
    main()