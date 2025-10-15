import streamlit as st
from transformer.app import AcademicTextHumanizer, NLP_GLOBAL, download_nltk_resources
from nltk.tokenize import word_tokenize



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
    uploaded_file = st.file_uploader("Or upload a .txt file:", type=["txt"])
    if uploaded_file is not None:
        file_text = uploaded_file.read().decode("utf-8", errors="ignore")
        user_text = file_text

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
                transformed = humanizer.humanize_text(
                    user_text,
                    use_passive=use_passive,
                    use_synonyms=use_synonyms
                )

                # Output
                st.subheader("Transformed Text:")
                st.write(transformed)

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