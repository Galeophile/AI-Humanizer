from __future__ import annotations

import ssl
import re
import random
import warnings
from typing import List, Tuple, Dict, TYPE_CHECKING

import nltk
import spacy
from nltk.tokenize import word_tokenize
from nltk.corpus import wordnet
from sentence_transformers import SentenceTransformer, util

if TYPE_CHECKING:
    from .document_parser import FormattedDocument, FormattedParagraph, FormattedRun, FormattedListItem, FormattedTable

warnings.filterwarnings("ignore", category=FutureWarning)

NLP_GLOBAL = spacy.load("en_core_web_sm")

def download_nltk_resources():
    """
    Download required NLTK resources if not already installed.
    """
    try:
        _create_unverified_https_context = ssl._create_unverified_context
    except AttributeError:
        pass
    else:
        ssl._create_default_https_context = _create_unverified_https_context

    resources = ['punkt', 'averaged_perceptron_tagger', 'punkt_tab','wordnet','averaged_perceptron_tagger_eng']
    for resource in resources:
        try:
            nltk.download(resource, quiet=True)
        except Exception as e:
            print(f"Error downloading {resource}: {str(e)}")


# This class  contains methods to humanize academic text, such as improving readability or
# simplifying complex language.
class AcademicTextHumanizer:
    """
    Transforms text into a more formal (academic) style while preserving formatting.
    
    Supports both plain text (via humanize_text()) and formatted documents (via humanize_document()):
      - Expands contractions
      - Adds academic transitions
      - Optionally converts some sentences to passive voice
      - Optionally replaces words with synonyms for more formality
      
    For formatted documents, uses token-level alignment to preserve all formatting metadata
    including bold, italic, fonts, colors, and paragraph styles. The same transformation
    probabilities and logic apply to both plain text and formatted document modes.
    """

    def __init__(
        self,
        model_name='paraphrase-MiniLM-L6-v2',
        p_passive=0.2,
        p_synonym_replacement=0.3,
        p_academic_transition=0.3,
        seed=None,
        ensure_change=True
    ):
        if seed is not None:
            random.seed(seed)

        self.nlp = spacy.load("en_core_web_sm")
        # Try to load embedding model; gracefully degrade if unavailable (offline/no cache)
        try:
            self.model = SentenceTransformer(model_name)
            self.model_error = None
        except Exception as e:
            self.model = None
            self.model_error = e

        # Transformation probabilities
        self.p_passive = max(0.0, min(1.0, p_passive))
        self.p_synonym_replacement = max(0.0, min(1.0, p_synonym_replacement))
        self.p_academic_transition = max(0.0, min(1.0, p_academic_transition))
        self.ensure_change = ensure_change

        # Common academic transitions
        self.academic_transitions = [
            "Moreover,", "Additionally,", "Furthermore,", "Hence,", 
            "Therefore,", "Consequently,", "Nonetheless,", "Nevertheless,"
        ]

    def humanize_text(self, text, use_passive=False, use_synonyms=False):
        # Sanitize input to avoid transforming diagnostic logs or stack traces
        text = self.sanitize_input(text)
        doc = self.nlp(text)
        transformed_sentences = []
        inserted_transition = False

        for sent in doc.sents:
            sentence_str = sent.text.strip()
            original_sentence = sentence_str
            changed = False

            # 1. Expand contractions
            expanded = self.expand_contractions(sentence_str)
            if expanded != sentence_str:
                changed = True
                sentence_str = expanded

            # 2. Possibly add academic transitions (or ensure at least one early transition)
            should_transition = (random.random() < self.p_academic_transition) or (self.ensure_change and not changed and not inserted_transition)
            if should_transition:
                transitioned = self.add_academic_transitions(sentence_str)
                if transitioned != sentence_str:
                    inserted_transition = True
                    changed = True
                    sentence_str = transitioned

            # 3. Optionally convert to passive (if probability triggers or still no change)
            if use_passive and (random.random() < self.p_passive or (self.ensure_change and not changed)):
                passive = self.convert_to_passive(sentence_str)
                if passive != sentence_str:
                    changed = True
                    sentence_str = passive

            # 4. Optionally replace words with synonyms (if probability triggers or still no change)
            if use_synonyms and (random.random() < self.p_synonym_replacement or (self.ensure_change and not changed)):
                with_syn = self.replace_with_synonyms(sentence_str)
                if with_syn != sentence_str:
                    changed = True
                    sentence_str = with_syn
                elif self.ensure_change and not changed:
                    forced_syn, did = self.replace_with_synonyms_force_one(sentence_str)
                    if did:
                        changed = True
                        sentence_str = forced_syn

            # 5. If still no change, gently add a transition on first eligible sentence
            if not changed and not inserted_transition:
                transitioned = self.add_academic_transitions(sentence_str)
                if transitioned != sentence_str:
                    inserted_transition = True
                    changed = True
                    sentence_str = transitioned

            transformed_sentences.append(sentence_str)

        final_text = ' '.join(transformed_sentences)
        return final_text

    def sanitize_input(self, text: str) -> str:
        """
        Remove common diagnostic/stack trace lines (e.g., OSError, Traceback, HF URLs)
        that may accidentally appear in the input text so the transformation output
        is clean and user-centric.
        """
        error_patterns = [
            r"^\s*Traceback\b",
            r"^\s*File\s+['\"]?.+?['\"]?,\s*line\s*\d+",
            r"^\s*OSError\b",
            r"^\s*EnvironmentError\b",
            r"https?://huggingface\.co",
            r"https?://.*transformers/installation#offline-mode",
            r"^\s*We couldn't connect to",
            r"^\s*Could not (connect|find)"
        ]
        lines = text.splitlines()
        kept = [line for line in lines if not any(re.search(p, line) for p in error_patterns)]
        sanitized = "\n".join(kept).strip()
        return sanitized if sanitized else text

    def expand_contractions(self, sentence):
        contraction_map = {
            "n't": " not", "'re": " are", "'s": " is", "'ll": " will",
            "'ve": " have", "'d": " would", "'m": " am"
        }
        tokens = word_tokenize(sentence)
        expanded_tokens = []
        for token in tokens:
            lower_token = token.lower()
            replaced = False
            for contraction, expansion in contraction_map.items():
                if contraction in lower_token and lower_token.endswith(contraction):
                    new_token = lower_token.replace(contraction, expansion)
                    if token[0].isupper():
                        new_token = new_token.capitalize()
                    expanded_tokens.append(new_token)
                    replaced = True
                    break
            if not replaced:
                expanded_tokens.append(token)

        return ' '.join(expanded_tokens)

    def add_academic_transitions(self, sentence):
        transition = random.choice(self.academic_transitions)
        return f"{transition} {sentence}"

    def convert_to_passive(self, sentence):
        doc = self.nlp(sentence)
        subj_tokens = [t for t in doc if t.dep_ == 'nsubj' and t.head.dep_ == 'ROOT']
        dobj_tokens = [t for t in doc if t.dep_ == 'dobj']

        if subj_tokens and dobj_tokens:
            subject = subj_tokens[0]
            dobj = dobj_tokens[0]
            verb = subject.head
            if subject.i < verb.i < dobj.i:
                passive_str = f"{dobj.text} {verb.lemma_} by {subject.text}"
                original_str = ' '.join(token.text for token in doc)
                chunk = f"{subject.text} {verb.text} {dobj.text}"
                if chunk in original_str:
                    sentence = original_str.replace(chunk, passive_str)
        return sentence

    def replace_with_synonyms(self, sentence):
        tokens = word_tokenize(sentence)
        pos_tags = nltk.pos_tag(tokens)

        new_tokens = []
        for (word, pos) in pos_tags:
            if pos.startswith(('J', 'N', 'V', 'R')) and wordnet.synsets(word):
                if random.random() < 0.5:
                    synonyms = self._get_synonyms(word, pos)
                    if synonyms:
                        best_synonym = self._select_closest_synonym(word, synonyms)
                        new_tokens.append(best_synonym if best_synonym else word)
                    else:
                        new_tokens.append(word)
                else:
                    new_tokens.append(word)
            else:
                new_tokens.append(word)

        return ' '.join(new_tokens)

    def replace_with_synonyms_force_one(self, sentence):
        """
        Attempt to replace at least one token with a reasonable synonym.
        Returns (new_sentence, did_replace: bool).
        """
        tokens = word_tokenize(sentence)
        pos_tags = nltk.pos_tag(tokens)

        new_tokens = tokens[:]
        for idx, (word, pos) in enumerate(pos_tags):
            # Map POS to WordNet POS
            wn_pos = None
            if pos.startswith('J'):
                wn_pos = wordnet.ADJ
            elif pos.startswith('N'):
                wn_pos = wordnet.NOUN
            elif pos.startswith('R'):
                wn_pos = wordnet.ADV
            elif pos.startswith('V'):
                wn_pos = wordnet.VERB

            if wn_pos is None:
                continue

            synonyms = self._get_synonyms(word, pos)
            if not synonyms:
                continue

            # Prefer embedding-based selection; else fallback heuristic
            if self.model is not None:
                candidate = self._select_closest_synonym(word, synonyms)
            else:
                candidate = self._select_closest_synonym(word, synonyms)  # falls back inside method

            if candidate and candidate.lower() != word.lower():
                new_tokens[idx] = candidate
                return ' '.join(new_tokens), True

        return ' '.join(new_tokens), False

    def _get_synonyms(self, word, pos):
        wn_pos = None
        if pos.startswith('J'):
            wn_pos = wordnet.ADJ
        elif pos.startswith('N'):
            wn_pos = wordnet.NOUN
        elif pos.startswith('R'):
            wn_pos = wordnet.ADV
        elif pos.startswith('V'):
            wn_pos = wordnet.VERB

        synonyms = set()
        for syn in wordnet.synsets(word, pos=wn_pos):
            for lemma in syn.lemmas():
                lemma_name = lemma.name().replace('_', ' ')
                if lemma_name.lower() != word.lower():
                    synonyms.add(lemma_name)
        return list(synonyms)

    def _select_closest_synonym(self, original_word, synonyms):
        if not synonyms:
            return None

        # If embedding model is available, select by cosine similarity
        if self.model is not None:
            original_emb = self.model.encode(original_word, convert_to_tensor=True)
            synonym_embs = self.model.encode(synonyms, convert_to_tensor=True)
            cos_scores = util.cos_sim(original_emb, synonym_embs)[0]
            max_score_index = cos_scores.argmax().item()
            max_score = cos_scores[max_score_index].item()
            if max_score >= 0.5:
                return synonyms[max_score_index]
            return None

        # Fallback: heuristic selection without embeddings
        # Prefer synonyms closest in length to the original word
        try:
            closest = min(
                (s for s in synonyms if s.lower() != original_word.lower()),
                key=lambda s: abs(len(s) - len(original_word))
            )
            # If the closest is wildly different, skip replacement
            if abs(len(closest) - len(original_word)) <= 3:
                return closest
            return None
        except ValueError:
            return None

    def _tokenize_runs_with_formatting(self, runs: List[FormattedRun]) -> Tuple[List[Tuple[str, FormattedRun, str]], str]:
        """
        Tokenize runs while preserving formatting metadata and original whitespace.
        
        Args:
            runs: List of FormattedRun objects
            
        Returns:
            Tuple of:
            - List of tuples: (token, source_run, following_whitespace)
            - Full concatenated text for reference
        """
        import re
        
        # Build char-level concatenation with run boundary tracking
        full_text = ""
        run_boundaries = []  # (start_char, end_char, run)
        
        for run in runs:
            start_pos = len(full_text)
            full_text += run.text
            end_pos = len(full_text)
            run_boundaries.append((start_pos, end_pos, run))
        
        if not full_text.strip():
            return [], full_text
        
        # Tokenize the full text while preserving whitespace info
        token_pattern = r'\S+|\s+'
        matches = list(re.finditer(token_pattern, full_text))
        
        token_format_map = []
        
        for i, match in enumerate(matches):
            token_text = match.group()
            start_char = match.start()
            end_char = match.end()
            
            # Skip pure whitespace tokens for processing, but track them
            if not token_text.strip():
                continue
            
            # Find which run(s) this token belongs to
            dominant_run = None
            max_overlap = 0
            
            for run_start, run_end, run in run_boundaries:
                overlap_start = max(start_char, run_start)
                overlap_end = min(end_char, run_end)
                overlap = max(0, overlap_end - overlap_start)
                
                if overlap > max_overlap:
                    max_overlap = overlap
                    dominant_run = run
            
            # Determine following whitespace
            following_whitespace = ""
            if i + 1 < len(matches):
                next_match = matches[i + 1]
                next_text = next_match.group()
                if not next_text.strip():  # Next token is whitespace
                    following_whitespace = next_text
                    # Look ahead to see if there's more whitespace
                    j = i + 2
                    while j < len(matches) and not matches[j].group().strip():
                        following_whitespace += matches[j].group()
                        j += 1
            
            if dominant_run:
                token_format_map.append((token_text, dominant_run, following_whitespace))
        
        return token_format_map, full_text

    def _reconstruct_runs_from_tokens(self, transformed_text: str, original_token_map: List[Tuple[str, FormattedRun, str]]) -> List[FormattedRun]:
        """
        Reconstruct formatted runs from transformed text using preserved whitespace.
        
        Args:
            transformed_text: The transformed text string
            original_token_map: Original token-to-formatting mapping with whitespace info
            
        Returns:
            List of FormattedRun objects with formatting preserved
        """
        from .document_parser import FormattedRun
        
        if not transformed_text.strip() or not original_token_map:
            return [FormattedRun(text=transformed_text)]
        
        transformed_tokens = word_tokenize(transformed_text)
        new_runs = []
        original_consumed = 0
        
        for i, token in enumerate(transformed_tokens):
            found_match = False
            
            # Try to find exact match in remaining original tokens
            for j in range(original_consumed, len(original_token_map)):
                orig_token, orig_run, orig_whitespace = original_token_map[j]
                
                if token.lower() == orig_token.lower():
                    # Exact match - use original formatting and whitespace
                    token_with_whitespace = token + orig_whitespace
                    
                    if new_runs and self._runs_have_same_formatting(new_runs[-1], orig_run):
                        # Merge with previous run if same formatting
                        new_runs[-1] = FormattedRun(
                            text=new_runs[-1].text + token_with_whitespace,
                            bold=orig_run.bold,
                            italic=orig_run.italic,
                            underline=orig_run.underline,
                            underline_style=orig_run.underline_style,
                            font_name=orig_run.font_name,
                            font_size=orig_run.font_size,
                            color=orig_run.color,
                            highlight=orig_run.highlight
                        )
                    else:
                        # Create new run
                        new_runs.append(FormattedRun(
                            text=token_with_whitespace,
                            bold=orig_run.bold,
                            italic=orig_run.italic,
                            underline=orig_run.underline,
                            underline_style=orig_run.underline_style,
                            font_name=orig_run.font_name,
                            font_size=orig_run.font_size,
                            color=orig_run.color,
                            highlight=orig_run.highlight
                        ))
                    original_consumed = j + 1
                    found_match = True
                    break
            
            # If no exact match, try fuzzy matching with WordNet lemmatization
            if not found_match:
                for j in range(original_consumed, len(original_token_map)):
                    orig_token, orig_run, orig_whitespace = original_token_map[j]
                    
                    # Check if tokens are related through lemmatization
                    if self._tokens_are_related(token, orig_token):
                        token_with_whitespace = token + orig_whitespace
                        
                        if new_runs and self._runs_have_same_formatting(new_runs[-1], orig_run):
                            new_runs[-1] = FormattedRun(
                                text=new_runs[-1].text + token_with_whitespace,
                                bold=orig_run.bold,
                                italic=orig_run.italic,
                                underline=orig_run.underline,
                                underline_style=orig_run.underline_style,
                                font_name=orig_run.font_name,
                                font_size=orig_run.font_size,
                                color=orig_run.color,
                                highlight=orig_run.highlight
                            )
                        else:
                            new_runs.append(FormattedRun(
                                text=token_with_whitespace,
                                bold=orig_run.bold,
                                italic=orig_run.italic,
                                underline=orig_run.underline,
                                underline_style=orig_run.underline_style,
                                font_name=orig_run.font_name,
                                font_size=orig_run.font_size,
                                color=orig_run.color,
                                highlight=orig_run.highlight
                            ))
                        original_consumed = j + 1
                        found_match = True
                        break
            
            # If still no match, create new run with default formatting
            if not found_match:
                # Use space as default whitespace except for last token
                default_whitespace = " " if i < len(transformed_tokens) - 1 else ""
                token_with_whitespace = token + default_whitespace
                
                if new_runs and self._is_default_formatting(new_runs[-1]):
                    # Merge with previous default-formatted run
                    new_runs[-1] = FormattedRun(text=new_runs[-1].text + token_with_whitespace)
                else:
                    # Create new default run
                    new_runs.append(FormattedRun(text=token_with_whitespace))
        
        return self._merge_consecutive_runs_whitespace_aware(new_runs)

    def _runs_have_same_formatting(self, run1: FormattedRun, run2: FormattedRun) -> bool:
        """Check if two runs have identical formatting."""
        return (run1.bold == run2.bold and
                run1.italic == run2.italic and
                run1.underline == run2.underline and
                run1.underline_style == run2.underline_style and
                run1.font_name == run2.font_name and
                run1.font_size == run2.font_size and
                run1.color == run2.color and
                run1.highlight == run2.highlight)

    def _tokens_are_related(self, token1: str, token2: str) -> bool:
        """Check if two tokens are related through lemmatization."""
        if token1.lower() == token2.lower():
            return True
        
        # Try WordNet lemmatization
        try:
            synsets1 = wordnet.synsets(token1.lower())
            synsets2 = wordnet.synsets(token2.lower())
            
            for syn1 in synsets1:
                for syn2 in synsets2:
                    if syn1 == syn2:
                        return True
                        
            # Check if one is a lemma of the other
            for syn1 in synsets1:
                for lemma in syn1.lemmas():
                    if lemma.name().lower() == token2.lower():
                        return True
                        
            for syn2 in synsets2:
                for lemma in syn2.lemmas():
                    if lemma.name().lower() == token1.lower():
                        return True
        except:
            pass
        
        return False

    def _is_default_formatting(self, run: FormattedRun) -> bool:
        """Check if a run has default formatting (no bold, italic, etc.)."""
        return (not run.bold and 
                not run.italic and 
                not run.underline and 
                run.underline_style is None and 
                run.font_name is None and 
                run.font_size is None and 
                run.color is None and 
                run.highlight is None)

    def _merge_consecutive_runs_whitespace_aware(self, runs: List[FormattedRun]) -> List[FormattedRun]:
        """Merge consecutive runs with identical formatting without adding extra spaces."""
        if not runs:
            return []
        
        merged = []
        current_run = runs[0]
        
        for next_run in runs[1:]:
            if self._runs_have_same_formatting(current_run, next_run):
                # Merge text without adding spaces (whitespace already included)
                current_run = FormattedRun(
                    text=current_run.text + next_run.text,
                    bold=current_run.bold,
                    italic=current_run.italic,
                    underline=current_run.underline,
                    underline_style=current_run.underline_style,
                    font_name=current_run.font_name,
                    font_size=current_run.font_size,
                    color=current_run.color,
                    highlight=current_run.highlight
                )
            else:
                merged.append(current_run)
                current_run = next_run
        
        merged.append(current_run)
        return merged

    def _merge_consecutive_runs(self, runs: List[FormattedRun]) -> List[FormattedRun]:
        """Merge consecutive runs with identical formatting (legacy method)."""
        if not runs:
            return []
        
        merged = []
        current_run = runs[0]
        
        for next_run in runs[1:]:
            if self._runs_have_same_formatting(current_run, next_run):
                # Merge text without unconditionally adding spaces
                separator = " " if not current_run.text.endswith(" ") and not next_run.text.startswith(" ") else ""
                current_run = FormattedRun(
                    text=current_run.text + separator + next_run.text,
                    bold=current_run.bold,
                    italic=current_run.italic,
                    underline=current_run.underline,
                    underline_style=current_run.underline_style,
                    font_name=current_run.font_name,
                    font_size=current_run.font_size,
                    color=current_run.color,
                    highlight=current_run.highlight
                )
            else:
                merged.append(current_run)
                current_run = next_run
        
        merged.append(current_run)
        return merged

    def _extract_plain_text_from_paragraph(self, para: FormattedParagraph) -> str:
        """
        Extract plain text from a FormattedParagraph.
        
        Args:
            para: The FormattedParagraph object
            
        Returns:
            Plain text string
        """
        return ''.join(run.text for run in para.runs)

    def _copy_paragraph_formatting(self, source_para: FormattedParagraph, new_runs: List[FormattedRun]) -> FormattedParagraph:
        """
        Copy paragraph-level formatting attributes to a new paragraph with different runs.
        
        Args:
            source_para: Source paragraph with original formatting
            new_runs: New runs to use in the paragraph
            
        Returns:
            New FormattedParagraph with copied formatting
        """
        from .document_parser import FormattedParagraph
        
        return FormattedParagraph(
            runs=new_runs,
            style=source_para.style,
            alignment=source_para.alignment,
            space_before=source_para.space_before,
            space_after=source_para.space_after,
            line_spacing=source_para.line_spacing,
            left_indent=source_para.left_indent,
            right_indent=source_para.right_indent,
            first_line_indent=source_para.first_line_indent
        )

    def humanize_paragraph(self, para: FormattedParagraph, use_passive: bool = False, use_synonyms: bool = False) -> FormattedParagraph:
        """
        Transform a single formatted paragraph while preserving formatting.
        
        Args:
            para: The FormattedParagraph to transform
            use_passive: Whether to apply passive voice transformations
            use_synonyms: Whether to apply synonym replacements
            
        Returns:
            Transformed FormattedParagraph with formatting preserved
        """
        # Handle empty paragraphs
        if not para.runs or not any(run.text.strip() for run in para.runs):
            return para
        
        try:
            # Extract plain text for transformation
            plain_text = self._extract_plain_text_from_paragraph(para)
            
            # Build token-to-formatting map with whitespace preservation
            token_format_map, original_full_text = self._tokenize_runs_with_formatting(para.runs)
            
            # Apply existing transformation logic
            transformed_text = self.humanize_text(plain_text, use_passive=use_passive, use_synonyms=use_synonyms)
            
            # Reconstruct formatted runs
            new_runs = self._reconstruct_runs_from_tokens(transformed_text, token_format_map)
            
            # Create new paragraph with transformed runs but original formatting
            return self._copy_paragraph_formatting(para, new_runs)
            
        except Exception:
            # If transformation fails, return original paragraph
            return para

    def humanize_document(self, formatted_doc: FormattedDocument, use_passive: bool = False, use_synonyms: bool = False) -> FormattedDocument:
        """
        Transform a formatted document while preserving all formatting.
        
        Args:
            formatted_doc: The FormattedDocument to transform
            use_passive: Whether to apply passive voice transformations
            use_synonyms: Whether to apply synonym replacements
            
        Returns:
            Transformed FormattedDocument with formatting preserved
        """
        from .document_parser import FormattedDocument, FormattedListItem, FormattedTable
        
        try:
            # Process regular paragraphs
            transformed_paragraphs = []
            for para in formatted_doc.paragraphs:
                transformed_para = self.humanize_paragraph(para, use_passive=use_passive, use_synonyms=use_synonyms)
                transformed_paragraphs.append(transformed_para)
            
            # Process list items
            transformed_list_items = []
            for list_item in formatted_doc.list_items:
                transformed_para = self.humanize_paragraph(list_item.paragraph, use_passive=use_passive, use_synonyms=use_synonyms)
                transformed_list_item = FormattedListItem(
                    paragraph=transformed_para,
                    level=list_item.level,
                    list_type=list_item.list_type,
                    number_format=list_item.number_format
                )
                transformed_list_items.append(transformed_list_item)
            
            # Process tables
            transformed_tables = []
            for table in formatted_doc.tables:
                transformed_rows = []
                for row in table.rows:
                    transformed_row = []
                    for cell in row:
                        transformed_cell = []
                        for cell_para in cell:
                            transformed_cell_para = self.humanize_paragraph(cell_para, use_passive=use_passive, use_synonyms=use_synonyms)
                            transformed_cell.append(transformed_cell_para)
                        transformed_row.append(transformed_cell)
                    transformed_rows.append(transformed_row)
                
                transformed_table = FormattedTable(
                    rows=transformed_rows,
                    style=table.style
                )
                transformed_tables.append(transformed_table)
            
            # Create and return new document with transformed content
            return FormattedDocument(
                paragraphs=transformed_paragraphs,
                tables=transformed_tables,
                list_items=transformed_list_items,
                styles=formatted_doc.styles
            )
            
        except Exception:
            # If transformation fails, return original document
            return formatted_doc


# Module-level convenience functions
def humanize_text(text: str, use_passive: bool = False, use_synonyms: bool = False, **kwargs) -> str:
    """
    Module-level convenience function to humanize plain text.
    
    Args:
        text: Plain text to transform
        use_passive: Whether to apply passive voice transformations
        use_synonyms: Whether to apply synonym replacements
        **kwargs: Additional arguments passed to AcademicTextHumanizer constructor
        
    Returns:
        Transformed text string
    """
    humanizer = AcademicTextHumanizer(**kwargs)
    return humanizer.humanize_text(text, use_passive=use_passive, use_synonyms=use_synonyms)


def humanize_paragraph(para: "FormattedParagraph", use_passive: bool = False, use_synonyms: bool = False, **kwargs) -> "FormattedParagraph":
    """
    Module-level convenience function to humanize a formatted paragraph.
    
    Args:
        para: FormattedParagraph to transform
        use_passive: Whether to apply passive voice transformations
        use_synonyms: Whether to apply synonym replacements
        **kwargs: Additional arguments passed to AcademicTextHumanizer constructor
        
    Returns:
        Transformed FormattedParagraph with formatting preserved
    """
    humanizer = AcademicTextHumanizer(**kwargs)
    return humanizer.humanize_paragraph(para, use_passive=use_passive, use_synonyms=use_synonyms)


def humanize_document(doc: "FormattedDocument", use_passive: bool = False, use_synonyms: bool = False, **kwargs) -> "FormattedDocument":
    """
    Module-level convenience function to humanize a formatted document.
    
    Args:
        doc: FormattedDocument to transform
        use_passive: Whether to apply passive voice transformations
        use_synonyms: Whether to apply synonym replacements
        **kwargs: Additional arguments passed to AcademicTextHumanizer constructor
        
    Returns:
        Transformed FormattedDocument with formatting preserved
    """
    humanizer = AcademicTextHumanizer(**kwargs)
    return humanizer.humanize_document(doc, use_passive=use_passive, use_synonyms=use_synonyms)