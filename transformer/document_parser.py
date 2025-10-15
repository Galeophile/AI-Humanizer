"""
Document parser for .docx files with comprehensive formatting preservation.

This module provides data classes and functions for parsing Microsoft Word documents
while preserving formatting information including fonts, colors, styles, and structure.
"""

from dataclasses import dataclass
from typing import List, Optional, Union, Dict, Any
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.dml import MSO_COLOR_TYPE
import io


@dataclass
class FormattedRun:
    """
    Represents a run of text with consistent formatting.
    
    Attributes:
        text: The text content of the run
        bold: Whether the text is bold
        italic: Whether the text is italic
        underline: Whether the text is underlined
        underline_style: The specific underline style if any
        font_name: The font family name
        font_size: The font size in points
        color: The text color in hex format (e.g., '#FF0000')
        highlight: The highlight color in hex format
    """
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    underline_style: Optional[str] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    color: Optional[str] = None
    highlight: Optional[str] = None


@dataclass
class FormattedParagraph:
    """
    Represents a paragraph with its formatting and runs.
    
    Attributes:
        runs: List of formatted runs in the paragraph
        style: The paragraph style name
        alignment: Text alignment as string
        space_before: Space before paragraph in points
        space_after: Space after paragraph in points
        line_spacing: Line spacing value
        left_indent: Left indent in points
        right_indent: Right indent in points
        first_line_indent: First line indent in points
    """
    runs: List[FormattedRun]
    style: Optional[str] = None
    alignment: Optional[str] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    line_spacing: Optional[float] = None
    left_indent: Optional[float] = None
    right_indent: Optional[float] = None
    first_line_indent: Optional[float] = None


@dataclass
class FormattedListItem:
    """
    Represents a list item with its level and type.
    
    Attributes:
        paragraph: The formatted paragraph content
        level: The nesting level (0-based)
        list_type: The type of list ('bullet', 'number', etc.)
        number_format: The numbering format if applicable
    """
    paragraph: FormattedParagraph
    level: int
    list_type: str
    number_format: Optional[str] = None


@dataclass
class FormattedTable:
    """
    Represents a table with its cell contents.
    
    Attributes:
        rows: List of rows, each containing a list of cell paragraphs
        style: The table style name
    """
    rows: List[List[List[FormattedParagraph]]]
    style: Optional[str] = None


@dataclass
class FormattedDocument:
    """
    Represents a complete document with all its content.
    
    Attributes:
        paragraphs: List of formatted paragraphs
        tables: List of formatted tables
        list_items: List of formatted list items
        styles: Document styles information
    """
    paragraphs: List[FormattedParagraph]
    tables: List[FormattedTable]
    list_items: List[FormattedListItem]
    styles: Optional[Dict[str, Any]] = None


def rgb_to_hex(rgb) -> Optional[str]:
    """
    Convert RGB color to hex format.
    
    Args:
        rgb: RGB color object or tuple
        
    Returns:
        Hex color string (e.g., '#FF0000') or None if conversion fails
    """
    try:
        if hasattr(rgb, 'rgb'):
            color = rgb.rgb
            if color is not None:
                return f"#{color:06X}"
        elif isinstance(rgb, tuple) and len(rgb) == 3:
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        elif isinstance(rgb, RGBColor):
            return f"#{rgb.rgb:06X}"
    except (AttributeError, TypeError, ValueError):
        pass
    return None


def alignment_to_string(alignment) -> Optional[str]:
    """
    Convert alignment enum to string representation.
    
    Args:
        alignment: WD_ALIGN_PARAGRAPH enum value
        
    Returns:
        String representation of alignment or None
    """
    try:
        if alignment is None:
            return None
        
        alignment_map = {
            WD_ALIGN_PARAGRAPH.LEFT: "left",
            WD_ALIGN_PARAGRAPH.CENTER: "center",
            WD_ALIGN_PARAGRAPH.RIGHT: "right",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
            WD_ALIGN_PARAGRAPH.DISTRIBUTE: "distribute"
        }
        return alignment_map.get(alignment)
    except (AttributeError, TypeError):
        return None


def extract_run_formatting(run) -> FormattedRun:
    """
    Extract formatting information from a document run.
    
    Args:
        run: The document run object
        
    Returns:
        FormattedRun object with extracted formatting
    """
    try:
        text = run.text or ""
        bold = run.bold or False
        italic = run.italic or False
        
        # Handle underline with style detection
        underline = False
        underline_style = None
        if hasattr(run, 'underline') and run.underline is not None:
            if run.underline is True or (hasattr(run.underline, 'val') and run.underline.val):
                underline = True
                if hasattr(run.underline, 'val') and run.underline.val != True:
                    underline_style = str(run.underline.val)
        
        # Extract font information
        font_name = None
        font_size = None
        if hasattr(run, 'font'):
            if hasattr(run.font, 'name') and run.font.name:
                font_name = run.font.name
            if hasattr(run.font, 'size') and run.font.size:
                # Convert from Emu to points (1 point = 12700 EMU)
                font_size = float(run.font.size.pt) if hasattr(run.font.size, 'pt') else None
        
        # Extract color information
        color = None
        if hasattr(run, 'font') and hasattr(run.font, 'color'):
            font_color = run.font.color
            if hasattr(font_color, 'rgb') and font_color.rgb is not None:
                color = rgb_to_hex(font_color.rgb)
            elif hasattr(font_color, 'theme_color') and font_color.theme_color is not None:
                # Handle theme colors - would need document theme for full resolution
                pass
        
        # Extract highlight information
        highlight = None
        if hasattr(run, 'font') and hasattr(run.font, 'highlight_color'):
            highlight_color = run.font.highlight_color
            if highlight_color is not None and hasattr(highlight_color, 'rgb'):
                highlight = rgb_to_hex(highlight_color.rgb)
        
        return FormattedRun(
            text=text,
            bold=bold,
            italic=italic,
            underline=underline,
            underline_style=underline_style,
            font_name=font_name,
            font_size=font_size,
            color=color,
            highlight=highlight
        )
    except Exception:
        # Return basic run with just text if formatting extraction fails
        return FormattedRun(text=getattr(run, 'text', ''))


def extract_paragraph_formatting(paragraph) -> FormattedParagraph:
    """
    Extract formatting information from a document paragraph.
    
    Args:
        paragraph: The document paragraph object
        
    Returns:
        FormattedParagraph object with extracted formatting
    """
    try:
        # Extract runs
        runs = [extract_run_formatting(run) for run in paragraph.runs]
        
        # Extract paragraph style
        style = None
        if hasattr(paragraph, 'style') and paragraph.style:
            style = paragraph.style.name
        
        # Extract alignment
        alignment = alignment_to_string(paragraph.alignment)
        
        # Extract spacing and indent information
        space_before = None
        space_after = None
        line_spacing = None
        left_indent = None
        right_indent = None
        first_line_indent = None
        
        if hasattr(paragraph, 'paragraph_format'):
            pf = paragraph.paragraph_format
            
            if hasattr(pf, 'space_before') and pf.space_before:
                space_before = float(pf.space_before.pt) if hasattr(pf.space_before, 'pt') else None
            
            if hasattr(pf, 'space_after') and pf.space_after:
                space_after = float(pf.space_after.pt) if hasattr(pf.space_after, 'pt') else None
            
            if hasattr(pf, 'line_spacing') and pf.line_spacing:
                line_spacing = float(pf.line_spacing)
            
            if hasattr(pf, 'left_indent') and pf.left_indent:
                left_indent = float(pf.left_indent.pt) if hasattr(pf.left_indent, 'pt') else None
            
            if hasattr(pf, 'right_indent') and pf.right_indent:
                right_indent = float(pf.right_indent.pt) if hasattr(pf.right_indent, 'pt') else None
            
            if hasattr(pf, 'first_line_indent') and pf.first_line_indent:
                first_line_indent = float(pf.first_line_indent.pt) if hasattr(pf.first_line_indent, 'pt') else None
        
        return FormattedParagraph(
            runs=runs,
            style=style,
            alignment=alignment,
            space_before=space_before,
            space_after=space_after,
            line_spacing=line_spacing,
            left_indent=left_indent,
            right_indent=right_indent,
            first_line_indent=first_line_indent
        )
    except Exception:
        # Return basic paragraph with just runs if formatting extraction fails
        runs = []
        try:
            runs = [extract_run_formatting(run) for run in paragraph.runs]
        except:
            pass
        return FormattedParagraph(runs=runs)


def detect_list_item(paragraph) -> Optional[tuple]:
    """
    Detect if a paragraph is a list item and extract list information.
    
    Args:
        paragraph: The document paragraph object
        
    Returns:
        Tuple of (level, list_type, number_format) or None if not a list item
    """
    try:
        # Check for numbering properties using xpath
        num_pr = paragraph._element.xpath('.//w:numPr')
        if not num_pr:
            return None
        
        level = 0
        list_type = "bullet"
        number_format = None
        
        # Extract numbering level
        level_elements = paragraph._element.xpath('.//w:numPr/w:ilvl')
        if level_elements:
            try:
                level = int(level_elements[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))
            except (ValueError, AttributeError):
                level = 0
        
        # Extract numbering ID to determine list type
        num_id_elements = paragraph._element.xpath('.//w:numPr/w:numId')
        if num_id_elements and hasattr(paragraph, '_parent') and hasattr(paragraph._parent, 'part'):
            try:
                num_id = num_id_elements[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if num_id and hasattr(paragraph._parent.part, 'numbering_definitions'):
                    # Try to determine list type from numbering definitions
                    numbering_defs = paragraph._parent.part.numbering_definitions
                    if hasattr(numbering_defs, 'num_having_numId'):
                        num_def = numbering_defs.num_having_numId(int(num_id))
                        if num_def and hasattr(num_def, 'abstractNum'):
                            abstract_num = num_def.abstractNum
                            if hasattr(abstract_num, 'lvl_lst'):
                                lvl_list = abstract_num.lvl_lst
                                if len(lvl_list) > level:
                                    lvl_def = lvl_list[level]
                                    if hasattr(lvl_def, 'numFmt') and lvl_def.numFmt:
                                        fmt_val = lvl_def.numFmt.val
                                        if fmt_val == 'bullet':
                                            list_type = "bullet"
                                        elif fmt_val in ['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter']:
                                            list_type = "number"
                                            number_format = fmt_val
                                        else:
                                            list_type = "number"
                                            number_format = fmt_val
            except (ValueError, AttributeError, IndexError):
                pass
        
        return (level, list_type, number_format)
    except Exception:
        return None


def extract_table_formatting(table) -> FormattedTable:
    """
    Extract formatting information from a document table.
    
    Args:
        table: The document table object
        
    Returns:
        FormattedTable object with extracted formatting
    """
    try:
        rows = []
        for table_row in table.rows:
            row_cells = []
            for cell in table_row.cells:
                cell_paragraphs = []
                for paragraph in cell.paragraphs:
                    cell_paragraphs.append(extract_paragraph_formatting(paragraph))
                row_cells.append(cell_paragraphs)
            rows.append(row_cells)
        
        # Extract table style
        style = None
        if hasattr(table, 'style') and table.style:
            style = table.style.name
        
        return FormattedTable(rows=rows, style=style)
    except Exception:
        return FormattedTable(rows=[])


def parse_docx_file(file_or_bytes: Union[str, bytes, io.BytesIO]) -> FormattedDocument:
    """
    Parse a .docx file and build a FormattedDocument.
    
    Args:
        file_or_bytes: File path, bytes, or BytesIO object containing the .docx file
        
    Returns:
        FormattedDocument object containing all document content with formatting
        
    Raises:
        Exception: If document parsing fails
    """
    try:
        # Load the document
        if isinstance(file_or_bytes, (str, bytes)):
            doc = Document(file_or_bytes)
        else:
            doc = Document(file_or_bytes)
        
        paragraphs = []
        tables = []
        list_items = []
        
        # Process document elements in order
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                # Find corresponding paragraph object
                for para in doc.paragraphs:
                    if para._element == element:
                        formatted_para = extract_paragraph_formatting(para)
                        
                        # Check if it's a list item
                        list_info = detect_list_item(para)
                        if list_info:
                            level, list_type, number_format = list_info
                            list_item = FormattedListItem(
                                paragraph=formatted_para,
                                level=level,
                                list_type=list_type,
                                number_format=number_format
                            )
                            list_items.append(list_item)
                        else:
                            paragraphs.append(formatted_para)
                        break
            
            elif element.tag.endswith('tbl'):  # Table
                # Find corresponding table object
                for table in doc.tables:
                    if table._element == element:
                        formatted_table = extract_table_formatting(table)
                        tables.append(formatted_table)
                        break
        
        # Extract document styles information
        styles = {}
        if hasattr(doc, 'styles'):
            try:
                for style in doc.styles:
                    if hasattr(style, 'name') and style.name:
                        styles[style.name] = {
                            'type': getattr(style, 'type', None),
                            'builtin': getattr(style, 'builtin', None)
                        }
            except Exception:
                pass
        
        return FormattedDocument(
            paragraphs=paragraphs,
            tables=tables,
            list_items=list_items,
            styles=styles if styles else None
        )
    
    except Exception as e:
        raise Exception(f"Failed to parse document: {str(e)}")


def get_plain_text(formatted_doc: FormattedDocument) -> str:
    """
    Extract plain text from a FormattedDocument.
    
    Args:
        formatted_doc: The FormattedDocument object
        
    Returns:
        Plain text string with paragraph breaks
    """
    text_parts = []
    
    # Add paragraph text
    for paragraph in formatted_doc.paragraphs:
        paragraph_text = ''.join(run.text for run in paragraph.runs)
        if paragraph_text.strip():
            text_parts.append(paragraph_text)
    
    # Add list item text
    for list_item in formatted_doc.list_items:
        paragraph_text = ''.join(run.text for run in list_item.paragraph.runs)
        if paragraph_text.strip():
            text_parts.append(paragraph_text)
    
    # Add table text
    for table in formatted_doc.tables:
        for row in table.rows:
            for cell in row:
                for paragraph in cell:
                    paragraph_text = ''.join(run.text for run in paragraph.runs)
                    if paragraph_text.strip():
                        text_parts.append(paragraph_text)
    
    return '\n'.join(text_parts)


def merge_consecutive_runs(runs: List[FormattedRun]) -> List[FormattedRun]:
    """
    Merge consecutive runs with identical formatting.
    
    Args:
        runs: List of FormattedRun objects
        
    Returns:
        List of merged FormattedRun objects
    """
    if not runs:
        return []
    
    merged = []
    current_run = runs[0]
    
    for next_run in runs[1:]:
        # Check if formatting is identical (excluding text)
        if (current_run.bold == next_run.bold and
            current_run.italic == next_run.italic and
            current_run.underline == next_run.underline and
            current_run.underline_style == next_run.underline_style and
            current_run.font_name == next_run.font_name and
            current_run.font_size == next_run.font_size and
            current_run.color == next_run.color and
            current_run.highlight == next_run.highlight):
            # Merge text
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
            # Different formatting, add current and start new
            merged.append(current_run)
            current_run = next_run
    
    # Add the last run
    merged.append(current_run)
    return merged