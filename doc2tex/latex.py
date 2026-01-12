# Handles converting docx to latex
# This part is a bit tricky because word and latex are very different

import os
import re
from typing import List, Dict, Optional, Tuple
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from .options import ConversionOptions, LineSpacing
from .utils import (
    escape_latex, logger, ensure_directory, 
    optimize_image, sanitize_filename, get_temp_dir
)
from .errors import ConversionError, ImageProcessingError


class LatexGenerator:
    # This class does the heavy lifting for latex generation
    
    def __init__(self, options: ConversionOptions):
        self.options = options
        self.bibliography_entries = []
        self.image_counter = 0
        self.footnote_counter = 0
        self.temp_dir = None
        
    def convert(self, input_path: str, output_path: str) -> str:
        # Main entry point for converting one file
        try:
            logger.info(f"Converting DOCX to LaTeX: {input_path}")
            
            # Use the python-docx library to read the file
            doc = Document(input_path)
            
            # Create a temp folder for images if we need to extract them
            if self.options.preserve_images:
                self.temp_dir = get_temp_dir()
            
            # Start gathering all the latex code
            latex_content = self._generate_latex(doc, output_path)
            
            # Write everything to the output file
            with open(output_path, 'w', encoding=self.options.output_encoding) as f:
                f.write(latex_content)
            
            # If user wanted bibliography, we handle that here
            if self.options.extract_bibliography and self.bibliography_entries:
                self._generate_bibliography(output_path)
            
            return output_path
            
        except Exception as e:
            # Just catch anything that goes wrong and wrap it in our error
            raise ConversionError(f"Something went wrong during conversion: {e}")
    
    def _generate_latex(self, doc: Document, output_path: str) -> str:
        # Puts together the whole latex document structure
        sections = []
        
        # Add the preamble (documentclass, packages, etc.)
        if self.options.include_preamble and self.options.standalone_document:
            sections.append(self._generate_preamble())
        
        # Start the actual document
        if self.options.standalone_document:
            sections.append("\\begin{document}\n")
        
        # Loop through paragraphs and tables
        body_content = self._process_document_body(doc, output_path)
        sections.append(body_content)
        
        # End it
        if self.options.standalone_document:
            sections.append("\n\\end{document}")
        
        return '\n'.join(sections)
    
    def _generate_preamble(self) -> str:
        # Setup the latex settings and packages
        lines = []
        
        # Document class (like article or report)
        doc_class = self.options.document_type.value
        font_size = self.options.font_size.value
        lines.append(f"\\documentclass[{font_size}]{{{doc_class}}}\n")
        
        # Standard packages that are always useful
        if self.options.unicode_support:
            lines.append("\\usepackage[T1]{fontenc}")
            lines.append("\\usepackage[utf8]{inputenc}")
        
        lines.append(f"\\usepackage[{self.options.page_margins}]{{geometry}}")
        
        if self.options.preserve_images:
            lines.append("\\usepackage{graphicx}")
            # we assume images will be in a folder named 'images'
            lines.append("\\graphicspath{{./images/}}")
        
        lines.append("\\usepackage{hyperref}")
        lines.append("\\hypersetup{colorlinks=true,linkcolor=blue}")
        
        lines.append("\\usepackage{amsmath}")
        lines.append("\\usepackage{amssymb}")
        lines.append("\\usepackage{booktabs}")
        lines.append("\\usepackage{longtable}")
        lines.append("\\usepackage{array}")
        
        # Set the line spacing (single, 1.5, or double)
        if self.options.line_spacing != LineSpacing.SINGLE:
            lines.append("\\usepackage{setspace}")
            if self.options.line_spacing == LineSpacing.ONE_HALF:
                lines.append("\\onehalfspacing")
            elif self.options.line_spacing == LineSpacing.DOUBLE:
                lines.append("\\doublespacing")
        
        if self.options.extract_bibliography:
            lines.append("\\usepackage{natbib}")
            lines.append(f"\\bibliographystyle{{{self.options.bibliography_style}}}")
        
        # Add any extra packages that weren't in my list
        for package in self.options.custom_packages:
            lines.append(f"\\usepackage{{{package}}}")
        
        return '\n'.join(lines) + '\n'
    
    def _process_document_body(self, doc: Document, output_path: str) -> str:
        # Iterate through every element in the word doc's body
        lines = []
        
        for element in doc.element.body:
            # Check if it's a paragraph
            if isinstance(element, CT_P):
                paragraph = Paragraph(element, doc)
                latex_para = self._process_paragraph(paragraph)
                if latex_para:
                    lines.append(latex_para)
            # Check if it's a table
            elif isinstance(element, CT_Tbl):
                table = Table(element, doc)
                latex_table = self._process_table(table)
                if latex_table:
                    lines.append(latex_table)
        
        return '\n\n'.join(lines)
    
    def _process_paragraph(self, paragraph: Paragraph) -> str:
        # Converts a word paragraph into latex syntax
        if not paragraph.text.strip():
            return ""
        
        # Heading styles are handled specially
        if paragraph.style.name.startswith('Heading'):
            return self._process_heading(paragraph)
        
        # Process each "run" (part of text with same formatting)
        text_parts = []
        for run in paragraph.runs:
            # Escape symbols like &, %, $, etc.
            text = escape_latex(run.text)
            
            # Check for bold, italics, etc.
            if run.bold:
                text = f"\\textbf{{{text}}}"
            if run.italic:
                text = f"\\textit{{{text}}}"
            if run.underline:
                text = f"\\underline{{{text}}}"
            
            # Simple hyperlink handling (partial)
            if hasattr(run, 'hyperlink') and run.hyperlink:
                url = run.hyperlink.address if hasattr(run.hyperlink, 'address') else ''
                if url:
                    text = f"\\href{{{url}}}{{{text}}}"
            
            text_parts.append(text)
        
        paragraph_text = ''.join(text_parts)
        
        # Horizontal alignment
        try:
            alignment = paragraph.alignment
            if alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
                paragraph_text = f"\\begin{{center}}\n{paragraph_text}\n\\end{{center}}"
            elif alignment == WD_PARAGRAPH_ALIGNMENT.RIGHT:
                paragraph_text = f"\\begin{{flushright}}\n{paragraph_text}\n\\end{{flushright}}"
        except:
            # If alignment check fails for some reason (like 'start' mapping issue),
            # we just treat it as default alignment
            pass
        
        return paragraph_text
    
    def _process_heading(self, paragraph: Paragraph) -> str:
        # Decides which section command to use based on level
        text = escape_latex(paragraph.text)
        style_name = paragraph.style.name
        
        if 'Heading 1' in style_name:
            return f"\\section{{{text}}}"
        elif 'Heading 2' in style_name:
            return f"\\subsection{{{text}}}"
        elif 'Heading 3' in style_name:
            return f"\\subsubsection{{{text}}}"
        elif 'Heading 4' in style_name:
            return f"\\paragraph{{{text}}}"
        else:
            return f"\\subparagraph{{{text}}}"
    
    def _process_table(self, table: Table) -> str:
        # Converts word tables to latex longtable or tabular
        if not table.rows:
            return ""
        
        num_cols = len(table.rows[0].cells)
        
        lines = []
        lines.append("\\begin{table}[h]")
        lines.append("\\centering")
        # Creating column structure like |l|l|l|
        lines.append(f"\\begin{{tabular}}{{{'|'.join(['l'] * num_cols)}}}")
        lines.append("\\toprule")
        
        for i, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                # Escape each cell content
                cell_text = escape_latex(cell.text.strip())
                cells.append(cell_text)
            
            # Join cells with & for latex
            row_text = " & ".join(cells) + " \\\\"
            lines.append(row_text)
            
            # Add line after header
            if i == 0:
                lines.append("\\midrule")
        
        lines.append("\\bottomrule")
        lines.append("\\end{tabular}")
        lines.append("\\end{table}")
        
        return '\n'.join(lines)
    
    def _generate_bibliography(self, output_path: str) -> None:
        # Simple placeholder for bibliography generation
        if not self.bibliography_entries:
            return
        
        # Just creating a renamed .bib file
        bib_path = output_path.replace('.tex', '.bib')
        
        with open(bib_path, 'w', encoding=self.options.output_encoding) as f:
            for entry in self.bibliography_entries:
                f.write(entry + '\n\n')
        
        logger.info(f"Created bib file at: {bib_path}")
