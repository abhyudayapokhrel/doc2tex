# The main class that ties everything together
# It detects if you're going DOCX -> LaTeX or LaTeX -> DOCX

import os
from pathlib import Path
from typing import Optional

from .options import ConversionOptions
from .latex import LatexGenerator
from .docx import DocxGenerator
from .utils import logger, is_valid_file, cleanup_temp_dir, get_file_info
from .errors import ConversionError, InvalidFileFormatError


class DocTeXConverter:
    # This class manages the whole conversion process
    # Just create an instance and call .convert()
    
    SUPPORTED_INPUT_FORMATS = ['docx', 'tex', 'latex']
    SUPPORTED_OUTPUT_FORMATS = ['docx', 'tex']
    
    def __init__(self, options: Optional[ConversionOptions] = None):
        # Initialize with default options if none provided
        self.options = options or ConversionOptions()
        
        # Validates that we didn't put weird values in options
        self.options.validate()
        
        if self.options.verbose:
            logger.setLevel('DEBUG')
    
    def convert(
        self, 
        input_path: str, 
        output_path: Optional[str] = None,
        direction: Optional[str] = None
    ) -> str:
        # Main function to convert a file
        if not os.path.exists(input_path):
            raise ConversionError(f"Input file doesn't exist: {input_path}")
        
        # Log basic file info so we know what's happening
        file_info = get_file_info(input_path)
        logger.info(f"Processing: {file_info['name']}")
        
        # Figure out if we are going to latex or to docx
        if direction is None:
            direction = self._detect_direction(input_path)
        
        # Create output path if user didn't specify one
        if output_path is None:
            output_path = self._generate_output_path(input_path, direction)
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # Run the actual conversion based on direction
        try:
            if direction == 'to_latex':
                result = self._convert_to_latex(input_path, output_path)
            elif direction == 'to_docx':
                result = self._convert_to_docx(input_path, output_path)
            else:
                raise ConversionError(f"Invalid direction: {direction}")
            
            return result
            
        except Exception as e:
            logger.error(f"Failed to convert: {e}")
            raise
    
    # Shortcut for docx -> latex
    def convert_to_latex(self, input_path: str, output_path: Optional[str] = None) -> str:
        return self.convert(input_path, output_path, direction='to_latex')
    
    # Shortcut for latex -> docx
    def convert_to_docx(self, input_path: str, output_path: Optional[str] = None) -> str:
        return self.convert(input_path, output_path, direction='to_docx')
    
    def _detect_direction(self, input_path: str) -> str:
        # Checks the extension to guess what user wants
        ext = Path(input_path).suffix.lower().lstrip('.')
        
        if ext == 'docx':
            return 'to_latex'
        elif ext in ['tex', 'latex']:
            return 'to_docx'
        else:
            raise InvalidFileFormatError(f"I don't know how to handle .{ext} files. Use DOCX or TEX.")
    
    def _generate_output_path(self, input_path: str, direction: str) -> str:
        # Rename input.docx to input.tex or vice versa
        input_path_obj = Path(input_path)
        stem = input_path_obj.stem
        
        if direction == 'to_latex':
            return str(input_path_obj.parent / f"{stem}.tex")
        else:
            return str(input_path_obj.parent / f"{stem}.docx")
    
    def _convert_to_latex(self, input_path: str, output_path: str) -> str:
        # Calls the LatexGenerator
        if not is_valid_file(input_path, ['docx']):
            raise InvalidFileFormatError("Need a .docx file for this")
        
        generator = LatexGenerator(self.options)
        result = generator.convert(input_path, output_path)
        
        if self.options.clean_temp_files and generator.temp_dir:
            cleanup_temp_dir(generator.temp_dir)
        
        return result
    
    def _convert_to_docx(self, input_path: str, output_path: str) -> str:
        # Calls the DocxGenerator
        if not is_valid_file(input_path, ['tex', 'latex']):
            raise InvalidFileFormatError("Need a .tex or .latex file for this")
        
        generator = DocxGenerator(self.options)
        result = generator.convert(input_path, output_path)
        
        return result
    
    def batch_convert(
        self, 
        input_files: list, 
        output_dir: Optional[str] = None,
        direction: Optional[str] = None
    ) -> list:
        # For converting multiple files at once
        results = []
        
        for input_file in input_files:
            try:
                # If they gave an output dir, put the result there
                if output_dir:
                    filename = Path(input_file).name
                    detected_direction = direction or self._detect_direction(input_file)
                    ext = '.tex' if detected_direction == 'to_latex' else '.docx'
                    output_path = os.path.join(output_dir, Path(filename).stem + ext)
                else:
                    output_path = None
                
                result = self.convert(input_file, output_path, direction)
                results.append(result)
                
            except Exception as e:
                logger.error(f"Skipping {input_file} because of error: {e}")
                results.append(None)
        
        return results
    
    def get_conversion_info(self) -> dict:
        # Just returns some info about the current setup
        return {
            'inputs': self.SUPPORTED_INPUT_FORMATS,
            'outputs': self.SUPPORTED_OUTPUT_FORMATS,
            'options_used': self.options.to_dict(),
        }
