#!/usr/bin/env python3
# Command line interface for doc2tex
# Used for batch processing or just quick converts from terminal

import argparse
import sys
import os
from pathlib import Path

from doc2tex import (
    DocTeXConverter,
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing,
    ConversionError
)
from doc2tex.utils import logger, setup_logger


def create_parser() -> argparse.ArgumentParser:
    # Setting up all the flags for terminal usage
    parser = argparse.ArgumentParser(
        description='doc2tex - Simple DOCX ↔ LaTeX Converter',
        epilog="""
Example usage:
  python cli.py file.docx -o result.tex
  python cli.py file.tex
        """
    )
    
    # Input files. You can pass multiple files here.
    parser.add_argument(
        'input',
        nargs='+',
        help='Files to convert'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Where to save the result. If not set, it guesses name.'
    )
    
    parser.add_argument(
        '-d', '--output-dir',
        help='Folder to save results if batching'
    )
    
    parser.add_argument(
        '--direction',
        choices=['to_latex', 'to_docx'],
        help='Force it to go one way'
    )
    
    # Document settings
    doc_group = parser.add_argument_group('Doc Settings')
    doc_group.add_argument(
        '--doc-type',
        type=str,
        choices=['article', 'report', 'thesis', 'letter', 'book'],
        default='article'
    )
    doc_group.add_argument(
        '--font-size',
        type=str,
        choices=['10pt', '11pt', '12pt'],
        default='12pt'
    )
    doc_group.add_argument(
        '--spacing',
        type=str,
        choices=['single', 'onehalf', 'double'],
        default='single'
    )
    doc_group.add_argument(
        '--margins',
        type=str,
        default='top=2.5cm,bottom=2.5cm,left=2.5cm,right=2.5cm'
    )
    
    # Bib and Image settings
    parser.add_argument('--extract-bib', action='store_true', help='Try to extract bibliography')
    parser.add_argument('--no-images', action='store_true', help='Dont handle images')
    parser.add_argument('--optimize-images', action='store_true', help='Shrink images')
    
    # General stuff
    parser.add_argument('-v', '--verbose', action='store_true', help='Show more logs')
    parser.add_argument('--version', action='version', version='doc2tex 1.0.0')
    
    return parser


def build_options(args: argparse.Namespace) -> ConversionOptions:
    # Map CLI args to our ConversionOptions class
    return ConversionOptions(
        document_type=DocumentType(args.doc_type),
        font_size=FontSize(args.font_size),
        line_spacing=LineSpacing(args.spacing),
        page_margins=args.margins,
        extract_bibliography=args.extract_bib,
        preserve_images=not args.no_images,
        optimize_images=args.optimize_images,
        verbose=args.verbose,
    )


def main():
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup the logger format
    setup_logger(verbose=args.verbose)
    
    try:
        options = build_options(args)
        converter = DocTeXConverter(options)
        
        # Batch conversion if more than one file
        if len(args.input) > 1 or args.output_dir:
            logger.info(f"Batch converting {len(args.input)} files...")
            
            results = converter.batch_convert(
                args.input,
                output_dir=args.output_dir,
                direction=args.direction
            )
            
            success = sum(1 for r in results if r is not None)
            logger.info(f"Done! {success}/{len(results)} worked.")
        
        # Single file conversion
        else:
            input_file = args.input[0]
            output_file = args.output
            
            result = converter.convert(
                input_file,
                output_file,
                direction=args.direction
            )
            
            print(f"Success! Saved to: {result}")
    
    except ConversionError as e:
        logger.error(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
