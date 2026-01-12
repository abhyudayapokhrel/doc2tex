# doc2tex package initialization
# Exports the main converter class and options for easy use

from .converter import DocTeXConverter
from .options import (
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing
)
from .errors import (
    DocTeXError,
    ConversionError
)

__version__ = "1.0.0"
