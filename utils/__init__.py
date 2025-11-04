"""
Utils package for WordStyler
"""

# from .docx_split_merge import DocxSplitMergeProcessor, quick_split_document
# from .docx_split_no_toc import DocxSplitNoTocProcessor, quick_split_document_no_toc
# from .docx_split_combined import split_document_combined, quick_split_document_combined
from .docx_split import DocxSplitProcessor, split_document_for_conversion, quick_split_for_conversion
from .docx_picture import format_pictures_in_document, format_pictures_with_advanced_settings

__all__ = [
    # "DocxSplitMergeProcessor",
    # "quick_split_document",
    # "DocxSplitNoTocProcessor",
    # "quick_split_document_no_toc",
    # "split_document_combined",
    # "quick_split_document_combined",
    "DocxSplitProcessor",
    "split_document_for_conversion",
    "quick_split_for_conversion",
    "format_pictures_in_document",
    "format_pictures_with_advanced_settings"
]