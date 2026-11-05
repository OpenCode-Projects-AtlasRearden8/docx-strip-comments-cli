"""docx_strip: copy a .docx and strip tracked changes, comments, and personal metadata.

Pure Python standard library. No network. Never modifies the original file.
"""

from .scrub import strip_docx, DEFAULT_SUFFIX, DocxStripError

__all__ = ["strip_docx", "DEFAULT_SUFFIX", "DocxStripError"]
__version__ = "1.0.0"
