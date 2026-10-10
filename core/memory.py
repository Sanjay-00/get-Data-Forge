import ctypes
import gc
import sys

# On Linux (where the app is hosted) glibc keeps freed heap memory inside the
# process instead of returning it to the OS, so a server that once parsed a
# big workbook keeps that peak RSS indefinitely even after every DataFrame is
# gone. malloc_trim(0) hands the free pages back. Elsewhere it's a no-op.
try:
    _malloc_trim = ctypes.CDLL("libc.so.6").malloc_trim if sys.platform.startswith("linux") else None
except (OSError, AttributeError):
    _malloc_trim = None


def release_memory():
    """Collect garbage and return freed heap pages to the OS. Call after
    large objects (parsed sheets, built workbooks) have been dropped."""
    gc.collect()
    if _malloc_trim is not None:
        try:
            _malloc_trim(0)
        except Exception:
            pass
