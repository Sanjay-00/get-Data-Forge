import hashlib
import io
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st

from core.io import get_sheet_names, read_single_file
from core.memory import release_memory

# Uploaded workbooks are parsed once per upload and kept for that browser
# session only:
# - cache_resource hands back the cached DataFrame itself. cache_data would
#   pickle/unpickle a full copy on every rerun, which for a 200k row sheet
#   costs both time and a second copy in memory on every widget click.
#   Callers must treat the returned frame as read-only (every tool already
#   copies or derives new frames instead of mutating in place).
# - scope="session" drops the entries when the tab disconnects, so one
#   user's uploads never sit in server memory after they leave.
# - Keyed on the uploader's per-upload file_id instead of hashing the raw
#   bytes, which on a multi-file merge meant re-hashing every file's full
#   content on every rerun.
# - Entries are also evicted as soon as their upload is gone (removed from
#   the uploader, replaced, or left behind on another tool page), see
#   release_unused_uploads(); the ttl is only a backstop.
_FILE_CACHE = dict(scope="session", max_entries=32, ttl="1h", show_spinner=False)

_REGISTRY_KEY = "_cached_upload_entries"


def ns_key(tool_name, key):
    return f"{tool_name}__{key}"


def file_key(uploaded_file):
    """Stable identity for one uploaded file's content."""
    file_id = getattr(uploaded_file, "file_id", None)
    if file_id:
        return file_id
    return hashlib.md5(uploaded_file.getvalue()).hexdigest()


def remember_cache_entry(key, cached_func, *args):
    """Record that cached_func(*args) holds data derived from the upload
    `key`, so it can be evicted when that upload goes away. args must be
    the call's arguments (underscore-prefixed ones may be passed as None,
    they aren't part of the cache key)."""
    try:
        registry = st.session_state.setdefault(_REGISTRY_KEY, {})
    except Exception:
        return  # outside a Streamlit session (tests, bare scripts)
    entries = registry.setdefault(key, [])
    entry = (cached_func, args)
    if entry not in entries:
        entries.append(entry)


def release_unused_uploads(*uploaded_files):
    """Evict every cached entry built from an upload that isn't among
    uploaded_files. Called by each tool page with the files it currently
    has: when the user removes or replaces a file, or comes from another
    tool page (whose uploads Streamlit discards on navigation), the parsed
    data for those files is dropped right away instead of lingering in
    memory until the cache ttl."""
    try:
        registry = st.session_state.get(_REGISTRY_KEY)
    except Exception:
        return
    if not registry:
        return
    active = {file_key(f) for f in uploaded_files if f is not None}
    stale = [k for k in registry if k not in active]
    for key in stale:
        for cached_func, args in registry.pop(key):
            try:
                cached_func.clear(*args)
            except Exception:
                pass
    if stale:
        release_memory()


def _buffer(content, filename):
    buf = io.BytesIO(content)
    buf.name = filename
    return buf


@st.cache_resource(**_FILE_CACHE)
def _cached_sheet_names(key, filename, _content):
    return get_sheet_names(_buffer(_content, filename))


# Results parsed ahead of time by cached_read_many's worker threads, waiting
# to be handed to _cached_read. Keyed by (upload key, sheet); upload keys
# are unique per upload, so sessions can't see each other's entries.
_prefetched = {}
_prefetched_lock = threading.Lock()

# Parallel parsing: fastexcel/calamine release the GIL, so files parse side
# by side. Measured on 4 x 200k-row (16 MB) files: 1 worker 5.3s / +367 MB
# peak, 2 workers 3.3s / +544 MB, 3 workers 3.3s / +635 MB. Two workers get
# all of the speedup; a third only adds memory. Each in-flight parse peaks
# at roughly 11x its file size, so big files are parsed one at a time to
# keep the peak safe on a small (~1 GB) host.
_READ_WORKERS = 2
_PARALLEL_MAX_FILE_BYTES = 20 * 2**20


@st.cache_resource(**_FILE_CACHE)
def _cached_read(key, filename, sheet_idx, _content):
    with _prefetched_lock:
        result = _prefetched.pop((key, sheet_idx), None)
    if result is not None:
        return result
    result = read_single_file(_buffer(_content, filename), sheet_idx)
    # The parse leaves large transient buffers behind; give them back now.
    release_memory()
    return result


def cached_sheet_names(uploaded_file):
    key = file_key(uploaded_file)
    remember_cache_entry(key, _cached_sheet_names, key, uploaded_file.name, None)
    return _cached_sheet_names(key, uploaded_file.name, uploaded_file.getvalue())


def cached_read(uploaded_file, sheet_idx):
    """(df, err) for one sheet of an uploaded file. df is shared across
    reruns: never mutate it in place."""
    key = file_key(uploaded_file)
    remember_cache_entry(key, _cached_read, key, uploaded_file.name, sheet_idx, None)
    return _cached_read(key, uploaded_file.name, sheet_idx, uploaded_file.getvalue())


def _is_cached(key, filename, sheet_idx):
    try:
        entries = st.session_state.get(_REGISTRY_KEY, {}).get(key, [])
    except Exception:
        return False
    return (_cached_read, (key, filename, sheet_idx, None)) in entries


def cached_read_many(requests, progress_cb=None):
    """cached_read for many (uploaded_file, sheet) pairs, parsing the ones
    not cached yet in parallel. Returns [(df, err), ...] in request order.
    progress_cb(done, total) is called from this (the script) thread as
    files finish. Worker threads only parse; the results are handed to the
    session cache from here, since Streamlit's cache can't be used from
    other threads."""
    todo = [(f, s) for f, s in requests if not _is_cached(file_key(f), f.name, s)]
    largest = max((len(f.getvalue()) for f, _ in todo), default=0)
    if len(todo) > 1 and largest <= _PARALLEL_MAX_FILE_BYTES:
        keys = [(file_key(f), s) for f, s in todo]
        try:
            with ThreadPoolExecutor(max_workers=min(_READ_WORKERS, len(todo))) as pool:
                futures = {
                    pool.submit(read_single_file, _buffer(f.getvalue(), f.name), s): k
                    for (f, s), k in zip(todo, keys)
                }
                for done, future in enumerate(as_completed(futures), start=1):
                    with _prefetched_lock:
                        _prefetched[futures[future]] = future.result()
                    if progress_cb:
                        progress_cb(done, len(todo))
            results = [cached_read(f, s) for f, s in requests]
        finally:
            # Anything not consumed (e.g. an exception above) must not linger.
            with _prefetched_lock:
                for k in keys:
                    _prefetched.pop(k, None)
        release_memory()
        return results
    results = []
    for i, (f, s) in enumerate(requests, start=1):
        results.append(cached_read(f, s))
        if progress_cb:
            progress_cb(i, len(requests))
    return results
