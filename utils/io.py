import os
from typing import Iterable, List
from pathlib import Path

def iter_xml_paths_from_dir(dir_path: str) -> Iterable[Path]:
    p = Path(dir_path)
    for ext in ("*.xml", "*.XML"):
        yield from p.rglob(ext)

def chunked(iterable, size: int):
    chunk = []
    for item in iterable:
        chunk.append(item)
        if len(chunk) >= size:
            yield chunk
            chunk = []
    if chunk:
        yield chunk
