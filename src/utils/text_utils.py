"""Text processing utilities."""

import html
import re
from collections import OrderedDict
from typing import Dict, Iterable, List, Sequence

from .constants import FOOTNOTE_LINE_RE, NON_ALNUM_RE, URL_RE


def normalize_text(value: object) -> str:
    """Normalize text by removing extra whitespace and special characters.
    
    Args:
        value: Text value to normalize
        
    Returns:
        Normalized text string
    """
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\xa0", " ").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def slugify(text: str) -> str:
    """Convert text to a URL-safe slug.
    
    Args:
        text: Text to slugify
        
    Returns:
        Slugified text with only alphanumeric and underscores
    """
    text = normalize_text(text).replace("/", " ")
    text = re.sub(r"[^A-Za-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def htmlize_text(text: str) -> str:
    """Escape HTML and convert URLs to links.
    
    Args:
        text: Text to process
        
    Returns:
        HTML-escaped text with clickable URLs
    """
    escaped = html.escape(normalize_text(text))
    return URL_RE.sub(r'<a href="\1" target="_blank">\1</a>', escaped)


def unique_preserve_order(items: Iterable[str]) -> List[str]:
    """Get unique items while preserving order.
    
    Args:
        items: Iterable of strings
        
    Returns:
        List of unique strings in original order
    """
    seen = OrderedDict()
    for item in items:
        item = normalize_text(item)
        if item:
            seen[item] = None
    return list(seen.keys())


def parse_footnote_map(text: str) -> Dict[str, str]:
    """Extract footnotes from text in format '1. Footnote text'.
    
    Args:
        text: Text containing footnotes
        
    Returns:
        Dictionary mapping footnote number to text
    """
    notes: Dict[str, str] = {}
    for line in normalize_text(text).split("\n"):
        m = FOOTNOTE_LINE_RE.match(line)
        if m:
            notes[m.group(1)] = normalize_text(m.group(2))
    return notes


def split_main_notes_and_bullets(text: str) -> tuple[str, Dict[str, str], List[str]]:
    """Split text into main content, footnotes, and bullet points.
    
    Args:
        text: Text to split
        
    Returns:
        Tuple of (main_text, footnotes_dict, bullet_list)
    """
    lines = [normalize_text(line) for line in normalize_text(text).split("\n") if normalize_text(line)]
    main_lines: List[str] = []
    notes: Dict[str, str] = {}
    bullets: List[str] = []
    in_notes = False
    
    for line in lines:
        m = FOOTNOTE_LINE_RE.match(line)
        if m:
            notes[m.group(1)] = normalize_text(m.group(2))
            in_notes = True
            continue
        if line.startswith("•"):
            bullets.append(normalize_text(line.lstrip("•").strip()))
            in_notes = True
            continue
        if not in_notes:
            main_lines.append(line)
        else:
            bullets.append(line)
    
    main_text = normalize_text(" ".join(main_lines))
    return main_text, notes, bullets


def sentence_chunks(text: str, max_words: int = 110) -> List[str]:
    """Split text into chunks by sentences, respecting word limit.
    
    Args:
        text: Text to chunk
        max_words: Maximum words per chunk
        
    Returns:
        List of text chunks
    """
    text = normalize_text(text)
    if not text:
        return []
    
    parts = [normalize_text(p) for p in re.split(r'(?<=[.!?])\s+|\n+', text) if normalize_text(p)]
    if not parts:
        return [text]
    
    chunks: List[str] = []
    current: List[str] = []
    current_words = 0
    
    for part in parts:
        words = len(part.split())
        if words > max_words and not current:
            # Split oversized part
            raw_words = part.split()
            for i in range(0, len(raw_words), max_words):
                chunks.append(' '.join(raw_words[i:i + max_words]))
            continue
        
        if current and current_words + words > max_words:
            chunks.append(normalize_text(' '.join(current)))
            current = [part]
            current_words = words
        else:
            current.append(part)
            current_words += words
    
    if current:
        chunks.append(normalize_text(' '.join(current)))
    
    return [c for c in chunks if c]


def chunk_list(items: Sequence[str], max_words: int = 110, max_items: int = 8) -> List[List[str]]:
    """Split items into chunks respecting word and item limits.
    
    Args:
        items: List of items to chunk
        max_words: Maximum words per chunk
        max_items: Maximum items per chunk
        
    Returns:
        List of item chunks
    """
    chunks: List[List[str]] = []
    current: List[str] = []
    current_words = 0
    
    for item in items:
        item = normalize_text(item)
        if not item:
            continue
        words = len(item.split())
        
        if current and (current_words + words > max_words or len(current) >= max_items):
            chunks.append(current)
            current = [item]
            current_words = words
        else:
            current.append(item)
            current_words += words
    
    if current:
        chunks.append(current)
    
    return chunks


def compact_text(text: str, max_words: int = 20) -> str:
    """Compact text to a maximum word count.
    
    Args:
        text: Text to compact
        max_words: Maximum words to keep
        
    Returns:
        Compacted text with ellipsis if truncated
    """
    text = normalize_text(text)
    if not text:
        return ''
    
    words = text.split()
    if len(words) <= max_words:
        return text
    
    return ' '.join(words[:max_words]).rstrip(',;:') + '...'
