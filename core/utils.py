"""
core/utils.py - Shared Text Formatting, Measurement, and Typography Utilities.
"""

import os
import re
import hashlib
import tempfile
from typing import List, Optional, Tuple

from services.image_service import (
    SUPPORTED_IMAGE_EXTENSIONS,
    ensure_word_compatible_image,
)
from renderers.common.units_and_colors import (
    cm_to_pt,
    pt_to_cm,
    COLOR_NAME_TO_RGB,
    HIGHLIGHT_NAME_TO_INDEX,
    parse_color_to_rgb_int,
)
from renderers.common.typography import (
    GdiTextMeasurer,
    get_gdi_text_measurer,
)


def natural_sort_key(s: str) -> list:
    """Returns alphanumeric sorting key for natural sorting (e.g. img1, img2, img10)."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def extract_question_prefix_and_body(text: str) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
    """
    Universally extracts (full_prefix, delim_char, q_num, body_text) from any question text.
    Handles markdown/ULN tags, explicit '#', question words (Question, Câu, Task, etc.),
    parentheses, and delimiters.
    """
    if not text or not text.strip():
        return None, None, None, text

    s = text.strip()

    lead_tags_m = re.match(r'^((?:\[(?:ins|b|i|u|hl|ans|q|opt|box|p0|p1|p2|gap)\]|<(?:b|i|u|strong|em|span)[^>]*>|\*\*|\*|_|\{u\})*)\s*', s, re.IGNORECASE)
    lead_tags = lead_tags_m.group(1) if lead_tags_m else ''
    s_work = s[len(lead_tags):].strip() if lead_tags else s

    pref_word_m = re.match(r'^(Question|Câu|Task|Activity|Ex|Exercise|Part|Section|Item|Sentence|Dialogue|Problem|Bài)\s+', s_work, re.IGNORECASE)
    pref_word = pref_word_m.group(1) if pref_word_m else ''
    if pref_word:
        s_work = s_work[pref_word_m.end():].strip()

    num_m = re.match(r'^(#\s*)?(?:(\(|\[)\s*)?(#\s*)?(\d+)([\.\:\/\-])?(?:\s*(\)|\]))?([\.\:\/\-])?\s*(.*)$', s_work, re.DOTALL)
    if not num_m:
        return None, None, None, text

    has_hash = bool(num_m.group(1) or num_m.group(3) or ('#' in lead_tags))
    open_p = num_m.group(2) or ''
    q_num = num_m.group(4)
    inner_d = num_m.group(5) or ''
    close_p = num_m.group(6) or ''
    outer_d = num_m.group(7) or ''
    raw_body = num_m.group(8) or ''

    has_paren = bool(open_p and close_p)
    has_delim = bool(inner_d or outer_d or close_p)

    if not (has_hash or pref_word or has_paren or has_delim):
        return None, None, None, text

    body = raw_body.strip()
    if body:
        if '**' in lead_tags and re.search(r'\*\*(?:\[\/(?:ins|b|i|u|hl)\])*$', body) and not re.match(r'^(?:\[(?:ins|b|i|u|hl)\])*\*\*', body):
            m_lead = re.match(r'^(?:\[(?:ins|b|i|u|hl)\])+', body, re.IGNORECASE)
            if m_lead:
                body = body[:m_lead.end()] + '**' + body[m_lead.end():]
            else:
                body = '**' + body

        if '*' in lead_tags and not ('**' in lead_tags) and re.search(r'\*(?:\[\/(?:ins|b|i|u|hl)\])*$', body) and not re.match(r'^(?:\[(?:ins|b|i|u|hl)\])*\*', body):
            m_lead = re.match(r'^(?:\[(?:ins|b|i|u|hl)\])+', body, re.IGNORECASE)
            if m_lead:
                body = body[:m_lead.end()] + '*' + body[m_lead.end():]
            else:
                body = '*' + body

        for tag in ['ins', 'b', 'i', 'u', 'hl']:
            open_t = f'[{tag}]'
            close_t = f'[/{tag}]'
            if open_t in lead_tags.lower() and close_t in body.lower() and not body.lower().startswith(open_t):
                body = open_t + body

        body = re.sub(r'^(?:[:\.\-\)])\s*', '', body).strip()

    if open_p == '(' and close_p == ')':
        full_prefix = f'{pref_word} (' if pref_word else '('
        delim_char = f'{inner_d}){outer_d}' if outer_d else (f'{inner_d})' if inner_d else ')')
        return full_prefix, delim_char, q_num, body
    elif open_p == '[' and close_p == ']':
        full_prefix = f'{pref_word} [' if pref_word else '['
        delim_char = f'{inner_d}]{outer_d}' if outer_d else (f'{inner_d}]' if inner_d else ']')
        return full_prefix, delim_char, q_num, body
    elif close_p == ')':
        full_prefix = f'{pref_word} ' if pref_word else ''
        delim_char = ')'
        return full_prefix, delim_char, q_num, body
    elif close_p == ']':
        full_prefix = f'{pref_word} ' if pref_word else ''
        delim_char = ']'
        return full_prefix, delim_char, q_num, body
    else:
        raw_delim = outer_d or inner_d
        delim_char = raw_delim if raw_delim else '.'
        full_prefix = f'{pref_word} ' if pref_word else ''
        return full_prefix, delim_char, q_num, body


def split_line_into_option_items(line_text: str) -> List[str]:
    """
    Splits line text into multi-column items automatically by detecting \t or
    sequential option letters (A. B. C. D. or 1. 2. 3. 4.).
    """
    if not line_text:
        return []

    if '\t' in line_text:
        return [x.strip() for x in line_text.split('\t') if x.strip()]

    pref, delim, q_num, rest_text = extract_question_prefix_and_body(line_text)
    q_prefix = f"{pref}#{q_num}{delim} " if q_num else ""

    pattern = r'(?:^|\s+)(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z0-9][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)(?=\s+|$)'
    matches = list(re.finditer(pattern, rest_text))
    if len(matches) >= 2:
        labels = [re.sub(r'[\*\_\`\[\]\(\)\.\s]', '', m.group(1)) for m in matches]

        is_seq = False
        if all(len(l) == 1 and l.isalpha() for l in labels):
            if labels[0] == 'A' and all(ord(l) == ord('A') + i for i, l in enumerate(labels)):
                is_seq = True
            elif labels[0] == 'a' and all(ord(l) == ord('a') + i for i, l in enumerate(labels)):
                is_seq = True
        elif all(l.isdigit() for l in labels):
            nums = [int(l) for l in labels]
            if nums[0] == 1 and all(n == 1 + i for i, n in enumerate(nums)):
                is_seq = True

        if is_seq:
            items = []
            for i in range(len(matches)):
                start = matches[i].start()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(rest_text)
                item_txt = rest_text[start:end].strip()
                if i == 0 and q_prefix:
                    item_txt = f"{q_prefix.strip()} {item_txt}"
                items.append(item_txt)
            return items

    return [line_text.strip()]


def apply_title_case_to_text(text: str) -> str:
    """Capitalizes first letter of each main word, preserving Roman numerals."""
    if not text:
        return text
    minor_words = {'a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'from', 'by', 'with', 'in', 'of', 'off'}
    words = text.split(' ')
    cased_words = []
    for idx, w in enumerate(words):
        if not w:
            cased_words.append(w)
            continue
        m = re.match(r'^([^\w]*)([\w\'-]+)([^\w]*)$', w)
        if not m:
            cased_words.append(w)
            continue
        pre, core, post = m.group(1), m.group(2), m.group(3)
        if re.match(r'^(?:[IVXLCDM]+|[A-Z]|\d+)$', core, re.IGNORECASE) and len(core) <= 4 and core.upper() in ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "A", "B", "C", "D", "E", "F", "G"]:
            cased_words.append(f"{pre}{core.upper()}{post}")
        elif idx == 0 or idx == len(words) - 1 or core.lower() not in minor_words:
            cased_words.append(f"{pre}{core.capitalize()}{post}")
        else:
            cased_words.append(f"{pre}{core.lower()}{post}")
    return ' '.join(cased_words)


def apply_sentence_case_to_text(text: str) -> str:
    """Capitalizes ONLY first letter of the sentence, preserving Roman/number prefix."""
    if not text:
        return text
    m = re.match(r'^(\s*(?:[IVXLCDM]+\.?|[A-Z]\.|\#?\d+[\.\:]?)\s*)(.*)$', text)
    if m:
        pref = m.group(1)
        body = m.group(2).strip()
        if body:
            body_cased = body[0].upper() + body[1:].lower()
            return f"{pref}{body_cased}"
        return text
    clean = text.strip()
    if clean:
        return clean[0].upper() + clean[1:].lower()
    return text


__all__ = [
    "natural_sort_key",
    "extract_question_prefix_and_body",
    "split_line_into_option_items",
    "apply_title_case_to_text",
    "apply_sentence_case_to_text",
    "SUPPORTED_IMAGE_EXTENSIONS",
    "ensure_word_compatible_image",
    "cm_to_pt",
    "pt_to_cm",
    "COLOR_NAME_TO_RGB",
    "HIGHLIGHT_NAME_TO_INDEX",
    "parse_color_to_rgb_int",
    "GdiTextMeasurer",
    "get_gdi_text_measurer",
]
