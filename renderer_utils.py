import re
from typing import List, Optional, Tuple

def cm_to_pt(cm: float) -> float:
    return float(cm) * 28.346456692913385

def pt_to_cm(pt: float) -> float:
    return float(pt) / 28.346456692913385

COLOR_NAME_TO_RGB = {
    "red": 255,                 # RGB(255, 0, 0)
    "blue": 16711680,           # RGB(0, 0, 255)
    "green": 32768,             # RGB(0, 128, 0)
    "yellow": 65535,            # RGB(255, 255, 0)
    "purple": 8388736,          # RGB(128, 0, 128)
    "orange": 42495,            # RGB(255, 165, 0)
    "black": 0,
    "white": 16777215,
    "darkblue": 9125196,        # RGB(12, 15, 139)
    "grey": 8421504,
    "gray": 8421504,
}

HIGHLIGHT_NAME_TO_INDEX = {
    "yellow": 7,    # wdYellow = 7
    "green": 4,     # wdGreen = 4
    "cyan": 3,      # wdTurquoise = 3
    "pink": 5,      # wdPink = 5
    "blue": 9,      # wdBlue = 9
    "red": 6,       # wdRed = 6
    "gray": 15,     # wdGray25 = 15
}

def parse_color_to_rgb_int(color_str: str) -> Optional[int]:
    if not color_str:
        return None
    clean = color_str.strip().lower()
    hex_match = re.search(r'#([0-9a-fA-F]{6})', color_str)
    if hex_match:
        hex_val = hex_match.group(1).lower()
        r = int(hex_val[0:2], 16)
        g = int(hex_val[2:4], 16)
        b = int(hex_val[4:6], 16)
        return r + (g * 256) + (b * 65536)
    if clean in COLOR_NAME_TO_RGB:
        return COLOR_NAME_TO_RGB[clean]
    return None

def extract_question_prefix_and_body(text: str) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
    """
    Extracts (full_prefix, delim_char, q_num, body_text) when formatted with '#' placeholder:
    - #1. -> ('', '.', '1', 'Mike...') -> Word list: '1.'
    - Question #1. -> ('Question ', '.', '1', 'What...') -> Word list: 'Question 1.'
    - Question #1: -> ('Question ', ':', '1', 'What...') -> Word list: 'Question 1:'
    - Câu #1: -> ('Câu ', ':', '1', '...') -> Word list: 'Câu 1:'
    - Task #1. -> ('Task ', '.', '1', '...') -> Word list: 'Task 1.'
    Returns (full_prefix, delim_char, q_num, body_text) if matched, else (None, None, None, original_text).
    """
    if not text or '#' not in text:
        return None, None, None, text

    pattern = r'^\s*(?:\*\*)?([A-Za-zÀ-ỹ\s]+?)?#(\d+)([\.\)\:\-]|\s*:\s*|\s*\.\s*)?\s*(?:\*\*)?[:\.\)]?\s*(.*)$'
    m = re.match(pattern, text, re.IGNORECASE)
    if m:
        prefix_word = m.group(1).strip() if m.group(1) else ""
        q_num = m.group(2)
        delimiter = m.group(3).strip() if m.group(3) else "."
        body = m.group(4).strip()
        body = re.sub(r'^(?:\*\*|[:\.\-\)])\s*', '', body).strip()

        full_prefix = f"{prefix_word} " if prefix_word else ""
        delim_char = ":" if ":" in delimiter else ("." if "." in delimiter else (delimiter if delimiter else "."))
        return full_prefix, delim_char, q_num, body
    return None, None, None, text

def split_line_into_option_items(line_text: str) -> List[str]:
    """
    Splits line text into multi-column items automatically by detecting:
    1. Explicit \\t characters
    2. Multiple sequential option choices like A. text B. text C. text D. text or a. b. c. d. or 1. 2. 3. 4.
    """
    if not line_text:
        return []

    if '\t' in line_text:
        return [x.strip() for x in line_text.split('\t') if x.strip()]

    # Extract potential question number prefix (e.g. "#1. ", "Question #1: ")
    pref, delim, q_num, rest_text = extract_question_prefix_and_body(line_text)
    q_prefix = f"{pref}#{q_num}{delim} " if q_num else ""

    # Search for option choices like A. B. C. D. or a. b. c. d. or (A) (B) (C) or 1. 2. 3. 4.
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
    """Capitalizes first letter of each main word, keeping minor words in lowercase and preserving Roman numerals."""
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
    """Capitalizes ONLY the first letter of the sentence, preserving Roman numeral / section letter prefix."""
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
