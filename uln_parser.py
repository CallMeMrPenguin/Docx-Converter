import re
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, field

@dataclass
class InlineSpan:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    uppercase: bool = False
    color: Optional[str] = None      # e.g., "red", "#FF0000"
    bg_color: Optional[str] = None   # e.g., "yellow", "#FFFF00"
    is_instruction: bool = False     # Flag for [ins]...[/ins] instruction formatting

@dataclass
class PicInfo:
    description: str
    pos: str = "center"   # "inline", "center", "right"
    size: str = "medium"  # "small", "medium", "large"
    filepath: Optional[str] = None

@dataclass
class ULNTableCell:
    content: str
    is_header: bool = False
    spans: List[InlineSpan] = field(default_factory=list)

@dataclass
class ULNTableRow:
    cells: List[ULNTableCell] = field(default_factory=list)
    is_header: bool = False

@dataclass
class ULNTableData:
    rows: List[ULNTableRow] = field(default_factory=list)
    borderless: bool = False

@dataclass
class ULNBlock:
    tag: str  # "H1", "H2", "H3", "H4", "H5", "H6", "P0", "P1", "P2", "TAB2", "BOX", "QUOTE", "PIC", "TABLE", "P", "INS"
    content: str = ""
    col1: str = ""
    col2: str = ""
    col1_spans: List[InlineSpan] = field(default_factory=list)
    col2_spans: List[InlineSpan] = field(default_factory=list)
    spans: List[InlineSpan] = field(default_factory=list)
    pic: Optional[PicInfo] = None
    table_data: Optional[ULNTableData] = None
    children: List['ULNBlock'] = field(default_factory=list)
    is_instruction: bool = False


def parse_pic_tag(text: str) -> Optional[PicInfo]:
    """
    Parses [PIC] or [PIC: "description" | pos:center | size:medium]
    """
    text_strip = text.strip()
    if text_strip.upper() == "[PIC]":
        return PicInfo(description="Activity Picture", pos="center", size="medium")

    m = re.search(r'\[PIC:\s*"(.*?)"(?:\s*\|\s*pos:(\w+))?(?:\s*\|\s*size:(\w+))?\s*\]', text, re.IGNORECASE)
    if not m:
        m = re.search(r'\[PIC:\s*([^\|\]]+)(?:\s*\|\s*pos:(\w+))?(?:\s*\|\s*size:(\w+))?\s*\]', text, re.IGNORECASE)
    if m:
        desc = m.group(1).strip()
        pos = m.group(2).lower() if m.group(2) else "center"
        size = m.group(3).lower() if m.group(3) else "medium"
        return PicInfo(description=desc, pos=pos, size=size)

    if re.search(r'\[PIC\]', text, re.IGNORECASE):
        return PicInfo(description="Activity Picture", pos="center", size="medium")

    return None


def parse_inline_spans(text: str, default_bold: bool = False, default_italic: bool = False, default_instruction: bool = False) -> List[InlineSpan]:
    """
    Parses inline formatting elements recursively, handling nested tags like **bold with *italic*** and [ins] instruction tags.
    """
    if not text:
        return []

    # Handle unclosed [ins] tag at the start of line/text (e.g. [ins]**II. Choose...**)
    if re.match(r'^\s*\[ins\]', text, re.IGNORECASE) and not re.search(r'\[/ins\]', text, re.IGNORECASE):
        cleaned_text = re.sub(r'^\s*\[ins\]\s*', '', text, flags=re.IGNORECASE)
        return parse_inline_spans(cleaned_text, default_bold=default_bold, default_italic=default_italic, default_instruction=True)

    pattern = re.compile(
        r'(?P<pic>\[PIC(?::[^\]]+)?\])|'
        r'(?P<ins>\[ins\](?P<ins_txt>.*?)\[/ins\])|'
        r'(?P<annot>\[(?P<ann_txt>[^\]]+)\]\{(?P<ann_mod>[^\}]+)\})|'
        r'(?P<bold_italic>\*\*\*(?P<bi_txt>.*?)\*\*\*)|'
        r'(?P<bold>\*\*(?P<b_txt>.*?)\*\*)|'
        r'(?P<italic>\*(?P<i_txt>.*?)\*)',
        re.IGNORECASE | re.DOTALL
    )

    spans = []
    last_idx = 0

    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            plain_txt = text[last_idx:start]
            plain_txt = re.sub(r'\[\/?ins\]', '', plain_txt, flags=re.IGNORECASE)
            if plain_txt:
                spans.append(InlineSpan(text=plain_txt, bold=default_bold, italic=default_italic, is_instruction=default_instruction))

        gd = match.groupdict()

        if gd['pic']:
            spans.append(InlineSpan(text=gd['pic'], bold=default_bold, italic=default_italic, is_instruction=default_instruction))

        elif gd['ins']:
            ins_inner = gd['ins_txt']
            inner_spans = parse_inline_spans(ins_inner, default_bold=default_bold, default_italic=default_italic, default_instruction=True)
            spans.extend(inner_spans)

        elif gd['bold_italic']:
            bi_inner = gd['bi_txt']
            inner_spans = parse_inline_spans(bi_inner, default_bold=True, default_italic=True, default_instruction=default_instruction)
            spans.extend(inner_spans)

        elif gd['annot']:
            ann_txt = gd['ann_txt']
            ann_mod = gd['ann_mod'].strip().lower()
            
            b_flag = default_bold
            i_flag = default_italic
            underline = False
            uppercase = False
            color = None
            bg_color = None

            for mod in ann_mod.split(','):
                mod = mod.strip()
                if mod in ['u', 'underline']:
                    underline = True
                elif mod in ['upper', 'uppercase']:
                    uppercase = True
                    ann_txt = ann_txt.upper()
                elif mod in ['b', 'bold']:
                    b_flag = True
                elif mod in ['i', 'italic']:
                    i_flag = True
                elif mod.startswith('color:'):
                    color = mod.split(':', 1)[1].strip()
                elif mod.startswith('bg:'):
                    bg_color = mod.split(':', 1)[1].strip()

            spans.append(InlineSpan(
                text=ann_txt,
                bold=b_flag,
                italic=i_flag,
                underline=underline,
                uppercase=uppercase,
                color=color,
                bg_color=bg_color,
                is_instruction=default_instruction
            ))

        elif gd['bold']:
            b_inner = gd['b_txt']
            inner_spans = parse_inline_spans(b_inner, default_bold=True, default_italic=default_italic, default_instruction=default_instruction)
            spans.extend(inner_spans)

        elif gd['italic']:
            i_inner = gd['i_txt']
            inner_spans = parse_inline_spans(i_inner, default_bold=default_bold, default_italic=True, default_instruction=default_instruction)
            spans.extend(inner_spans)

        last_idx = end

    if last_idx < len(text):
        remaining = text[last_idx:]
        remaining = re.sub(r'\[\/?ins\]', '', remaining, flags=re.IGNORECASE)
        if remaining:
            spans.append(InlineSpan(text=remaining, bold=default_bold, italic=default_italic, is_instruction=default_instruction))

    return spans


class ULNParser:
    """Parser to convert Universal Layout Notation (ULN) text into ULNBlock objects."""

    @staticmethod
    def parse(uln_text: str) -> List[ULNBlock]:
        lines = uln_text.splitlines()
        blocks: List[ULNBlock] = []
        
        in_num = False
        num_lines: List[str] = []

        in_box = False
        box_lines: List[str] = []

        in_opt = False
        opt_lines: List[str] = []

        in_quote = False
        quote_lines: List[str] = []

        in_table = False
        table_lines: List[str] = []
        table_borderless = False

        in_pic_grid = False
        pic_grid_lines: List[str] = []

        for line in lines:
            trimmed = line.strip()

            # Handle multi-line [PIC_GRID] ... [/PIC_GRID]
            if trimmed.upper().startswith("[PIC_GRID"):
                in_pic_grid = True
                pic_grid_lines = []
                continue

            if in_pic_grid:
                if trimmed.upper() == "[/PIC_GRID]":
                    in_pic_grid = False
                    grid_children: List[ULNBlock] = []
                    for gline in pic_grid_lines:
                        gline_trim = gline.strip()
                        if not gline_trim:
                            continue
                        g_items = gline_trim.split('|')
                        for g_item in g_items:
                            g_clean = g_item.strip()
                            if g_clean:
                                pic_info = parse_pic_tag(g_clean)
                                g_spans = parse_inline_spans(g_clean)
                                grid_children.append(ULNBlock(tag="PIC_ITEM", content=g_clean, spans=g_spans, pic=pic_info))
                    
                    blocks.append(ULNBlock(tag="PIC_GRID", children=grid_children))
                    pic_grid_lines = []
                else:
                    pic_grid_lines.append(line)
                continue

            # Handle multi-line [TABLE] ... [/TABLE]
            if trimmed.upper().startswith("[TABLE") and (trimmed.endswith("]") or "TABLE" in trimmed.upper()):
                in_table = True
                table_lines = []
                table_borderless = "borderless" in trimmed.lower()
                continue

            if in_table:
                if trimmed.upper() == "[/TABLE]":
                    in_table = False
                    # Parse table rows
                    rows: List[ULNTableRow] = []
                    for tline in table_lines:
                        tline_trim = tline.strip()
                        if not tline_trim:
                            continue
                        
                        is_hdr = False
                        row_content = tline_trim

                        if tline_trim.upper().startswith("[TH]"):
                            is_hdr = True
                            row_content = tline_trim[4:].strip()
                        elif tline_trim.upper().startswith("[TR]"):
                            row_content = tline_trim[4:].strip()

                        cell_texts = row_content.split('|')
                        cells: List[ULNTableCell] = []
                        for ctxt in cell_texts:
                            c_clean = ctxt.strip()
                            c_spans = parse_inline_spans(c_clean, default_bold=is_hdr)
                            cells.append(ULNTableCell(content=c_clean, is_header=is_hdr, spans=c_spans))

                        rows.append(ULNTableRow(cells=cells, is_header=is_hdr))

                    tbl_data = ULNTableData(rows=rows, borderless=table_borderless)
                    blocks.append(ULNBlock(tag="TABLE", table_data=tbl_data))
                    table_lines = []
                else:
                    table_lines.append(line)
                continue

            # Handle multi-line [NUM] ... [/NUM]
            if trimmed.upper() == "[NUM]" or trimmed.upper().startswith("[NUM] ") or trimmed.upper().startswith("[NUM:"):
                if trimmed.upper() == "[NUM]":
                    in_num = True
                    num_lines = []
                    continue
                else:
                    rest = trimmed[5:].lstrip(": ").strip()
                    if rest.upper().endswith("[/NUM]"):
                        num_content = rest[:-6].strip()
                        child_blocks = ULNParser.parse(num_content)
                        blocks.append(ULNBlock(tag="NUM", children=child_blocks))
                        continue
                    else:
                        in_num = True
                        num_lines = [rest]
                        continue

            if in_num:
                if trimmed.upper() == "[/NUM]":
                    in_num = False
                    child_blocks = ULNParser.parse("\n".join(num_lines))
                    blocks.append(ULNBlock(tag="NUM", children=child_blocks))
                    num_lines = []
                else:
                    num_lines.append(line)
                continue

            # Handle multi-line [OPT] ... [/OPT]
            if trimmed.upper() == "[OPT]" or trimmed.upper().startswith("[OPT] ") or trimmed.upper().startswith("[OPT:"):
                if trimmed.upper() == "[OPT]":
                    in_opt = True
                    opt_lines = []
                    continue
                else:
                    rest = trimmed[5:].lstrip(": ").strip()
                    if rest.upper().endswith("[/OPT]"):
                        opt_content = rest[:-6].strip()
                        blocks.append(ULNBlock(tag="OPT", content=opt_content, spans=parse_inline_spans(opt_content)))
                        continue
                    else:
                        in_opt = True
                        opt_lines = [rest]
                        continue
            
            if in_opt:
                if trimmed.upper() == "[/OPT]":
                    in_opt = False
                    opt_content = "\n".join(opt_lines)
                    blocks.append(ULNBlock(tag="OPT", content=opt_content, spans=parse_inline_spans(opt_content)))
                    opt_lines = []
                else:
                    opt_lines.append(line)
                continue

            # Handle multi-line [BOX] ... [/BOX] or [WORDBANK] ... [/WORDBANK]
            is_box_start = (
                trimmed.upper() == "[BOX]" or trimmed.upper().startswith("[BOX] ") or trimmed.upper().startswith("[BOX:") or
                trimmed.upper() == "[WORDBANK]" or trimmed.upper().startswith("[WORDBANK] ") or trimmed.upper().startswith("[WORDBANK:")
            )
            if is_box_start:
                end_bracket_idx = trimmed.find("]")
                box_tag = trimmed[1:end_bracket_idx].strip() if end_bracket_idx != -1 else "BOX"
                rest = trimmed[end_bracket_idx + 1:].strip() if end_bracket_idx != -1 else ""

                if not rest:
                    in_box = True
                    box_lines = []
                    continue
                else:
                    if rest.upper().endswith("[/BOX]"):
                        box_content = rest[:-6].strip()
                        blocks.append(ULNBlock(tag=box_tag, content=box_content, spans=parse_inline_spans(box_content)))
                        continue
                    elif rest.upper().endswith("[/WORDBANK]"):
                        box_content = rest[:-11].strip()
                        blocks.append(ULNBlock(tag=box_tag, content=box_content, spans=parse_inline_spans(box_content)))
                        continue
                    else:
                        in_box = True
                        box_lines = [rest]
                        continue
            
            if in_box:
                if trimmed.upper() in ("[/BOX]", "[/WORDBANK]"):
                    in_box = False
                    box_content = "\n".join(box_lines)
                    blocks.append(ULNBlock(tag="BOX", content=box_content, spans=parse_inline_spans(box_content)))
                    box_lines = []
                else:
                    box_lines.append(line)
                continue

            # Handle multi-line [QUOTE] ... [/QUOTE]
            if trimmed.upper() == "[QUOTE]" or trimmed.upper().startswith("[QUOTE] "):
                if trimmed.upper() == "[QUOTE]":
                    in_quote = True
                    quote_lines = []
                    continue
                else:
                    rest = trimmed[7:].strip()
                    if rest.upper().endswith("[/QUOTE]"):
                        raw_q_text = rest[:-8].strip()
                        raw_paras = re.split(r'\n\s*\n+', raw_q_text)
                        for p_txt in raw_paras:
                            clean_p = " ".join(p_txt.split())
                            if clean_p.strip():
                                blocks.append(ULNBlock(tag="QUOTE", content=clean_p, spans=parse_inline_spans(clean_p)))
                        continue
                    else:
                        in_quote = True
                        quote_lines = [rest]
                        continue

            if in_quote:
                if trimmed.upper() == "[/QUOTE]":
                    in_quote = False
                    raw_q_text = "\n".join(quote_lines)
                    raw_paras = re.split(r'\n\s*\n+', raw_q_text.strip())
                    for p_txt in raw_paras:
                        clean_p = " ".join(p_txt.split())
                        if clean_p.strip():
                            blocks.append(ULNBlock(tag="QUOTE", content=clean_p, spans=parse_inline_spans(clean_p)))
                    quote_lines = []
                else:
                    quote_lines.append(line)
                continue

            if not trimmed:
                continue

            tag_match = re.match(r'^\[(H1|H2|H3|H4|H5|H6|P0|P1|P2|TAB2|PIC|INS)\]\s*(.*)$', trimmed, re.IGNORECASE)

            if tag_match:
                tag = tag_match.group(1).upper()
                rest_content = tag_match.group(2)

                if tag == "INS":
                    tag = "P0"
                    spans = parse_inline_spans(rest_content, default_instruction=True)
                    pic_info = parse_pic_tag(rest_content)
                    blocks.append(ULNBlock(tag=tag, content=rest_content, spans=spans, pic=pic_info, is_instruction=True))

                elif tag == "TAB2":
                    pic_start_match = re.match(r'^\s*(\[PIC:[^\]]+\])\s+(.*)$', rest_content, re.IGNORECASE)
                    if pic_start_match:
                        col1_raw = pic_start_match.group(1).strip()
                        col2_raw = pic_start_match.group(2).strip()
                    elif '\t' in rest_content:
                        parts = rest_content.split('\t', 1)
                        col1_raw = parts[0].strip() if len(parts) > 0 else ""
                        col2_raw = parts[1].strip() if len(parts) > 1 else ""
                    elif ' | ' in rest_content:
                        parts = rest_content.split(' | ', 1)
                        col1_raw = parts[0].strip() if len(parts) > 0 else ""
                        col2_raw = parts[1].strip() if len(parts) > 1 else ""
                    else:
                        parts = re.split(r'\s{2,}', rest_content, 1)
                        col1_raw = parts[0].strip() if len(parts) > 0 else ""
                        col2_raw = parts[1].strip() if len(parts) > 1 else ""

                    c1_tag_match = re.match(r'^\s*\[(P0|P1|P2)\]\s*(.*)$', col1_raw, re.IGNORECASE)
                    if c1_tag_match:
                        col1_raw = c1_tag_match.group(2)

                    c2_tag_match = re.match(r'^\s*\[(P0|P1|P2)\]\s*(.*)$', col2_raw, re.IGNORECASE)
                    if c2_tag_match:
                        col2_raw = c2_tag_match.group(2)

                    col1_spans = parse_inline_spans(col1_raw)
                    col2_spans = parse_inline_spans(col2_raw)

                    pic_info = parse_pic_tag(col2_raw) or parse_pic_tag(col1_raw)

                    blocks.append(ULNBlock(
                        tag="TAB2",
                        col1=col1_raw,
                        col2=col2_raw,
                        col1_spans=col1_spans,
                        col2_spans=col2_spans,
                        pic=pic_info
                    ))

                elif tag == "PIC":
                    pic_info = parse_pic_tag(f"[PIC] {rest_content}") or parse_pic_tag(rest_content)
                    blocks.append(ULNBlock(tag="PIC", pic=pic_info))

                else:
                    spans = parse_inline_spans(rest_content)
                    pic_info = parse_pic_tag(rest_content)
                    is_ins = any(s.is_instruction for s in spans) or bool(re.search(r'\[ins\]', rest_content, re.IGNORECASE))
                    blocks.append(ULNBlock(tag=tag, content=rest_content, spans=spans, pic=pic_info, is_instruction=is_ins))

            else:
                pic_info = parse_pic_tag(trimmed)
                if pic_info and trimmed.upper().startswith("[PIC:"):
                    blocks.append(ULNBlock(tag="PIC", pic=pic_info))
                else:
                    spans = parse_inline_spans(trimmed)
                    is_ins = any(s.is_instruction for s in spans) or bool(re.search(r'\[ins\]', trimmed, re.IGNORECASE))
                    blocks.append(ULNBlock(tag="P0", content=trimmed, spans=spans, pic=pic_info, is_instruction=is_ins))

        # Merge empty P0/P1 question-number-only blocks preceding OPT blocks or Dialogue lines
        merged_blocks: List[ULNBlock] = []
        i = 0
        while i < len(blocks):
            curr = blocks[i]
            if i + 1 < len(blocks) and curr.tag in ["P0", "P1"]:
                num_only_match = re.match(r'^\s*(?:\*\*)?(?:(?:Question|Câu|Task|Exercise|Ex|Activity)\s+)?#(\d+)[\.\)\:\-]?\s*(?:\*\*)?[:\.\)]?\s*$', curr.content, re.IGNORECASE)
                if num_only_match:
                    nxt = blocks[i + 1]
                    if nxt.tag == "OPT":
                        nxt.content = f"{curr.content.strip()} {nxt.content}"
                        nxt.spans = parse_inline_spans(nxt.content)
                        merged_blocks.append(nxt)
                        i += 2
                        continue
                    elif nxt.tag in ["P0", "P1"]:
                        # Check if next block starts with dialogue/speaker pattern e.g. A:, Speaker A:, Q:
                        if re.match(r'^\s*(?:(?:\*\*|\*|\[)?(?:Speaker\s+)?[A-Za-z0-9]+\s*[:\.\-](?:\*\*|\*|\])?)\s*', nxt.content, re.IGNORECASE):
                            nxt.tag = "P0"
                            nxt.content = f"{curr.content.strip()} {nxt.content}"
                            nxt.spans = parse_inline_spans(nxt.content)
                            merged_blocks.append(nxt)
                            i += 2
                            continue

            merged_blocks.append(curr)
            i += 1




        return merged_blocks
