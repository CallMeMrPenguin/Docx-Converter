import re
from typing import List, Optional
from core.models import ULNBlock, InlineSpan, PicInfo, ULNTableRow, ULNTableData, ULNTableCell


def parse_pic_tag(text: str) -> Optional[PicInfo]:
    """Parses [PIC] or [PIC: "description" | pos:center | size:medium]"""
    if not text:
        return None
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


def parse_inline_spans(
    text: str,
    default_bold: bool = False,
    default_italic: bool = False,
    default_instruction: bool = False,
    _depth: int = 0
) -> List[InlineSpan]:
    """Parses inline formatting elements recursively, handling nested tags like **bold *italic*** and [ins]."""
    if not text or _depth > 8:
        return [InlineSpan(text=text, bold=default_bold, italic=default_italic, is_instruction=default_instruction)] if text else []

    # Handle unclosed [ins] tag at start of text
    if re.match(r'^\s*\[ins\]', text, re.IGNORECASE) and not re.search(r'\[/ins\]', text, re.IGNORECASE):
        cleaned_text = re.sub(r'^\s*\[ins\]\s*', '', text, flags=re.IGNORECASE)
        return parse_inline_spans(
            cleaned_text,
            default_bold=default_bold,
            default_italic=default_italic,
            default_instruction=True,
            _depth=_depth + 1
        )

    pattern = re.compile(
        r'(?P<pic>\[PIC(?::[^\]]+)?\])|'
        r'(?P<ins>\[ins\](?P<ins_txt>.*?)\[/ins\])|'
        r'(?P<annot>\[(?P<ann_txt>[^\]]+)\]\{(?P<ann_mod>[^\}]+)\})|'
        r'(?P<bold_italic>\*\*\*(?P<bi_txt>.*?)\*\*\*)|'
        r'(?P<bold>\*\*(?P<b_txt>.*?)\*\*)|'
        r'(?P<italic>\*(?P<i_txt>.*?)\*)',
        re.IGNORECASE | re.DOTALL
    )

    spans: List[InlineSpan] = []
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
            spans.extend(parse_inline_spans(ins_inner, default_bold=default_bold, default_italic=default_italic, default_instruction=True, _depth=_depth + 1))

        elif gd['bold_italic']:
            bi_inner = gd['bi_txt']
            spans.extend(parse_inline_spans(bi_inner, default_bold=True, default_italic=True, default_instruction=default_instruction, _depth=_depth + 1))

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
            spans.extend(parse_inline_spans(b_inner, default_bold=True, default_italic=default_italic, default_instruction=default_instruction, _depth=_depth + 1))

        elif gd['italic']:
            i_inner = gd['i_txt']
            spans.extend(parse_inline_spans(i_inner, default_bold=default_bold, default_italic=True, default_instruction=default_instruction, _depth=_depth + 1))

        last_idx = end

    if last_idx < len(text):
        remaining = text[last_idx:]
        remaining = re.sub(r'\[\/?ins\]', '', remaining, flags=re.IGNORECASE)
        if remaining:
            spans.append(InlineSpan(text=remaining, bold=default_bold, italic=default_italic, is_instruction=default_instruction))

    return spans


def split_bracket_safe_columns(text: str) -> List[str]:
    """Splits a multi-column string on '|' while ignoring '|' characters inside brackets like [PIC: ... | ...]."""
    if not text:
        return []

    # Check if there is any tab character
    if '\t' in text:
        return [p.strip() for p in text.split('\t') if p.strip()]

    parts = []
    current = []
    bracket_level = 0

    for ch in text:
        if ch == '[':
            bracket_level += 1
            current.append(ch)
        elif ch == ']':
            if bracket_level > 0:
                bracket_level -= 1
            current.append(ch)
        elif ch == '|' and bracket_level == 0:
            token = "".join(current).strip()
            if token:
                parts.append(token)
            current = []
        else:
            current.append(ch)

    last_token = "".join(current).strip()
    if last_token:
        parts.append(last_token)

    # Fallback to multi-space split if no top-level pipe was found
    if len(parts) <= 1 and not ('|' in text):
        space_split = [p.strip() for p in re.split(r'\s{3,}', text) if p.strip()]
        if len(space_split) > 1:
            return space_split

    return parts if parts else [text.strip()]


class ULNParser:
    """Parser to convert Universal Layout Notation (ULN) text into ULNBlock objects."""

    @staticmethod
    def parse(uln_text: str) -> List[ULNBlock]:
        lines = uln_text.splitlines()
        blocks: List[ULNBlock] = []

        in_num = False
        num_lines: List[str] = []

        in_box = False
        box_tag = "BOX"
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

        def flush_table(tbl_lines: List[str], borderless: bool):
            rows: List[ULNTableRow] = []
            for tline in tbl_lines:
                tline_trim = tline.strip()
                if not tline_trim:
                    continue
                is_hdr = False
                row_content = tline_trim
                if tline_trim.upper().startswith("[TH]"):
                    is_hdr = True
                    row_content = tline_trim[4:].strip()
                elif tline_trim.upper().startswith("[TR]") or tline_trim.upper().startswith("[TD]"):
                    row_content = tline_trim[4:].strip()
                row_content = re.sub(r'\[\/(?:TH|TR|TD)\]\s*$', '', row_content, flags=re.IGNORECASE).strip()
                cell_texts = split_bracket_safe_columns(row_content)
                cells: List[ULNTableCell] = []
                for ctxt in cell_texts:
                    c_clean = ctxt.strip()
                    c_clean = re.sub(r'^\s*\[(?:TH|TR|TD)\]\s*', '', c_clean, flags=re.IGNORECASE)
                    c_clean = re.sub(r'\s*\[\/(?:TH|TR|TD)\]\s*$', '', c_clean, flags=re.IGNORECASE).strip()
                    c_spans = parse_inline_spans(c_clean, default_bold=is_hdr)
                    cells.append(ULNTableCell(content=c_clean, is_header=is_hdr, spans=c_spans))
                rows.append(ULNTableRow(cells=cells, is_header=is_hdr))
            tbl_data = ULNTableData(rows=rows, borderless=borderless)
            blocks.append(ULNBlock(tag="TABLE", table_data=tbl_data))

        def flush_pic_grid(grid_lines: List[str]):
            grid_children: List[ULNBlock] = []
            for gline in grid_lines:
                gline_trim = gline.strip()
                if not gline_trim:
                    continue
                g_items = split_bracket_safe_columns(gline_trim)
                for g_item in g_items:
                    g_clean = g_item.strip()
                    if g_clean:
                        pic_info = parse_pic_tag(g_clean)
                        g_spans = parse_inline_spans(g_clean)
                        grid_children.append(ULNBlock(tag="PIC_ITEM", content=g_clean, spans=g_spans, pic=pic_info))
            blocks.append(ULNBlock(tag="PIC_GRID", children=grid_children))

        def flush_quote(q_lines: List[str]):
            for q_l in q_lines:
                clean_l = q_l.strip()
                if clean_l:
                    m_sub = re.match(r'^\s*\[(P0|P1|P2|INS)\]\s*(.*)$', clean_l, flags=re.IGNORECASE)
                    if m_sub:
                        s_tag = m_sub.group(1).upper()
                        body_c = m_sub.group(2).strip()
                        blocks.append(ULNBlock(tag="QUOTE", content=body_c, sub_tag=s_tag, spans=parse_inline_spans(body_c)))
                    else:
                        blocks.append(ULNBlock(tag="QUOTE", content=clean_l, sub_tag=None, spans=parse_inline_spans(clean_l)))

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
                    flush_pic_grid(pic_grid_lines)
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
                    flush_table(table_lines, table_borderless)
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

            # Handle multi-line or single-line [OPT] ... [/OPT]
            if trimmed.upper() == "[OPT]" or trimmed.upper().startswith("[OPT] ") or trimmed.upper().startswith("[OPT:"):
                if trimmed.upper() == "[OPT]":
                    in_opt = True
                    opt_lines = []
                    continue
                else:
                    rest = trimmed[5:].lstrip(": ").strip()
                    if "[/OPT]" in rest.upper():
                        opt_content = re.sub(r'\[\/OPT\]', '', rest, flags=re.IGNORECASE).strip()
                        blocks.append(ULNBlock(tag="OPT", content=opt_content, spans=parse_inline_spans(opt_content)))
                        continue
                    else:
                        blocks.append(ULNBlock(tag="OPT", content=rest, spans=parse_inline_spans(rest)))
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

            # Standalone container markers
            if trimmed.upper() in ("[NUM]", "[/NUM]"):
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
                    blocks.append(ULNBlock(tag=box_tag, content=box_content, spans=parse_inline_spans(box_content)))
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
                        flush_quote(raw_q_text.splitlines())
                        continue
                    else:
                        in_quote = True
                        quote_lines = [rest]
                        continue

            if in_quote:
                if trimmed.upper() == "[/QUOTE]":
                    in_quote = False
                    flush_quote(quote_lines)
                    quote_lines = []
                else:
                    quote_lines.append(line)
                continue

            if not trimmed:
                continue

            tag_match = re.match(r'^\[(H1|H2|H3|H4|H5|H6|P0|P1|P2|TAB\d*|PIC|INS)\]\s*(.*)$', trimmed, re.IGNORECASE)

            if tag_match:
                tag = tag_match.group(1).upper()
                rest_content = tag_match.group(2)

                if tag == "INS":
                    tag = "P0"
                    spans = parse_inline_spans(rest_content, default_instruction=True)
                    pic_info = parse_pic_tag(rest_content)
                    blocks.append(ULNBlock(tag=tag, content=rest_content, spans=spans, pic=pic_info, is_instruction=True))

                elif tag.startswith("TAB"):
                    parts = split_bracket_safe_columns(rest_content)
                    if len(parts) < 2:
                        parts = [rest_content.strip(), ""]

                    clean_cols = []
                    cols_spans = []
                    for part in parts:
                        c_match = re.match(r'^\s*\[(P0|P1|P2)\]\s*(.*)$', part, re.IGNORECASE)
                        c_text = c_match.group(2) if c_match else part
                        clean_cols.append(c_text)
                        cols_spans.append(parse_inline_spans(c_text))

                    num_cols = len(clean_cols)
                    actual_tag = f"TAB{num_cols}" if tag in ("TAB", "TAB2") and num_cols > 2 else tag
                    if actual_tag == "TAB":
                        actual_tag = f"TAB{num_cols}"

                    col1_raw = clean_cols[0] if len(clean_cols) > 0 else ""
                    col2_raw = clean_cols[1] if len(clean_cols) > 1 else ""
                    col1_spans = cols_spans[0] if len(cols_spans) > 0 else []
                    col2_spans = cols_spans[1] if len(cols_spans) > 1 else []

                    pic_info = None
                    for c_raw in clean_cols:
                        pic_info = parse_pic_tag(c_raw)
                        if pic_info:
                            break

                    blocks.append(ULNBlock(
                        tag=actual_tag,
                        cols=clean_cols,
                        cols_spans=cols_spans,
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

        # ── CRITICAL RECOVERY FOR UNCLOSED CONTAINERS AT END OF FILE ──
        if in_pic_grid and pic_grid_lines:
            flush_pic_grid(pic_grid_lines)
        if in_table and table_lines:
            flush_table(table_lines, table_borderless)
        if in_box and box_lines:
            box_content = "\n".join(box_lines)
            blocks.append(ULNBlock(tag=box_tag, content=box_content, spans=parse_inline_spans(box_content)))
        if in_opt and opt_lines:
            opt_content = "\n".join(opt_lines)
            blocks.append(ULNBlock(tag="OPT", content=opt_content, spans=parse_inline_spans(opt_content)))
        if in_quote and quote_lines:
            flush_quote(quote_lines)
        if in_num and num_lines:
            child_blocks = ULNParser.parse("\n".join(num_lines))
            blocks.append(ULNBlock(tag="NUM", children=child_blocks))

        # Merge empty question-number-only blocks preceding OPT or dialogue lines
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
                    elif nxt.tag == "P0":
                        if re.match(r'^\s*(?:(?:\*\*|\*|\[)?(?:Speaker\s+[A-Za-z0-9]+|[A-Z][a-zA-Z0-9]*)\s*:(?:\*\*|\*|\])?)\s*', nxt.content, re.IGNORECASE):
                            nxt.tag = "P0"
                            nxt.content = f"{curr.content.strip()} {nxt.content}"
                            nxt.spans = parse_inline_spans(nxt.content)
                            merged_blocks.append(nxt)
                            i += 2
                            continue

            merged_blocks.append(curr)
            i += 1

        return merged_blocks
