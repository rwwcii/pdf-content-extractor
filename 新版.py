
import os
import re
import sys
import yaml
import io
import unicodedata
from pathlib import Path
from collections import Counter
import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ==================== 字体工具 ====================
def set_run_font(run, font_name="宋体", font_size=10.5, bold=False, italic=False):
    """设置 run 的字体、大小、加粗、斜体（中文字体也同步设置）"""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        from lxml import etree
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), font_name)

# ==================== 表格/线条区域检测 ====================
def get_table_areas(page, min_length=25, merge_gap=20):
    """
    通过页面中的线条（drawings）检测表格区域。
    返回：list of fitz.Rect，每个矩形代表一个表格区域。
    """
    lines = []
    for d in page.get_drawings():
        for item in d.get("items", []):
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            x0, y0 = p1.x, p1.y
            x1, y1 = p2.x, p2.y
            length = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
            if length < min_length:
                continue
            is_h = abs(y1 - y0) <= 2
            is_v = abs(x1 - x0) <= 2
            if not (is_h or is_v):
                continue
            rect = fitz.Rect(min(x0, x1) - 2, min(y0, y1) - 2,
                             max(x0, x1) + 2, max(y0, y1) + 2)
            lines.append({"rect": rect, "type": "h" if is_h else "v"})

    areas = []
    for line in lines:
        rect = line["rect"]
        added = False
        for area in areas:
            expanded = fitz.Rect(area["rect"].x0 - merge_gap, area["rect"].y0 - merge_gap,
                                 area["rect"].x1 + merge_gap, area["rect"].y1 + merge_gap)
            if expanded.intersects(rect):
                area["rect"].x0 = min(area["rect"].x0, rect.x0)
                area["rect"].y0 = min(area["rect"].y0, rect.y0)
                area["rect"].x1 = max(area["rect"].x1, rect.x1)
                area["rect"].y1 = max(area["rect"].y1, rect.y1)
                area["types"].append(line["type"])
                added = True
                break
        if not added:
            areas.append({"rect": fitz.Rect(rect), "types": [line["type"]]})

    table_areas = []
    for area in areas:
        # 表格通常包含至少2条横线和2条竖线
        if area["types"].count("h") >= 2 and area["types"].count("v") >= 2:
            r = area["rect"]
            # 适当外扩以便完整截图
            table_areas.append(fitz.Rect(
                max(0, r.x0 - 10), max(0, r.y0 - 10),
                min(page.rect.width, r.x1 + 10),
                min(page.rect.height, r.y1 + 10)
            ))
    return table_areas


# ==================== 参考代码表格识别辅助函数 ====================
def rect_overlap_ratio(rect1, rect2):
    inter = rect1 & rect2
    if inter.is_empty:
        return 0
    area1 = rect1.get_area()
    if area1 == 0:
        return 0
    return inter.get_area() / area1


def is_table_like_line(line):
    """
    参考代码逻辑：
    一行中如果多个词之间存在明显空隙，认为它可能是表格式排版。
    """
    words = line.get("words", [])
    if len(words) < 2:
        return False

    words = sorted(words, key=lambda w: w["x0"])
    big_gap_count = 0

    for i in range(1, len(words)):
        gap = words[i]["x0"] - words[i - 1]["x1"]
        if gap > 8:
            big_gap_count += 1

    return big_gap_count >= 1


def is_layout_special_line(line, default_left, default_right, gap_threshold=18):
    """
    参考代码逻辑：
    通过左右边距是否异常 + 行内是否存在大间隔，识别特殊排版区域。
    """
    words = line.get("words", [])

    if len(words) < 2:
        return False

    words = sorted(words, key=lambda w: w["x0"])

    left_blank = round(line["x0"] / 10) * 10
    right_blank = round((line["page_width"] - line["x1"]) / 10) * 10

    left_normal = abs(left_blank - default_left) <= 30
    right_normal = abs(right_blank - default_right) <= 30

    left_abnormal = not left_normal
    right_abnormal = not right_normal

    has_left_content = False
    has_right_content = False

    for i in range(1, len(words)):
        gap = words[i]["x0"] - words[i - 1]["x1"]

        if gap > gap_threshold:
            has_left_content = True
            has_right_content = True
            break

    if not (has_left_content and has_right_content):
        return False

    if left_normal and right_abnormal:
        return True

    if left_abnormal and right_normal:
        return True

    if left_abnormal and right_abnormal:
        return True

    return False

# ==================== PDF 文本行提取 ====================
def group_words_into_lines(words, page_width, y_tolerance=4):
    lines = []

    for word in words:
        placed = False

        for line in lines:
            if abs(line["y"] - word["y0"]) <= y_tolerance:
                line["words"].append(word)
                line["ys"].append(word["y0"])
                placed = True
                break

        if not placed:
            lines.append({
                "y": word["y0"],
                "ys": [word["y0"]],
                "words": [word]
            })

    result_lines = []

    for line in lines:
        line_words = sorted(line["words"], key=lambda x: x["x0"])

        text_parts = []
        previous_x1 = None

        for w in line_words:
            if previous_x1 is None:
                text_parts.append(w["text"])
            else:
                gap = w["x0"] - previous_x1
                if gap > 20:
                    text_parts.append("    " + w["text"])
                else:
                    text_parts.append(" " + w["text"])

            previous_x1 = w["x1"]

        text = "".join(text_parts).strip()

        result_lines.append({
            "text": text,
            "y": sum(line["ys"]) / len(line["ys"]),
            "y1": max(w["y1"] for w in line_words),
            "x0": min(w["x0"] for w in line_words),
            "x1": max(w["x1"] for w in line_words),
            "page_width": page_width,
            "words": line_words
        })

    result_lines.sort(key=lambda x: x["y"])
    return result_lines

def extract_lines(pdf_path):
    doc = fitz.open(pdf_path)
    all_lines = []

    for page_num, page in enumerate(doc):
        words = page.get_text("words")
        if not words:
            continue

        selected_words = []

        for w in words:
            x0, y0, x1, y1, text = w[:5]

            if y0 < 60 or y1 > page.rect.height - 60:
                continue

            text = str(text).strip()
            if not text:
                continue

            selected_words.append({
                "x0": x0,
                "y0": y0,
                "x1": x1,
                "y1": y1,
                "text": text
            })

        if not selected_words:
            continue

        selected_words.sort(key=lambda x: (x["y0"], x["x0"]))

        lines = group_words_into_lines(
            selected_words,
            page_width=page.rect.width
        )

        for line in lines:
            line["page_index"] = page_num
            line["y0"] = line["y"]
            line["y1"] = line.get("y1", line["y"] + 5)

        all_lines.extend(lines)

    doc.close()
    all_lines.sort(key=lambda l: (l["page_index"], l["y"], l["x0"]))
    return all_lines
def make_clip_rect(page, table_lines):
    x0 = min(line["x0"] for line in table_lines)
    x1 = max(line["x1"] for line in table_lines)
    y0 = min(line["y"] for line in table_lines)
    y1 = max(line.get("y1", line["y"]) for line in table_lines)

    return fitz.Rect(
        max(0, x0 - 40),
        max(0, y0 - 8),
        min(page.rect.width, x1 + 30),
        min(page.rect.height, y1 + 7)
    )


def line_in_block(line, block):
    if line["page_index"] != block["page_index"]:
        return False

    line_rect = fitz.Rect(
        line["x0"],
        line["y"],
        line["x1"],
        line.get("y1", line["y"] + 5)
    )

    return rect_overlap_ratio(line_rect, block["clip"]) > 0.02
# ==================== 智能段落合并（含编号行处理） ====================
import re
from collections import Counter

import re
from collections import Counter


def merge_lines(lines):
    if not lines:
        return []

    right_blanks = []
    left_blanks = []

    for line in lines:
        text = line["text"].strip()
        if not text:
            continue
        if re.fullmatch(r"<<TABLE_IMG_\d+>>", text):
            continue

        right_blank = line["page_width"] - line["x1"]
        right_blanks.append(round(right_blank / 10) * 10)
        left_blanks.append(line["x0"])

    default_right_blank = Counter(right_blanks).most_common(1)[0][0] if right_blanks else 0
    default_left = min(left_blanks) if left_blanks else 0

    right_tolerance = 8
    left_tolerance = 6

    paragraphs = []
    current_para = ""

    def is_table_marker(text):
        return bool(re.fullmatch(r"<<TABLE_IMG_\d+>>", text))

    def is_field_line(text):
        return bool(re.match(r'^[\u4e00-\u9fffA-Za-z（）()]{2,30}[:：].*$', text))

    def is_formula_line(text):
        return bool(re.match(r'^[^\s]{1,30}[=＝].+', text))

    def is_circled_number_title(text):
        return bool(re.match(r'^\s*[①②③④⑤⑥⑦⑧⑨⑩]', text))

    def is_short_number_title(text):
        return bool(
            re.fullmatch(r'\s*\d+[、.．]\s*[^。；！？，,]{1,35}', text)
            and len(text) <= 40
        )

    def is_numbered_para_start(text):
        return bool(re.match(r'^\s*\d+[、.．]\s*\S+', text))

    def is_cn_subsection_start(text):
        return bool(re.match(r'^\s*[（(]\s*[一二三四五六七八九十]+\s*[）)]', text))

    def is_other_heading(text):
        patterns = [
            r'^\s*[（(]\s*\d+\s*[）)]',
            r'^\s*第[一二三四五六七八九十百]+[章节]',
        ]
        return any(re.match(p, text) for p in patterns)

    def is_sentence_end(text):
        return bool(re.search(r'[。！？；;】]$', text))

    def join_text(a, b):
        if not a:
            return b
        if not b:
            return a

        if re.search(r'[\u4e00-\u9fff]$', a):
            return a + b

        if re.match(r'^[，。；：！？、,.!?;:）)]', b):
            return a + b

        return a + " " + b

    for line in lines:
        text = line["text"].strip()
        if not text:
            continue

        if is_table_marker(text):
            if current_para:
                paragraphs.append(current_para)
                current_para = ""
            paragraphs.append(text)
            continue

        if is_field_line(text):
            if current_para:
                paragraphs.append(current_para)
                current_para = ""
            paragraphs.append(text)
            continue

        if is_formula_line(text):
            if current_para:
                paragraphs.append(current_para)
                current_para = ""
            paragraphs.append(text)
            continue

        # 圈圈编号标题独立，例如：① xxx、② xxx
        if is_circled_number_title(text):
            if current_para:
                paragraphs.append(current_para)
                current_para = ""
            paragraphs.append(text)
            continue

        if is_short_number_title(text):
            if current_para:
                paragraphs.append(current_para)
                current_para = ""
            paragraphs.append(text)
            continue

        if is_numbered_para_start(text) or is_cn_subsection_start(text) or is_other_heading(text):
            if current_para:
                paragraphs.append(current_para)
            current_para = text
            continue

        right_blank = line["page_width"] - line["x1"]
        left_blank = line["x0"]
        right_key = round(right_blank / 10) * 10

        right_equal_default = abs(right_key - default_right_blank) <= right_tolerance
        left_equal_min = abs(left_blank - default_left) <= left_tolerance

        if right_equal_default and left_equal_min:
            current_para = join_text(current_para, text)
            continue

        if right_equal_default and not left_equal_min:
            if current_para:
                paragraphs.append(current_para)
            current_para = text
            continue

        if not right_equal_default:
            current_para = join_text(current_para, text)

            if is_sentence_end(text):
                paragraphs.append(current_para)
                current_para = ""

            continue

    if current_para:
        paragraphs.append(current_para)

    return paragraphs
# ==================== 表格占位符与图片信息提取 ====================
def get_default_margins(lines):
    left_keys = []
    right_keys = []

    for line in lines:
        if not line["text"].strip():
            continue

        left_keys.append(round(line["x0"] / 10) * 10)
        right_keys.append(round((line["page_width"] - line["x1"]) / 10) * 10)

    default_left = max(set(left_keys), key=left_keys.count) if left_keys else 0
    default_right = max(set(right_keys), key=right_keys.count) if right_keys else 0

    return default_left, default_right

def get_table_area_for_line(line, table_areas):
    line_rect = fitz.Rect(
        line["x0"],
        line["y"],
        line["x1"],
        line.get("y1", line["y"] + 5)
    )

    for area in table_areas:
        if area.intersects(line_rect):
            return area

    return None

def read_pdf_with_tables(pdf_path):
    lines = extract_lines(pdf_path)

    if not lines:
        return "", []

    pdf_doc = fitz.open(pdf_path)

    page_table_areas = {}
    for page in pdf_doc:
        page_table_areas[page.number] = get_table_areas(page)

    default_left, default_right = get_default_margins(lines)

    blocks = []
    used_table_clips = []

    i = 0

    while i < len(lines):
        line = lines[i]
        page_index = line["page_index"]

        table_area = get_table_area_for_line(
            line,
            page_table_areas.get(page_index, [])
        )

        if table_area:
            duplicated = False

            for used in used_table_clips:
                if used["page_index"] == page_index and used["clip"].intersects(table_area):
                    duplicated = True
                    break

            if not duplicated:
                blocks.append({
                    "page_index": page_index,
                    "start_index": i,
                    "end_index": i + 1,
                    "clip": table_area
                })

                used_table_clips.append({
                    "page_index": page_index,
                    "clip": table_area
                })

            i += 1
            continue

        if is_table_like_line(line) or is_layout_special_line(line, default_left, default_right):
            table_lines = [line]
            last_y = line["y"]
            j = i + 1

            while j < len(lines):
                next_line = lines[j]

                if next_line.get("page_index") != page_index:
                    break

                y_close = abs(next_line["y"] - last_y) <= 28

                next_table_area = get_table_area_for_line(
                    next_line,
                    page_table_areas.get(page_index, [])
                )

                same_kind = (
                    next_table_area
                    or is_table_like_line(next_line)
                    or is_layout_special_line(next_line, default_left, default_right)
                )

                if y_close and same_kind:
                    table_lines.append(next_line)
                    last_y = next_line["y"]
                    j += 1
                else:
                    break

            page = pdf_doc[page_index]
            clip = make_clip_rect(page, table_lines)

            blocks.append({
                "page_index": page_index,
                "start_index": i,
                "end_index": j,
                "clip": clip
            })

            i = j
        else:
            i += 1

    merged_blocks = []

    for block in blocks:
        if not merged_blocks:
            merged_blocks.append(block)
            continue

        last = merged_blocks[-1]
        same_page = block["page_index"] == last["page_index"]
        vertical_gap = block["clip"].y0 - last["clip"].y1

        if same_page and vertical_gap <= 25:
            last["end_index"] = max(last["end_index"], block["end_index"])
            last["clip"] = fitz.Rect(
                min(last["clip"].x0, block["clip"].x0),
                min(last["clip"].y0, block["clip"].y0),
                max(last["clip"].x1, block["clip"].x1),
                max(last["clip"].y1, block["clip"].y1)
            )
        else:
            merged_blocks.append(block)

    blocks = merged_blocks

    processed_lines = []
    table_images = []
    img_counter = 0

    i = 0

    while i < len(lines):
        line = lines[i]

        block_start = None

        for block in blocks:
            if block["start_index"] == i:
                block_start = block
                break

        if block_start:
            marker = f"<<TABLE_IMG_{img_counter}>>"

            processed_lines.append({
                "text": marker,
                "x0": line["x0"],
                "x1": line["x1"],
                "y": line["y"],
                "y0": line["y0"],
                "y1": line["y1"],
                "page_width": line["page_width"],
                "page_index": block_start["page_index"],
                "words": []
            })

            table_images.append((
                block_start["page_index"],
                block_start["clip"],
                marker
            ))

            img_counter += 1
            i = block_start["end_index"]
            continue

        if any(line_in_block(line, block) for block in blocks):
            i += 1
            continue

        processed_lines.append(line)
        i += 1

    pdf_doc.close()

    paragraphs = merge_lines(processed_lines)
    merged_text = "\n".join(paragraphs)

    return merged_text, table_images

def read_pdf_smart(pdf_path):
    """兼容旧接口，只返回文本"""
    txt, _ = read_pdf_with_tables(pdf_path)
    return txt

# ==================== 读合同（PDF/DOCX） ====================
def read_pdf(path):
    return read_pdf_smart(path)

def read_contract(path):
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        return read_pdf(path)
    else:
        raise ValueError("不支持的文件类型（仅支持 .pdf）")

# ==================== 文本清洗 ====================
def clean_text(text):
    text = text.replace("\r", "\n")
    # 删除独立页码行
    text = re.sub(r'\n\s*\d{1,4}\s*\n', '\n', text)
    # 删除目录行（大量点+数字）
    text = re.sub(r'^.*?\.{5,}\s*\d+\s*$', '', text, flags=re.M)
    # 合并连续多个空行
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

# ==================== 字符归一化 ====================
def normalize(text):
    return unicodedata.normalize('NFKC', text)

# ==================== 核心搜索 ====================
def robust_search(text, keywords):
    norm_text = normalize(text)
    best_start = len(text)
    best_end = len(text)
    for kw in keywords:
        norm_kw = normalize(kw)
        compact = re.sub(r'\s+', '', norm_kw)
        pattern = r'\s*'.join(re.escape(c) for c in compact)
        match = re.search(pattern, norm_text, re.IGNORECASE)
        if match and match.start() < best_start:
            best_start = match.start()
            best_end = match.end()
    if best_start == len(text):
        return None
    return (best_start, best_end)

def robust_search_last(text, keywords):
    norm_text = normalize(text)

    best_start = -1
    best_end = -1

    for kw in keywords:
        norm_kw = normalize(kw)
        compact = re.sub(r'\s+', '', norm_kw)
        pattern = r'\s*'.join(re.escape(c) for c in compact)

        matches = list(re.finditer(pattern, norm_text, re.IGNORECASE))

        if matches:
            m = matches[-1]
            if m.start() > best_start:
                best_start = m.start()
                best_end = m.end()

    if best_start == -1:
        return None

    return best_start, best_end


def extract_heading_last(text, start_markers, end_markers):
    pos = robust_search_last(text, start_markers)

    if not pos:
        return ""

    start_pos = pos[0]
    remaining = text[start_pos:]

    end_pos_info = robust_search(remaining, end_markers)

    if end_pos_info:
        end_pos = end_pos_info[0]
        return remaining[:end_pos].strip()

    return remaining.strip()

def find_block(text, start_markers, end_markers):
    pos = robust_search(text, start_markers)
    if not pos:
        return ""
    start_pos = pos[1]
    remaining = text[start_pos:]
    end_pos_info = robust_search(remaining, end_markers)
    if end_pos_info:
        end_pos = end_pos_info[0]
        return remaining[:end_pos].strip()
    return remaining.strip()

def find_block_after(text, after_markers, start_markers, end_markers):
    """
    从 after_markers 之后开始找 start_markers，
    避免匹配到目录里的标题。
    """
    search_text = text

    after_pos = robust_search(text, after_markers)
    if after_pos:
        search_text = text[after_pos[1]:]

    return extract_heading(search_text, start_markers, end_markers)

def extract_heading(text, start_markers, end_markers):
    pos = robust_search(text, start_markers)
    if not pos:
        return ""
    start_pos = pos[0]
    remaining = text[start_pos:]
    end_pos_info = robust_search(remaining, end_markers)
    if end_pos_info:
        end_pos = end_pos_info[0]
        return remaining[:end_pos].strip()
    return remaining.strip()

def extract_next_sentence(text, start_markers):
    pos = robust_search(text, start_markers)
    if not pos:
        return ""
    after = text[pos[1]:].lstrip()
    m = re.search(r'[。；\n]', after)
    if m:
        return after[:m.end()].strip()
    return after.strip()

def extract_paragraph(text, start_markers):
    pos = robust_search(text, start_markers)
    if not pos:
        return ""
    block_start = text.rfind('\n', 0, pos[0])
    if block_start == -1:
        block_start = 0
    block_end = text.find('\n', pos[1])
    if block_end == -1:
        block_end = len(text)
    return text[block_start:block_end].strip()


def extract_subitem(item_text, sub_marker, stop_marker=None):
    escaped = re.escape(sub_marker)
    # 匹配从 sub_marker 开始，直到下一个符合条件的编号（①~⑩、(1)、(2)等）或文本结束
    pattern = re.compile(
        rf'(?ms)^\s*{escaped}\s*.*?'
        rf'(?=^\s*[①②③④⑤⑥⑦⑧⑨⑩]|^\s*[（(]\s*\d+\s*[）)]|\Z)'
    )
    m = pattern.search(item_text)
    if not m:
        return ""
    full_text = m.group(0).strip()

    # 应用 stop_marker
    if stop_marker:
        idx = full_text.find(stop_marker)
        if idx != -1:
            full_text = full_text[:idx].strip()

    # 防止提取的文本包含第二个相同标记（例如两个"②"），截断到第二个之前
    if full_text.count(sub_marker) > 1:
        # 从第一个标记之后查找下一个
        second_start = full_text.find(sub_marker, len(sub_marker))
        if second_start != -1:
            full_text = full_text[:second_start].strip()

    return full_text

# ==================== 编号条目提取 ====================
def extract_block_until_next_level(item_text, sub_marker):
    """
    提取从 sub_marker 开始，直到下一个同级编号（如（3））或文末的完整块。
    不把 ①、② 等视为结束符，确保子内容被包含。
    """
    escaped = re.escape(sub_marker)
    # 结束符只匹配与 sub_marker 同级的编号（如（1）、（2）、（3）等）
    pattern = re.compile(
        rf'(?ms)^\s*{escaped}\s*.*?'
        rf'(?=^\s*[（(]\s*\d+\s*[）)]|\Z)'
    )
    m = pattern.search(item_text)
    return m.group(0).strip() if m else ""

def is_garbled(text, threshold=0.3):
    """判断文本是否主要为乱码"""
    if not text:
        return False
    allowed = re.findall(r'[\u4e00-\u9fffA-Za-z0-9\s，。；：、．（）《》【】“”‘’—\-%.=+×÷…]', text)
    ratio = len(allowed) / len(text) if text else 1
    return ratio < (1 - threshold)

def extract_numbered_items(block, numbered_list):
    if not block:
        return []

    # 一级编号匹配
    pattern = re.compile(r'(?m)^\s*(\d+)\s*[、.．]')
    matches = list(pattern.finditer(block))
    results = []

    if numbered_list.get("include_intro", False):
        if matches:
            intro = block[:matches[0].start()].strip()
        else:
            intro = block.strip()
        if intro:
            results.append(intro)

    item_map = {}
    for i, m in enumerate(matches):
        num = int(m.group(1))
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(block)
        item_map[num] = block[start:end].strip()

    # 递归处理子规则
    def process_subrules(parent_text, rules, seen=None):
        if seen is None:
            seen = set()
        res = []
        for rule in rules:
            marker = rule.get("marker")
            stop_marker = rule.get("stop_marker")
            rule_id = (marker or "") + (stop_marker or "")
            if rule_id in seen:
                continue
            seen.add(rule_id)

            sub_subrules = rule.get("subrules", [])
            if sub_subrules:
                # 用宽提取获取包含子项的完整块
                wide_text = extract_block_until_next_level(parent_text, marker)
                if not wide_text:
                    continue
                # 是否保留当前标题
                if rule.get("include_title", True):
                    res.append(wide_text.splitlines()[0].strip())
                # 递归处理子子规则
                res.extend(process_subrules(wide_text, sub_subrules, seen))
            else:
                # 无子规则，用原始 extract_subitem（会正确识别①、②作为结束符）
                sub_text = extract_subitem(parent_text, marker, stop_marker)
                if not sub_text:
                    continue
                if rule.get("title_only"):
                    res.append(sub_text.splitlines()[0].strip())
                elif rule.get("full") or rule.get("include_title", False):
                    res.append(sub_text)
        return res

    # 处理每个一级规则
    for rule in numbered_list.get("rules", []):
        num = rule.get("number")
        item_text = item_map.get(num, "")
        if not item_text:
            continue

        if rule.get("title_only"):
            results.append(item_text.splitlines()[0].strip())
            continue

        if rule.get("full") and not rule.get("subrules"):
            results.append(item_text)
            continue

        subrules = rule.get("subrules", [])
        if subrules:
            sub_results = []

            if rule.get("include_title", True):
                sub_results.append(item_text.splitlines()[0].strip())

            if rule.get("include_intro", False):
                first_marker = subrules[0].get("marker")
                if first_marker:
                    idx = item_text.find(first_marker)
                    if idx != -1:
                        intro = item_text.splitlines()[0].strip()
                        before_first_sub = item_text[:idx].strip()

                        before_first_sub = before_first_sub.replace(intro, "", 1).strip()

                        if before_first_sub:
                            sub_results.append(before_first_sub)

            sub_results.extend(process_subrules(item_text, subrules))
            results.append("\n\n".join(sub_results))
        else:
            results.append(item_text)

    return results

# ==================== 中文编号提取 ====================
def extract_subsections(block, wanted_items):
    if not block:
        return []

    pattern = re.compile(
        r'(?s)(（[一二三四五六七八九十]+）.*?)(?=（[一二三四五六七八九十]+）|\Z)'
    )

    results = []

    for m in pattern.finditer(block):
        section_text = m.group(1).strip()
        marker = re.match(r'（[一二三四五六七八九十]+）', section_text)

        if marker and marker.group(0) in wanted_items:
            results.append(section_text)

    return results

def extract_subsections_or_block(block, wanted_items):
    items = extract_subsections(block, wanted_items)
    if items:
        return "\n\n".join(items)
    return block.strip()

def extract_cn_numbered_sections(block, wanted_markers):
    pattern = re.compile(
        r'(?ms)^\s*(（[一二三四五六七八九十]+）.*?)'
        r'(?=^\s*（[一二三四五六七八九十]+）|\Z)'
    )
    results = []
    for m in pattern.finditer(block):
        section_text = m.group(1).strip()
        marker = re.match(r'^（[一二三四五六七八九十]+）', section_text)
        if marker and marker.group(0) in wanted_markers:
            results.append(section_text)
    return results

def extract_text_before_table_and_table(text, section):
    start_kw = section.get("start_keywords", [])
    end_kw = section.get("end_keywords", [])
    table_kw = section.get("table_keywords", [])
    block = extract_heading(text, start_kw, end_kw)
    if not block:
        return ""
    table_pos = robust_search(block, table_kw)
    if not table_pos:
        return block
    return block[:table_pos[0]].strip() + "\n\n【表格内容】\n" + block[table_pos[0]:].strip()

def extract_by_path(text, level2_num=None, level3_marker=None):
    pattern_l2 = re.compile(
        rf'^\s*{level2_num}[、.．]\s*(.*?)(?=^\s*\d+[、.．]\s*|\Z)',
        re.S | re.M
    )
    m2 = pattern_l2.search(text)
    if not m2:
        return ""
    block_l2 = m2.group(0)
    if level3_marker:
        pattern_l3 = re.compile(
            rf'^\s*{re.escape(level3_marker)}\s*(.*?)(?=^\s*（\d+）|\Z)',
            re.S | re.M
        )
        m3 = pattern_l3.search(block_l2)
        if not m3:
            return ""
        return m3.group(0).strip()
    return block_l2.strip()

# ==================== 提取任务执行器 ====================
def execute_extraction(text, config):
    results = []
    for section in config.get('sections', []):
        name = section.get('name', '')
        output_title = section.get('output_title', name)
        mode = section.get('mode', 'block')
        start_kw = section.get('start_keywords', [name])
        end_kw = section.get('end_keywords', [])
        numbered_list = section.get('numbered_list', {})
        content = ""

        if mode == 'heading_only':
            content = ""
        elif mode == 'block':
            after_kw = section.get("after_keywords", [])

            if after_kw:
                content = find_block_after(text, after_kw, start_kw, end_kw)
            else:
                content = find_block(text, start_kw, end_kw)
        elif mode == 'heading':
            content = extract_heading(text, start_kw, end_kw)
        elif mode == 'next_sentence':
            content = extract_next_sentence(text, start_kw)
        elif mode == 'paragraph':
            content = extract_paragraph(text, start_kw)
        elif mode == "subsections":
            use_last = section.get("use_last", False)

            if use_last:
                block = extract_heading_last(text, start_kw, end_kw)
            else:
                block = extract_heading(text, start_kw, end_kw)

            subsection_cfg = section.get("subsections", {})
            wanted_items = subsection_cfg.get("items", [])

            content = extract_subsections_or_block(block, wanted_items)
        elif mode == "text_before_table_and_table":
            content = extract_text_before_table_and_table(text, section)
        elif mode == 'numbered_items':
            block = find_block(text, start_kw, end_kw)
            if block and numbered_list:
                items = extract_numbered_items(block, numbered_list)
                content = '\n\n'.join(items)
            else:
                content = block
        elif mode == "cn_numbered_sections":
            block = extract_heading(text, start_kw, end_kw)
            cn_cfg = section.get("cn_numbered_sections", {})
            items = extract_cn_numbered_sections(block, cn_cfg.get("items", []))
            content = "\n\n".join(items)
        elif mode == 'numbered_all':
            block = find_block(text, start_kw, end_kw)
            content = block
        else:
            content = find_block(text, start_kw, end_kw)

        results.append({
            'title': output_title,
            'level': section.get('level', 1),
            'content': content
        })
    return results

# ==================== Word 生成（含图片占位符替换） ====================
def generate_word(output_path, all_results, doc_config, table_images, pdf_path):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(10.5)
    style.paragraph_format.first_line_indent = Cm(0.74)
    style.paragraph_format.space_after = Pt(2)

    # 大标题
    title = doc_config.get('title', '合同重点提取')
    doc.add_heading(title, level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER

    preface = doc_config.get('preface', '')
    if preface:
        p = doc.add_paragraph()
        run = p.add_run(preface)
        set_run_font(run, bold=True)

    # 将 table_images 列表转为字典，便于通过占位符查找
    img_dict = {marker: (page, clip) for (page, clip, marker) in table_images}

    # 输出每个章节
    for item in all_results:
        heading = item['title']
        level = item['level']
        content = item['content']

        if heading:
            # 二级标题自动加粗，字体稍大
            h = doc.add_heading(heading, level=min(level, 2))
            for run in h.runs:
                set_run_font(run, "宋体", 11 if level == 2 else 14, bold=True)

        if not content:
            continue

        # 将内容按表格占位符拆分
        parts = re.split(r'(<<TABLE_IMG_\d+>>)', content)
        for part in parts:
            if part.startswith('<<TABLE_IMG_') and part.endswith('>>'):
                marker = part.strip()
                if marker in img_dict:
                    page_idx, clip_rect = img_dict[marker]
                    try:
                        pdf_doc = fitz.open(pdf_path)
                        page = pdf_doc[page_idx]
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip_rect)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        pdf_doc.close()
                        with io.BytesIO() as img_stream:
                            img.save(img_stream, format='PNG')
                            img_stream.seek(0)
                            doc.add_picture(img_stream, width=Inches(5.5))
                    except Exception as e:
                        p = doc.add_paragraph(f"[表格图片提取失败: {e}]")
                        set_run_font(p.runs[0], "宋体", 10.5)
                else:
                    p = doc.add_paragraph("[表格图片丢失]")
                    set_run_font(p.runs[0], "宋体", 10.5)
            elif part.strip():
                # 普通文本分行
                for line in part.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    set_run_font(run, "宋体", 10.5)

    disclaimer = doc_config.get('disclaimer', '')
    if disclaimer:
        p = doc.add_paragraph()
        run = p.add_run(disclaimer)
        set_run_font(run, "宋体", 10.5, italic=True)

    doc.save(output_path)

# ==================== 内置配置（私募基金合同示例） ====================
BUILTIN_CONFIG = r"""
document:
  title: "基金合同内容提取"
  preface: ""
  disclaimer: ""

sections:
  # ========== 风险揭示书 ==========
  - name: "风险揭示书标题"
    mode: "block"
    start_keywords: ["自行承担投资风险。"]
    end_keywords: ["尊敬的投资者"]
    output_title: ""
    level: 1

  - name: "特殊风险揭示"
    mode: "numbered_items"
    start_keywords: ["（一）特殊风险揭示"]
    end_keywords:
      - "1、投资标的所涉市场风险"
      - "（二）私募基金投资标的所涉风险"
      - "三、投资者声明"
    numbered_list:
      rules:
        - number: 1
          title_only: true
        - number: 14
          include_title: true
          subrules:
            - marker: "（1）"
              title_only: true
            - marker: "（11）"
              full: true
        - number: 18
          full: true
    output_title: "（一）特殊风险揭示"
    level: 2

  - name: "私募基金投资标的所涉风险"
    mode: "numbered_items"
    start_keywords: ["（二）私募基金投资标的所涉风险"]
    end_keywords:
      - "（三）一般风险揭示"
      - "三、投资者声明"
    numbered_list:
      rules:
        - number: 1
          title_only: true
        - number: 2
          include_title: true
          subrules:
            - marker: "（1）"
              title_only: true
            - marker: "（7）"
              title_only: true
        - number: 7
          title_only: true
    output_title: "（二）私募基金投资标的所涉风险"
    level: 2

  # ========== 投资者声明 ==========
  - name: "投资者声明"
    mode: "block"
    start_keywords: ["三、投资者声明"]
    end_keywords: ["合格投资者承诺书", "签署页", "基金的基本情况"]
    output_title: "三、投资者声明"
    level: 1

  # ========== 基金基本情况 ==========
  - name: "基金基本情况大标题"
    mode: "heading_only"
    output_title: "四、基金的基本情况"
    level: 1

  - name: "基金名称"
    mode: "next_sentence"
    start_keywords: ["（一）基金的名称"]
    output_title: "（一）基金的名称"
    level: 2

  - name: "产品类型"
    mode: "next_sentence"
    start_keywords: ["（二）产品类型"]
    output_title: "（二）产品类型"
    level: 2

  - name: "运作方式"
    mode: "next_sentence"
    start_keywords: ["（三）基金的运作方式"]
    output_title: "（三）基金的运作方式"
    level: 2

  - name: "投资目标和投资范围简述"
    mode: "next_sentence"
    start_keywords: ["（四）基金的投资目标和投资范围"]
    output_title: "（四）基金的投资目标和投资范围"
    level: 2

  - name: "存续期限"
    mode: "next_sentence"
    start_keywords: ["（五）基金的存续期限"]
    output_title: "（五）基金的存续期限"
    level: 2

  - name: "初始募集面值"
    mode: "next_sentence"
    start_keywords: ["（六）基金份额的初始募集面值"]
    output_title: "（六）基金份额的初始募集面值"
    level: 2

  - name: "托管事项"
    mode: "next_sentence"
    start_keywords: ["（七）基金的托管事项"]
    output_title: "（七）基金的托管事项"
    level: 2

  # ========== 基金的分级分类 ==========
  - name: "基金的分级分类"
    mode: "subsections"
    use_last: true
    start_keywords: 
     - "五、基金的分级分类"
    end_keywords: 
      - "六、基金的募集"
    subsections:
      items: ["（一）"]  
    output_title: "五、基金的分级分类"
    level: 1

  # ========== 基金认购 ==========
  - name: "基金的认购事项"
    mode: "numbered_items"
    start_keywords: ["（二）基金的认购事项"]
    end_keywords:
      - "（三）募集期间募集资金的管理"
    numbered_list:
      rules:
        - number: 2
          full: true
        - number: 4
          full: true
    output_title: ""
    level: 2

  # ========== 基金的申购、赎回与转让 ==========
  - name: "申购赎回与转让大标题"
    mode: "heading_only"
    output_title: "八、基金的申购、赎回与转让"
    level: 1

  - name: "申购和赎回的开放日及时间"
    mode: "numbered_items"
    start_keywords: ["（二）申购和赎回的开放日及时间"]
    end_keywords:
      - "（三）基金的申购事项"
    numbered_list:
      include_intro: true
      rules:
        - number: 1
          full: true
        - number: 2
          full: true
        - number: 3
          include_title: true
          subrules:
            - marker: "（1）"
              full: true
    output_title: "（二）申购和赎回的开放日及时间"
    level: 2

  - name: "基金的申购事项"
    mode: "numbered_items"
    start_keywords: ["（三）基金的申购事项"]
    end_keywords:
      - "（四）基金的赎回事项"
    numbered_list:
      rules:
        - number: 2
          full: true
        - number: 4
          full: true
    output_title: "（三）基金的申购事项"
    level: 2

  - name: "基金的赎回事项"
    mode: "numbered_items"
    start_keywords: ["（四）基金的赎回事项"]
    end_keywords:
      - "（五）基金份额的转让"
    numbered_list:
      rules:
        - number: 2
          include_title: true
          include_intro: true
          subrules:
            - marker: "（1）"
              title_only: true
            - marker: "（2）"
              title_only: true
        - number: 4
          full: true
    output_title: "（四）基金的赎回事项"
    level: 2

  # ========== 当事人的权利和义务 ==========
  - name: "当事人的权利和义务"
    mode: "heading_only"
    output_title: "九、当事人的权利和义务"
    level: 1

  - name: "基金管理人"
    mode: "numbered_items"
    start_keywords: ["（二）基金管理人"]
    end_keywords:
      - "（三）基金托管人"
    numbered_list:
      rules:
        - number: 1
          full: true
    output_title: "（二）基金管理人"
    level: 2

  # ========== 基金的投资 ==========
  - name: "基金的投资"
    mode: "heading_only"
    output_title: "十二、基金的投资"
    level: 1

  - name: "投资范围"
    mode: "block"
    start_keywords: ["（三）投资范围"]
    end_keywords: ["（四）投资策略"]
    output_title: "（三）投资范围"
    level: 2

  - name: "投资策略"
    mode: "block"
    start_keywords: ["（四）投资策略"]
    end_keywords: ["（五）投资限制"]
    output_title: "（四）投资策略"
    level: 2

  - name: "投资限制"
    mode: "numbered_items"
    start_keywords: ["（五）投资限制"]
    end_keywords:
      - "（六）投资嵌套层级的限制"
      - "（六）"
    numbered_list:
      include_intro: true
      rules:
        - number: 1
          full: true
        - number: 2
          full: true
        - number: 3
          full: true
        - number: 4
          full: true
        - number: 5
          full: true
        - number: 6
          full: true
        - number: 7
          title_only: true
    output_title: "（五）投资限制"
    level: 2

  - name: "预警止损机制"
    mode: "block"
    start_keywords: ["（十二）预警止损机制"]
    end_keywords: ["（十三）业绩比较基准（如有）"]
    output_title: "（十二）预警止损机制"
    level: 2

  # ========== 资金清算 ==========
  - name: "资金清算交收安排"
    mode: "heading_only"
    output_title: "十四、资金清算交收安排"
    level: 1

  - name: "基金成立"
    mode: "numbered_items"
    start_keywords:
      - "（四）基金成立、申购或赎回的资金清算"
    end_keywords:
      - "十五、投资指令的发送、确认和执行"
    numbered_list:
      rules:
        - number: 4
          full: true
    output_title: ""
    level: 2

  - name: "基金费用"
    mode: "numbered_items"
    start_keywords:
      - "（二）费用计提方法、计提标准和支付方式"
    end_keywords:
      - "本基金的销售服务费自本基金成立之日起"
    numbered_list:
      rules:
        - number: 1
          include_title: true
          subrules:
            - marker: "（1）"
              full: true
            - marker: "（2）"
              include_title: true
              subrules:
                - marker: "①"
                  full: true
                - marker: "②"
                  full: true
                  stop_marker: "其中"
        - number: 2
          full: true
        - number: 3
          full: true
        - number: 4
          full: true
    output_title: "（二）费用计提方法、计提标准和支付方式"
    level: 2

  - name: "基金的收益分配"
    mode: "heading_only"
    output_title: "十八、基金的收益分配"
    level: 1

  - name: "基金的收益分配"
    mode: "numbered_items"
    start_keywords: ["（二）基金收益分配原则"]
    end_keywords:
      - "（三）收益分配方案的确定与通知"
    numbered_list:
      rules:
        - number: 3
          full: true
    output_title: ""
    level: 2
"""

def load_config(config_path):
    if config_path and os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    print("使用内置默认配置（私募基金合同）")
    return yaml.safe_load(BUILTIN_CONFIG)

# ==================== 单文件处理 ====================
def process_file(input_path, config, output_path=None):
    # 1. 读取 PDF 并得到合并后的文本和表格图片信息
    merged_text, table_images = read_pdf_with_tables(input_path)
    # 2. 清洗文本
    text = clean_text(merged_text)
    # 3. 执行提取规则
    results = execute_extraction(text, config)
    # 4. 生成 Word（传入表格图片列表和 PDF 路径，以便截图）
    if not output_path:
        output_path = Path(input_path).stem + "_重点提取.docx"
    generate_word(output_path, results, config.get('document', {}), table_images, input_path)
    print(f"✅ 已生成：{output_path}")

# ==================== 主程序 ====================
# ==================== 主程序 ====================
if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    config_path = script_dir / 'config.yaml'
    if not config_path.exists():
        config_path = None
    config = load_config(str(config_path) if config_path else None)

    # 只查找 PDF 文件
    files = list(script_dir.glob("*.pdf"))
    files = [f for f in files if "重点提取" not in f.name
             and not f.name.startswith("~$")
             and not f.name.startswith(".~")]

    if not files:
        print("当前文件夹没有找到可处理的 PDF 合同文件")
        input("按回车退出...")
        sys.exit(0)

    for f in files:
        print(f"▶ 处理：{f.name}")
        try:
            process_file(str(f), config)
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"❌ 处理失败：{f.name}，错误：{e}")

    input("全部完成，按回车键退出...")