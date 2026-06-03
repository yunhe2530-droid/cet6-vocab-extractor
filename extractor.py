"""六级黄标单词提取引擎 v2 — 按原始需求重写"""

import io
import re
import time

import fitz
import openpyxl
import pytesseract
import requests
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image

from cet6_phrases import CET6_PHRASES

# ── 配置 ──────────────────────────────────────────
YELLOW = (0.973, 0.890, 0.518)
YELLOW_TOLERANCE = 0.001
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# 仅过滤最基础的功能词（用户主动标黄的不应过滤）
MINIMAL_STOP = {
    "the", "a", "an", "and", "or", "but", "so", "if", "than",
    "of", "in", "on", "at", "to", "for", "by", "with", "from",
    "be", "is", "are", "was", "were", "been", "am",
    "it", "its", "it's", "itself",
    "i", "you", "he", "she", "we", "they", "me", "him", "her", "us", "them",
    "my", "your", "his", "my", "our", "their",
    "this", "that", "these", "those",
    "what", "which", "who", "whom", "whose",
    "no", "not", "nor", "neither",
    "as", "just", "very", "too", "also", "even",
    "there", "here", "where", "when", "why", "how",
    "do", "does", "did", "done",
    "have", "has", "had",
    "will", "would", "shall", "should", "can", "could", "may", "might", "must",
    "say", "said", "says", "tell", "told", "make", "made", "get", "got",
}


# ══════════════════════════════════════════════════
#  颜色检测
# ══════════════════════════════════════════════════

def is_yellow(fill):
    if fill is None:
        return False
    try:
        return (abs(fill[0] - YELLOW[0]) < YELLOW_TOLERANCE and
                abs(fill[1] - YELLOW[1]) < YELLOW_TOLERANCE and
                abs(fill[2] - YELLOW[2]) < YELLOW_TOLERANCE)
    except (TypeError, IndexError):
        return False


# ══════════════════════════════════════════════════
#  辅助工具
# ══════════════════════════════════════════════════

def _merge_rects(rects):
    if not rects:
        return []
    merged = []
    sorted_rects = sorted(rects, key=lambda r: (r.y0, r.x0))
    cur = sorted_rects[0]
    for r in sorted_rects[1:]:
        if (abs(r.y0 - cur.y0) < 5 and abs(r.y1 - cur.y1) < 5
                and r.x0 <= cur.x1 + 5):
            cur = fitz.Rect(cur.x0, cur.y0, max(cur.x1, r.x1), max(cur.y1, r.y1))
        else:
            merged.append(cur)
            cur = r
    merged.append(cur)
    return merged


def _get_yellow_rects(page):
    rects = []
    for drawing in page.get_drawings():
        f = drawing.get("fill")
        if f and is_yellow(f):
            rects.append(drawing["rect"])
        for item in drawing.get("items", []):
            if len(item) >= 3:
                fill = item[2] if isinstance(item[2], (tuple, list)) else None
                if fill and is_yellow(fill) and len(item) > 1:
                    try:
                        rects.append(item[1])
                    except Exception:
                        pass
    return _merge_rects(rects)


def _clean_words(text):
    """提取字母单词，返回 (words, original_spans)"""
    text = text.replace('‘', "'").replace('’', "'")
    text = text.replace('“', '"').replace('”', '"')
    spans = re.findall(r"[A-Za-z]+(?:['-][A-Za-z]+)*", text)
    return [s for s in spans if len(s) > 1], spans


# ══════════════════════════════════════════════════
#  句子提取（原生 PDF）
# ══════════════════════════════════════════════════

def _split_sentences(text):
    """按 .?! 分割句子，返回句子列表"""
    sents = re.split(r'(?<=[.?!])\s+', text)
    return [s.strip() for s in sents if len(s.strip()) > 5]


def _extract_page_text(page):
    """获取页面纯文本"""
    return page.get_text("text")


def find_sentence_for_word(word, page_text, bbox_y, page_num=0):
    """
    在页面文本中找到包含 word 的完整句子。
    bbox_y 用于消歧（选择 y 坐标最近的段落）
    """
    sents = _split_sentences(page_text)
    if not sents:
        return ""

    w_lower = word.strip(".,!?;:'\"").lower()

    # 找所有包含该词的句子
    candidates = []
    for i, s in enumerate(sents):
        if w_lower in s.lower():
            candidates.append((i, s))

    if not candidates:
        return word  # fallback
    if len(candidates) == 1:
        return candidates[0][1]

    # 多句包含，按句子在页面中的位置（行号）消歧
    # 用 bbox_y 估计该词在第几行，选择最近的段落
    # 简化：返回第一个
    return candidates[0][1]


# ══════════════════════════════════════════════════
#  短语检测（结合句子上下文 + CET6 短语库）
# ══════════════════════════════════════════════════

def detect_phrase_in_sentence(highlighted_word, sentence):
    """
    在句子中检测 highlighted_word 是否属于 CET-6 短语。
    返回 (phrase, words_in_phrase) 或 (highlighted_word, 1)
    """
    hl = highlighted_word.lower().strip(".,!?;:'\"")
    words, _ = _clean_words(sentence)
    words_lower = [w.lower() for w in words]

    # 找 highlighted_word 在句子中的所有位置
    positions = [i for i, w in enumerate(words_lower) if w == hl]
    if not positions:
        return highlighted_word, words  # fallback: 返回原词

    # 对每个位置，尝试最长短语匹配（从长到短）
    for pos in positions:
        for n in range(5, 1, -1):  # 5→2 词短语
            for start in range(max(0, pos - n + 1),
                               min(pos + 1, len(words) - n + 1)):
                if start + n > len(words):
                    continue
                candidate = ' '.join(words_lower[start:start + n])
                if candidate in CET6_PHRASES:
                    # 找到了！返回原文格式的短语
                    original = ' '.join(words[start:start + n])
                    return original, words[start:start + n]
    return highlighted_word, [highlighted_word]


# ══════════════════════════════════════════════════
#  原生 PDF 提取
# ══════════════════════════════════════════════════

def extract_native(pdf_bytes):
    """提取原生 PDF 中的黄标单词，返回每页的 (text, bbox, page_text)"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = {}

    for pn, page in enumerate(doc):
        yellow_rects = _get_yellow_rects(page)
        if not yellow_rects:
            continue

        full_text = _extract_page_text(page)
        page_words = []

        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    bbox = fitz.Rect(span["bbox"])
                    text = span["text"].strip()
                    if not text:
                        continue
                    for yr in yellow_rects:
                        overlap = bbox.intersect(yr)
                        if overlap and overlap.get_area() > 0:
                            page_words.append({
                                "text": text,
                                "bbox": bbox,
                            })
                            break

        results[pn] = {
            "words": page_words,
            "full_text": full_text,
        }

    doc.close()
    return results


# ══════════════════════════════════════════════════
#  扫描件 PDF 提取（OCR）
# ══════════════════════════════════════════════════

def extract_scanned(pdf_bytes):
    """从扫描件 PDF 提取黄标区域→OCR"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = {}

    for pn, page in enumerate(doc):
        yellow_rects = _get_yellow_rects(page)
        if not yellow_rects:
            continue

        # 取底图
        images = page.get_images(full=True)
        if not images:
            continue

        best = max(images, key=lambda x: (x[2], x[3]))
        pix = fitz.Pixmap(doc, best[0])
        if pix.n > 4:
            pix = fitz.Pixmap(fitz.csRGB, pix)

        pil_img = Image.open(io.BytesIO(pix.tobytes("png")))
        img_w, img_h = pil_img.size

        # 图片在页面的放置位置
        img_info = page.get_image_info()
        img_bbox = fitz.Rect(img_info[0]["bbox"]) if img_info else page.rect
        pdf_img_w, pdf_img_h = img_bbox.width, img_bbox.height

        page_words = []
        for yr in yellow_rects:
            if yr.intersect(img_bbox).get_area() <= 0:
                continue

            sx, sy = img_w / pdf_img_w, img_h / pdf_img_h
            x0 = int((yr.x0 - img_bbox.x0) * sx)
            y0 = int((yr.y0 - img_bbox.y0) * sy)
            x1 = int((yr.x1 - img_bbox.x0) * sx)
            y1 = int((yr.y1 - img_bbox.y0) * sy)
            x0, y0 = max(0, x0), max(0, y0)
            x1, y1 = min(img_w, x1), min(img_h, y1)
            if x1 <= x0 or y1 <= y0:
                continue

            pad = 5
            crop = pil_img.crop((
                max(0, x0 - pad), max(0, y0 - pad),
                min(img_w, x1 + pad), min(img_h, y1 + pad),
            ))
            crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)

            try:
                text = pytesseract.image_to_string(crop, lang="eng",
                                                   config="--psm 7 --oem 3")
                text = text.strip()
                if text:
                    page_words.append({"text": text, "bbox": yr})
            except Exception:
                continue

        results[pn] = {
            "words": page_words,
            "full_text": "\n".join(w["text"] for w in page_words),
        }

    doc.close()
    return results


# ══════════════════════════════════════════════════
#  文本类型判断
# ══════════════════════════════════════════════════

def is_scanned_pdf(pdf_bytes, sample_pages=2):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text_len = sum(len(doc[i].get_text().strip()) for i in range(min(len(doc), sample_pages)))
    doc.close()
    return text_len < 50


# ══════════════════════════════════════════════════
#  有道词典查询
# ══════════════════════════════════════════════════

def query_youdao(word, retries=3):
    """查询有道词典，返回中文释义（含词性）"""
    url = f"https://dict.youdao.com/w/eng/{requests.utils.quote(word)}"
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Accept-Language": "zh-CN,zh;q=0.9",
    }

    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code != 200:
                continue
            html = resp.text

            meanings = []

            # 方法A: 找 .trans-container > ul > li
            # 这是有道最常规的释义容器
            for item in re.findall(
                    r'<div[^>]*class="[^"]*trans-container[^"]*"[^>]*>'
                    r'.*?<ul>(.*?)</ul>', html, re.DOTALL):
                for li in re.findall(r'<li>(.*?)</li>', item, re.DOTALL):
                    # 提取 span 或直接文本
                    text = re.sub(r'<[^>]+>', '', li).strip()
                    if text and re.search(r'[一-鿿]', text):
                        # 提取词性（如果有）
                        clean = re.sub(r'\s+', ' ', text).strip()
                        if clean not in meanings:
                            meanings.append(clean)

            # 方法B: 词性 + 释义模式（<span class="pos">v.</span>  <span class="trans">转移</span>）
            if not meanings:
                for pos, trans in re.findall(
                        r'<span[^>]*class="[^"]*pos[^"]*"[^>]*>(.*?)</span>\s*'
                        r'<span[^>]*class="[^"]*trans[^"]*"[^>]*>(.*?)</span>',
                        html, re.DOTALL):
                    p = pos.strip()
                    t = re.sub(r'<[^>]+>', '', trans.strip())
                    if re.search(r'[一-鿿]', t):
                        meanings.append(f"{p}{t}")

            # 方法C: <div class="word-define">...<li>释义</li>...
            if not meanings:
                for li in re.findall(
                        r'<div[^>]*class="[^"]*word-define[^"]*"[^>]*>'
                        r'.*?<li[^>]*>(.*?)</li>', html, re.DOTALL):
                    t = re.sub(r'<[^>]+>', '', li).strip()
                    if t and re.search(r'[一-鿿]', t) and t not in meanings:
                        meanings.append(t)

            # 方法D: 短语释义（.<div class="wordGroup">）
            if not meanings:
                for span in re.findall(
                        r'<div[^>]*class="[^"]*wordGroup[^"]*"[^>]*>'
                        r'.*?<span[^>]*>(.*?)</span>', html, re.DOTALL):
                    t = re.sub(r'<[^>]+>', '', span).strip()
                    if t and re.search(r'[一-鿿]', t) and t not in meanings:
                        meanings.append(t)

            if meanings:
                return "；".join(meanings[:3])

            # 方法E: 尝试 JSON API 作为后备
            try:
                api_url = f"https://dict.youdao.com/jsonapi?q={requests.utils.quote(word)}"
                api_resp = requests.get(api_url, headers=headers, timeout=5)
                if api_resp.status_code == 200:
                    data = api_resp.json()
                    # 查 ec 字段（英中词典）
                    for section in (data.get("ec", {}).get("exam", []) +
                                    data.get("ec", {}).get("source", [])):
                        for tr in section.get("trs", []):
                            for tran in tr.get("tr", []):
                                t = tran.get("l", {}).get("i", "")
                                if re.search(r'[一-鿿]', t):
                                    pos = tr.get("pos", "")
                                    meanings.append(f"{pos} {t}".strip())
                    if meanings:
                        return "；".join(meanings[:3])
            except Exception:
                pass

            return "未找到释义"

        except requests.RequestException:
            if attempt < retries - 1:
                time.sleep(1)
                continue
            return "查询失败"

    return "查询失败"


# ══════════════════════════════════════════════════
#  主处理流程
# ══════════════════════════════════════════════════

def process_pdf(pdf_bytes, filename="", on_status=None):
    """
    提取单个 PDF 中的黄标单词。
    返回 [(word, chinese, sentence, source_word, is_phrase), ...]
    """
    if on_status:
        on_status(f"分析 {filename}...")

    scanned = is_scanned_pdf(pdf_bytes)
    if on_status:
        on_status(f"{'扫描件' if scanned else '文本'} PDF: {filename}")

    raw = extract_scanned(pdf_bytes) if scanned else extract_native(pdf_bytes)

    rows = []  # [(source_word, sentence_fragment, full_text)]

    for pn in sorted(raw.keys()):
        data = raw[pn]
        page_text = data.get("full_text", "")
        sents = _split_sentences(page_text) if page_text else []

        for item in data["words"]:
            text = item["text"]
            bbox = item["bbox"]

            # 提取单词
            words, _ = _clean_words(text)

            for w in words:
                w_lower = w.lower()
                if w_lower in MINIMAL_STOP:
                    continue

                # 找完整句子
                sentence = ""
                if sents:
                    sentence = find_sentence_for_word(w, page_text,
                                                      bbox.y0, pn)
                if not sentence:
                    sentence = text  # fallback

                # 检测短语
                phrase, phrase_words = detect_phrase_in_sentence(w, sentence)
                is_phrase = phrase.lower() != w.lower()

                if is_phrase:
                    rows.append({
                        "word": phrase,
                        "source_word": w,
                        "sentence": sentence,
                        "is_phrase": True,
                    })
                else:
                    rows.append({
                        "word": w,
                        "source_word": w,
                        "sentence": sentence,
                        "is_phrase": False,
                    })

    if on_status:
        on_status(f"  共提取 {len(rows)} 条")

    # 查词典（去重）
    unique_words = set()
    for r in rows:
        unique_words.add(r["word"].lower())

    cache = {}
    total = len(unique_words)
    for i, key in enumerate(sorted(unique_words)):
        if on_status:
            on_status(f"  查词典 ({i + 1}/{total}): {key}")
        cache[key] = query_youdao(key)
        time.sleep(0.3)

    # 填充结果
    for r in rows:
        r["chinese"] = cache.get(r["word"].lower(), "未找到释义")

    return rows


def process_multiple_pdfs(file_dict, on_status=None):
    """处理多个 PDF，返回所有行的扁平列表"""
    all_rows = []
    total = len(file_dict)

    for idx, (fname, data) in enumerate(file_dict.items()):
        if on_status:
            on_status(f"[{idx + 1}/{total}] {fname}")
        rows = process_pdf(data, fname, on_status=on_status)
        for r in rows:
            r["source"] = fname
        all_rows.extend(rows)

    return all_rows


# ══════════════════════════════════════════════════
#  Excel 生成（带 SUMIF 公式）
# ══════════════════════════════════════════════════

def to_excel(rows):
    """生成 Excel，每行为 (word/phrase, 释义, 原句, SUMIF公式, 辅助列E)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "单词本"

    # ── 样式 ──
    hfont = Font(bold=True, size=11, name="Arial", color="FFFFFF")
    hfill = PatternFill("solid", fgColor="4472C4")
    h_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )

    # 表头
    headers = ["单词/短语", "中文意思和词性", "原文原句", "累计出现次数"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = h_align
        cell.border = thin_border

    # E列隐藏（辅助列）
    col_e = ws.column_dimensions['E']
    col_e.hidden = True

    # ── 数据 ──
    data_font = Font(size=10, name="Arial")
    bold_font = Font(size=10, name="Arial", bold=True)
    highlight_fill = PatternFill("solid", fgColor="FFF2CC")  # 浅黄底色表示来源词

    for i, r in enumerate(rows, 2):
        # A列：单词/短语
        if r["is_phrase"]:
            # 短语：用浅黄底色标出原始高亮词
            cell_a = ws.cell(row=i, column=1,
                             value=r["word"])
            cell_a.font = data_font
            cell_a.fill = highlight_fill
        else:
            # 单个词：加粗表示这就是来源词
            cell_a = ws.cell(row=i, column=1,
                             value=r["word"])
            cell_a.font = bold_font

        # B列：中文释义
        cell_b = ws.cell(row=i, column=2, value=r["chinese"])
        cell_b.font = data_font

        # C列：原文原句
        cell_c = ws.cell(row=i, column=3, value=r["sentence"])
        cell_c.font = data_font
        cell_c.alignment = Alignment(wrap_text=True)

        # D列：SUMIF公式 =SUMIF(A:A, A{i}, E:E)
        ws.cell(row=i, column=4).value = f'=SUMIF(A:A,A{i},E:E)'
        ws.cell(row=i, column=4).font = data_font
        ws.cell(row=i, column=4).alignment = Alignment(horizontal="center")

        # E列（隐藏）：辅助计数
        ws.cell(row=i, column=5, value=1)

        # 边框
        for col in range(1, 6):
            ws.cell(row=i, column=col).border = thin_border

    # ── 列宽 ──
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 65
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 3

    # ── 冻结首行 + 自动筛选 ──
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{len(rows) + 1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
