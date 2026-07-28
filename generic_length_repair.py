import json
import os
import re
import sys
import time
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI


BASE_DIR = Path(__file__).parent
OUT_DIR = BASE_DIR / "output"
CONFIG_PATH = BASE_DIR / "config.local.json"

INPUT_CANDIDATES = [
    OUT_DIR / "综述最终稿_通用引文修复版.docx",
    OUT_DIR / "综述定稿候选_硬约束通过.docx",
]

OUTPUT_DOCX = OUT_DIR / "综述最终稿_字数修正版.docx"
REPORT_PATH = OUT_DIR / "字数预算修复报告.txt"


def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


CONFIG = load_config()

API_KEY = os.getenv("API_KEY") or os.getenv("OPENAI_API_KEY") or CONFIG.get("api_key", "")
BASE_URL = os.getenv("BASE_URL") or os.getenv("OPENAI_BASE_URL") or CONFIG.get("base_url", "https://api.302.ai/v1")
WRITE_MODEL = os.getenv("WRITE_MODEL") or CONFIG.get("write_model", "gpt-5.4")

MIN_TOTAL = int(CONFIG.get("min_total_count", 3000))
MAX_TOTAL = int(CONFIG.get("max_total_count", 8000))
TARGET_TOTAL = int((MIN_TOTAL + MAX_TOTAL) / 2)


def find_input_docx():
    for p in INPUT_CANDIDATES:
        if p.exists():
            return p
    docxs = sorted(OUT_DIR.glob("*.docx"), key=lambda x: x.stat().st_mtime, reverse=True)
    for p in docxs:
        if p.name.startswith("~$"):
            continue
        if any(k in p.name for k in ["材料包", "资料卡片", "参考文献清单", "审稿", "评分"]):
            continue
        return p
    return None


def count_text(text):
    return len(re.sub(r"\s+", "", text))


def read_docx_sections(path):
    doc = Document(path)
    title = ""
    body_items = []
    refs = []
    in_refs = False

    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue

        if t == "参考文献":
            in_refs = True
            continue

        if in_refs:
            refs.append(t)
        else:
            if not title:
                title = t
            else:
                body_items.append(t)

    return title, body_items, refs


def is_heading(t):
    if t in ["摘要", "引言", "结语", "结论", "讨论"]:
        return True
    if re.match(r"^\d+(\.\d+)?\s+.+", t):
        return True
    return False


def split_units(body_items):
    units = []
    current_heading = None
    current_texts = []

    for t in body_items:
        if is_heading(t):
            if current_heading is not None:
                units.append((current_heading, "\n".join(current_texts).strip()))
            current_heading = t
            current_texts = []
        else:
            if current_heading is None:
                current_heading = "正文"
            current_texts.append(t)

    if current_heading is not None:
        units.append((current_heading, "\n".join(current_texts).strip()))

    return units


def set_doc_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5


def set_para_font(p):
    p.paragraph_format.line_spacing = 1.5
    for run in p.runs:
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(12)


def add_heading(doc, text, level):
    p = doc.add_heading(text, level=level)
    set_para_font(p)


def add_para(doc, text):
    p = doc.add_paragraph(text)
    set_para_font(p)


def heading_level(h):
    if re.match(r"^\d+\.\d+\s+", h):
        return 2
    return 1


def write_docx(title, units, refs):
    doc = Document()
    set_doc_style(doc)

    add_heading(doc, title, 0)

    for h, text in units:
        add_heading(doc, h, heading_level(h))
        if text:
            add_para(doc, text)

    add_heading(doc, "参考文献", 1)
    for r in refs:
        add_para(doc, r)

    doc.save(OUTPUT_DOCX)


def call_model(prompt):
    if not API_KEY:
        raise RuntimeError("API_KEY 为空，请在界面填写并保存配置。")

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL, max_retries=5, timeout=120.0)

    last_err = None
    for i in range(1, 4):
        try:
            resp = client.chat.completions.create(
                model=WRITE_MODEL,
                temperature=0.2,
                messages=[
                    {
                        "role": "system",
                        "content": "你是严谨的中文学术论文压缩助手。只能压缩已有内容，不得新增事实、数据、文献或结论。"
                    },
                    {"role": "user", "content": prompt},
                ],
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            time.sleep(3 * i)

    raise last_err


def compress_unit(heading, text, target_chars):
    if not text.strip():
        return text

    prompt = f"""
请压缩以下论文小节。

章节标题：{heading}
目标长度：约 {target_chars} 个汉字，可上下浮动 10%。

硬性要求：
1. 只压缩已有内容，不得新增观点、数据、文献、案例；
2. 必须保留原有数字引文，如[1]、[2]、[3]；
3. 不得编造新的引文编号；
4. 不得删除该节的核心论点；
5. 语言自然、克制，减少重复、套话和空泛表述；
6. 不要输出标题，只输出压缩后的正文。

原文：
{text}
"""
    out = call_model(prompt)
    out = out.strip()
    out = re.sub(r"^#+\s*", "", out)
    return out


def main():
    input_path = find_input_docx()
    if not input_path:
        print("没有找到可进行字数修复的 Word 文件。")
        sys.exit(1)

    title, body_items, refs = read_docx_sections(input_path)
    units = split_units(body_items)

    full_text = "\n".join([title] + body_items + ["参考文献"] + refs)
    current_count = count_text(full_text)

    lines = []
    lines.append("字数预算修复报告")
    lines.append("=" * 50)
    lines.append(f"输入文件：{input_path}")
    lines.append(f"输出文件：{OUTPUT_DOCX}")
    lines.append(f"字数范围：{MIN_TOTAL}—{MAX_TOTAL}")
    lines.append(f"当前全文统计：{current_count}")
    lines.append("")

    if MIN_TOTAL <= current_count <= MAX_TOTAL:
        write_docx(title, units, refs)
        lines.append("结论：当前字数已经在范围内，未调用模型压缩，仅复制为字数修正版。")
        REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        return

    if current_count < MIN_TOTAL:
        write_docx(title, units, refs)
        lines.append(f"结论：全文偏少：{current_count} < {MIN_TOTAL}。为避免无证据扩写，本阶段不自动扩写。")
        REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        sys.exit(1)

    # 超字数：压缩
    target_after = int(MAX_TOTAL * 0.96)
    body_count = count_text("\n".join(body_items))
    ref_count = count_text("\n".join(refs)) + 20
    title_count = count_text(title)

    target_body = max(1000, target_after - ref_count - title_count)

    lines.append(f"目标压缩后全文：约 {target_after}")
    lines.append(f"当前正文统计：{body_count}")
    lines.append(f"目标正文统计：约 {target_body}")
    lines.append("")

    unit_lengths = [(h, count_text(t)) for h, t in units]
    total_unit_len = sum(x for _, x in unit_lengths) or 1

    new_units = []

    for h, text in units:
        old_len = count_text(text)
        if old_len == 0:
            new_units.append((h, text))
            continue

        target_len = max(160, int(target_body * old_len / total_unit_len))

        # 只压缩明显超长的小节；短小节原样保留
        if old_len <= target_len * 1.12:
            new_units.append((h, text))
            lines.append(f"{h}：{old_len} 字，未压缩")
            continue

        lines.append(f"{h}：{old_len} 字 → 目标约 {target_len} 字，开始压缩")
        new_text = compress_unit(h, text, target_len)
        lines.append(f"{h}：压缩后 {count_text(new_text)} 字")
        new_units.append((h, new_text))

    write_docx(title, new_units, refs)

    new_full = "\n".join(
        [title]
        + sum(([h, t] for h, t in new_units), [])
        + ["参考文献"]
        + refs
    )
    new_count = count_text(new_full)

    lines.append("")
    lines.append(f"压缩后全文统计：{new_count}")

    if new_count <= MAX_TOTAL:
        lines.append("结论：通过。全文已压缩到上限以内。")
    else:
        lines.append("结论：仍超字数。可再次运行本脚本，或进一步降低大纲章节数量。")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
