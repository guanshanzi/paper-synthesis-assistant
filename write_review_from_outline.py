import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI


BASE_DIR = Path(__file__).parent
OUT_DIR = BASE_DIR / "output"
OUT_DIR.mkdir(exist_ok=True)

CONFIG_PATH = BASE_DIR / "config.local.json"
REQ_PATH = BASE_DIR / "论文要求与大纲.txt"
MATRIX_PATH = OUT_DIR / "证据矩阵.csv"

DRAFT_DOCX = OUT_DIR / "综述初稿_带引文.docx"
MATERIAL_DOCX = OUT_DIR / "综述写作材料包_带编号.docx"
REF_DOCX = OUT_DIR / "参考文献清单_初稿.docx"
PROGRESS_PATH = OUT_DIR / "review_progress.txt"
ERROR_LOG = OUT_DIR / "review_error_log.txt"


def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


CONFIG = load_config()

API_KEY = os.getenv("API_KEY") or os.getenv("OPENAI_API_KEY") or CONFIG.get("api_key", "")
BASE_URL = os.getenv("BASE_URL") or os.getenv("OPENAI_BASE_URL") or CONFIG.get("base_url", "https://api.302.ai/v1")
WRITE_MODEL = os.getenv("WRITE_MODEL") or CONFIG.get("write_model", "gpt-5.4")

PAPER_TITLE = CONFIG.get("title", "综述论文")
MIN_TOTAL = int(CONFIG.get("min_total_count", 3000))
MAX_TOTAL = int(CONFIG.get("max_total_count", 7000))
MIN_REFS = int(CONFIG.get("min_references", 2))
MAX_REFS = int(CONFIG.get("max_references", 8))


def log_error(text):
    with ERROR_LOG.open("a", encoding="utf-8") as f:
        f.write(text + "\n")


def read_text(path):
    if not path.exists():
        return ""
    for enc in ["utf-8", "utf-8-sig", "gbk"]:
        try:
            return path.read_text(encoding=enc)
        except Exception:
            pass
    return path.read_text(encoding="utf-8", errors="replace")


def extract_outline_block(text):
    markers = ["大纲：", "大纲:", "文章大纲：", "文章大纲:", "目录：", "目录:"]
    for m in markers:
        if m in text:
            return text.split(m, 1)[1]
    return text


def clean_line(line):
    line = line.strip()
    line = line.replace("　", " ")
    line = re.sub(r"^[\-—•·*]+\s*", "", line)
    line = re.sub(r"^(正文|章节|标题)[:：]\s*", "", line)
    return line.strip()


def is_heading(line):
    if not line:
        return False
    if line in ["摘要", "关键词", "引言", "结语", "结论", "讨论"]:
        return line != "关键词"
    if re.match(r"^\d+(\.\d+)?\s+.+", line):
        return True
    if re.match(r"^[一二三四五六七八九十]+、\s*.+", line):
        return True
    return False


def parse_outline():
    req_text = read_text(REQ_PATH)
    block = extract_outline_block(req_text)
    lines = [clean_line(x) for x in block.splitlines()]
    headings = [x for x in lines if is_heading(x)]

    if len(headings) >= 3:
        return headings

    return [
        "摘要",
        "引言",
        "1 研究背景与发展现状",
        "2 主要问题",
        "3 对策建议",
        "4 讨论与结语",
    ]


def read_matrix():
    if not MATRIX_PATH.exists():
        raise FileNotFoundError(f"没有找到证据矩阵：{MATRIX_PATH}")

    for enc in ["utf-8-sig", "utf-8", "gbk"]:
        try:
            return pd.read_csv(MATRIX_PATH, encoding=enc)
        except Exception:
            pass

    return pd.read_csv(MATRIX_PATH, encoding="utf-8", errors="replace")


def get_cell(row, keywords):
    for col in row.index:
        c = str(col)
        if any(k in c for k in keywords):
            v = row[col]
            if pd.notna(v) and str(v).strip():
                return str(v).strip()
    return ""


def row_to_material(row, idx):
    title = get_cell(row, ["题名", "标题", "文献名称", "论文名称", "资料卡片", "card_title"])
    author = get_cell(row, ["作者"])
    year = get_cell(row, ["年份", "年"])
    source = get_cell(row, ["来源", "期刊", "刊名"])
    section = get_cell(row, ["适合章节", "可支撑", "章节"])
    point = get_cell(row, ["核心观点", "观点", "结论", "核心结论"])
    material = get_cell(row, ["可用材料", "正文材料", "证据", "摘录", "内容"])
    limitation = get_cell(row, ["局限", "使用局限"])

    if not title:
        title = f"文献{idx}"

    all_text = "；".join(
        str(row[col]).strip()
        for col in row.index
        if pd.notna(row[col]) and str(row[col]).strip()
    )

    return {
        "num": idx,
        "title": title,
        "author": author,
        "year": year,
        "source": source,
        "section": section,
        "point": point or all_text[:500],
        "material": material or all_text[:900],
        "limitation": limitation,
        "raw": all_text[:1600],
    }


def build_reference_text(ref):
    parts = []
    if ref["author"]:
        parts.append(ref["author"])
    parts.append(ref["title"])
    if ref["source"]:
        parts.append(ref["source"])
    if ref["year"]:
        parts.append(str(ref["year"]))
    return "，".join(parts)


def load_references():
    df = read_matrix()
    refs = []

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        refs.append(row_to_material(row, i))

    if MAX_REFS > 0:
        refs = refs[:MAX_REFS]

    return refs


def set_doc_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5


def add_para(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(12)


def add_heading(doc, text, level):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(12)


def heading_level(h):
    if h == "摘要":
        return 1
    if h in ["引言", "结语", "结论", "讨论"]:
        return 1
    if re.match(r"^\d+\.\d+\s+", h):
        return 2
    return 1


def strip_heading_num(h):
    h = re.sub(r"^\d+(\.\d+)?\s+", "", h)
    h = re.sub(r"^[一二三四五六七八九十]+、\s*", "", h)
    return h


def keyword_match_score(heading, ref):
    h = strip_heading_num(heading)
    score = 0
    text = (ref["section"] + " " + ref["point"] + " " + ref["material"] + " " + ref["title"])
    for token in re.split(r"[、，,；;：:\s]+", h):
        token = token.strip()
        if len(token) >= 2 and token in text:
            score += 2
    return score


def select_refs_for_heading(heading, refs, quota=4):
    scored = []
    for ref in refs:
        scored.append((keyword_match_score(heading, ref), ref))
    scored.sort(key=lambda x: x[0], reverse=True)

    selected = [r for s, r in scored if s > 0][:quota]
    if not selected:
        selected = refs[:min(quota, len(refs))]
    return selected


def allocate_targets(headings):
    total_target = int((MIN_TOTAL + MAX_TOTAL) / 2)
    weights = []

    for h in headings:
        if h == "摘要":
            weights.append(0.45)
        elif h in ["引言", "结语", "结论", "讨论"]:
            weights.append(0.75)
        elif re.match(r"^\d+\.\d+\s+", h):
            weights.append(0.75)
        else:
            weights.append(1.0)

    s = sum(weights) or 1
    targets = []

    for h, w in zip(headings, weights):
        t = int(total_target * w / s)
        if MAX_TOTAL <= 7000:
            t = max(220, min(t, 650))
        else:
            t = max(260, min(t, 850))
        if h == "摘要":
            t = min(t, 350)
        targets.append(t)

    return targets


def call_model(prompt, model):
    if not API_KEY:
        raise RuntimeError("API_KEY 为空，请在界面填写并保存配置。")

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL, max_retries=5, timeout=120.0)

    last_err = None
    for i in range(1, 4):
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=0.35,
                messages=[
                    {"role": "system", "content": "你是严谨的中文学术论文写作助手。要求语言自然、稳妥，避免夸张和明显AI腔。"},
                    {"role": "user", "content": prompt},
                ],
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            time.sleep(3 * i)
    raise last_err


def build_material_prompt(heading, selected_refs):
    lines = []
    for r in selected_refs:
        lines.append(f"[{r['num']}] {build_reference_text(r)}")
        lines.append(f"可用材料：{r['material'] or r['point'] or r['raw']}")
        if r["limitation"]:
            lines.append(f"使用边界：{r['limitation']}")
        lines.append("")
    return "\n".join(lines)


def generate_section(heading, target, refs):
    selected = select_refs_for_heading(heading, refs, quota=min(4, len(refs)))
    available_nums = ", ".join(f"[{r['num']}]" for r in refs)

    materials = build_material_prompt(heading, selected)
    requirements = read_text(REQ_PATH)

    if heading == "摘要":
        prompt = f"""
论文题目：{PAPER_TITLE}

请写论文摘要，控制在约 {target} 字。
只能使用这些文献编号：{available_nums}。
摘要中可以少量使用数字引文，但不要堆砌。
语言要自然、稳妥，不要夸张。

论文要求：
{requirements}

可用材料：
{materials}
"""
    else:
        prompt = f"""
论文题目：{PAPER_TITLE}
当前章节标题：{heading}
本节目标字数：约 {target} 字，允许上下浮动 15%。

请严格围绕当前标题写作，不要写旧题目的固定大纲。
只能使用这些文献编号：{available_nums}。
凡引用文献观点、案例、数据、结论，请在句末加数字引文，如[1]、[2]。
不要编造不存在的文献编号。
不要写三级标题。
不要输出标题本身，只输出本节正文。

论文要求和大纲：
{requirements}

可用材料：
{materials}
"""
    return call_model(prompt, WRITE_MODEL)


def write_material_package(refs):
    doc = Document()
    set_doc_style(doc)
    add_heading(doc, "综述写作材料包_带编号", 0)

    for r in refs:
        add_heading(doc, f"[{r['num']}] {r['title']}", 1)
        add_para(doc, f"作者：{r['author']}")
        add_para(doc, f"年份：{r['year']}")
        add_para(doc, f"来源：{r['source']}")
        add_para(doc, f"核心观点：{r['point']}")
        add_para(doc, f"可用材料：{r['material'] or r['raw']}")
        add_para(doc, f"使用局限：{r['limitation']}")

    doc.save(MATERIAL_DOCX)


def write_ref_doc(refs):
    doc = Document()
    set_doc_style(doc)
    add_heading(doc, "参考文献清单_初稿", 0)
    for r in refs:
        add_para(doc, f"[{r['num']}] {build_reference_text(r)}")
    doc.save(REF_DOCX)


def write_draft(sections, refs):
    doc = Document()
    set_doc_style(doc)
    add_heading(doc, PAPER_TITLE, 0)

    for heading, text in sections:
        if heading != "摘要":
            add_heading(doc, heading, heading_level(heading))
        else:
            add_heading(doc, "摘要", 1)
        add_para(doc, text)

    add_heading(doc, "参考文献", 1)
    for r in refs:
        add_para(doc, f"[{r['num']}] {build_reference_text(r)}")

    doc.save(DRAFT_DOCX)


def main():
    try:
        headings = parse_outline()
        refs = load_references()

        print(f"已启用界面大纲驱动写作，共识别 {len(headings)} 个写作单元。")
        for h in headings:
            print(f"- {h}")

        print(f"证据矩阵总条数：{len(refs)}")
        print(f"进入写作材料条数：{len(refs)}")

        targets = allocate_targets(headings)

        sections = []
        for heading, target in zip(headings, targets):
            print(f"\n正在写作：{heading}，目标约 {target} 字")
            text = generate_section(heading, target, refs)
            sections.append((heading, text))
            print(f"已写入：{heading}")
            PROGRESS_PATH.write_text("\n".join(h for h, _ in sections), encoding="utf-8")

        write_draft(sections, refs)
        write_material_package(refs)
        write_ref_doc(refs)

        print("\n综述初稿生成完成。")
        print(f"综述初稿：{DRAFT_DOCX}")
        print(f"材料包：{MATERIAL_DOCX}")
        print(f"参考文献清单：{REF_DOCX}")
        print(f"写作进度：{PROGRESS_PATH}")
        print(f"错误日志：{ERROR_LOG}")

    except Exception as e:
        err = f"程序出错：{e}"
        print(err)
        log_error(err)
        sys.exit(1)


if __name__ == "__main__":
    main()
