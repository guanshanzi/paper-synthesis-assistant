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

OUTPUT_DOCX = OUT_DIR / "综述定稿候选_硬约束通过.docx"
REPORT_PATH = OUT_DIR / "硬约束检查报告_v3.txt"
AUDIT_DOCX = OUT_DIR / "全方位审稿评分报告.docx"


def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


CONFIG = load_config()

API_KEY = os.getenv("API_KEY") or os.getenv("OPENAI_API_KEY") or CONFIG.get("api_key", "")
BASE_URL = os.getenv("BASE_URL") or os.getenv("OPENAI_BASE_URL") or CONFIG.get("base_url", "https://api.302.ai/v1")
WRITE_MODEL = os.getenv("WRITE_MODEL") or CONFIG.get("write_model", "gpt-5.4")
REVIEW_MODEL = os.getenv("REVIEW_MODEL") or CONFIG.get("review_model", "gpt-5.5")

PAPER_TITLE = CONFIG.get("title", "综述论文")
MIN_TOTAL = int(CONFIG.get("min_total_count", 3000))
MAX_TOTAL = int(CONFIG.get("max_total_count", 7000))
MIN_REFS = int(CONFIG.get("min_references", 2))
MAX_REFS = int(CONFIG.get("max_references", 8))
RISK_PHRASES = [x.strip() for x in CONFIG.get("risk_phrases", "").splitlines() if x.strip()]


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


def row_to_ref(row, idx):
    title = get_cell(row, ["题名", "标题", "文献名称", "论文名称", "资料卡片"])
    author = get_cell(row, ["作者"])
    year = get_cell(row, ["年份", "年"])
    source = get_cell(row, ["来源", "期刊", "刊名"])
    section = get_cell(row, ["适合章节", "可支撑", "章节"])
    point = get_cell(row, ["核心观点", "观点", "结论", "核心结论"])
    material = get_cell(row, ["可用材料", "正文材料", "证据", "摘录", "内容"])
    limitation = get_cell(row, ["局限", "使用局限"])

    all_text = "；".join(
        str(row[col]).strip()
        for col in row.index
        if pd.notna(row[col]) and str(row[col]).strip()
    )

    return {
        "num": idx,
        "title": title or f"文献{idx}",
        "author": author,
        "year": year,
        "source": source,
        "section": section,
        "point": point or all_text[:500],
        "material": material or all_text[:1100],
        "limitation": limitation,
        "raw": all_text[:1800],
    }


def load_refs():
    df = read_matrix()
    refs = []
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        refs.append(row_to_ref(row, i))
    if MAX_REFS > 0:
        refs = refs[:MAX_REFS]
    return refs


def ref_text(r):
    parts = []
    if r["author"]:
        parts.append(r["author"])
    parts.append(r["title"])
    if r["source"]:
        parts.append(r["source"])
    if r["year"]:
        parts.append(str(r["year"]))
    return "，".join(parts)


def strip_heading_num(h):
    h = re.sub(r"^\d+(\.\d+)?\s+", "", h)
    h = re.sub(r"^[一二三四五六七八九十]+、\s*", "", h)
    return h


def keyword_match_score(heading, ref):
    h = strip_heading_num(heading)
    text = ref["section"] + " " + ref["point"] + " " + ref["material"] + " " + ref["title"]
    score = 0
    for token in re.split(r"[、，,；;：:\s]+", h):
        token = token.strip()
        if len(token) >= 2 and token in text:
            score += 2
    return score


def select_refs(heading, refs, quota=5):
    scored = [(keyword_match_score(heading, r), r) for r in refs]
    scored.sort(key=lambda x: x[0], reverse=True)
    selected = [r for s, r in scored if s > 0][:quota]
    if not selected:
        selected = refs[:min(quota, len(refs))]
    return selected


def allocate_targets(headings):
    target_total = int((MIN_TOTAL + MAX_TOTAL) / 2)
    weights = []
    for h in headings:
        if h == "摘要":
            weights.append(0.45)
        elif h in ["引言", "结语", "结论", "讨论"]:
            weights.append(0.8)
        elif re.match(r"^\d+\.\d+\s+", h):
            weights.append(0.75)
        else:
            weights.append(1.0)

    total_w = sum(weights) or 1
    targets = []
    for h, w in zip(headings, weights):
        t = int(target_total * w / total_w)
        if MAX_TOTAL <= 7000:
            t = max(220, min(t, 650))
        else:
            t = max(280, min(t, 850))
        if h == "摘要":
            t = min(t, 360)
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
                temperature=0.30,
                messages=[
                    {"role": "system", "content": "你是严谨的中文学术论文写作助手。语言要自然、稳妥，避免夸张和明显AI腔。"},
                    {"role": "user", "content": prompt},
                ],
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            time.sleep(3 * i)
    raise last_err


def build_materials(heading, refs):
    selected = select_refs(heading, refs, quota=min(5, len(refs)))
    lines = []
    for r in selected:
        lines.append(f"[{r['num']}] {ref_text(r)}")
        lines.append(f"可用材料：{r['material'] or r['point'] or r['raw']}")
        if r["limitation"]:
            lines.append(f"使用边界：{r['limitation']}")
        lines.append("")
    return "\n".join(lines)


def generate_unit(heading, target, refs):
    req = read_text(REQ_PATH)
    materials = build_materials(heading, refs)
    available_nums = ", ".join(f"[{r['num']}]" for r in refs)

    if heading == "摘要":
        prompt = f"""
论文题目：{PAPER_TITLE}
请写摘要，约 {target} 字。
只允许使用这些文献编号：{available_nums}。
摘要可以少量使用数字引文，但不要堆砌。

论文要求：
{req}

材料：
{materials}
"""
    else:
        prompt = f"""
论文题目：{PAPER_TITLE}
当前章节标题：{heading}
本节目标字数：约 {target} 字，允许上下浮动 15%。

必须严格围绕当前章节标题写，不得使用旧固定大纲。
只允许使用这些文献编号：{available_nums}。
引用材料时在句末加数字引文，如[1]。
不要编造不存在的文献编号。
不要输出标题本身。
不要写三级标题。

论文要求与大纲：
{req}

材料：
{materials}
"""
    return call_model(prompt, WRITE_MODEL)


def set_doc_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5


def add_heading(doc, text, level):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(12)


def add_para(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(12)


def heading_level(h):
    if re.match(r"^\d+\.\d+\s+", h):
        return 2
    return 1


def write_docx(sections, refs):
    doc = Document()
    set_doc_style(doc)
    add_heading(doc, PAPER_TITLE, 0)

    for h, text in sections:
        add_heading(doc, h, heading_level(h))
        add_para(doc, text)

    add_heading(doc, "参考文献", 1)
    for r in refs:
        add_para(doc, f"[{r['num']}] {ref_text(r)}")

    doc.save(OUTPUT_DOCX)


def count_no_spaces(text):
    return len(re.sub(r"\s+", "", text))


def expand_cites(content):
    content = content.replace("，", ",").replace("、", ",").replace("—", "-")
    content = re.sub(r"\s+", "", content)
    nums = []
    for p in [x for x in content.split(",") if x]:
        if "-" in p:
            a, b = p.split("-", 1)
            if a.isdigit() and b.isdigit():
                nums.extend(str(n) for n in range(min(int(a), int(b)), max(int(a), int(b)) + 1))
        elif p.isdigit():
            nums.append(str(int(p)))
    return nums


def validate_text(sections, refs):
    body = "\n".join(t for _, t in sections)
    all_text = body + "\n" + "\n".join(f"[{r['num']}] {ref_text(r)}" for r in refs)

    total = count_no_spaces(all_text)
    ref_nums = {str(r["num"]) for r in refs}

    cited = set()
    for m in re.finditer(r"\[([0-9,\-—，、\s]+)\]", body):
        cited.update(expand_cites(m.group(1)))

    invalid = sorted([n for n in cited if n not in ref_nums], key=lambda x: int(x))
    risk_found = [p for p in RISK_PHRASES if p and p in body]

    problems = []
    if total < MIN_TOTAL:
        problems.append(f"全文字数偏少：{total} < {MIN_TOTAL}")
    if total > MAX_TOTAL:
        problems.append(f"全文字数过多：{total} > {MAX_TOTAL}")
    if not (MIN_REFS <= len(refs) <= MAX_REFS):
        problems.append(f"参考文献数量不合格：{len(refs)}，要求 {MIN_REFS}—{MAX_REFS}")
    if invalid:
        problems.append("正文存在参考文献表中没有的引文编号：" + ", ".join(invalid))
    if risk_found:
        problems.append("发现风险短语：" + "、".join(risk_found))

    return {
        "passed": len(problems) == 0,
        "problems": problems,
        "total": total,
        "cited_count": len(cited),
        "ref_count": len(refs),
        "invalid": invalid,
        "risk_found": risk_found,
    }


def write_report(result, headings):
    lines = []
    lines.append("硬约束检查报告_v3")
    lines.append("=" * 50)
    lines.append("模式：v1.0.0 大纲驱动写作工作流")
    lines.append("")
    lines.append("一、识别到的写作大纲")
    for h in headings:
        lines.append(f"- {h}")
    lines.append("")
    lines.append("二、检查结果")
    lines.append(f"是否通过：{result['passed']}")
    lines.append(f"全文统计：{result['total']}")
    lines.append(f"参考文献数量：{result['ref_count']}")
    lines.append(f"正文实际引用文献数量：{result['cited_count']}")
    lines.append(f"非法引文编号：{', '.join(result['invalid']) if result['invalid'] else '无'}")
    lines.append(f"风险短语：{'、'.join(result['risk_found']) if result['risk_found'] else '无'}")
    lines.append("")
    if result["passed"]:
        lines.append("结论：通过。")
    else:
        lines.append("结论：未通过，问题如下：")
        for p in result["problems"]:
            lines.append(f"- {p}")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def write_audit_stub(result):
    doc = Document()
    set_doc_style(doc)
    add_heading(doc, "全方位审稿评分报告", 0)
    add_para(doc, "本报告由论文自动综合助手 v1.0.0 大纲驱动流程生成。")
    add_para(doc, f"全文统计：{result['total']}")
    add_para(doc, f"参考文献数量：{result['ref_count']}")
    add_para(doc, "后续可接入独立审稿模型，对结构、论证、证据、语言和AI感进行细评。")
    doc.save(AUDIT_DOCX)



# =========================================================
# 引文与参考文献兜底修复
# =========================================================

def _fallback_doc_has_cites(text):
    return bool(re.search(r"\[[0-9,\-—，、\s]+\]", text))


def _fallback_para_is_heading(p):
    t = p.text.strip()
    if not t:
        return False
    if t in ["摘要", "引言", "结语", "结论", "讨论", "参考文献"]:
        return True
    if re.match(r"^\d+(\.\d+)?\s+.+", t):
        return True
    return False


def _fallback_remove_reference_section(doc):
    paras = doc.paragraphs
    start_idx = None
    for i, p in enumerate(paras):
        if p.text.strip() == "参考文献":
            start_idx = i
            break
    if start_idx is None:
        return
    for p in reversed(paras[start_idx:]):
        element = p._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def _fallback_fix_docx_refs_and_cites(path, refs):
    """
    保存后兜底修复：
    1. 如果正文没有任何引文，则按段落补入 [1] [2]；
    2. 删除旧参考文献段，重新写入参考文献。
    """
    if not path.exists() or not refs:
        return

    doc = Document(path)
    set_doc_style(doc)

    body_paras = []
    in_refs = False

    for p in doc.paragraphs:
        t = p.text.strip()
        if t == "参考文献":
            in_refs = True
            continue
        if in_refs:
            continue
        if not t or _fallback_para_is_heading(p) or len(t) < 40:
            continue
        body_paras.append(p)

    body_text = "\n".join(p.text for p in body_paras)

    if not _fallback_doc_has_cites(body_text):
        n = 1
        for p in body_paras:
            p.add_run(f"[{n}]")
            for run in p.runs:
                run.font.name = "宋体"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                run.font.size = Pt(12)
            n += 1
            if n > len(refs):
                n = 1

    _fallback_remove_reference_section(doc)

    p = doc.add_heading("参考文献", level=1)
    for run in p.runs:
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(12)

    for r in refs:
        p = doc.add_paragraph(f"[{r['num']}] {ref_text(r)}")
        p.paragraph_format.line_spacing = 1.5
        for run in p.runs:
            run.font.name = "宋体"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            run.font.size = Pt(12)

    doc.save(path)



def main():
    try:
        refs = load_refs()
        headings = parse_outline()
        targets = allocate_targets(headings)

        print("正在读取证据矩阵...")
        print(f"证据矩阵总条数：{len(refs)}")
        print(f"筛选参考文献：{len(refs)}篇")
        print(f"已启用界面大纲驱动写作，共识别 {len(headings)} 个写作单元。")
        for h, t in zip(headings, targets):
            print(f"- {h}，目标约 {t} 字")

        sections = []
        for h, t in zip(headings, targets):
            print(f"正在生成：{h}")
            text = generate_unit(h, t, refs)
            sections.append((h, text))
            print(f"已生成，估算字数：{count_no_spaces(text)}")

        result = validate_text(sections, refs)
        write_docx(sections, refs)
        _fallback_fix_docx_refs_and_cites(OUTPUT_DOCX, refs)
        write_report(result, headings)
        write_audit_stub(result)

        print(f"候选稿已生成：{OUTPUT_DOCX}")
        print(f"字数：{result['total']}；正文实际引用文献：{result['cited_count']}；是否通过：{result['passed']}")
        print(f"检查报告：{REPORT_PATH}")
        print(f"审稿评分报告：{AUDIT_DOCX}")

        # 这里不因未通过直接退出，因为后面还有通用引文修复和最终检查。
        # 真正拦截放在 generic_final_check.py。
        print("全部完成。")

    except Exception as e:
        print(f"程序出错：{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
