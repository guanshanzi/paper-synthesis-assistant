import json
import re
from pathlib import Path
from docx import Document


BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.local.json"

DEFAULT_OUTPUT = BASE_DIR / "output"


def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


def find_latest_docx(output_dir: Path):
    candidates = [
        "综述最终稿_字数修正版.docx",
        "综述最终稿_通用引文修复版.docx",
        "综述最终稿_V10_引文顺序修正版.docx",
        "综述最终稿_V9_标点修正版.docx",
        "综述最终稿_V8.docx",
        "综述定稿候选_V7_小修定稿.docx",
        "综述定稿候选_V6_微修稿.docx",
        "综述定稿候选_V5_中修稿.docx",
        "综述定稿候选_V4_补齐一级标题.docx",
        "综述定稿候选_硬约束通过.docx",
    ]

    for name in candidates:
        path = output_dir / name
        if path.exists():
            return path

    docxs = sorted(output_dir.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if docxs:
        return docxs[0]

    return None


def read_docx_texts(path):
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def count_text_no_spaces(text):
    return len(re.sub(r"\s+", "", text))


def split_body_and_refs(texts):
    body = []
    refs = []
    in_refs = False

    for t in texts:
        if t == "参考文献":
            in_refs = True
            continue

        if in_refs:
            refs.append(t)
        else:
            body.append(t)

    return "\n".join(body), refs


def is_reference_item(text):
    return bool(re.match(r"^\[\d+\]", text.strip()))


def extract_reference_numbers(refs):
    nums = set()
    for r in refs:
        m = re.match(r"^\[(\d+)\]", r.strip())
        if m:
            nums.add(m.group(1))
    return nums


def expand_cite_content(content):
    content = content.replace("，", ",").replace("、", ",").replace("—", "-")
    content = re.sub(r"\s+", "", content)

    nums = []

    for p in [x for x in content.split(",") if x]:
        if "-" in p:
            a, b = p.split("-", 1)
            if a.isdigit() and b.isdigit():
                for n in range(min(int(a), int(b)), max(int(a), int(b)) + 1):
                    nums.append(str(n))
        elif p.isdigit():
            nums.append(str(int(p)))

    return nums


def extract_cited_numbers(text):
    cited = set()
    for m in re.finditer(r"\[([0-9,\-—，、\s]+)\]", text):
        for n in expand_cite_content(m.group(1)):
            cited.add(n)
    return cited


def first_citation_number(text):
    m = re.search(r"\[([0-9,\-—，、\s]+)\]", text)
    if not m:
        return None
    nums = expand_cite_content(m.group(1))
    return nums[0] if nums else None


def has_third_or_deeper_heading(texts):
    return [t for t in texts if re.match(r"^\d+\.\d+\.\d+(\.\d+)*\s+.+", t.strip())]


def has_chinese_heading(texts):
    bad = []
    for t in texts:
        if re.match(r"^[一二三四五六七八九十]+、", t.strip()):
            bad.append(t)
        elif re.match(r"^（[一二三四五六七八九十]+）", t.strip()):
            bad.append(t)
    return bad


def check_risk_phrases(body_text, risk_text):
    risks = []
    phrases = [x.strip() for x in risk_text.splitlines() if x.strip()]

    for phrase in phrases:
        if phrase in body_text:
            risks.append(phrase)

    return risks


def check_docx(path, config):
    texts = read_docx_texts(path)
    all_text = "\n".join(texts)
    body_text, refs = split_body_and_refs(texts)

    total_count = count_text_no_spaces(all_text)
    body_count = count_text_no_spaces(body_text)

    min_total = int(config.get("min_total_count", 7500))
    max_total = int(config.get("max_total_count", 8500))
    min_refs = int(config.get("min_references", 30))
    max_refs = int(config.get("max_references", 35))

    ref_nums = extract_reference_numbers(refs)
    cited_nums = extract_cited_numbers(body_text)

    ref_count = len(ref_nums)
    cited_count = len(cited_nums)

    invalid_cites = sorted([n for n in cited_nums if n not in ref_nums], key=lambda x: int(x))
    uncited_refs = sorted([n for n in ref_nums if n not in cited_nums], key=lambda x: int(x))

    third_headings = has_third_or_deeper_heading(texts)
    chinese_headings = has_chinese_heading(texts)

    first_cite = first_citation_number(body_text)

    ref_sequence_ok = True
    cited_sequence_ok = True

    if ref_nums:
        ref_sequence_ok = sorted([int(n) for n in ref_nums]) == list(range(1, ref_count + 1))

    if cited_nums:
        cited_sequence_ok = sorted([int(n) for n in cited_nums]) == list(range(1, cited_count + 1))

    risk_found = check_risk_phrases(body_text, config.get("risk_phrases", ""))

    problems = []

    if total_count < min_total:
        problems.append(f"全文字数偏少：{total_count} < {min_total}")

    if total_count > max_total:
        problems.append(f"全文字数过多：{total_count} > {max_total}")

    if ref_count < min_refs or ref_count > max_refs:
        problems.append(f"参考文献数量不合格：{ref_count}，要求 {min_refs}—{max_refs}")

    if invalid_cites:
        problems.append("正文存在参考文献表中没有的引文编号：" + ", ".join(invalid_cites))

    if uncited_refs:
        problems.append("参考文献列出但正文未引用：" + ", ".join(uncited_refs))

    if config.get("forbid_third_heading", True) and third_headings:
        problems.append("存在三级或更深标题")

    heading_style = config.get("heading_style", "1 / 1.1")
    if "1" in heading_style and chinese_headings:
        problems.append("存在中文编号标题，不符合当前标题格式设置")

    if config.get("require_reference_order", True):
        if first_cite != "1":
            problems.append(f"正文第一个引文不是[1]，而是[{first_cite}]")
        if not ref_sequence_ok:
            problems.append("参考文献编号不是从[1]开始连续排列")
        if not cited_sequence_ok:
            problems.append("正文引文编号不是从[1]开始连续覆盖")

    if risk_found:
        problems.append("发现风险短语：" + "、".join(risk_found))

    return {
        "passed": len(problems) == 0,
        "problems": problems,
        "path": str(path),
        "total_count": total_count,
        "body_count": body_count,
        "ref_count": ref_count,
        "cited_count": cited_count,
        "first_cite": first_cite,
        "invalid_cites": invalid_cites,
        "uncited_refs": uncited_refs,
        "third_headings": third_headings,
        "chinese_headings": chinese_headings,
        "ref_sequence_ok": ref_sequence_ok,
        "cited_sequence_ok": cited_sequence_ok,
        "risk_found": risk_found,
    }


def write_report(result, output_dir: Path):
    report_path = output_dir / "通用最终检查报告.txt"

    lines = []
    lines.append("通用最终检查报告")
    lines.append("=" * 50)
    lines.append(f"检查文件：{result['path']}")
    lines.append("")
    lines.append("一、字数检查")
    lines.append(f"全文去空白字符数：{result['total_count']}")
    lines.append(f"正文去空白字符数：{result['body_count']}")
    lines.append("")
    lines.append("二、参考文献与引文检查")
    lines.append(f"参考文献数量：{result['ref_count']}")
    lines.append(f"正文实际引用文献数量：{result['cited_count']}")
    lines.append(f"正文第一个引文编号：[{result['first_cite']}]")
    lines.append(f"参考文献编号连续：{'是' if result['ref_sequence_ok'] else '否'}")
    lines.append(f"正文引文编号连续覆盖：{'是' if result['cited_sequence_ok'] else '否'}")
    lines.append(f"非法引文编号：{', '.join(result['invalid_cites']) if result['invalid_cites'] else '无'}")
    lines.append(f"参考文献列出但正文未引用：{', '.join(result['uncited_refs']) if result['uncited_refs'] else '无'}")
    lines.append("")
    lines.append("三、标题检查")
    lines.append(f"三级或更深标题：{'有' if result['third_headings'] else '无'}")
    lines.append(f"中文编号标题：{'有' if result['chinese_headings'] else '无'}")
    lines.append("")
    lines.append("四、风险短语检查")
    lines.append(f"发现风险短语：{'、'.join(result['risk_found']) if result['risk_found'] else '无'}")
    lines.append("")
    lines.append("五、结论")

    if result["passed"]:
        lines.append("通过：该稿件符合当前界面设置的通用硬约束。")
    else:
        lines.append("未通过，问题如下：")
        for p in result["problems"]:
            lines.append(f"- {p}")

    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def main():
    config = load_config()
    output_dir = Path(config.get("output_dir", str(DEFAULT_OUTPUT)))
    output_dir.mkdir(exist_ok=True)

    target = find_latest_docx(output_dir)

    if not target:
        print("没有找到可检查的 Word 文件。")
        raise SystemExit(1)

    print(f"正在执行通用最终检查：{target}")

    result = check_docx(target, config)
    report = write_report(result, output_dir)

    print(f"检查报告：{report}")
    print(f"是否通过：{result['passed']}")

    if not result["passed"]:
        print("问题：")
        for p in result["problems"]:
            print(f"- {p}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()