import re
import sys
from pathlib import Path

import pandas as pd
from docx import Document


BASE_DIR = Path(__file__).parent
OUT_DIR = BASE_DIR / "output"
CARDS_DIR = BASE_DIR / "cards_input"

CARD_DOCX_1 = OUT_DIR / "论文资料卡片汇总.docx"
CARD_DOCX_2 = CARDS_DIR / "新资料卡片汇总.docx"
MATRIX_CSV = OUT_DIR / "证据矩阵.csv"
REPORT_PATH = OUT_DIR / "证据门控检查报告.txt"


MIN_VALID_REFS = 2
MIN_MATERIAL_CHARS = 80


def read_docx_text(path: Path):
    if not path.exists():
        return ""
    doc = Document(path)
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())


def read_csv_any(path: Path):
    if not path.exists():
        return None

    for enc in ["utf-8-sig", "utf-8", "gbk"]:
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            pass

    return None


def get_cell(row, keys):
    for col in row.index:
        c = str(col)
        if any(k in c for k in keys):
            v = row[col]
            if pd.notna(v) and str(v).strip():
                return str(v).strip()
    return ""


def row_value_join(row):
    values = []
    for col in row.index:
        v = row[col]
        if pd.notna(v) and str(v).strip():
            values.append(str(v).strip())
    return "\n".join(values)


def judge_row(row):
    title = get_cell(row, ["题名", "标题", "文献名称", "论文名称", "资料卡片"])
    author = get_cell(row, ["作者"])
    year = get_cell(row, ["年份", "年"])
    source = get_cell(row, ["来源", "期刊", "刊名"])
    point = get_cell(row, ["核心观点", "观点", "结论", "核心结论"])
    material = get_cell(row, ["可用材料", "正文材料", "证据", "摘录", "内容", "材料"])
    limitation = get_cell(row, ["局限", "使用局限"])

    all_text = row_value_join(row)

    # 兜底只用于“判断是否有内容”，不是用于生成参考文献
    if not material and len(all_text) >= MIN_MATERIAL_CHARS:
        material = all_text

    problems = []

    if not title:
        problems.append("缺少题名/文献名称")
    if not year:
        problems.append("缺少年份")
    if not point and not material:
        problems.append("缺少核心观点或可用材料")
    if material and len(material) < MIN_MATERIAL_CHARS:
        problems.append(f"可用材料过短：{len(material)} 字")

    valid = len(problems) == 0

    return {
        "valid": valid,
        "title": title,
        "author": author,
        "year": year,
        "source": source,
        "point": point,
        "material": material,
        "limitation": limitation,
        "problems": problems,
    }


def main():
    lines = []
    lines.append("证据门控检查报告")
    lines.append("=" * 50)

    card_text_1 = read_docx_text(CARD_DOCX_1)
    card_text_2 = read_docx_text(CARD_DOCX_2)

    lines.append("一、资料卡片检查")
    lines.append(f"output 资料卡片存在：{CARD_DOCX_1.exists()}")
    lines.append(f"cards_input 资料卡片存在：{CARD_DOCX_2.exists()}")
    lines.append(f"output 资料卡片文本长度：{len(card_text_1)}")
    lines.append(f"cards_input 资料卡片文本长度：{len(card_text_2)}")

    if len(card_text_1) < 200 and len(card_text_2) < 200:
        lines.append("")
        lines.append("结论：未通过。资料卡片内容过少，不能进入写作。")
        REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        sys.exit(1)

    lines.append("")
    lines.append("二、证据矩阵检查")

    df = read_csv_any(MATRIX_CSV)

    if df is None or df.empty:
        lines.append("证据矩阵不存在或为空。")
        lines.append("")
        lines.append("结论：未通过。没有证据矩阵，不能进入写作。")
        REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        sys.exit(1)

    lines.append(f"证据矩阵行数：{len(df)}")
    lines.append(f"证据矩阵列名：{', '.join(str(c) for c in df.columns)}")

    valid_rows = []
    invalid_rows = []

    for idx, row in df.iterrows():
        result = judge_row(row)
        if result["valid"]:
            valid_rows.append((idx + 1, result))
        else:
            invalid_rows.append((idx + 1, result))

    lines.append("")
    lines.append("三、有效证据检查")
    lines.append(f"有效文献数：{len(valid_rows)}")
    lines.append(f"无效文献数：{len(invalid_rows)}")

    if valid_rows:
        lines.append("")
        lines.append("有效文献示例：")
        for no, r in valid_rows[:5]:
            lines.append(f"- 第{no}行：{r['title']}，年份：{r['year']}，材料长度：{len(r['material'])}")

    if invalid_rows:
        lines.append("")
        lines.append("无效文献问题：")
        for no, r in invalid_rows[:10]:
            title = r["title"] or "未识别题名"
            lines.append(f"- 第{no}行：{title}；问题：{'；'.join(r['problems'])}")

    if len(valid_rows) < MIN_VALID_REFS:
        lines.append("")
        lines.append(f"结论：未通过。有效文献少于 {MIN_VALID_REFS} 篇，不能进入写作。")
        REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        sys.exit(1)

    lines.append("")
    lines.append("结论：通过。证据矩阵具备最低写作条件，可以进入大纲驱动写作。")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))


if __name__ == "__main__":
    main()