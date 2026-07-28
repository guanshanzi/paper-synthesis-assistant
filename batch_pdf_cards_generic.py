import json
import os
import re
import sys
import time
import csv
from pathlib import Path

import fitz
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI


BASE_DIR = Path(__file__).parent
PDF_DIR = Path(os.getenv("PDF_DIR") or BASE_DIR / "pdfs")
OUT_DIR = Path(os.getenv("OUTPUT_DIR") or BASE_DIR / "output")
OUT_DIR.mkdir(exist_ok=True)

CONFIG_PATH = BASE_DIR / "config.local.json"
CARD_PROMPT_PATH = BASE_DIR / "资料卡片指令.txt"

RESULT_CSV = OUT_DIR / "第一轮筛选结果.csv"
CARD_DOCX = OUT_DIR / "论文资料卡片汇总.docx"
PROGRESS_PATH = OUT_DIR / "progress.txt"
ERROR_LOG = OUT_DIR / "error_log.txt"

MAX_PDF_CHARS = 18000
SECOND_ROUND_THRESHOLD = 3


def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


CONFIG = load_config()

API_KEY = (
    os.getenv("API_KEY")
    or os.getenv("OPENAI_API_KEY")
    or CONFIG.get("api_key", "")
)

BASE_URL = (
    os.getenv("BASE_URL")
    or os.getenv("OPENAI_BASE_URL")
    or CONFIG.get("base_url", "https://api.302.ai/v1")
)

CARD_MODEL = (
    os.getenv("CARD_MODEL")
    or CONFIG.get("card_model", "gpt-5.4")
)


def read_text_file(path: Path):
    if not path.exists():
        return ""
    for enc in ["utf-8", "utf-8-sig", "gbk"]:
        try:
            return path.read_text(encoding=enc)
        except Exception:
            pass
    return path.read_text(encoding="utf-8", errors="replace")


def get_card_prompt():
    text = read_text_file(CARD_PROMPT_PATH).strip()
    if text:
        return text

    text = CONFIG.get("card_prompt", "").strip()
    if text:
        return text

    return "请阅读 PDF，并整理成服务于后续综述写作的论文资料卡片。"


CARD_PROMPT = get_card_prompt()


def log_error(msg):
    with ERROR_LOG.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def read_done_files():
    if not PROGRESS_PATH.exists():
        return set()
    return set(
        x.strip()
        for x in PROGRESS_PATH.read_text(encoding="utf-8", errors="replace").splitlines()
        if x.strip()
    )


def mark_done(filename):
    with PROGRESS_PATH.open("a", encoding="utf-8") as f:
        f.write(filename + "\n")


def extract_pdf_text(pdf_path: Path, max_chars=MAX_PDF_CHARS):
    texts = []
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            txt = page.get_text("text")
            if txt:
                texts.append(txt)
            if sum(len(x) for x in texts) >= max_chars:
                break
        doc.close()
    except Exception as e:
        log_error(f"读取 PDF 失败：{pdf_path.name}；{e}")
        return ""

    text = "\n".join(texts)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text[:max_chars]


def call_model(prompt, model=CARD_MODEL):
    if not API_KEY:
        raise RuntimeError("API_KEY 为空，请在界面填写并保存配置。")

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL, max_retries=5, timeout=120.0)

    last_err = None
    for i in range(1, 4):
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=0.25,
                messages=[
                    {
                        "role": "system",
                        "content": "你是严谨的中文学术文献筛选与资料卡片整理助手。必须严格按照用户给定选题和资料卡片指令判断相关性，不得沿用旧主题。"
                    },
                    {"role": "user", "content": prompt},
                ],
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            time.sleep(3 * i)

    raise last_err


def parse_json_loose(text):
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass

    m = re.search(r"\{[\s\S]*\}", text)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass

    return None


def first_round_screen(pdf_name, pdf_text):
    prompt = f"""
请根据“本次资料卡片指令”判断这篇 PDF 是否适合进入资料卡片整理。

特别强调：
1. 必须只根据下面的“本次资料卡片指令”判断相关性；
2. 不允许沿用任何旧题目、旧关键词或旧主题；
3. 如果本次指令是红曲主题，就按红曲降脂、安全性、质量控制、合理用药来判断；
4. 不要出现与本次指令无关的判断理由，例如“耕地非粮化、土地资源约束、中药材种植产业”等，除非本次指令本身明确要求。

本次资料卡片指令：
{CARD_PROMPT}

PDF 文件名：
{pdf_name}

PDF 文本节选：
{pdf_text[:12000]}

请只输出 JSON，不要输出其他解释：
{{
  "score": 0到5的整数,
  "direction": "相关方向，多个方向用/分隔",
  "enter_second_round": true或false,
  "reason": "判断理由，必须围绕本次指令"
}}

评分标准：
0 = 完全无关；
1 = 只有极弱背景关系；
2 = 可作为一般背景；
3 = 有一定支撑价值，建议进入资料卡片；
4 = 重要相关文献；
5 = 核心文献。
"""
    ans = call_model(prompt)
    data = parse_json_loose(ans)

    if not data:
        return {
            "score": 0,
            "direction": "解析失败",
            "enter_second_round": False,
            "reason": "模型未按 JSON 返回，暂不进入第二轮。原始返回：" + ans[:300],
        }

    try:
        score = int(data.get("score", 0))
    except Exception:
        score = 0

    score = max(0, min(score, 5))

    enter = data.get("enter_second_round", score >= SECOND_ROUND_THRESHOLD)
    if isinstance(enter, str):
        enter = enter.lower() in ["true", "yes", "是", "1"]

    # 双保险：分数达到阈值即进入第二轮
    enter = bool(enter or score >= SECOND_ROUND_THRESHOLD)

    return {
        "score": score,
        "direction": str(data.get("direction", "")).strip() or "未说明",
        "enter_second_round": enter,
        "reason": str(data.get("reason", "")).strip() or "未说明",
    }


def generate_card(pdf_name, pdf_text, screen):
    prompt = f"""
请根据以下 PDF 内容和本次资料卡片指令，整理成“论文资料卡片”。

本次资料卡片指令：
{CARD_PROMPT}

PDF 文件名：
{pdf_name}

第一轮筛选结果：
相关度评分：{screen['score']}
相关方向：{screen['direction']}
判断理由：{screen['reason']}

PDF 文本节选：
{pdf_text[:16000]}

请生成资料卡片。要求：
1. 必须服务于本次选题，不要泛泛摘要；
2. 必须提取可用于后续综述写作的证据；
3. 所有材料都要有使用边界，不能过度推出；
4. 如果能识别题名、作者、年份、来源，请尽量补全；
5. 输出中必须保留固定小标题“论文基本信息”，以便后续程序拆分；
6. 不要使用 Markdown 表格。

请按以下格式输出：

论文基本信息
题名：
作者：
年份：
来源：
文件名：
相关度评分：
使用优先级：
适合章节：

研究对象与核心问题：
与本文主题相关的关键观点：
可用于正文写作的材料：
可支撑的论证部分：
安全性、有效性或质量控制相关信息：
使用局限：
对本文的价值判断：
"""
    return call_model(prompt)


def init_docx():
    doc = Document()
    set_doc_style(doc)
    p = doc.add_heading("论文资料卡片汇总", level=0)
    set_para_font(p)
    p = doc.add_paragraph("以下内容为批量处理 PDF 后生成的论文资料卡片。")
    set_para_font(p)
    return doc


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


def append_card_to_docx(doc, card_text):
    p = doc.add_paragraph("")
    set_para_font(p)

    p = doc.add_heading("论文基本信息", level=1)
    set_para_font(p)

    # 如果模型已经输出“论文基本信息”，避免重复太怪，删掉第一处
    card_text = re.sub(r"^\s*论文基本信息\s*", "", card_text.strip())

    for line in card_text.splitlines():
        line = line.strip()
        if not line:
            continue

        if re.match(r"^(题名|作者|年份|来源|文件名|相关度评分|使用优先级|适合章节|研究对象与核心问题|与本文主题相关的关键观点|可用于正文写作的材料|可支撑的论证部分|安全性、有效性或质量控制相关信息|使用局限|对本文的价值判断)[:：]", line):
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.bold = True
            set_para_font(p)
        else:
            p = doc.add_paragraph(line)
            set_para_font(p)


def write_csv_header_if_needed():
    if RESULT_CSV.exists():
        return

    with RESULT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["文件名", "相关度评分", "相关方向", "是否进入第二轮", "判断理由"])


def append_csv_row(pdf_name, screen):
    write_csv_header_if_needed()
    with RESULT_CSV.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            pdf_name,
            screen["score"],
            screen["direction"],
            "是" if screen["enter_second_round"] else "否",
            screen["reason"],
        ])


def main():
    print("请输入本次资料卡片处理指令。")
    print("可以粘贴多行。输入完成后，单独输入一行 END，然后回车。")
    print("--------------------------------------------------")

    # 兼容 run_pipeline 自动传入。这里实际优先用 资料卡片指令.txt / config.local.json。
    stdin_lines = []
    try:
        for line in sys.stdin:
            if line.strip() == "END":
                break
            stdin_lines.append(line.rstrip("\n"))
    except Exception:
        pass

    global CARD_PROMPT
    stdin_prompt = "\n".join(stdin_lines).strip()
    if stdin_prompt:
        CARD_PROMPT = stdin_prompt

    pdfs = sorted(PDF_DIR.glob("*.pdf"))

    print()
    print(f"共发现 {len(pdfs)} 篇 PDF。")

    done = read_done_files()
    print(f"已完成 {len(done)} 篇，将自动跳过。")
    print("--------------------------------------------------")

    if CARD_DOCX.exists():
        try:
            doc = Document(CARD_DOCX)
            set_doc_style(doc)
        except Exception:
            doc = init_docx()
    else:
        doc = init_docx()

    generated_cards = 0

    for i, pdf in enumerate(pdfs, start=1):
        if pdf.name in done:
            print(f"跳过已完成：{pdf.name}")
            continue

        print()
        print(f"正在处理第 {i}/{len(pdfs)} 篇：{pdf.name}")

        try:
            pdf_text = extract_pdf_text(pdf)

            if len(pdf_text.strip()) < 200:
                screen = {
                    "score": 0,
                    "direction": "文本不足",
                    "enter_second_round": False,
                    "reason": "PDF 可读取文本过少，无法可靠整理资料卡片。",
                }
            else:
                screen = first_round_screen(pdf.name, pdf_text)

            print(f"第一轮评分：{screen['score']}")
            print(f"相关方向：{screen['direction']}")
            print(f"判断理由：{screen['reason']}")

            append_csv_row(pdf.name, screen)

            if screen["enter_second_round"]:
                card = generate_card(pdf.name, pdf_text, screen)
                append_card_to_docx(doc, card)
                doc.save(CARD_DOCX)
                generated_cards += 1
                print("已生成资料卡片并写入 Word。")
            else:
                print("相关度较低，不进入第二轮。")

            mark_done(pdf.name)

        except Exception as e:
            msg = f"处理失败：{pdf.name}；{e}"
            print(msg)
            log_error(msg)

    doc.save(CARD_DOCX)

    print()
    print("处理结束。")
    print(f"本次生成资料卡片数量：{generated_cards}")
    print(f"筛选结果：{RESULT_CSV}")
    print(f"资料卡片 Word：{CARD_DOCX}")
    print(f"进度记录：{PROGRESS_PATH}")
    print(f"错误日志：{ERROR_LOG}")


if __name__ == "__main__":
    main()
