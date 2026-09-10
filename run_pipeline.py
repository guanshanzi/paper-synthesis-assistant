import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime


BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.local.json"

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

PIPELINE_LOG = LOG_DIR / f"pipeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"


STAGES = [
    {
        "name": "01 PDF生成资料卡片",
        "script": "batch_pdf_cards_generic.py",
        "required": True,
        "stdin_key": "card_prompt",
        "max_retries": None,
    },
    {
        "name": "02 资料卡片生成证据矩阵",
        "script": "cards_to_matrix.py",
        "required": True,
        "stdin_key": None,
        "max_retries": None,
    },
    {
    "name": "02.5 证据门控检查",
    "script": "validate_evidence_gate.py",
    "required": True,
    "stdin_key": None,
    "max_retries": 1,
    },
    {
        "name": "03 按界面大纲写初稿",
        "script": "write_review_from_outline.py",
        "required": True,
        "stdin_key": None,
        "max_retries": None,
    },
    {
        "name": "04 按界面大纲生成候选稿",
        "script": "finalize_outline_docx.py",
        "required": True,
        "stdin_key": None,
        "max_retries": None,
    },
    {
        "name": "05 通用引文修复",
        "script": "generic_reference_repair.py",
        "required": True,
        "stdin_key": None,
        "max_retries": 1,
    },
    {
        "name": "05.5 字数预算修复",
        "script": "generic_length_repair.py",
        "required": True,
        "stdin_key": None,
        "max_retries": 1,
    },
    {
        "name": "06 通用最终检查",
        "script": "generic_final_check.py",
        "required": True,
        "stdin_key": None,
        "max_retries": 1,
    },
]


def log(msg):
    print(msg, flush=True)
    with PIPELINE_LOG.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError("没有找到 config.local.json。请先运行 app.py 并保存配置。")
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def prepare_env(config):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    env["API_KEY"] = config.get("api_key", "")
    env["OPENAI_API_KEY"] = config.get("api_key", "")
    env["BASE_URL"] = config.get("base_url", "")
    env["OPENAI_BASE_URL"] = config.get("base_url", "")

    env["CARD_MODEL"] = config.get("card_model", "")
    env["WRITE_MODEL"] = config.get("write_model", "")
    env["REVIEW_MODEL"] = config.get("review_model", "")

    env["PDF_DIR"] = config.get("pdf_dir", str(BASE_DIR / "pdfs"))
    env["OUTPUT_DIR"] = config.get("output_dir", str(BASE_DIR / "output"))
    env["PAPER_TITLE"] = config.get("title", "")

    return env


def write_runtime_inputs(config):
    (BASE_DIR / "资料卡片指令.txt").write_text(config.get("card_prompt", ""), encoding="utf-8")
    (BASE_DIR / "论文要求与大纲.txt").write_text(config.get("requirements", ""), encoding="utf-8")
    (BASE_DIR / "论文题目.txt").write_text(config.get("title", ""), encoding="utf-8")
    (BASE_DIR / "风险短语.txt").write_text(config.get("risk_phrases", ""), encoding="utf-8")


def build_stdin_text(stage, config):
    key = stage.get("stdin_key")

    if key == "card_prompt":
        prompt = config.get("card_prompt", "").strip()
        if not prompt:
            prompt = "请阅读 PDF 并整理为论文资料卡片。"
        return prompt + "\nEND\n"

    if key == "requirements":
        req = config.get("requirements", "").strip()
        if not req:
            req = "请按论文大纲和老师要求写作。"
        return req + "\nEND\n"

    return None


def prepare_cards_input():
    cards_dir = BASE_DIR / "cards_input"
    cards_dir.mkdir(exist_ok=True)

    src_candidates = [
        BASE_DIR / "output" / "论文资料卡片汇总.docx",
        BASE_DIR / "output" / "新资料卡片汇总.docx",
    ]

    dst = cards_dir / "新资料卡片汇总.docx"

    for src in src_candidates:
        if src.exists():
            shutil.copy2(src, dst)
            log(f"已自动复制资料卡片到 cards_input：{dst}")
            return True

    log("未找到可复制的资料卡片 Word。cards_to_matrix.py 可能会失败。")
    return False


def run_script(script_path, env, stdin_text=None):
    use_stdin = stdin_text is not None

    process = subprocess.Popen(
        [sys.executable, str(script_path)],
        cwd=str(BASE_DIR),
        stdin=subprocess.PIPE if use_stdin else None,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )

    if use_stdin:
        try:
            process.stdin.write(stdin_text)
            process.stdin.flush()
            process.stdin.close()
            log("已自动向脚本输入指令，并发送 END。")
        except Exception as e:
            log(f"自动输入指令失败：{e}")

    for line in process.stdout:
        log(line.rstrip())

    return process.wait()


def run_stage(stage, env, config, global_max_retries):
    script_path = BASE_DIR / stage["script"]

    if stage.get("script") == "cards_to_matrix.py":
        prepare_cards_input()

    if not script_path.exists():
        msg = f"未找到脚本：{stage['script']}"
        if stage.get("required", True):
            log("错误：" + msg)
            return False
        log("跳过：" + msg)
        return True

    stage_max_retries = stage.get("max_retries")
    if stage_max_retries is None:
        stage_max_retries = global_max_retries
    stage_max_retries = int(stage_max_retries)

    log("")
    log("=" * 80)
    log(f"开始阶段：{stage['name']}")
    log(f"使用脚本：{stage['script']}")
    log("=" * 80)

    stdin_text = build_stdin_text(stage, config)

    for attempt in range(1, stage_max_retries + 1):
        log(f"第 {attempt}/{stage_max_retries} 次尝试：{stage['name']}")

        start = time.time()
        code = run_script(script_path, env, stdin_text=stdin_text)
        elapsed = round(time.time() - start, 1)

        if code == 0:
            log(f"阶段通过：{stage['name']}，耗时 {elapsed} 秒。")
            return True

        log(f"阶段失败：{stage['name']}，退出码 {code}，耗时 {elapsed} 秒。")

        if attempt < stage_max_retries:
            log("准备重试当前阶段。")
            time.sleep(2)

    log(f"阶段最终失败：{stage['name']}。流程停止。")
    return False


def main():
    try:
        config = load_config()
        write_runtime_inputs(config)

        max_retries = int(config.get("max_retries", 3))
        env = prepare_env(config)

        pdf_dir = Path(config.get("pdf_dir", str(BASE_DIR / "pdfs")))
        output_dir = Path(config.get("output_dir", str(BASE_DIR / "output")))

        pdf_dir.mkdir(exist_ok=True)
        output_dir.mkdir(exist_ok=True)
        (BASE_DIR / "cards_input").mkdir(exist_ok=True)
        (BASE_DIR / "logs").mkdir(exist_ok=True)

        log("论文自动综合助手流程启动。")
        log("当前流程：v1.0.0 大纲驱动写作工作流")
        log(f"项目目录：{BASE_DIR}")
        log(f"PDF目录：{pdf_dir}")
        log(f"输出目录：{output_dir}")
        log(f"日志文件：{PIPELINE_LOG}")

        pdf_count = len(list(pdf_dir.glob("*.pdf")))
        log(f"当前 PDF 数量：{pdf_count}")

        if pdf_count == 0:
            log("错误：PDF 文件夹为空。请先放入 PDF。")
            sys.exit(1)

        for stage in STAGES:
            ok = run_stage(stage, env, config, max_retries)
            if not ok:
                sys.exit(1)

        log("")
        log("全部阶段运行完成。")
        log("请检查 output 文件夹中的最终稿和最终检查报告。")

    except Exception as e:
        log(f"程序出错：{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
