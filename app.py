import json
import os
import subprocess
import sys
import threading
import shutil
import re
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


# =========================================================
# 一、DPI 适配，解决 Windows 下 Tkinter 字体发虚
# =========================================================

def enable_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # Windows 8.1+
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


enable_dpi_awareness()


BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.local.json"

PDF_DIR = BASE_DIR / "pdfs"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"


DEFAULT_CARD_PROMPT = """请阅读我上传的 PDF，并整理成“论文资料卡片”。整理目的不是普通摘要，而是服务于后续综述写作。

请重点提取：
1. 论文基本信息：题名、作者、年份、来源；
2. 研究对象、研究背景和核心问题；
3. 与我的论文主题直接相关的观点、数据、案例和结论；
4. 可支撑的论文部分：发展现状、主要困境、发展对策、讨论；
5. 可直接用于正文写作的材料；
6. 使用局限：该文献只能支撑什么，不能过度推出什么；
7. 对本文选题的价值判断：核心文献、重要文献、一般背景文献或不建议使用。

不要泛泛总结，要围绕我的论文主题筛选材料。
"""

DEFAULT_REQUIREMENTS = """请严格按照我填写的论文题目、研究逻辑、大纲和老师要求写作。

基本要求：
1. 围绕主题筛选证据，不要把资料卡片机械堆砌成正文；
2. 所有来自文献的观点、案例、数据和结论都要加数字引文；
3. 对策必须和前文困境对应；
4. 注意避免过度因果，不能把间接证据写成直接证明；
5. 语言尽量自然、稳妥，避免明显 AI 腔。
"""

DEFAULT_RISK_PHRASES = """直接导致
必然
完全解决
显著证明
决定性证明
毫无疑问
众所周知
"""


def load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass

    return {
        "base_url": "https://api.302.ai/v1",
        "api_key": "",
        "card_model": "gpt-5.4",
        "write_model": "gpt-5.4",
        "review_model": "gpt-5.5",

        "title": "防止耕地非粮化背景下中药材种植业的发展困境与对策研究",
        "pdf_dir": str(PDF_DIR),
        "output_dir": str(OUTPUT_DIR),
        "max_retries": 3,

        "min_total_count": 7500,
        "max_total_count": 8500,
        "min_references": 30,
        "max_references": 35,
        "min_year": 2021,
        "heading_style": "1 / 1.1",
        "forbid_third_heading": True,
        "require_reference_order": True,

        "card_prompt": DEFAULT_CARD_PROMPT,
        "requirements": DEFAULT_REQUIREMENTS,
        "risk_phrases": DEFAULT_RISK_PHRASES,
    }


def save_config(config):
    PDF_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    LOG_DIR.mkdir(exist_ok=True)

    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    (BASE_DIR / "资料卡片指令.txt").write_text(config["card_prompt"], encoding="utf-8")
    (BASE_DIR / "论文要求与大纲.txt").write_text(config["requirements"], encoding="utf-8")
    (BASE_DIR / "论文题目.txt").write_text(config["title"], encoding="utf-8")
    (BASE_DIR / "风险短语.txt").write_text(config["risk_phrases"], encoding="utf-8")

    gitignore = BASE_DIR / ".gitignore"
    if not gitignore.exists():
        gitignore.write_text(
            "config.local.json\n"
            "代码全集备份/\n"
            "logs/\n",
            encoding="utf-8"
        )


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("本地综述论文自动写作助手 MVP 1.1")
        self.root.geometry("1600x980")
        self.root.minsize(1450, 850)
        try:
            self.root.state("zoomed")
        except Exception:
            pass

        self.config = load_config()
        self.process = None

        self.setup_style()
        self.build_ui()

    def setup_style(self):
        # 字体放大，减少模糊和拥挤
        self.font_main = ("Microsoft YaHei UI", 20)
        self.font_label = ("Microsoft YaHei UI", 20)
        self.font_title = ("Microsoft YaHei UI", 22, "bold")
        self.font_text = ("Microsoft YaHei UI", 20)
        self.font_log = ("Consolas", 18)

        self.root.option_add("*Font", self.font_main)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background="#F7F8FA")
        style.configure("TLabelframe", background="#F7F8FA")
        style.configure("TLabelframe.Label", font=self.font_title, foreground="#1F2937")
        style.configure("TLabel", background="#F7F8FA", foreground="#111827", font=self.font_label)
        style.configure("TButton", font=self.font_main, padding=(18, 10))
        style.configure("TEntry", font=self.font_main, padding=4)
        style.configure("TNotebook", background="#F7F8FA")
        style.configure("TNotebook.Tab", font=self.font_main, padding=(24, 12))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 20, "bold"), padding=(20, 12))

        # Tk 缩放
        try:
            self.root.tk.call("tk", "scaling", 1.0)
        except Exception:
            pass

    def build_ui(self):
        self.root.configure(bg="#F7F8FA")

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=10)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 8))

        ttk.Label(
            header,
            text="本地综述论文自动写作助手",
            font=("Microsoft YaHei UI", 30, "bold")
        ).pack(side="left")

        ttk.Label(
            header,
            text="PDF → 资料卡片 → 证据矩阵 → 初稿 → 检查 → 引文排序 → 最终稿",
            foreground="#4B5563"
        ).pack(side="left", padx=16)

        # 顶部：文件夹和 API
        top = ttk.Frame(main)
        top.pack(fill="x")

        left_top = ttk.LabelFrame(top, text="文件夹设置")
        left_top.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.pdf_dir_var = tk.StringVar(value=self.config.get("pdf_dir", str(PDF_DIR)))
        self.output_dir_var = tk.StringVar(value=self.config.get("output_dir", str(OUTPUT_DIR)))

        self.add_path_row(left_top, "PDF 文件夹", self.pdf_dir_var, self.choose_pdf_dir, self.open_pdf_dir, 0)
        self.add_path_row(left_top, "输出文件夹", self.output_dir_var, self.choose_output_dir, self.open_output_dir, 1)

        right_top = ttk.LabelFrame(top, text="API 与模型")
        right_top.pack(side="right", fill="x", padx=(8, 0))

        self.base_url_var = tk.StringVar(value=self.config.get("base_url", "https://api.302.ai/v1"))
        self.api_key_var = tk.StringVar(value=self.config.get("api_key", ""))
        self.card_model_var = tk.StringVar(value=self.config.get("card_model", "gpt-5.4"))
        self.write_model_var = tk.StringVar(value=self.config.get("write_model", "gpt-5.4"))
        self.review_model_var = tk.StringVar(value=self.config.get("review_model", "gpt-5.5"))
        self.max_retries_var = tk.StringVar(value=str(self.config.get("max_retries", 3)))

        self.add_labeled_entry(right_top, "Base URL", self.base_url_var, 0, 0, width=38)
        self.add_labeled_entry(right_top, "API Key", self.api_key_var, 1, 0, width=38, show="*")
        self.add_labeled_entry(right_top, "资料卡片模型", self.card_model_var, 2, 0, width=18)
        self.add_labeled_entry(right_top, "写作模型", self.write_model_var, 2, 2, width=18)
        self.add_labeled_entry(right_top, "审稿模型", self.review_model_var, 3, 0, width=18)
        self.add_labeled_entry(right_top, "重试次数", self.max_retries_var, 3, 2, width=8)

        # Notebook
        notebook = ttk.Notebook(main)
        notebook.pack(fill="both", expand=True, pady=10)

        tab_basic = ttk.Frame(notebook)
        tab_prompts = ttk.Frame(notebook)
        tab_constraints = ttk.Frame(notebook)
        tab_run = ttk.Frame(notebook)

        notebook.add(tab_basic, text="① 论文信息")
        notebook.add(tab_prompts, text="② 指令与大纲")
        notebook.add(tab_constraints, text="③ 硬约束与风险词")
        notebook.add(tab_run, text="④ 运行日志")

        self.build_basic_tab(tab_basic)
        self.build_prompts_tab(tab_prompts)
        self.build_constraints_tab(tab_constraints)
        self.build_run_tab(tab_run)

        # 底部按钮
        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(2, 0))

        ttk.Button(actions, text="保存配置", command=self.save_current_config).pack(side="left", padx=4)
        ttk.Button(actions, text="开始全流程", style="Accent.TButton", command=self.start_pipeline).pack(side="left", padx=4)
        ttk.Button(actions, text="停止", command=self.stop_pipeline).pack(side="left", padx=4)
        ttk.Button(actions, text="打开 PDF 文件夹", command=self.open_pdf_dir).pack(side="left", padx=4)
        ttk.Button(actions, text="打开输出文件夹", command=self.open_output_dir).pack(side="left", padx=4)
        ttk.Button(actions, text="打开日志文件夹", command=self.open_log_dir).pack(side="left", padx=4)
        ttk.Button(actions, text="测试模式", command=self.apply_test_mode).pack(side="left", padx=4)
        ttk.Button(actions, text="正式模式", command=self.apply_formal_mode).pack(side="left", padx=4)
        ttk.Button(actions, text="新建任务/归档清空", command=self.archive_current_task).pack(side="left", padx=4)

    def add_path_row(self, parent, label, var, choose_cmd, open_cmd, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(parent, textvariable=var, width=66).grid(row=row, column=1, sticky="we", padx=4, pady=6)
        ttk.Button(parent, text="选择", command=choose_cmd).grid(row=row, column=2, padx=3)
        ttk.Button(parent, text="打开", command=open_cmd).grid(row=row, column=3, padx=3)
        parent.columnconfigure(1, weight=1)

    def add_labeled_entry(self, parent, label, var, row, col, width=20, show=None):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=8, pady=5)
        ttk.Entry(parent, textvariable=var, width=width, show=show).grid(row=row, column=col + 1, sticky="w", padx=4, pady=5)

    def build_basic_tab(self, tab):
        frame = ttk.LabelFrame(tab, text="论文基本信息")
        frame.pack(fill="x", padx=10, pady=10)

        self.title_var = tk.StringVar(value=self.config.get("title", ""))

        ttk.Label(frame, text="论文题目").pack(anchor="w", padx=8, pady=(8, 2))
        ttk.Entry(frame, textvariable=self.title_var, width=120).pack(fill="x", padx=8, pady=(0, 8))

        hint = (
            "建议：第一次测试只放 2—3 篇 PDF；正式跑几百篇前，先确认每一阶段脚本都能通过。"
            "API Key 只保存在本地 config.local.json，不要上传或发给别人。"
        )
        ttk.Label(frame, text=hint, foreground="#6B7280", wraplength=1000).pack(anchor="w", padx=8, pady=8)

    def build_prompts_tab(self, tab):
        paned = ttk.PanedWindow(tab, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.LabelFrame(paned, text="资料卡片指令")
        right = ttk.LabelFrame(paned, text="论文大纲与老师要求")
        paned.add(left, weight=1)
        paned.add(right, weight=1)

        self.card_prompt_text = ScrolledText(left, wrap="word", height=22, font=self.font_text, padx=8, pady=8)
        self.card_prompt_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.card_prompt_text.insert("1.0", self.config.get("card_prompt", DEFAULT_CARD_PROMPT))

        self.requirements_text = ScrolledText(right, wrap="word", height=22, font=self.font_text, padx=8, pady=8)
        self.requirements_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.requirements_text.insert("1.0", self.config.get("requirements", DEFAULT_REQUIREMENTS))

    def build_constraints_tab(self, tab):
        top = ttk.LabelFrame(tab, text="通用硬约束")
        top.pack(fill="x", padx=10, pady=10)

        self.min_total_count_var = tk.StringVar(value=str(self.config.get("min_total_count", 7500)))
        self.max_total_count_var = tk.StringVar(value=str(self.config.get("max_total_count", 8500)))
        self.min_refs_var = tk.StringVar(value=str(self.config.get("min_references", 30)))
        self.max_refs_var = tk.StringVar(value=str(self.config.get("max_references", 35)))
        self.min_year_var = tk.StringVar(value=str(self.config.get("min_year", 2021)))
        self.heading_style_var = tk.StringVar(value=self.config.get("heading_style", "1 / 1.1"))

        self.forbid_third_heading_var = tk.BooleanVar(value=bool(self.config.get("forbid_third_heading", True)))
        self.require_reference_order_var = tk.BooleanVar(value=bool(self.config.get("require_reference_order", True)))

        self.add_labeled_entry(top, "最小字数", self.min_total_count_var, 0, 0, width=10)
        self.add_labeled_entry(top, "最大字数", self.max_total_count_var, 0, 2, width=10)
        self.add_labeled_entry(top, "参考文献最少", self.min_refs_var, 1, 0, width=10)
        self.add_labeled_entry(top, "参考文献最多", self.max_refs_var, 1, 2, width=10)
        self.add_labeled_entry(top, "近五年起始年份", self.min_year_var, 2, 0, width=10)
        self.add_labeled_entry(top, "标题格式", self.heading_style_var, 2, 2, width=16)

        ttk.Checkbutton(top, text="禁止三级标题，例如 2.1.1", variable=self.forbid_third_heading_var).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=8)
        ttk.Checkbutton(top, text="要求参考文献按正文首次出现顺序排列", variable=self.require_reference_order_var).grid(row=3, column=2, columnspan=3, sticky="w", padx=8, pady=8)

        risk = ttk.LabelFrame(tab, text="风险短语 / 禁用表达")
        risk.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ttk.Label(
            risk,
            text="每行一个。这里是泛用检查，不再写死某个题目。留空则不检查风险短语。",
            foreground="#6B7280"
        ).pack(anchor="w", padx=8, pady=(8, 2))

        self.risk_text = ScrolledText(risk, wrap="word", height=12, font=self.font_text, padx=8, pady=8)
        self.risk_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.risk_text.insert("1.0", self.config.get("risk_phrases", DEFAULT_RISK_PHRASES))

    def build_run_tab(self, tab):
        self.log_text = ScrolledText(tab, wrap="word", height=22, font=self.font_log, padx=8, pady=8)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)


    def apply_test_mode(self):
        """
        少量 PDF 调试用，不追求正式论文要求。
        """
        self.min_total_count_var.set("3000")
        self.max_total_count_var.set("6500")
        self.min_refs_var.set("2")
        self.max_refs_var.set("8")
        self.max_retries_var.set("5")
        self.log("已切换为测试模式：3000—6500字，参考文献2—8篇。请点击保存配置。")

    def apply_formal_mode(self):
        """
        正式论文写作用。
        """
        self.min_total_count_var.set("7500")
        self.max_total_count_var.set("8500")
        self.min_refs_var.set("30")
        self.max_refs_var.set("35")
        self.max_retries_var.set("3")
        self.log("已切换为正式模式：7500—8500字，参考文献30—35篇。请点击保存配置。")

    def safe_task_name(self, title):
        title = title.strip() or "未命名任务"
        title = re.sub(r'[\\/:*?"<>|]+', "_", title)
        title = re.sub(r"\s+", "_", title)
        return title[:30]

    def move_folder_to_archive(self, src_path, archive_dir, name):
        src = Path(src_path)
        if not src.exists():
            src.mkdir(parents=True, exist_ok=True)
            return

        # 空文件夹也一起归档，保证新任务干净
        dst = archive_dir / name
        if dst.exists():
            dst = archive_dir / f"{name}_{datetime.now().strftime('%H%M%S')}"

        try:
            shutil.move(str(src), str(dst))
        except Exception as e:
            self.log(f"归档 {src} 失败：{e}")
            return

        src.mkdir(parents=True, exist_ok=True)

    def archive_current_task(self):
        """
        归档当前 pdfs/output/cards_input/logs，并重新创建空工作区。
        不删除 config.local.json，不影响 API Key。
        """
        if self.process and self.process.poll() is None:
            messagebox.showwarning("正在运行", "当前流程还在运行，不能新建任务。")
            return

        ok = messagebox.askyesno(
            "确认新建任务",
            "将把当前 pdfs、output、cards_input、logs 归档到 archive 文件夹，并重新创建空工作区。\\n\\n不会删除 API Key 配置。是否继续？"
        )

        if not ok:
            return

        task_name = self.safe_task_name(self.title_var.get())
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archive_dir = BASE_DIR / "archive" / f"{stamp}_{task_name}"
        archive_dir.mkdir(parents=True, exist_ok=True)

        self.move_folder_to_archive(self.pdf_dir_var.get(), archive_dir, "pdfs")
        self.move_folder_to_archive(self.output_dir_var.get(), archive_dir, "output")
        self.move_folder_to_archive(BASE_DIR / "cards_input", archive_dir, "cards_input")
        self.move_folder_to_archive(BASE_DIR / "logs", archive_dir, "logs")

        # 重新保证默认文件夹存在
        Path(self.pdf_dir_var.get()).mkdir(parents=True, exist_ok=True)
        Path(self.output_dir_var.get()).mkdir(parents=True, exist_ok=True)
        (BASE_DIR / "cards_input").mkdir(parents=True, exist_ok=True)
        (BASE_DIR / "logs").mkdir(parents=True, exist_ok=True)

        self.log(f"当前任务已归档到：{archive_dir}")
        self.log("已新建空 pdfs/output/cards_input/logs。现在可以放入新 PDF，填写新题目并开始。")


    def choose_pdf_dir(self):
        folder = filedialog.askdirectory(initialdir=str(BASE_DIR))
        if folder:
            self.pdf_dir_var.set(folder)

    def choose_output_dir(self):
        folder = filedialog.askdirectory(initialdir=str(BASE_DIR))
        if folder:
            self.output_dir_var.set(folder)

    def open_pdf_dir(self):
        folder = Path(self.pdf_dir_var.get())
        folder.mkdir(exist_ok=True)
        os.startfile(str(folder))

    def open_output_dir(self):
        folder = Path(self.output_dir_var.get())
        folder.mkdir(exist_ok=True)
        os.startfile(str(folder))

    def open_log_dir(self):
        LOG_DIR.mkdir(exist_ok=True)
        os.startfile(str(LOG_DIR))

    def int_value(self, var, default):
        try:
            return int(var.get())
        except Exception:
            return default

    def get_current_config(self):
        return {
            "base_url": self.base_url_var.get().strip(),
            "api_key": self.api_key_var.get().strip(),
            "card_model": self.card_model_var.get().strip(),
            "write_model": self.write_model_var.get().strip(),
            "review_model": self.review_model_var.get().strip(),

            "title": self.title_var.get().strip(),
            "pdf_dir": self.pdf_dir_var.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
            "max_retries": self.int_value(self.max_retries_var, 3),

            "min_total_count": self.int_value(self.min_total_count_var, 7500),
            "max_total_count": self.int_value(self.max_total_count_var, 8500),
            "min_references": self.int_value(self.min_refs_var, 30),
            "max_references": self.int_value(self.max_refs_var, 35),
            "min_year": self.int_value(self.min_year_var, 2021),
            "heading_style": self.heading_style_var.get().strip(),
            "forbid_third_heading": bool(self.forbid_third_heading_var.get()),
            "require_reference_order": bool(self.require_reference_order_var.get()),

            "card_prompt": self.card_prompt_text.get("1.0", "end").strip(),
            "requirements": self.requirements_text.get("1.0", "end").strip(),
            "risk_phrases": self.risk_text.get("1.0", "end").strip(),
        }

    def save_current_config(self):
        config = self.get_current_config()

        if not config["api_key"]:
            self.log("提醒：API Key 为空。可以先保存配置，但开始全流程前需要填写。")

        save_config(config)
        self.log("配置已保存。config.local.json 含 API Key，不要上传。")

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def start_pipeline(self):
        self.save_current_config()

        if self.process and self.process.poll() is None:
            messagebox.showwarning("正在运行", "当前流程还在运行。")
            return

        pipeline_path = BASE_DIR / "run_pipeline.py"
        if not pipeline_path.exists():
            messagebox.showerror("缺少文件", f"没有找到 {pipeline_path}")
            return

        t = threading.Thread(target=self.run_pipeline_thread, daemon=True)
        t.start()

    def run_pipeline_thread(self):
        self.log("开始运行全流程。")

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        cmd = [sys.executable, str(BASE_DIR / "run_pipeline.py")]

        self.process = subprocess.Popen(
            cmd,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env
        )

        for line in self.process.stdout:
            self.log(line.rstrip())

        code = self.process.wait()

        if code == 0:
            self.log("全流程完成。请打开 output 文件夹查看最终稿和检查报告。")
        else:
            self.log(f"流程中断，退出码：{code}。请查看日志最后几十行。")

    def stop_pipeline(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.log("已请求停止当前流程。")
        else:
            self.log("当前没有正在运行的流程。")


def main():
    PDF_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    LOG_DIR.mkdir(exist_ok=True)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()