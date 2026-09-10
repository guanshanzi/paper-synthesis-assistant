# 论文自动综合助手 | Paper Synthesis Assistant

把一套“先读文献、再做资料卡片、再建立证据矩阵、最后按大纲写作并做硬约束检查”的个人综述写作方法，封装成可以本地运行的软件工作流。

它不是单纯的“把文字改写一遍”，也不是替代人工判断的论文生成器。它更接近一个可执行的写作 Skill：

**PDF → 文献筛选 → 资料卡片 → 证据矩阵 → 大纲驱动写作 → 引文/参考文献修复 → 字数修复 → 最终检查**

当前公开版本：**v1.0.0**

## 最简单的使用方法（Windows）

1. 下载并解压整个项目。
2. 双击 **`START.bat`**。
3. 第一次运行时，程序会自动检查运行环境：
   - 如果电脑已有兼容的 Python 3.10–3.13，会直接使用它创建项目自己的 `.venv`；
   - 如果没有，会先询问你是否同意自动下载 Python 3.12；
   - 同意后只从 Python 官方网站下载，并安装到本软件目录 `.runtime/python`；
   - 不加入系统 PATH，不要求管理员权限，也不会修改你已有的 Python 环境；
   - 随后自动安装本项目依赖并做一次环境自检。
4. 软件界面出现后，填写 API Base URL、API Key、模型、论文题目、写作要求和大纲。
5. 把要分析的 PDF 放入 `pdfs/`，点击“运行全流程”。
6. 在 `output/` 中查看结果。

首次启动需要联网下载 Python/依赖，时间取决于网络速度。以后再启动时会自动复用现有环境，不会重复安装。

如果环境损坏，可双击 **`REPAIR_ENV.bat`** 重新创建 `.venv` 并安装依赖。

## 这个项目做什么

- 批量读取 PDF 文献；
- 根据当前任务要求判断文献是否值得进入资料卡片；
- 生成服务于后续写作的结构化资料卡片；
- 将资料卡片压缩成证据矩阵；
- 根据论文题目、要求和大纲组织综述初稿；
- 检查证据是否足够；
- 修复正文引文与参考文献顺序；
- 按字数预算进行修正；
- 输出最终检查报告。

项目的重点不是某一个固定题目，而是这套写作流程本身。v1.0.0 已清理旧项目中残留的固定研究主题，证据矩阵会跟随当前论文题目组织材料。

## API 与隐私

本项目不包含任何 API Key、论文 PDF、运行日志或个人历史输出。

第一次在界面保存配置后，会在本地生成：

```text
config.local.json
```

其中可能包含你的 API Key。该文件已被 `.gitignore` 排除。**不要把它上传到 GitHub，也不要转发给别人。**

程序目前使用 OpenAI-compatible API 方式调用模型。默认 Base URL 只是一个可修改的示例，用户可在界面中替换为自己实际使用的兼容服务。

## 项目结构

```text
paper-synthesis-assistant/
├── START.bat                   # 推荐入口：自动准备环境并启动
├── REPAIR_ENV.bat              # 环境异常时重建依赖
├── bootstrap.ps1               # Windows 自动环境引导
├── app.py                      # 图形界面入口
├── run_pipeline.py             # 全流程调度
├── batch_pdf_cards_generic.py  # PDF → 资料卡片
├── cards_to_matrix.py          # 资料卡片 → 证据矩阵
├── validate_evidence_gate.py   # 证据门控
├── write_review_from_outline.py
├── finalize_outline_docx.py
├── generic_reference_repair.py
├── generic_length_repair.py
├── generic_final_check.py
├── requirements.txt
├── config.example.json
└── VERSION.txt
```

运行后会自动创建：

```text
pdfs/
output/
logs/
cards_input/
archive/
.venv/
.runtime/        # 仅在需要自动安装本地 Python 时出现
bootstrap_logs/
```

## 手动运行（开发者）

如果你已经有 Python 3.10–3.13，也可以手动执行：

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

普通 Windows 用户不需要执行这些命令，直接双击启动脚本即可。

为避免不同 ZIP/解压工具对中文文件名编码处理不一致，启动入口统一使用 ASCII 文件名 `START.bat` / `REPAIR_ENV.bat`。

## 版本说明

### v1.0.0

- 从原先的内部 MVP 命名切换到正式版本号；
- 项目名称统一为“论文自动综合助手”；
- 新增首次运行自动环境引导；
- 缺少 Python 时，可在用户确认后自动从 Python.org 下载 Python 3.12；
- 依赖安装改为项目独立 `.venv`，不再污染系统 Python；
- 新增环境自检与修复入口；
- 修复 README 中存在启动脚本但实际缺失的问题；
- 清理证据矩阵中遗留的固定旧论文主题与章节分类；
- 统一界面、日志和报告中的公开版本名称。

## 免责声明

本工具用于辅助文献整理、证据组织与综述草稿生成。模型输出仍需人工核对，尤其包括事实准确性、引用真实性、参考文献格式、论证边界和学术规范。

如果这个项目对你的学习或科研有帮助，欢迎在 GitHub 右上角点一个 **Star**。这会帮助项目被更多真正需要的人看到。
