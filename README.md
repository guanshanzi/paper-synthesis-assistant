# 论文自动综述助手

一个本地运行的综述论文自动写作辅助工具。程序会读取 `pdfs/` 中的 PDF 文献，生成资料卡片、证据矩阵、综述初稿，并进行引文顺序、参考文献和字数约束检查。

> 当前整理版来自 MVP 1.5 稳定分享版。此仓库不包含 API Key、论文 PDF、运行日志或历史输出稿。

## 功能概览

- 批量读取 PDF 文献并生成资料卡片
- 将资料卡片整理为证据矩阵
- 根据题目、要求与大纲生成综述初稿
- 自动检查证据是否足够
- 修复参考文献与正文引文顺序
- 按字数预算进行修正
- 输出最终检查报告
- 提供 Tkinter 图形界面，便于填写 API Key、模型和任务要求

## 项目结构

```text
paper-review-assistant/
├── app.py                         # 图形界面入口
├── run_pipeline.py                # 全流程调度脚本
├── batch_pdf_cards_generic.py      # PDF → 资料卡片
├── cards_to_matrix.py             # 资料卡片 → 证据矩阵
├── validate_evidence_gate.py       # 证据门控检查
├── write_review_from_outline.py    # 根据大纲写综述
├── finalize_outline_docx.py        # 生成/整理 DOCX
├── generic_reference_repair.py     # 引文与参考文献修复
├── generic_length_repair.py        # 字数修复
├── generic_final_check.py          # 最终检查
├── config.example.json             # 配置模板，不含密钥
├── requirements.txt                # Python 依赖
├── install_deps.bat                # Windows 一键安装依赖
├── 启动软件.bat                    # Windows 一键启动
├── pdfs/                           # 放入待处理 PDF，本目录内容不上传
├── output/                         # 输出结果，本目录内容不上传
├── logs/                           # 运行日志，本目录内容不上传
├── cards_input/                    # 可选资料卡片输入，本目录内容不上传
└── archive/                        # 本地归档，本目录内容不上传
```

## 安装环境

建议使用 Python 3.10 或以上版本。

### Windows 快速安装

双击：

```text
install_deps.bat
```

### 手动安装

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

说明：代码中使用 `import fitz` 读取 PDF，对应的 pip 包名是 `pymupdf`，不是 `fitz`。

## 使用方法

1. 将待分析的 PDF 放入 `pdfs/`。
2. 运行图形界面：

```bash
python app.py
```

或者在 Windows 下双击：

```text
启动软件.bat
```

3. 在界面中填写：
   - Base URL
   - API Key
   - 资料卡片模型
   - 写作模型
   - 审稿模型
   - 论文题目、写作要求、大纲、风险短语等
4. 点击“保存配置”。程序会在本地生成 `config.local.json`。
5. 点击“运行全流程”。
6. 在 `output/` 中查看生成结果。

## 配置说明

仓库只提供 `config.example.json`，不提供 `config.local.json`。

首次运行并保存配置后，会在本地生成：

```text
config.local.json
```

这个文件包含 API Key，已经被 `.gitignore` 排除，**不要上传到 GitHub，不要转发给别人**。

## GitHub 上传前检查

上传前建议运行：

```bash
git status
```

确认没有这些内容被加入 Git：

```text
config.local.json
pdfs/ 中的论文 PDF
output/ 中的生成稿
logs/ 中的日志
archive/ 中的历史归档
任何包含 API Key 的文件
```

## GitHub 上传命令

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/你的用户名/paper-review-assistant.git
git push -u origin main
```

第一次上传建议把 GitHub 仓库设为 Private。确认没有密钥、PDF、日志和个人文件后，再考虑是否公开。

## 免责声明

本工具用于辅助整理文献和生成综述草稿，输出内容仍需人工核对，包括事实准确性、引用真实性、参考文献格式、引文顺序和学术规范。
