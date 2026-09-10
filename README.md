# 论文自动综合助手 | Paper Synthesis Assistant

一个把个人综述写作方法封装成**本地可执行工作流**的小工具。

它会把“读文献 → 做资料卡片 → 建证据矩阵 → 按大纲写作 → 修复引文与字数 → 最终检查”串成一条流程，帮助减少重复整理工作，把更多精力留给选题、判断和修改。

**工作流：**

**PDF → 文献筛选 → 资料卡片 → 证据矩阵 → 大纲驱动写作 → 引文/参考文献修复 → 字数修复 → 最终检查**

当前版本：**v1.0.0**

## 快速开始（Windows）

1. 下载并解压 Windows 版。
2. 双击 **`START.bat`**。
3. 首次运行会自动检查 Python 和依赖；缺少 Python 时会询问是否自动准备本地运行环境。
4. 打开软件后，填写或修改 **API Base URL、API Key、模型、论文题目、写作要求和大纲**。
5. 把 PDF 放入 `pdfs/`，点击 **“开始全流程”**。
6. 在 `output/` 查看结果。

环境异常时可双击 **`REPAIR_ENV.bat`** 重建运行环境。

## 主要功能

- 批量读取 PDF 并筛选可用材料；
- 生成面向综述写作的资料卡片；
- 汇总形成证据矩阵；
- 根据题目、要求和大纲组织初稿；
- 检查证据、引文、参考文献顺序和字数；
- 输出最终稿与检查结果。

选题、写作要求、大纲、API 和模型都可以自行配置，**不绑定固定论文主题**。

## API 与模型

程序使用 **OpenAI-compatible API**。

在软件界面的 **“API 与模型”** 区域，可以直接修改：

- Base URL
- API Key
- 资料卡片模型
- 写作模型
- 审稿模型

需要更换 API 服务商或中转站时，直接改这些输入框并保存即可，**不需要修改代码**。

<details>
<summary><b>开发者 / 项目结构</b></summary>

```text
paper-synthesis-assistant/
├── START.bat
├── REPAIR_ENV.bat
├── bootstrap.ps1
├── app.py
├── run_pipeline.py
├── batch_pdf_cards_generic.py
├── cards_to_matrix.py
├── validate_evidence_gate.py
├── write_review_from_outline.py
├── finalize_outline_docx.py
├── generic_reference_repair.py
├── generic_length_repair.py
├── generic_final_check.py
├── requirements.txt
├── config.example.json
└── VERSION.txt
```

如果已经安装 Python 3.10–3.13，也可以手动运行：

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

普通 Windows 用户不需要手动执行这些命令。

</details>

## 说明

这个工具用于辅助文献整理、证据组织和综述写作，不替代人工判断。模型生成内容仍建议人工核对事实、引用和学术规范。

如果它对你的学习或科研有帮助，欢迎点一个 **Star ⭐**。
