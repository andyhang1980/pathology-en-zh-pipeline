---
name: pathology-en-zh-pipeline
version: 1.0.0
description: 病理学专著 EN→ZH 出版级一体化流水线。集成 翻译/意译 → 校对 → 润色 → 仿人工精读 → 再精校 → 排版交付，一键产出可直接出版的 Word 文档。适用于 WHO 及同类英文病理/医学专著的中译出版级排版。
type: skill
---

# 病理学 EN→ZH 出版级一体化流水线

把一本英文病理学专著（docx）变成**直接可用于出版的 Word 中文文档**，全流程六阶段一体驱动。

## 何时用
- 拿到英文 WHO / 病理专著 docx，需要出版级中文版。
- 已有中文稿但需按统一红线（术语/标点/句式）校对、润色、排版交付。
- 做双语对照版（中文 ×2 行对齐）。

## 六阶段总览
| 阶段 | 名称 | 执行方 | 脚本/动作 |
|---|---|---|---|
| 0 | 准备 | 脚本 | `pipeline.py setup` → inventory + 抽取英文段落 JSON |
| 1 | 翻译/意译 | **LLM（agent）** | 译 JSON 文本 → `apply-translation` 回注 |
| 2 | 校对 | 脚本 | who_docx 子命令 + `proof25` 红线（跑两遍吸收级联） |
| 3 | 润色 | **LLM（agent）** | 按 `references/02` 清单精修句式/语序 |
| 4 | 仿人工精读 | **LLM（agent）** | 对照英文逐段核查漏译/误译/术语 |
| 5 | 再精校 + 排版交付 | 脚本 | `proof25` 复扫收敛 + typeset + images + verify |

## 一键驱动（agent 按此顺序）
```bash
# 1) 准备 + 抽取英文文本（打印翻译/润色提示词）
python scripts/pipeline.py setup <英文.docx> <中文.docx>

# 2) agent 翻译 extracted/<名>_en.json → <名>_zh.json（见 references/01 术语红线）
python scripts/pipeline.py apply-translation <英文.docx> <中文_zh.json> <中文.docx> --apply
#    （润色：agent 直接改 <中文.docx>）

# 3) 确定性校对 + 仿人工精读提示词 + 排版交付（出版级 docx）
python scripts/pipeline.py deliver <中文.docx> --apply
#    精读后若还有修订，再次 `deliver --apply`（幂等）。
```

细粒度子命令：`inventory` / `translate` / `proofread` / `typeset`，默认**干跑**，落盘须 `--apply`。

## 红线与铁律（务必先读）
- `references/01-translation.md`：零容错铁律、术语铁律（黏液/尤因/冬眠瘤/伴或不伴…）、标点规范、不译清单。
- `references/02-polish-closeread.md`：润色清单 + 仿人工精读清单。
- `references/03-delivery.md`：交付验收清单（字体/字号/缩进/行距/图片/体积对照）。

## 已固化的确定性红线（proof25，R1–R28 节选）
非侵袭性→非浸润性；显著/突出→明显；监测→随访；报告(文献义)→报道；关联→相关；
黏蛋白→黏液；胚系/种系/生殖细胞（体质性）→（体细胞）；它们→其；
形态学后缀前 外观→形态；分级 中至高→中或高；人/年 语序；分数 1/3；中文岁数→阿拉伯；
标题末尾去句号；图注补句号/去面板标号；范围 `-`/`至`→`～`。详见脚本内 `RULES` 与
上游技能 `who-en-zh-publishing-pipeline/references/09-proofreading-redlines.md`。

## 工程约束（fail-closed）
- docx = zip 容器：改 `word/document.xml`，整 zip 重建，逐部件保留 `compress_type`/时间戳/属性；
  写前良构校验 + `zip.testzip()` + 临时文件 + `os.replace` + 写后字节比对。
- 图片：只换 `word/media/` 字节，绝不改 `document.xml`；400 DPI 上限 + JPEG q95 + 4:2:2。
- 备份：每次落盘前自动 `_bak_*.docx`。
- **Word 占用锁定**：若 `os.replace` 报 WinError 32，只能写侧车文件（如 `<名>_校订版.docx`），
  待关闭占用进程后再覆盖回原名。

## 依赖
本技能自带确定性引擎 `scripts/who_docx.py`（inventory/punct-spaces/superscripts/captions/terms/
heading-periods/lagfix/images/typeset/verify）与 `scripts/proof25.py`（28 条红线），自包含可用。
翻译/润色/精读为 LLM 阶段，由 agent 依 references 执行。
