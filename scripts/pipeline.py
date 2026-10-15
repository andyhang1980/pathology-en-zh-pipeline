#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pipeline.py — 病理学 EN→ZH 出版级一体化流水线编排器（一体脚本）。

六大阶段：
  0 准备      inventory + 术语红线加载
  1 翻译/意译  extract 英文段落 → JSON（agent 译中并回注）   ← LLM 阶段
  2 校对       who_docx 子命令(punct/super/captions/terms/heading/lagfix) + proof25 红线
  3 润色       agent 按润色清单精修（句式/语序/通俗化）        ← LLM 阶段
  4 仿人工精读  agent 逐段对照英文精读，补漏纠错               ← LLM 阶段
  5 再精校      proof25 复跑至收敛 + typeset + images + verify → 出版级 docx

用法（agent 按 SKILL.md 顺序驱动）：
  python pipeline.py setup   <en.docx> <zh.docx>
        → 结构盘点 + 抽取英文段落 JSON + 打印「翻译/意译」「润色」提示词。
        agent 据此翻译 JSON、执行 apply-translation，再做润色。
  python pipeline.py apply-translation <en.docx> <zh_translated.json> <zh.docx> [--apply]
        → 把译文回注 docx（首文本 run 替换，保留样式/图片）。
  python pipeline.py deliver  <zh.docx> [--apply]
        → 确定性校对(2) + 打印「仿人工精读」提示词 + 排版+图片+验收 → 出版级 docx。
        agent 精读后若有修订，可再次 `deliver --apply`（幂等）。

也可用细粒度子命令单独跑：inventory / translate / proofread / typeset。
默认全部干跑；落盘须显式 --apply。fail-closed（每步 zip 校验 + 备份）。
"""
import sys, os, json, argparse, subprocess
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
WHO   = os.path.join(HERE, 'who_docx.py')
PROOF = os.path.join(HERE, 'proof25.py')
TRANS = os.path.join(HERE, 'translate_docx.py')

def run(cmd):
    print('\n$ ' + ' '.join(cmd))
    return subprocess.run([sys.executable] + cmd).returncode

def stage_proofread(src, apply):
    flags = ['--apply'] if apply else []
    print('\n========== 阶段2 确定性校对 ==========')
    for sub in ['punct-spaces', 'superscripts', 'captions', 'terms', 'heading-periods', 'lagfix']:
        run([WHO, sub, src, src] + flags)
    # proof25 红线：跑两遍吸收级联
    run([PROOF, src] + flags)
    if apply:
        run([PROOF, src] + flags)

def stage_typeset(src, apply):
    flags = ['--apply'] if apply else []
    print('\n========== 阶段5 排版 / 图片 / 验收 ==========')
    run([WHO, 'typeset', src, src] + flags)
    run([WHO, 'images', src, src] + flags)
    run([WHO, 'verify', src, src] + flags)

# ------------------------- 子命令 -------------------------
def cmd_inventory(a):
    run([WHO, 'inventory', a.src])

def cmd_translate(a):
    os.makedirs('extracted', exist_ok=True)
    jp = a.out
    run([TRANS, 'extract', a.src, jp])
    print(f"\n>>> 请翻译 {jp} 中每个 text 字段为中文，写入 'translation' 字段，保存。")
    print(">>> 术语红线见 SKILL.md / references/01-translation.md（黏液/尤因/冬眠瘤/伴或不伴…）。")

def cmd_apply_translation(a):
    run([TRANS, 'reinject', a.src, a.translated, a.dst] + (['--apply'] if a.apply else []))

def cmd_proofread(a):
    stage_proofread(a.src, a.apply)

def cmd_typeset(a):
    stage_typeset(a.src, a.apply)

# ------------------------- 一体流程 -------------------------
def cmd_setup(a):
    en, zh = a.src, a.out
    print('========== 阶段0 准备 / 结构盘点 ==========')
    run([WHO, 'inventory', en])
    os.makedirs('extracted', exist_ok=True)
    jp = os.path.join('extracted', os.path.splitext(os.path.basename(en))[0] + '_en.json')
    print('\n========== 阶段1 抽取英文段落 ==========')
    run([TRANS, 'extract', en, jp])
    print('\n' + '='*60)
    print('【LLM 阶段 · 翻译/意译】请 agent 执行：')
    print(f'  1. 读取 {jp}，将每个 text 译为规范中文，写入 translation 字段并保存。')
    print('  2. 严守术语红线（见 references/01-translation.md）：')
    print('      黏液(非粘液)/尤因肉瘤(非尤文)/冬眠瘤(非嗜睡细胞瘤)/伴或不伴(非伴有或不伴有)')
    print('      纤维母细胞/造釉细胞瘤/预后不良；基因名·抗体名·数字·ICD·分期分级 一律不译。')
    print('  3. 回注：python pipeline.py apply-translation ' + en + ' ' + jp.replace('_en.json','_zh.json') + ' ' + zh + ' --apply')
    print('  4. 【LLM 阶段 · 润色】对照英文逐句润色（句式/语序/通俗化），可直接改 ' + zh + '。')
    print('  5. 完成后：python pipeline.py deliver ' + zh + ' --apply')
    print('='*60)

def cmd_deliver(a):
    zh = a.src
    stage_proofread(zh, a.apply)
    print('\n' + '='*60)
    print('【LLM 阶段 · 仿人工精读】请 agent 执行：')
    print('  对照英文原版逐段精读，重点查：漏译/误译/术语不一致/数字单位(人/年)/')
    print('  标题层级/图注与正文对应。发现问题直接改 ' + zh + '，或再跑 proof25 红线。')
    print('  修订后再次：python pipeline.py deliver ' + zh + ' --apply （幂等）。')
    print('='*60)
    stage_typeset(zh, a.apply)
    print('\n✅ 交付（出版级）：' + zh)

def main():
    ap = argparse.ArgumentParser(prog='pipeline.py', description='病理学 EN→ZH 出版级一体化流水线')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('inventory'); p.add_argument('src'); p.set_defaults(func=cmd_inventory)
    p = sub.add_parser('translate'); p.add_argument('src'); p.add_argument('out'); p.set_defaults(func=cmd_translate)
    p = sub.add_parser('apply-translation'); p.add_argument('src'); p.add_argument('translated'); p.add_argument('dst'); p.add_argument('--apply', action='store_true'); p.set_defaults(func=cmd_apply_translation)
    p = sub.add_parser('proofread'); p.add_argument('src'); p.add_argument('--apply', action='store_true'); p.set_defaults(func=cmd_proofread)
    p = sub.add_parser('typeset'); p.add_argument('src'); p.add_argument('--apply', action='store_true'); p.set_defaults(func=cmd_typeset)
    p = sub.add_parser('setup'); p.add_argument('src'); p.add_argument('out'); p.set_defaults(func=cmd_setup)
    p = sub.add_parser('deliver'); p.add_argument('src'); p.add_argument('--apply', action='store_true'); p.set_defaults(func=cmd_deliver)

    a = ap.parse_args()
    a.func(a)

if __name__ == '__main__':
    main()
