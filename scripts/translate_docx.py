#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
translate_docx.py — 病理学 EN→ZH 流水线的「翻译/意译」阶段支撑脚本。

两个子命令：
  extract   <src.docx> <out.json>
      抽取主文档（含表格内段落）每个段落的纯文本，按段落顺序编号写 JSON。
      - 跳过含 w:drawing/w:object 的 run（图片本身不译）。
      - 记录 pStyle（标题/小标题等），供翻译提示词区分处理。
      - 空段落 / 纯图片段落标记 skip=true。

  reinject  <src.docx> <translated.json> <dst.docx> [--apply]
      把翻译好的 JSON 按段落编号写回 docx：
      - 段落首文本 run 的 <w:t> 设为译文，删除其余文本 run（保留图片 run）。
      - 段落 pPr（缩进/对齐/样式）原样保留，仅换文字。
      默认干跑（打印前后对照），加 --apply 落盘。

工程要点：用 lxml 序列化，保留 w: 命名空间前缀（与 who_docx.py / proof25.py 一致），
绝不用 xml.etree（会把 w: 改写成 ns0: 导致下游工具失效）。fail-closed（校验+备份）。
"""
import sys, os, json, zipfile, shutil, argparse, time
from lxml import etree

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DOC = 'word/document.xml'
etree.register_namespace('w', WNS)

def q(t):
    return f'{{{WNS}}}{t}'

def para_text(p):
    """拼接段落内所有文本 run 的文字（跳过含 drawing/object 的 run）。"""
    parts = []
    for r in p.iter(q('r')):
        if r.find(q('drawing')) is not None or r.find(q('object')) is not None:
            continue
        t = r.find(q('t'))
        if t is not None and t.text:
            parts.append(t.text)
    return ''.join(parts)

def para_style(p):
    ppr = p.find(q('pPr'))
    if ppr is None:
        return ''
    ps = ppr.find(q('pStyle'))
    if ps is None:
        return ''
    return ps.get(q('val'), '')

def iter_paras(root):
    return root.iter(q('p'))

def load_root(src):
    z = zipfile.ZipFile(src)
    data = z.read(DOC)
    parser = etree.XMLParser(resolve_entities=False, huge_tree=True)
    return etree.fromstring(data, parser)

def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def cmd_extract(args):
    root = load_root(args.src)
    out = []
    for i, p in enumerate(iter_paras(root)):
        txt = para_text(p)
        style = para_style(p)
        has_drawing = any(r.find(q('drawing')) is not None or r.find(q('object')) is not None
                          for r in p.iter(q('r')))
        rec = {'id': i, 'style': style, 'text': txt, 'skip': (txt.strip() == '') or has_drawing}
        out.append(rec)
    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    n = len(out)
    n_skip = sum(1 for r in out if r['skip'])
    print(f"✅ 抽取 {n} 个段落（其中 {n_skip} 个空/纯图片，翻译时跳过） -> {args.out}")

def cmd_reinject(args):
    root = load_root(args.src)
    with open(args.translated, encoding='utf-8') as f:
        data = json.load(f)
    by_id = {r['id']: r for r in data}

    changed = 0
    for i, p in enumerate(iter_paras(root)):
        rec = by_id.get(i)
        if rec is None or rec.get('skip') or not rec.get('text'):
            continue
        zh = rec.get('translation', '').strip()
        if not zh:
            print(f"  ⚠ 段落 {i} 缺 translation，跳过")
            continue
        text_runs = [r for r in p.iter(q('r'))
                     if r.find(q('t')) is not None
                     and r.find(q('drawing')) is None
                     and r.find(q('object')) is None]
        if not text_runs:
            continue
        first = text_runs[0]
        ft = first.find(q('t'))
        ft.text = zh
        for r in text_runs[1:]:
            p.remove(r)
        changed += 1
        if not args.apply:
            print(f"  [{i}] ({rec.get('style','')})  EN: {rec['text'][:80]}")
            print(f"            ZH: {zh[:80]}")

    if not args.apply:
        print(f"\n[干跑] 将改写 {changed} 个段落；加 --apply 落盘。")
        return

    new_doc = serialize(root)
    ts = time.strftime('%Y%m%d_%H%M%S')
    bak = f"{args.src}.bak_translate_{ts}.docx"
    shutil.copy2(args.src, bak)
    tmp = args.dst + '.tmp'
    rebuild_zip(args.src, tmp, new_doc)
    if zipfile.ZipFile(tmp).testzip() is not None:
        print("[ERR] zip 校验失败，未落盘"); os.remove(tmp); return
    os.replace(tmp, args.dst)
    print(f"✅ 已落盘: {args.dst} | 改写 {changed} 段 | 备份 {os.path.basename(bak)}")

def rebuild_zip(src, dst, new_document_xml_bytes):
    """整 zip 重建，逐部件保留 compress_type / 元数据；仅替换 document.xml。"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == DOC:
                data = new_document_xml_bytes
            zi = zipfile.ZipInfo(item.filename)
            zi.compress_type = item.compress_type
            zi.date_time = item.date_time
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zout.writestr(zi, data)

def main():
    ap = argparse.ArgumentParser(description="EN→ZH 翻译抽取/回注")
    sub = ap.add_subparsers(dest='cmd', required=True)
    pe = sub.add_parser('extract'); pe.add_argument('src'); pe.add_argument('out'); pe.set_defaults(func=cmd_extract)
    pr = sub.add_parser('reinject'); pr.add_argument('src'); pr.add_argument('translated'); pr.add_argument('dst'); pr.add_argument('--apply', action='store_true'); pr.set_defaults(func=cmd_reinject)
    a = ap.parse_args()
    a.func(a)

if __name__ == '__main__':
    main()
