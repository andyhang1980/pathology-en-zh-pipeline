#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
who_docx.py —— WHO 英→中出版级流水线工具箱（统一 CLI）

子命令：
  inventory       结构盘点（段落/样式/run/媒体/表格）
  punct-spaces    清除标点前后多余空格
  superscripts    mm² 与表格脚注 a/b/c/d 设为真上标
  captions        图注间距 / 面板标号加句号 / 项目符号 · 黑体加粗
  terms           批量术语替换（仅改 <w:t> 文本）
  heading-periods 清理标题末尾句号
  lagfix          Word 卡顿修复（rsid / proofState / 目录转静态 / 压缩 run 级 rFonts）
  images          图片盘点（analyze）与按显示 DPI 瘦身（shrink）
  typeset         按出版规范写版式（标题/正文/图注间距与缩进）
  verify          六项完整性验收

约定：
  * 默认【干跑】；确认无误再加 --apply 才会落盘。
  * 只重写 word/document.xml（或指定部件），其余按原字节拷贝，保留 compress_type 等 zip 元数据。
  * --apply 前自动建立 <原名>_bak_<子命令>_<时间戳>.docx 备份。

用法：
  python who_docx.py inventory book.docx --report inv.txt
  python who_docx.py punct-spaces a.docx b.docx --apply
"""
import argparse
import copy
import datetime
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

WURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RURI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XMLNS = "{http://www.w3.org/XML/1998/namespace}space"


def Q(t):
    """正确构造带命名空间的标签名（写成 W+'tag' 会静默匹配 0 处）。"""
    return "{%s}%s" % (WURI, t)


def emit(msg):
    """Windows 控制台打印中文可能崩溃，这里统一容错。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        sys.stdout.buffer.write((str(msg) + "\n").encode("utf-8", "replace"))


# ---------------------------------------------------------------- 读写基础

class DocxIO(object):
    """低风险读写：整包元数据保真 + 校验后置 + 原子替换。"""

    def __init__(self, src):
        self.src = src
        self.zf = zipfile.ZipFile(src, "r")
        self.infos = self.zf.infolist()

    def read(self, name):
        return self.zf.read(name)

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def write(self, dst, replacements, apply=False, backup=True):
        """replacements: {部件名: 新的 bytes}。其余部件按原字节+原元数据拷贝。"""
        if not apply:
            emit("  [干跑] 未 --apply，跳过落盘：%s" % dst)
            return None

        if backup and os.path.exists(dst):
            stamp = datetime.datetime.now().strftime("%m%d%H%M")
            bak = "%s_bak_%s.docx" % (os.path.splitext(dst)[0], stamp)
            shutil.copy2(dst, bak)
            emit("  备份 -> %s" % bak)

        tmp = dst + ".tmp"
        zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
        for item in self.infos:
            data = replacements.get(item.filename)
            if data is None:
                data = self.zf.read(item.filename)
            zout.writestr(item, data)          # 保留 compress_type / 时间戳等元数据
        zout.close()

        # 后置校验
        zt = zipfile.ZipFile(tmp, "r")
        bad_all = zt.testzip()
        n_media = sum(1 for n in zt.namelist() if n.startswith("word/media/"))
        declar_new = zt.read("word/document.xml").count(b"xmlns:")
        zt.close()
        if bad_all is not None:
            os.remove(tmp)
            raise RuntimeError("testzip 失败：%s（原文件未改动）" % bad_all)
        emit("  zip OK，媒体 %d 张， xmlns 声明 %d" % (n_media, declar_new))

        os.replace(tmp, dst)
        emit("  落盘 -> %s (%d bytes)" % (dst, os.path.getsize(dst)))
        return dst


def load_root(io, part="word/document.xml"):
    parser = etree.XMLParser(resolve_entities=False, huge_tree=True)
    return etree.fromstring(io.read(part), parser)


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------- 段落工具

def pstyle(p):
    pPr = p.find(Q("pPr"))
    if pPr is None:
        return None
    ps = pPr.find(Q("pStyle"))
    if ps is None:
        return None
    return ps.get(Q("val"))


def ptext(p):
    return "".join(t.text or "" for r in p.findall(Q("r")) for t in r.findall(Q("t")))


def run_text(r):
    return "".join(t.text or "" for t in r.findall(Q("t")))


def set_run_text(r, new):
    ts = r.findall(Q("t"))
    if not ts:
        t = etree.SubElement(r, Q("t"))
        t.set(XMLNS, "preserve")
        t.text = new
        return
    ts[0].text = new
    ts[0].set(XMLNS, "preserve")
    for extra in ts[1:]:
        r.remove(extra)


def iter_paras(root):
    return list(root.iter(Q("p")))


# ================================================================ 子命令

def cmd_inventory(args):
    io = DocxIO(args.src)
    root = load_root(io)
    paras = iter_paras(root)
    styles = {}
    for p in paras:
        st = pstyle(p) or "(none)"
        styles[st] = styles.get(st, 0) + 1
    media = [n for n in io.zf.namelist() if n.startswith("word/media/")]
    stored = sum(1 for n in media if io.zf.getinfo(n).compress_type == zipfile.ZIP_STORED)
    tbls = root.iter(Q("tbl"))
    n_tbl = sum(1 for _ in tbls)
    lines = []
    lines.append("文件: %s  (%.1f MB)" % (args.src, os.path.getsize(args.src) / 1048576.0))
    lines.append("段落总数: %d" % len(paras))
    lines.append("run 总数 : %d" % len(list(root.iter(Q("r")))))
    lines.append("w:t  总数: %d" % len(list(root.iter(Q("t")))))
    lines.append("表格总数 : %d" % n_tbl)
    lines.append("媒体总数 : %d  (STORED=%d, 其余 IMAGE 已压缩)" % (len(media), stored))
    lines.append("media 体积: %.1f MB" % (sum(io.zf.getinfo(n).file_size for n in media) / 1048576.0))
    lines.append("")
    lines.append("段落样式分布:")
    for k, v in sorted(styles.items(), key=lambda kv: -kv[1]):
        lines.append("  %-22s %d" % (k, v))
    out = "\n".join(lines)
    emit(out)
    if args.report:
        open(args.report, "w", encoding="utf-8").write(out)
        emit("报告 -> %s" % args.report)
    io.close()


# ------------------------------------------------------------ punct-spaces

SP = " \u00a0\u3000"
BEFORE_SET = "（。，：—·"
AFTER_SET = "·—"
RE_BEFORE = re.compile("[" + re.escape(SP) + "]+(?=[" + re.escape(BEFORE_SET) + "])")
RE_AFTER = re.compile("(?<=[" + re.escape(AFTER_SET) + "])[" + re.escape(SP) + "]+")


def cmd_punct_spaces(args):
    io = DocxIO(args.src)
    root = load_root(io)
    changed = 0
    paras = set()

    for p in root.iter(Q("p")):
        runs = p.findall(Q("r"))
        for r in runs:
            old = run_text(r)
            if not old:
                continue
            new = RE_BEFORE.sub("", old)
            new = RE_AFTER.sub("", new)
            if new != old:
                set_run_text(r, new)
                changed += 1
                paras.add(p)
        # 跨 run 边界
        runs = p.findall(Q("r"))
        for i in range(len(runs) - 1):
            ta, tb = run_text(runs[i]), run_text(runs[i + 1])
            if not ta or not tb:
                continue
            if ta[-1] in SP and tb[0] in BEFORE_SET:
                na = ta.rstrip(SP)
                if na != ta:
                    set_run_text(runs[i], na)
                    changed += 1
                    paras.add(p)
            if ta[-1] in AFTER_SET and tb[0] in SP:
                nb = tb.lstrip(SP)
                if nb != tb:
                    set_run_text(runs[i + 1], nb)
                    changed += 1
                    paras.add(p)

    emit("修正 run 数: %d   涉及段落: %d" % (changed, len(paras)))
    io.write(args.dst, {"word/document.xml": serialize(root)}, args.apply)
    io.close()


# ------------------------------------------------------------- superscripts

RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
]

LATIN = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
TARGET = set("abcd")
STAGE_RE = re.compile(
    r"pT\d[a-d]|TRG\s*\d[a-d]|(?<![A-Za-z])T\d[a-d]|(?<![A-Za-z])N\d[a-d]"
    r"|(?<![A-Za-z])M\d[a-d]|(?<![A-Za-z])G\d[a-d]", re.I)


def build_sup_rpr(src_pr):
    """按 CT_RPr 规范顺序重建 rPr，并插入 superscript。"""
    rpr = etree.Element(Q("rPr"))
    if src_pr is not None:
        children = list(src_pr)
        seen = []
        for name in RPR_ORDER:
            for c in children:
                if c.tag == Q(name) and c not in seen:
                    rpr.append(copy.deepcopy(c))
                    seen.append(c)
        for c in children:
            if c not in seen:
                seen.append(c)
                rpr.append(copy.deepcopy(c))
        # 已存在 vertAlign 时删除旧值
    for old in rpr.findall(Q("vertAlign")):
        rpr.remove(old)
    va = etree.Element(Q("vertAlign"))
    va.set(Q("val"), "superscript")
    ordered = list(rpr)
    names = [c.tag for c in ordered]
    try:
        idx = names.index(Q("fitText")) + 1
    except ValueError:
        try:
            idx = names.index(Q("rtl"))
        except ValueError:
            idx = len(ordered)
    rpr.insert(min(idx, len(ordered)), va)
    return rpr


def make_run(text, rpr_el):
    r = etree.Element(Q("r"))
    if rpr_el is not None:
        r.append(rpr_el)
    t = etree.SubElement(r, Q("t"))
    t.set(XMLNS, "preserve")
    t.text = text
    return r


def is_target_footnote(text, i):
    c = text[i]
    if c not in TARGET:
        return False
    if i > 0 and text[i - 1] in LATIN:
        return False
    if i + 1 < len(text) and text[i + 1] in LATIN:
        return False
    if i + 1 < len(text) and text[i + 1] == "-" and i + 2 < len(text) and text[i + 2] in LATIN:
        return False
    if i > 0 and text[i - 1] == "/":            # n/a
        return False
    if c == "c" and i + 1 < len(text) and text[i + 1] == ".":
        return False                            # HGVS c.
    if i > 0 and text[i - 1].isdigit():         # 整数下标：4a/1a级/解剖分段
        prev2 = text[i - 2] if i >= 2 else ""
        if prev2 not in LATIN and prev2 != ".":
            return False
    if STAGE_RE.search(text[max(0, i - 6):i + 1]):
        return False
    return True


def cmd_superscripts(args):
    io = DocxIO(args.src)
    root = load_root(io)
    n_mm2 = 0
    n_fn = 0
    TABLE_STYLES = {"TableText", "TableNote", "TableTitle"}

    for p in root.iter(Q("p")):
        st = pstyle(p)
        runs = p.findall(Q("r"))
        new_runs = []
        touched = False
        for r in runs:
            txt = run_text(r)
            # (1) mm² 字形 -> 真上标
            if args.mm2 and "\u00b2" in txt:
                parts = txt.split("\u00b2")
                seg = parts[0]
                if seg:
                    new_runs.append(make_run(seg, copy.deepcopy(r.find(Q("rPr")))))
                for k in range(1, len(parts)):
                    new_runs.append(make_run("2", build_sup_rpr(r.find(Q("rPr")))))
                    n_mm2 += 1
                    if parts[k]:
                        new_runs.append(make_run(parts[k], copy.deepcopy(r.find(Q("rPr")))))
                touched = True
                continue
            # (2) 表格脚注 a/b/c/d
            if args.footnotes and st in TABLE_STYLES:
                hits = [i for i, ch in enumerate(txt) if is_target_footnote(txt, i)]
                if hits:
                    seg = ""
                    prev = 0
                    for i in hits:
                        if prev < i:
                            seg += txt[prev:i]
                        if seg:
                            new_runs.append(make_run(seg, copy.deepcopy(r.find(Q("rPr")))))
                            seg = ""
                        new_runs.append(make_run(txt[i], build_sup_rpr(r.find(Q("rPr")))))
                        n_fn += 1
                        prev = i + 1
                    if prev < len(txt):
                        seg += txt[prev:]
                    if seg:
                        new_runs.append(make_run(seg, copy.deepcopy(r.find(Q("rPr")))))
                    touched = True
                    continue
            new_runs.append(r)
        if touched:
            pPr = p.find(Q("pPr"))
            for r in runs:
                p.remove(r)
            if pPr is not None:
                idx = list(p).index(pPr)
                for nr in new_runs:
                    idx += 1
                    p.insert(idx, nr)
            else:
                for nr in new_runs:
                    p.append(nr)

    emit("mm² 新上标: %d   表格脚注 abc 新上标: %d" % (n_mm2, n_fn))
    io.write(args.dst, {"word/document.xml": serialize(root)}, args.apply)
    io.close()


# ----------------------------------------------------------------- captions

PANEL = set("ABCDEF")
NOTBEFORE = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") | \
            set(chr(c) for c in range(0x4e00, 0x9fff + 1))
RE_PANEL = re.compile(r"(?<![A-Za-z\u4e00-\u9fff])([A-F])[" + re.escape(SP) + r"]+")
RE_PANEL_SP = re.compile(r"(?<![A-Za-z\u4e00-\u9fff])([A-F])\.[" + re.escape(SP) + r"]+")


def cmd_captions(args):
    io = DocxIO(args.src)
    root = load_root(io)
    n_space = n_panel = n_bullet = 0
    n_sep = 0

    for p in root.iter(Q("p")):
        if pstyle(p) != "ImageCaption":
            continue
        # (a) 段前段后 + 行距
        if args.spacing:
            pPr = p.find(Q("pPr"))
            if pPr is None:
                pPr = etree.SubElement(p, Q("pPr"))
            sp = pPr.find(Q("spacing"))
            if sp is None:
                sp = etree.SubElement(pPr, Q("spacing"))
            sp.set(Q("before"), str(args.before))
            sp.set(Q("after"), str(args.after))
            sp.set(Q("line"), str(args.line))
            sp.set(Q("lineRule"), args.line_rule)
        # (b) 面板标号加句号（两遍：同 run + 跨 run）
        if args.panel_labels:
            for r in p.findall(Q("r")):
                t = run_text(r)
                if not t:
                    continue
                n = RE_PANEL.sub(r"\1.", t)
                if n != t:
                    set_run_text(r, n)
                    n_panel += 1
                t2 = run_text(r)
                n2 = RE_PANEL_SP.sub(r"\1.", t2)
                if n2 != t2:
                    set_run_text(r, n2)
        # (c) 项目符号 · -> 黑体加粗 11pt
        if args.bullet_dot:
            T = ptext(p)
            dots = [i for i, ch in enumerate(T) if ch == "\u00b7"]
            if dots:
                is_bullet = set(dots) if len(dots) >= 2 else ({0} if dots[0] == 0 else set())
                off = 0
                for r in p.findall(Q("r")):
                    txt = run_text(r)
                    if not txt:
                        continue
                    local = [j for j, ch in enumerate(txt) if ch == "\u00b7" and (off + j) in is_bullet]
                    off += len(txt)
                    if not local:
                        continue
                    src_pr = r.find(Q("rPr"))
                    seg = ""
                    pieces = []
                    for j, ch in enumerate(txt):
                        if j in local:
                            if seg:
                                pieces.append((seg, False))
                                seg = ""
                            pieces.append((ch, True))
                        else:
                            seg += ch
                    if seg:
                        pieces.append((seg, False))
                    dot_runs = []
                    for s_txt, is_dot in pieces:
                        if is_dot:
                            pr = build_bullet_rpr()
                            n_bullet += 1
                        else:
                            pr = copy.deepcopy(src_pr)
                        dot_runs.append(make_run(s_txt, pr))
                    pPr = p.find(Q("pPr"))
                    idx = list(p).index(r)
                    p.remove(r)
                    for nr in dot_runs:
                        p.insert(idx, nr)
                        idx += 1
        # (d) 图注内标点前后空格
        if args.punct_spaces:
            for r in p.findall(Q("r")):
                t = run_text(r)
                if not t:
                    continue
                n = RE_BEFORE.sub("", t)
                n = RE_AFTER.sub("", n)
                if n != t:
                    set_run_text(r, n)
                    n_space += 1

    emit("间距更新段=%d  面板标号=%d  项目符号·=%d  标点空格=%d"
         % (count_style(root, "ImageCaption") if args.spacing else 0,
            n_panel, n_bullet, n_space))
    io.write(args.dst, {"word/document.xml": serialize(root)}, args.apply)
    io.close()


def count_style(root, style):
    return sum(1 for p in root.iter(Q("p")) if pstyle(p) == style)


def build_bullet_rpr():
    rpr = etree.Element(Q("rPr"))
    rf = etree.SubElement(rpr, Q("rFonts"))
    for a in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(Q(a), "黑体")
    b = etree.SubElement(rpr, Q("b"))
    b.set(Q("val"), "1")
    sz = etree.SubElement(rpr, Q("sz"))
    sz.set(Q("val"), "22")
    szc = etree.SubElement(rpr, Q("szCs"))
    szc.set(Q("val"), "22")
    return rpr


# ------------------------------------------------------------------- terms

def cmd_terms(args):
    io = DocxIO(args.src)
    rules = list(args.rule or [])
    if args.rules_file:
        for line in open(args.rules_file, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                rules.append(line)
    pairs = []
    for spec in rules:
        old, new = spec.split("=", 1)
        pairs.append((old, new))
    if not pairs:
        io.close()
        raise SystemExit("未提供替换规则（用 --rule 旧=新，可重复）")

    for name in ("word/document.xml",):
        xml = io.read(name).decode("utf-8")
        total = 0
        detail = []
        for old, new in pairs:            # 按给出顺序执行：先特殊/更长的，再一般的
            pat = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)

            def repl(m):
                nonlocal total
                if old not in m.group(2):
                    return m.group(0)
                cnt = m.group(2).count(old)
                total += cnt
                return m.group(1) + m.group(2).replace(old, new) + m.group(3)
            xml = pat.sub(repl, xml)
            detail.append("  %s -> %s" % (old, new))
        emit("替换规则:")
        emit("\n".join(detail))
        emit("命中总数: %d" % total)
        etree.fromstring(xml.encode("utf-8"))   # 良构校验
        io.write(args.dst, {name: xml.encode("utf-8")}, args.apply)
    io.close()


# ---------------------------------------------------------- heading-periods

HEAD_STYLES = {"1", "2", "3"}
SECTION_WHITELIST = re.compile(
    r"^(ICD-O\s*编码|ICD-11\s*编码|相关术语|部位|临床特征|流行病学|发病机制|肉眼观|诊断分子病理学)。$")


def cmd_heading_periods(args):
    io = DocxIO(args.src)
    root = load_root(io)
    n = 0
    dist = {}
    for p in root.iter(Q("p")):
        st = pstyle(p)
        if st is None:
            continue
        T = ptext(p)
        if not T.endswith("。"):
            continue
        ok = False
        if st in HEAD_STYLES:
            ok = True
        elif st == "SectionTitle" and SECTION_WHITELIST.match(T):
            ok = True
        if not ok:
            continue
        # 从后往前定位含「。」的那个 w:t
        ts = [t for r in p.findall(Q("r")) for t in r.findall(Q("t"))]
        for t in reversed(ts):
            if (t.text or "").endswith("。"):
                t.text = (t.text or "")[:-1]
                t.set(XMLNS, "preserve")
                n += 1
                dist[st] = dist.get(st, 0) + 1
                break
    emit("删除标题末尾句号: %d 处，分布 %s" % (n, dist))
    io.write(args.dst, {"word/document.xml": serialize(root)}, args.apply)
    io.close()


# ------------------------------------------------------------------ lagfix

def cmd_lagfix(args):
    io = DocxIO(args.src)
    parts = {}

    doc = io.read("word/document.xml").decode("utf-8")
    n_rsid = 0
    if args.rsids:
        before = doc
        doc = re.sub(r'\s+w:rsidR(?:Default|P|Pr)?="[^"]*"', "", doc)
        n_rsid = len(before) - len(doc)
    if args.strip_rfonts:
        before = doc
        doc = re.sub(r"<w:rFonts[^>]*/>", "", doc)
        n_rsid += (len(before) - len(doc))
    if args.toc in ("static", "keepfield"):
        doc = re.sub(r'<w:fldChar[^>]*w:fldCharType="begin"[^>]*>.*?</w:fldChar>', "", doc, flags=re.S)
        if args.toc == "static":
            doc = re.sub(r"<w:instrText[^>]*>.*?</w:instrText>", "", doc, flags=re.S)
    parts["word/document.xml"] = doc.encode("utf-8")
    emit("document.xml 清理字节: %d" % n_rsid)

    try:
        st = io.read("word/settings.xml").decode("utf-8")
    except KeyError:
        st = None
    if st is not None:
        if args.proof_state and "w:proofState" not in st:
            st = re.sub(r"(<w:settings[^>]*>)",
                        r'\1<w:proofState w:spelling="none" w:grammar="none"/>', st, count=1)
        st = re.sub(r'<w:rsids>.*?</w:rsids>', "", st, flags=re.S)
        parts["word/settings.xml"] = st.encode("utf-8")

    if args.lang_zh:
        try:
            th = io.read("word/theme/theme1.xml").decode("utf-8")
            if 'w:themeFontLang' in th or True:
                pass
        except KeyError:
            pass

    for name, data in parts.items():
        etree.fromstring(data)      # 良构校验
    io.write(args.dst, parts, args.apply)
    io.close()


# ------------------------------------------------------------------ images

EMU_PER_INCH = 914400
DXA = 635


def display_sizes(root, io):
    """返回 {media部件名: 最大显示宽度 EMU}。"""
    disp = {}
    for dw in root.iter(Q("drawing")):
        ext = dw.find(".//{%s}extent" % "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
        cx = None
        if ext is not None:
            cx = int(ext.get("cx") or 0)
        for blip in dw.iter("{%s}blip" % "http://schemas.openxmlformats.org/drawingml/2006/main"):
            rid = blip.get("{%s}embed" % RURI)
            if not rid or not cx:
                continue
            try:
                target = io.zf.read("word/_rels/document.xml.rels").decode("utf-8")
            except Exception:
                continue
            m = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid), target)
            if m:
                t = m.group(1)
                if not t.startswith("word/"):
                    t = "word/" + t.lstrip("/")
                disp[t] = max(disp.get(t, 0), cx)
    return disp


def jpeg_size(im, quality, subsampling, icc):
    from io import BytesIO
    b = BytesIO()
    kw = {"quality": quality, "subsampling": subsampling, "optimize": True}
    if icc:
        kw["icc_profile"] = icc
    im.save(b, "JPEG", **kw)
    return b.getvalue()


def cmd_images(args):
    io = DocxIO(args.src)
    if args.action == "analyze":
        media = [n for n in io.zf.namelist() if n.startswith("word/media/")]
        root = load_root(io)
        disp = display_sizes(root, io)
        lines = ["%-42s %10s %10s %8s" % ("部件", "体积KB", "像素宽", "有效DPI")]
        try:
            from PIL import Image
            have_pil = True
        except ImportError:
            have_pil = False
        for n in media:
            size = io.zf.getinfo(n).file_size
            w = "-"
            dpi = "-"
            if have_pil:
                from io import BytesIO
                try:
                    im = Image.open(BytesIO(io.zf.read(n)))
                    w = str(im.size[0])
                    if n in disp and disp[n]:
                        dpi = "%.0f" % (im.size[0] / (disp[n] / float(EMU_PER_INCH)))
                    im.close()
                except Exception as e:
                    w = "ERR"
            lines.append("%-42s %10.1f %10s %8s" % (n, size / 1024.0, w, dpi))
        out = "\n".join(lines)
        emit(out)
        if args.report:
            open(args.report, "w", encoding="utf-8").write(out)
        io.close()
        return

    if args.action == "shrink":
        from PIL import Image
        from io import BytesIO
        Image.MAX_IMAGE_PIXELS = None
        root = load_root(io)
        disp = display_sizes(root, io)
        media = [n for n in io.zf.namelist() if n.startswith("word/media/")]
        parts = {}
        saved = 0.0
        touched = 0
        for n in media:
            raw = io.zf.read(n)
            cx = disp.get(n, 0)
            if not cx:
                continue
            try:
                im = Image.open(BytesIO(raw))
            except Exception:
                continue
            cur = im.size[0] / (cx / float(EMU_PER_INCH))
            new_bytes = None
            if cur > args.dpi:
                tw = int(cx / float(EMU_PER_INCH) * args.dpi)
                if tw < im.size[0]:
                    ratio = tw / float(im.size[0])
                    im2 = im.resize((tw, max(1, int(im.size[1] * ratio))), Image.LANCZOS)
                else:
                    im2 = im
                if im2.mode not in ("RGB", "L", "CMYK"):
                    im2 = im2.convert("RGB")
                nb = jpeg_size(im2, args.quality, args.subsampling, im2.info.get("icc_profile"))
                # 收益 <5% 则保留原字节
                if len(nb) < len(raw) * 0.95:
                    new_bytes = nb
                if im2 is not im:
                    im2.close()
            im.close()
            if new_bytes is not None:
                parts[n] = new_bytes
                saved += (len(raw) - len(new_bytes)) / 1048576.0
                touched += 1
        emit("将重编码: %d 张，预计节省 %.1f MB" % (touched, saved))
        io.write(args.dst, parts, args.apply)
    io.close()


# ----------------------------------------------------------------- typeset

PRESETS = {
    # styleId -> (before, after, line, lineRule, firstLine, jc)
    "1": (480, 360, None, None, 0, "center"),
    "2": (360, 240, None, None, 0, "center"),
    "3": (240, 120, None, None, 0, "center"),
    "4": (120, 120, None, None, 0, "left"),
    "SectionTitle": (0, 0, None, None, 0, "left"),
    "ImageCaption": (120, 120, 240, "exact", 0, "center"),
}


def cmd_typeset(args):
    io = DocxIO(args.src)
    root = load_root(io)
    n = 0
    for p in root.iter(Q("p")):
        st = pstyle(p)
        spec = PRESETS.get(st)
        if st is None and args.body:
            spec = (0, 0, 360, "exact", 420, None)   # 正文：首行 2 字符 + 固定 18 磅
        if not spec:
            continue
        before, after, line, rule, first, jc = spec
        pPr = p.find(Q("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, Q("pPr"))
        sp = pPr.find(Q("spacing"))
        if sp is None:
            sp = etree.SubElement(pPr, Q("spacing"))
        sp.set(Q("before"), str(before))
        sp.set(Q("after"), str(after))
        # ⚠️ 图片段绝不能套固定行距
        if p.find(".//" + Q("drawing")) is not None:
            for attr in (Q("line"), Q("lineRule")):
                if sp.get(attr) is not None:
                    del sp.attrib[attr]
        elif line:
            sp.set(Q("line"), str(line))
            sp.set(Q("lineRule"), rule)
        if first is not None:
            ind = pPr.find(Q("ind"))
            if ind is None:
                ind = etree.SubElement(pPr, Q("ind"))
            ind.set(Q("firstLine"), str(first))
            if first:
                ind.set(Q("firstLineChars"), "200")
        if jc:
            j = pPr.find(Q("jc"))
            if j is None:
                j = etree.SubElement(pPr, Q("jc"))
            j.set(Q("val"), jc)
        reorder_ppr(pPr)
        n += 1
    emit("应用版式段落: %d" % n)
    io.write(args.dst, {"word/document.xml": serialize(root)}, args.apply)
    io.close()


PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
             "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
             "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
             "pPrChange"]


def reorder_ppr(pPr):
    """pPr 子元素顺序必须合规，否则 Word 报『无法读取的内容』。"""
    children = list(pPr)
    keyed = []
    rest = []
    for c in children:
        tag = c.tag.split("}")[-1]
        if tag in PPR_ORDER:
            keyed.append((PPR_ORDER.index(tag), c))
        else:
            rest.append(c)
    keyed.sort(key=lambda t: t[0])
    for c in children:
        pPr.remove(c)
    for _, c in keyed:
        pPr.append(c)
    for c in rest:
        pPr.append(c)


# ------------------------------------------------------------------ verify

def cmd_verify(args):
    io = DocxIO(args.src)
    rep = []

    zt = zipfile.ZipFile(args.src, "r")
    bad = zt.testzip()
    rep.append("1) ZIP testzip            : %s" % ("OK" if bad is None else "FAIL(%s)" % bad))
    n_media_src = sum(1 for n in zt.namelist() if n.startswith("word/media/"))
    decl_src = zt.read("word/document.xml").count(b"xmlns:")
    zt.close()

    try:
        root = load_root(io)
        n_p = len(list(root.iter(Q("p"))))
        n_r = len(list(root.iter(Q("r"))))
        rep.append("2) XML 良构               : OK（段落 %d / run %d）" % (n_p, n_r))
    except Exception as e:
        rep.append("2) XML 良构               : FAIL(%s)" % e)

    zero_ext = sum(1 for e in root.iter(
        "{%s}extent" % "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
        if not int(e.get("cx") or 0) or not int(e.get("cy") or 0))
    rep.append("3) 图片 extent 为 0       : %d（应为 0）" % zero_ext)

    # 注意：wp:docPr 的 @id 是无前缀属性，不能用 Q("id")（会全部返回 None）
    ids = [d.get("id") for d in root.iter(
        "{%s}docPr" % "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")]
    dup = len(ids) - len(set(ids))
    rep.append("4) docPr id 重复          : %d（应为 0）" % dup)

    bad_order = 0
    for pPr in root.iter(Q("pPr")):
        tags = [c.tag.split("}")[-1] for c in pPr]
        idxs = [PPR_ORDER.index(t) for t in tags if t in PPR_ORDER]
        if idxs != sorted(idxs):
            bad_order += 1
    rep.append("5) pPr 子元素顺序违规     : %d（应为 0）" % bad_order)

    if os.path.exists(args.dst):
        try:
            from xml.etree import ElementTree as ET
            io2 = DocxIO(args.dst)
            dec = serialize(load_root(io2)).count(b"xmlns:")
            z2 = zipfile.ZipFile(args.dst, "r")
            n_media_dst = sum(1 for n in z2.namelist() if n.startswith("word/media/"))
            z2.close()
            io2.close()
            rep.append("6) 与 dst 对比            : 媒体 %d→%d，xmlns %d→%d"
                       % (n_media_src, n_media_dst, decl_src, dec))
        except Exception as e:
            rep.append("6) 与 dst 对比            : FAIL(%s)" % e)

    out = "\n".join(rep)
    emit(out)
    if args.report:
        open(args.report, "w", encoding="utf-8").write(out)
    io.close()


# -------------------------------------------------------------------- main

def build_parser():
    ap = argparse.ArgumentParser(prog="who_docx.py", description="WHO 英→中出版级流水线工具箱")
    sub = ap.add_subparsers(dest="cmd")

    def add_io(sp, need_dst=True):
        sp.add_argument("src")
        if need_dst:
            sp.add_argument("dst")
        sp.add_argument("--apply", action="store_true", help="默认干跑，加此项才落盘")
        sp.add_argument("--report", help="报告输出文件")

    p = sub.add_parser("inventory", help="结构盘点")
    add_io(p, False)

    p = sub.add_parser("punct-spaces", help="清除标点前后多余空格")
    add_io(p)

    p = sub.add_parser("superscripts", help="mm² 与表格脚注 a/b/c/d 设为真上标")
    add_io(p)
    p.add_argument("--no-mm2", dest="mm2", action="store_false")
    p.add_argument("--no-footnotes", dest="footnotes", action="store_false")

    p = sub.add_parser("captions", help="图注间距/面板标号/项目符号")
    add_io(p)
    p.add_argument("--spacing", action="store_true")
    p.add_argument("--before", default=120)
    p.add_argument("--after", default=120)
    p.add_argument("--line", default=240)
    p.add_argument("--line-rule", default="exact")
    p.add_argument("--panel-labels", action="store_true")
    p.add_argument("--bullet-dot", action="store_true")
    p.add_argument("--punct-spaces", action="store_true")

    p = sub.add_parser("terms", help="批量术语替换")
    add_io(p)
    p.add_argument("--rule", action="append", help="旧=新，可重复")
    p.add_argument("--rules-file")

    p = sub.add_parser("heading-periods", help="清理标题末尾句号")
    add_io(p)

    p = sub.add_parser("lagfix", help="Word 卡顿修复")
    add_io(p)
    p.add_argument("--toc", choices=["static", "keepfield", "keep"], default="keep")
    p.add_argument("--rsids", action="store_true")
    p.add_argument("--strip-rfonts", action="store_true")
    p.add_argument("--proof-state", action="store_true")
    p.add_argument("--lang-zh", action="store_true")

    p = sub.add_parser("images", help="图片盘点与瘦身")
    p.add_argument("action", choices=["analyze", "shrink"])
    p.add_argument("src")
    p.add_argument("dst", nargs="?")
    p.add_argument("--dpi", default=400, type=int)
    p.add_argument("--quality", default=95, type=int)
    p.add_argument("--subsampling", default=1, type=int, choices=[0, 1, 2])
    p.add_argument("--apply", action="store_true")
    p.add_argument("--report")

    p = sub.add_parser("typeset", help="应用出版版式")
    add_io(p)
    p.add_argument("--body", action="store_true", help="同时处理正文（首行2字符+固定18磅）")

    p = sub.add_parser("verify", help="完整性验收")
    add_io(p)

    return ap


def main():
    ap = build_parser()
    args = ap.parse_args()
    if not args.cmd:
        ap.print_help()
        return
    fn = {
        "inventory": cmd_inventory,
        "punct-spaces": cmd_punct_spaces,
        "superscripts": cmd_superscripts,
        "captions": cmd_captions,
        "terms": cmd_terms,
        "heading-periods": cmd_heading_periods,
        "lagfix": cmd_lagfix,
        "images": cmd_images,
        "typeset": cmd_typeset,
        "verify": cmd_verify,
    }[args.cmd]
    fn(args)


if __name__ == "__main__":
    main()
