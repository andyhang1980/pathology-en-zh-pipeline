# -*- coding: utf-8 -*-
"""fix_proof25.py — 按「用户人工校对归纳的 25 条规则」做二次校对（出版级流水线阶段 4 补刀）。

场景：用户人工校对中文稿后说"根据我校对的问题，进行归纳总结，然后进行再次校对"，
并给出一批规则。本脚本是这套工作流的参考实现，可直接复用到其它书卷。

实证（WHO6 消化系统肿瘤中文版，2026-10-14）：
  第 1 遍 1732 处 / 37 规则 → 第 2 遍 170 处 / 11 规则（吸收级联）→ 第 3 遍干跑 0 处收敛。

用法：
  python fix_proof25.py <src.docx> [<out.docx>] [--apply]
  不写 <out.docx> 则原地覆盖（先建 .bak_proof25_<时间戳>.docx 备份）。默认干跑。

⚠️ 必须跑至少两遍（级联吸收）：
  某些规则会**新暴露**其它规则的目标。例：原文「1/3至三分之二」，R12 先把「三分之二」
  变成「2/3」，此后 R16d 才看得见「1/3至2/3」。第 2 遍干跑 = 0 才算收敛。

三个坑（改脚本前必读，本脚本已含对策）：
  1) 哨兵占位法失效：在原串上只记录 span、从不真正替换 → 还原 pair 永远找不到目标，
     保护词被写成字面哨兵；且替换词起点可能早于保护词起点，按 (start, order) 排序
     反而挤掉保护。**对策**：等长占位符 \\uE000 逐字替换保护词（位置与原文 1:1 对齐），
     规则跑完后按原位置还原。
  2) 「而且」被拆成「而，且」：环视放在含 `([而并])?` 可选组的整体之前，导致从「且」
     起匹配、group(1)=None。**对策**：拆两条 —— `([而并])且`（环视在 而/并 之前）
     + `(?<![而并])且`。
  3) 图注漏改：图注段只走图注规则、不走文本规则，残留近百处。**对策**：图注先走文本
     规则再走图注规则。另「A、B.」一次 finditer 只摘一半（分隔符被消费）→ 改为标签串
     整体匹配，且要求标签串以「．」收尾，从而避开 `chromogranin A、synaptophysin` 误报。

安全约定（与 SKILL.md 全局铁律一致）：
  整 zip 重建，逐部件保留 compress_type / date_time / external_attr / internal_attr /
  create_system（媒体图多为 STORED，一律改 DEFLATED 会让几百 MB 图片白跑一趟）；
  落盘前 minidom 良构 + zip.testzip() 校验；先备份；写后字节比对。幂等（复跑 0 改动）。
"""
import re, sys, os, io, shutil, zipfile, time, argparse, xml.dom.minidom as md
sys.stdout.reconfigure(encoding='utf-8')

PUNCT_CLASS = r"[，。；：、！？…“”‘’\"'()（）【】〔〕《》〈〉—－~～·\s\]\-]"
PH = '\uE000'

# 保护词：这些词里的「报告 / 关联 / 黏蛋白」不是规则目标，绝不能被改。
# 用等长占位符逐字替换，位置与原文 1:1 对齐；规则跑完后按原位置还原。
PROTECT = ['病理报告', '诊断报告', '报告规范', '报告标准', '报告系统',
           '关联研究',
           '黏蛋白卡红']

RULES = [
    # ---- R24 人/年语序（放在最前：先归一分母形态，后面的范围规则才能看见）----
    ("R24x 千分位空格", re.compile(r'(?<![\d])100000(?![\d])'), '100 000', None),
    ("R24d 人口区间", re.compile(r'每\s*([\d\s～]+)\s*人\s*口\s*([\d.]+)\s*例?\s*(?:～|至|到)\s*([\d.]+)\s*例'),
        r'\2例～\3例/\1人口', None),
    ("R24e 人口单值", re.compile(r'每\s*([\d\s～]+)\s*人\s*口\s*([\d.]+)\s*例'), r'\2例/\1人口', None),
    ("R24a 人/年区间", re.compile(r'每\s*([\d\s]+)\s*人\s*/\s*年\s*([\d.]+)\s*例?\s*(?:～|至|到)\s*([\d.]+)\s*例'),
        r'\2例～\3例/\1人/年', None),
    ("R24b 人/年单值", re.compile(r'每\s*([\d\s～]+)\s*人\s*/\s*年\s*([\d.]+)\s*例'), r'\2例/\1人/年', None),
    ("R24c 万人年区间", re.compile(r'每\s*([\d\s]+)\s*万\s*人\s*年\s*([\d.]+)\s*例?\s*(?:～|至|到)\s*([\d.]+)\s*例'),
        r'\2例～\3例/\1万人年', None),
    ("R24f 万人年单值", re.compile(r'每\s*([\d\s]+)\s*万\s*人\s*年\s*([\d.]+)\s*例'), r'\2例/\1万人年', None),
    ("R24g 万人单值", re.compile(r'每\s*([\d\s]+)\s*万\s*人\s*(?:中|内)?\s*([\d.]+)\s*例'), r'\2例/\1万人', None),
    ("R24h 约X/年", re.compile(r'约(\d[\d\s]*?)\s*/\s*年'), r'约\1例/年', None),

    # ---- R15 胚系 ----
    ("R15 胚系/种系/生殖细胞（体质性）→（体细胞）", re.compile(r'(胚系|种系|生殖细胞)（体质性）'), r'\1（体细胞）', None),
    ("R15b 致病性变异，引起的", re.compile(r'致病性变异，引起的'), '致病性变异引起的', None),

    # ---- R2 转折/连接词前加逗号 ----
    ("R2a 但前加逗号", re.compile(r'(?<!%s)(?<=[\s\S])但' % PUNCT_CLASS), '，但', None),
    ("R2b 而且/并且前加逗号", re.compile(r'(?<!%s)(?<=[\s\S])([而并])且' % PUNCT_CLASS), r'，\1且', None),
    ("R2c 且前加逗号", re.compile(r'(?<![而并])(?<!%s)(?<=[\s\S])且' % PUNCT_CLASS), '，且', None),

    # ---- R3 突出→明显（仅"显著/显眼"义；"外凸"义保留）----
    ("R3 突出→明显", None, None, [
        ('尤为突出', '尤为明显'), ('更加突出', '更加明显'), ('突出和可见', '明显和可见'),
        ('核仁突出', '核仁明显'), ('突出的绒毛状', '明显的绒毛状'), ('突出的类挖空细胞', '明显的类挖空细胞'),
        ('细胞边界突出', '细胞边界明显'), ('显示突出', '显示明显')]),

    # ---- R4 显著→明显（排除统计学术语）----
    ("R4 显著→明显", re.compile(r'(?<![无未不并学差])显著(?!相关|异质|性差异|显著性)'), '明显', None),

    # ---- R5 分层→复层 ----
    ("R5 核分层→复层", None, None, [
        ('核分层', '核复层'), ('分层拉长', '复层拉长'), ('杂乱分层', '杂乱复层'), ('无分层', '无复层')]),

    # ---- R6 监测→随访 ----
    ("R6 监测→随访", re.compile(r'监测'), '随访', None),

    # ---- R9 侵袭→浸润（定向，禁止把 aggressive 义的"侵袭性"一律改掉）----
    ("R9 非侵袭性→非浸润性", re.compile(r'非侵袭性'), '非浸润性', None),
    ("R9b 浸润性侵袭→浸润性生长", re.compile(r'浸润性(?:类型的?)?侵袭'), '浸润性生长', None),
    ("R9c 不依赖侵袭→不依赖浸润", re.compile(r'不依赖侵袭'), '不依赖浸润', None),
    ("R9d 在侵袭性肿瘤中", re.compile(r'在侵袭性肿瘤中'), '在浸润性肿瘤中', None),

    # ---- R12 分数数学式 ----
    ("R12 分数数学式", None, None, [
        ('三分之一', '1/3'), ('三分之二', '2/3'), ('四分之三', '3/4'), ('四分之一', '1/4'), ('二分之一', '1/2')]),
    ("R12b 半数→50%", None, None, [('半数病例', '50%的病例'), ('半数患者', '50%的患者'), ('半数', '50%')]),

    # ---- R13 报告→报道（仅文献叙述义；病理报告义受 PROTECT 保护）----
    ("R13 报告→报道", None, None, [
        ('荟萃分析报告', '荟萃分析报道'), ('病例报告', '病例报道'), ('研究报告', '研究报道'), ('文献报告', '文献报道'),
        ('篇报告', '篇报道'),
        ('报告患病率', '报道患病率'), ('报告发生率', '报道发生率'),
        ('据报告', '据报道'), ('已有报告', '已有报道'), ('报告不足', '报道不足'),
        ('尚未报告', '尚未报道'), ('未报告', '未报道'), ('所报告', '所报道'), ('一些报告', '一些报道'),
        ('不同报告中', '不同报道中'), ('报告的', '报道的'),
        ('患者报告有', '患者有'), ('就诊时报告有', '就诊时有'),
        ('已报告', '已报道'), ('报告了', '报道了'), ('报告称', '报道称'),
        ('报告显示', '报道显示'), ('有报告', '有报道'), ('的报告', '的报道')]),

    # ---- R14 语序 ----
    ("R14 发病率报道中病例的增加", re.compile(r'发病率报道中的增加'), '发病率报道中病例的增加', None),

    # ---- R16 范围→～（必须带单位可选组，否则 1%至12% 全漏）----
    ("R16a 第X至第Y", re.compile(r'第(\d+)至第(\d+)'), r'第\1～第\2', None),
    ("R16b 第X到第Y", re.compile(r'第(\d+)到第(\d+)'), r'第\1～第\2', None),
    ("R16c X期到Y期", re.compile(r'(\d)期到([IVX]+)期'), r'\1期～\2期', None),
    ("R16d 范围至→～", re.compile(r'(\d+(?:\.\d+)?)([%％]?)\s*(?:至|到)\s*(近|约)?(\d+(?:\.\d+)?)([%％]?)'),
        r'\1\2～\3\4\5', None),

    # ---- R17 MSI 连接符 ----
    ("R17 MSI连接符→半角", re.compile(r'MSI[\uFF0D\u2010\u2011\u2012\u2013\u2014\u2015\u2212]H'), 'MSI-H', None),
    ("R17b MSI-high→MSI-H", None, None, [('MSI-high', 'MSI-H')]),

    # ---- R19 关联→相关（"关联研究"受 PROTECT 保护；兜底 pair 放最后）----
    ("R19 关联→相关", None, None, [
        ('关联性', '相关性'), ('相关联', '相关'), ('相关的关联因素', '相关因素'),
        ('之间的关联', '之间的相关性'), ('这些关联', '这些相关性'), ('流行病学关联', '流行病学相关性'),
        ('明确关联', '明确相关性'), ('直接关联', '直接相关性'), ('分子关联', '分子相关性'),
        ('表型关联', '表型相关性'), ('综合征关联', '综合征相关性'), ('临床关联', '临床相关性'),
        ('影像学关联', '影像学相关性'), ('存在关联', '存在相关性'), ('的关联', '的相关性'),
        ('该关联', '该相关性'), ('无关联', '无相关性'), ('关联', '相关')]),

    # ---- R21 黏蛋白→黏液 ----
    ("R21 黏蛋白→黏液", None, None, [('黏蛋白池', '黏液湖'), ('黏蛋白', '黏液')]),

    # ---- R23 被识别 ----
    ("R23 被识别→尚未诊断", re.compile(r'未被识别的综合征性疾病'), '尚未诊断的综合征性疾病', None),

    # ---- R25 中文岁数→阿拉伯 ----
    ("R25 平均年龄六十多岁", re.compile(r'(平均年龄[为在])六十多岁'), r'\g<1>60岁左右', None),
    ("R25 中文岁数", None, None, [
        ('四五十岁', '40～50岁'), ('五六十岁', '50～60岁'), ('二十多岁', '20多岁'), ('三十多岁', '30多岁'),
        ('四十多岁', '40多岁'), ('五十多岁', '50多岁'), ('六十多岁', '60多岁'), ('七十多岁', '70多岁'),
        ('八十多岁', '80多岁'), ('九十多岁', '90多岁')]),

    # ---- R26 它们→其（用户2026-10-14：它们可删或用其替代；取确定性更强的“其”）----
    ("R26 它们→其", re.compile(r'它们'), '其', None),

    # ---- R27 外观→形态（用户2026-10-14：外观有时译形态更好）
    #       仅当“外观”前为形态学后缀（样/状/型/性）时改，避免污染“大体/肉眼/切面外观”等肉眼/大体语境。
    #       模糊的 148 处（如“细胞学外观”“表面呈斑驳外观”）暂不处理，待用户确认是否扩大范围。----
    ("R27 外观→形态(样/状/型/性+外观)", re.compile(r'(样|状|型|性)外观'), r'\1形态', None),

    # ---- R28 分级“至”→“或”（用户2026-10-14：“至”使用不规范，如“中至高”应为“中或高”）
    #       仅当“至”连接两个等级词（中/中等/中间级/中级别/中低度/低度/低 + 高/高级别/高度/高等）时改；
    #       真实范围（如“黏液样至高度胶原化”）不被命中。----
    ("R28 分级至→或(等级词+至+高*)", re.compile(r'(中等|中间级|中级别|中低度|中度|低度|低|中)至(高(?:级别|度|等)?)'), r'\1或\2', None),
]

# ==================== run 感知文本替换引擎 ====================

def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def collect_segs(p):
    """抽取段内每个 <w:t> 的绝对偏移与文本（Word 常把一个短语拆进多个 run）。"""
    segs = []
    for rm in re.finditer(r'<w:r\b[^>]*>.*?</w:r>', p, flags=re.DOTALL):
        base = rm.start(); inner = rm.group(0)
        if re.search(r'<w:(br|tab|object|drawing|pict|sym|commentR)\b', inner):
            continue
        for tm in re.finditer(r'<w:t(?:\s[^>]*)?>.*?</w:t>', inner, flags=re.DOTALL):
            cm = re.match(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', tm.group(0), flags=re.DOTALL)
            if not cm:
                continue
            a = base + tm.start() + len(cm.group(1))
            segs.append((a, a + len(cm.group(2)), cm.group(2)))
    return segs

def protect(full):
    """把保护词逐字替换为等长占位符，位置与原文 1:1 对齐。"""
    if PH in full:
        return full
    chars = list(full); i = 0; n = len(full)
    while i < n:
        for a in PROTECT:
            if full.startswith(a, i):
                for k in range(len(a)):
                    chars[i + k] = PH
                i += len(a); break
        else:
            i += 1
    return ''.join(chars)

def unprotect(s, start, full):
    """按原位置还原占位符。"""
    if PH not in s:
        return s
    return ''.join(full[start + k] if c == PH else c for k, c in enumerate(s))

def _ctx(work, m):
    s0 = max(0, m.start() - 16)
    return work[s0:m.end() + 16]

def apply_text_rules(p, counts, samples):
    segs = collect_segs(p)
    if not segs:
        return p
    full = ''.join(s[2] for s in segs)
    work = protect(full)
    ranges = []; pos = 0
    for s in segs:
        ranges.append((pos, pos + len(s[2]))); pos += len(s[2])
    spans = []; order = 0
    for name, rx, repl, lex in RULES:
        if lex:
            for a, b in lex:
                if a == b:
                    order += 1; continue
                ms = list(re.finditer(re.escape(a), work))
                if ms:
                    counts[name] = counts.get(name, 0) + len(ms)
                    samples.setdefault(name, []).append(_ctx(work, ms[0]))
                    for m in ms:
                        spans.append((m.start(), m.end(), b, order))
                order += 1
        elif rx:
            ms = list(rx.finditer(work))
            if ms:
                counts[name] = counts.get(name, 0) + len(ms)
                samples.setdefault(name, []).append(_ctx(work, ms[0]))
                for m in ms:
                    t = repl(m) if callable(repl) else m.expand(repl)
                    spans.append((m.start(), m.end(), t, order))
            order += 1
    if not spans:
        return p
    # order = 规则在 RULES 中的序号：解决"关联研究必须先于关联"这类顺序依赖
    spans.sort(key=lambda x: (x[0], x[3]))
    merged = []
    for s, e, t, o in spans:
        if merged and s < merged[-1][1]:
            continue
        merged.append((s, e, t))
    emitted = set(); newtexts = [''] * len(segs)
    for si, (a, b) in enumerate(ranges):
        pieces = []; cur = a
        for k, (s, e, t) in enumerate(merged):
            if e <= a or s >= b:
                continue
            os_, oe = max(s, a), min(e, b)
            if os_ > cur:
                pieces.append(unprotect(work[cur:os_], cur, full))
            if k not in emitted:
                pieces.append(t); emitted.add(k)
            cur = oe
        if cur < b:
            pieces.append(unprotect(work[cur:b], cur, full))
        newtexts[si] = ''.join(pieces)
    edits = [(a, b, nt) for (a, b, old), nt in zip(segs, newtexts) if nt != old]
    if not edits:
        return p
    res = []; last = 0
    for a, b, nt in edits:
        res.append(p[last:a]); res.append(esc(nt)); last = b
    res.append(p[last:])
    return ''.join(res)

# ==================== 图注（R10 / R11）====================

CAP_END_OK = '。！？…”：）)』」.%、,'
CAP_MARK = re.compile(
    r'由|显示|可见|注意|该患者|该病变|该病例|该肿瘤|此例|表现为|可见于|提示|染色|镜下|见图|伴有|存在|'
    r'覆以|被覆|衬覆|具有|显示出|表现出|其特征|相似|诊断为|可见到|显示为|构成|组成|覆盖|位于|发生于|'
    r'呈现|呈|累及|侵犯|浸润|形成')
# 图注面板标签串：A. / A、B. / A～C. / D、E. ；要求以「．」收尾，避开 chromogranin A、synaptophysin 误报
CAP_LABEL = re.compile(r'(?:^|[。；;，,\s])((?:[A-Z](?:[～～~\-–—][A-Z])?[.、])*[A-Z](?:[～～~\-–—][A-Z])?\.)')

def fix_caption(p, counts, samples):
    ts = list(re.finditer(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', p, flags=re.DOTALL))
    if not ts:
        return p
    full = ''.join(x.group(2) for x in ts)
    new2 = full; n = 0
    while True:  # 循环到不动点：一次 finditer 会消费掉标签间的分隔符
        out = []; last = 0; hit = 0
        for mm in CAP_LABEL.finditer(new2):
            out.append(new2[last:mm.start(1)]); last = mm.end(1); hit += 1
        out.append(new2[last:]); nxt = ''.join(out)
        if not hit or nxt == new2:
            break
        new2 = nxt; n += hit
    if n:
        counts["R11 图注去A./B."] = counts.get("R11 图注去A./B.", 0) + n
        samples.setdefault("R11 图注去A./B.", []).append(full[:44])
    if new2 and new2[-1] not in CAP_END_OK and len(new2) >= 8 \
       and (CAP_MARK.search(new2) or '，' in new2):
        new2 += '。'
        counts["R10 图注加句号"] = counts.get("R10 图注加句号", 0) + 1
        samples.setdefault("R10 图注加句号", []).append(new2[:44])
    if new2 == full:
        return p
    res = []; last = 0
    for k, x in enumerate(ts):
        res.append(p[last:x.start(2)])
        if k == 0:
            res.append(esc(new2))
        last = x.end(2)
    res.append(p[last:])
    return ''.join(res)

def is_caption(p):
    return bool(re.search(r'<w:pStyle[^>]*w:val="ImageCaption"', p))

# R8 图片前空行（line=400 exact，与正文视觉间隔一致）
EMPTY_P = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:p>'

# ==================== 主流程 ====================

def rebuild_zip(src, out_path, new_docxml_bytes):
    """整 zip 重建，仅替换 word/document.xml，其余部件按原字节 + 原元数据拷贝。"""
    with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(out_path, 'w') as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                data = new_docxml_bytes
            else:
                data = zin.read(item.filename)
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zout.writestr(zi, data)

def process(src, out, apply):
    with zipfile.ZipFile(src) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    matches = list(re.finditer(r'<w:p\b.*?</w:p>', xml, flags=re.DOTALL))
    starts = [m.start() for m in matches]
    hasimg = [bool(re.search(r'<w:drawing>|<w:pict>|<w:object>', m.group(0))) for m in matches]

    counts = {}; samples = {}
    new_paras = {}; insert_before = {}

    tbl_spans = [(m.start(), m.end()) for m in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, flags=re.DOTALL)]
    def enclosing_tbl_start(pos):
        for a, b in tbl_spans:
            if a < pos < b:
                return a
        return None

    for i, m in enumerate(matches):
        p = m.group(0); cap = is_caption(p)
        # 坑3 对策：图注先走文本规则，再走图注专属规则
        if cap:
            p2 = fix_caption(apply_text_rules(p, counts, samples), counts, samples)
        else:
            p2 = apply_text_rules(p, counts, samples)
        if i + 1 < len(matches) and hasimg[i + 1] and not hasimg[i] and not cap:
            prev = re.sub(r'<[^>]+>', '', p).strip()
            if len(prev) > 5:
                # 必须定位外层 <w:tbl> 起始偏移：图前插 <w:p> 会破坏表格子元素结构
                ts = enclosing_tbl_start(starts[i + 1])
                if ts is not None and ts not in insert_before:
                    insert_before[ts] = EMPTY_P
                    counts["R8 正文后空行"] = counts.get("R8 正文后空行", 0) + 1
                    samples.setdefault("R8 正文后空行", []).append(prev[:44])
        if p2 != p:
            new_paras[i] = p2

    print("=" * 76)
    print(f"FILE: {os.path.basename(src)}  段数={len(matches)}")
    print("=" * 76)
    tot = 0
    for k in sorted(counts):
        print(f"  [{k}] {counts[k]} 处")
        for s in samples.get(k, [])[:3]:
            print(f"        · {s}")
        tot += counts[k]
    print(f"\n  合计 {tot} 处命中 / {len(counts)} 条规则")

    if not apply:
        print("  [干跑] 加 --apply 生效")
        return True
    if tot == 0:
        print("  无改动，未重写文件（幂等收敛）")
        return True

    ops = []
    for idx, m in enumerate(matches):
        if idx in new_paras:
            ops.append((m.start(), m.end(), new_paras[idx]))
    for ts, txt in insert_before.items():
        ops.append((ts, ts, txt))
    ops.sort(key=lambda x: (x[0], x[1]))
    out_ch = []; last = 0
    for s, e, t in ops:
        if s < last:
            continue
        out_ch.append(xml[last:s]); out_ch.append(t); last = e
    out_ch.append(xml[last:])
    xml2 = ''.join(out_ch)

    # 1) XML 良构
    try:
        md.parseString(xml2.encode('utf-8'))
    except Exception as e:
        print(f"[ERR] XML 良构失败，原文件不动: {e}")
        return False
    new_bytes = xml2.encode('utf-8')

    # 2) 备份（同分区 copy2）
    ts_ = time.strftime('%Y%m%d_%H%M%S')
    bak = f"{src}.bak_proof25_{ts_}.docx"
    shutil.copy2(src, bak)
    print(f"  备份 -> {os.path.basename(bak)}")

    # 3) 写临时文件 + zip 完整性校验 + 替换
    target = out or src
    tmp = target + '.tmp'
    rebuild_zip(src, tmp, new_bytes)
    with zipfile.ZipFile(tmp) as zc:
        if zc.testzip() is not None:
            print("[ERR] zip.testzip 失败，原文件不动")
            try: os.remove(tmp)
            except OSError: pass
            return False
    os.replace(tmp, target)

    # 4) 写后字节比对
    with open(target, 'rb') as f:
        disk = f.read()
    with zipfile.ZipFile(io.BytesIO(disk)) as zc:
        if zc.read('word/document.xml') != new_bytes:
            print("[ERR] 写后字节校验失败")
            return False
    print(f"  ✅ 已落盘: {target} | zip_ok | well_formed | {len(disk):,} 字节")
    print("  ⚠️ 级联吸收：请再跑一遍干跑，命中应为 0；若非 0 则再 --apply 一次。")
    return True

def main():
    ap = argparse.ArgumentParser(description="按用户人工校对归纳的 25 条规则做二次校对（默认干跑）")
    ap.add_argument("src")
    ap.add_argument("out", nargs="?", default=None)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    ok = process(a.src, a.out, a.apply)
    sys.exit(0 if ok else 1)

if __name__ == "__main__":
    main()
