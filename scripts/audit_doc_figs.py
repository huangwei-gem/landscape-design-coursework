# -*- coding: utf-8 -*-
"""读回你的 Word 方案书，核四件事：图齐不齐、贴在对不对、印出来清不清楚、编号断没断。

这套检查是从一次真实返工里长出来的：目录里列了、正文里贴了，都不等于贴对了地方——
三张现状图全排在 1.2 末尾，1.3《现状植物》一张都没有，而图注自己写着「判读依据见 1.3」，
读者顺着图注去找就扑空。肉眼翻 50 页看不出这件事，机器一条正则就能抓出来。

四件事各自的红字：
  1) 齐   —— 图纸目录里登记的每一张都得以图注出现在正文；磁盘上出过的图都没用进报告也算不齐
  2) 对   —— 每张图落在哪一节，那一节的正文（不含图注行）就得点到这个图号或图名。
             图注行不算「正文提到」：让每张图靠自己的图注作证，这道门永远是绿的
  3) 清   —— 印到纸上的分辨率：像素数除以这个图在 Word 里实际占的宽度（从 docx 的
             wp:extent 读，不是猜版心）。宽像素不等于清楚，缩到 6cm 宽的 1000px 图只有 120dpi
  4) 断   —— 图号跳号／重号；正文引用了某个图号但全篇没有这张图的图注

用法：
    python audit_doc_figs.py --doc 方案书.docx --figdir 图纸目录 --registry 图纸目录.csv
    python audit_doc_figs.py --doc 方案书.docx --dpi-min 150 --json audit.json
registry CSV 列（可选）：图号,图名
"""
import argparse
import hashlib
import io
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
EMU_IN = 914400.0
IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp")


def tight(s):
    return re.sub(r"\s+", "", str(s).replace("　", ""))


def _blocks(doc):
    """正文段与表格行拉平；表格行给 ('row', 文本, [图, …])，正文段给 ('p', …)。"""
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for ch in doc.element.body.iterchildren():
        if isinstance(ch, CT_P):
            p = Paragraph(ch, doc)
            yield "p", p.text.strip(), _drawings(ch, doc)
        elif isinstance(ch, CT_Tbl):
            for row in Table(ch, doc).rows:
                cell = row.cells[0]
                txt = "|".join(x.text.strip() for x in cell.paragraphs).strip()
                yield "row", txt, _drawings(cell._tc, doc)


def _drawings(el, doc):
    """这一段落里的内联图：[{md5, 宽in, 高in, 像素宽, 像素高, 文件名?}, …]（按 docx 里存的字节算）。"""
    out = []
    for box in el.iter():
        if not (box.tag.endswith("}inline") or box.tag.endswith("}anchor")):
            continue
        ext = box.find("wp:extent", NS)
        blip = box.find(".//a:blip", NS)
        if ext is None or blip is None:
            continue
        rid = blip.get("{%s}embed" % NS["r"])
        part = doc.part.related_parts.get(rid)
        if part is None:
            continue
        blob = part.blob
        try:
            from PIL import Image
            px, py = Image.open(io.BytesIO(blob)).size
        except Exception:
            px = py = 0
        out.append({"md5": hashlib.md5(blob).hexdigest(), "宽in": float(ext.get("cx")) / EMU_IN,
                    "高in": float(ext.get("cy")) / EMU_IN, "px": px, "py": py})
    return out


def index_figdir(figdir):
    """磁盘上的图：md5→路径，另给 stem→[路径] 用来认同名 png/jpg 双份。"""
    by_md5, by_stem = {}, defaultdict(list)
    if not figdir or not Path(figdir).is_dir():
        return by_md5, by_stem
    from PIL import Image
    for p in sorted(Path(figdir).rglob("*")):
        if not (p.is_file() and p.suffix.lower() in IMG_EXT):
            continue
        try:
            by_md5[hashlib.md5(p.read_bytes()).hexdigest()] = p
        except OSError:
            continue
        by_stem[p.stem].append(p)
    return by_md5, by_stem


def split_headings(lines, h1, h2):
    """把行按章／节切开，返回 {章节路径: 正文拼接} 与 {图注: 落点章节}。图注行不进正文。"""
    cur, body, cap = ["", ""], defaultdict(str), []
    for l in lines:
        if h1.match(l):
            cur = [l, ""]
        elif h2.match(l):
            cur[1] = l
        elif CAP_RE.match(l):
            cap.append((l, "／".join(x for x in cur if x)))
            continue
        body["／".join(x for x in cur if x)] += l
    return body, cap


CAP_RE = re.compile(r"^图\s*\d+(?:[-－]\d+)?[　\s]")


def audit(doc_path, figdir=None, registry=None, dpi_min=150.0,
          codes=r"(?:PL|S|JD|XC)-?\d{2,3}", h1=r"^[一二三四五六七八九十]{1,3}、.{2,60}$",
          h2=r"^\d{1,2}\.\d{1,2}\s*[一-鿿].{1,60}$"):
    import docx
    doc = docx.Document(str(doc_path))
    CREE = re.compile(codes)
    rows = list(_blocks(doc))
    lines = [t for _k, t, _d in rows if t]
    body, cap = split_headings(lines, re.compile(h1), re.compile(h2))
    by_md5, by_stem = index_figdir(figdir)
    hard, warn, info = [], [], []
    reg = {}
    if registry:
        with open(registry, newline="", encoding="utf-8-sig") as f:
            for r in csv_rows(f):
                reg[r.get("图号", "").strip()] = (r.get("图名") or "").strip()

    # 图注 ↔ 它上面那几张图（真版式都是"图在上、号在下"）——用来查贴错文件
    pend, own_imgs = [], {}
    for _k, t, ds in rows:
        if t and CAP_RE.match(t):
            m = CREE.search(t)
            if m:
                own_imgs.setdefault(m.group(0), []).extend(pend)
            pend = []
        else:
            pend.extend(ds)

    # ── 1) 齐 ────────────────────────────────────────────────────────────────
    cap_code = {}
    for txt, sec in cap:
        m = CREE.search(txt)
        cap_code.setdefault(m.group(0) if m else txt[:12], (txt, sec))
    for code, nm in sorted(reg.items()):
        if code not in cap_code:
            hard.append("图纸未入册｜%s《%s》在方案书里没有以图注出现" % (code, nm))
        # 登记了却在磁盘上找不到同名图号的文件：目录与磁盘双向不符
        if figdir and not [p for ps in by_stem.values() for p in ps if code in p.stem]:
            hard.append("图纸目录与磁盘不符｜%s《%s》登记了，%s 里没有带这个图号的图"
                        % (code, nm, Path(figdir).name))
    # 贴错文件：图注写 PL-04，实际嵌进来的那张图的文件名却是 PL-03
    for code, ds in sorted(own_imgs.items()):
        for d in ds:
            src = by_md5.get(d["md5"])
            if src is None:
                continue
            mm = CREE.search(src.stem)
            if mm and mm.group(0) != code:
                hard.append("图文不符｜图注写 %s，实际嵌进来的文件是 %s（%s）"
                            % (code, mm.group(0), src.name))
    used = {d["md5"] for _k, _t, ds in rows for d in ds}
    if figdir:
        for stem, ps in sorted(by_stem.items()):
            if stem.endswith(("_thumb", "_预览")):
                continue
            if not any(hashlib.md5(p.read_bytes()).hexdigest() in used for p in ps
                       if p.is_file()):
                warn.append("出过的图没进报告｜%s（同名 png/jpg 双份，用过任一份即算用过）" % stem)

    # ── 2) 对：图注落点那节的正文必须点名这张图 ──────────────────────────────
    for own, (txt, sec) in sorted(cap_code.items()):
        if sec == "":
            hard.append("插图落点｜%s 落在任何章节之前（正文开头），没有一节在说它" % own)
            continue
        t = tight(body[sec])
        nm = reg.get(own, "")
        if own not in t and (not nm or tight(nm) not in t):
            hard.append("插图落点｜%s 贴在「%s」，这一节正文从头到尾没提过它" % (own, sec[:36]))

    # ── 3) 清：印到纸上的分辨率 ──────────────────────────────────────────────
    dpis = []
    for _k, _t, ds in rows:
        for d in ds:
            if d["宽in"] <= 0 or d["px"] <= 0:
                continue
            dpis.append((d["px"] / d["宽in"], d["px"], round(d["宽in"] * 2.54, 1),
                         by_md5.get(d["md5"], Path("(重编码过的图)"))))
    if dpis:
        dpis.sort()
        info.append("印到纸上的分辨率：最低 %.0fdpi（宽 %.1fcm／%dpx）—最高 %.0fdpi，共 %d 张"
                    % (dpis[0][0], dpis[0][2], dpis[0][1], dpis[-1][0], len(dpis)))
        for dpi, px, cm, p in dpis:
            if dpi < dpi_min:
                hard.append("印刷分辨率不够｜%.0fdpi＜%.0f｜%s（%dpx 印成 %.1fcm 宽）"
                            % (dpi, dpi_min, p.name, px, cm))

    # ── 4) 断：编号跳号／重号／正文引用悬空 ─────────────────────────────────
    labs = [CAP_RE.match(l).group(0).strip(" 　") for l in lines if CAP_RE.match(l)]
    seen, dup = set(), []
    for l in labs:
        if l in seen:
            dup.append(l)
        seen.add(l)
    if dup:
        hard.append("图号重号｜%s" % "、".join(sorted(set(dup))))
    tops = sorted({int(re.match(r"图\s*(\d+)", l).group(1)) for l in seen})
    gaps = [n for n in range(tops[0], tops[-1] + 1) if n not in tops] if tops else []
    if gaps:
        warn.append("图号跳号｜缺 图%s（子号 图4-1/4-2 这种不算跳）" % "、".join(map(str, gaps)))
    capjoin = " ".join(l for l, _s in cap)
    prose = " ".join(l for l in lines if not CAP_RE.match(l))
    for code in sorted(set(CREE.findall(prose))):
        if code not in capjoin:
            warn.append("正文引用了 %s，但全篇没有以它为图号的图注（贴了没写号，或根本没贴）" % code)

    nchar = sum(len(t) for _k, t, _d in rows)
    return {"插图": len(dpis), "字数": nchar, "图注": len(seen), "章节": len(body),
            "硬伤": hard, "提醒": warn, "信息": info}


def csv_rows(f):
    import csv
    return [{(k or "").strip(): (v or "") for k, v in r.items()} for r in csv.DictReader(f)]


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--doc", required=True, help="方案书／报告 .docx")
    ap.add_argument("--figdir", help="出过的图所在目录（用来查「出过的图全用上了没」）")
    ap.add_argument("--registry", help="图纸目录 CSV：图号,图名（权威清单，缺了就是硬伤）")
    ap.add_argument("--dpi-min", type=float, default=150.0)
    ap.add_argument("--codes", default=r"(?:PL|S|JD|XC)-?\d{2,3}",
                    help="图号的形状，默认 PL-01／S12／JD-03 这类；纯中文图名就留宽一点")
    ap.add_argument("--json", metavar="OUT")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    r = audit(a.doc, a.figdir, a.registry, a.dpi_min, a.codes)
    if a.json:
        Path(a.json).write_text(json.dumps(r, ensure_ascii=False, indent=1), encoding="utf-8")
    print("%s：%d 页内插图／%d 条图注／%d 字／%d 个章节段落" % (
        Path(a.doc).name, r["插图"], r["图注"], r["字数"], r["章节"]))
    for x in r["信息"]:
        print("  · " + x)
    for tag in ("硬伤", "提醒"):
        print("\n%s %d 条" % (tag, len(r[tag])))
        for q in r[tag]:
            print(("  ✗ " if tag == "硬伤" else "  · ") + q)
    return 1 if r["硬伤"] else 0


if __name__ == "__main__":
    sys.exit(main())
