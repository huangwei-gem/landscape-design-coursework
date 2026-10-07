# -*- coding: utf-8 -*-
"""ezdxf 图纸脚手架：图框、图层、中文字体、指北针、直线比例尺、图签、表格一次配好，
再自带一张"这张图画坏了没有"的检查——越框、压图签、比例装不下、标注字太小、图签缺项。

为什么用它而不是在 CAD 里手画：图幅与比例是靠人记住的，而错误恰恰长在记性上（本项目
真出过"图签写 A3、图幅其实是 A4"、"标 1:500 但场地按 1:800 画"）。这里 mm_per_m =
1000/分母 只算一次，图面尺寸、比例尺分划、图签自述三处都由它派生，对不上就出红字。

三条口径（其余都能改）：
  · 模型单位＝毫米（$INSUNITS=4），图框按 ISO 幅面的真实毫米画；实地一米在图面上 1000/分母 mm
  · 幅面不够就换大幅面，不要靠打印时"适合页面"缩放——那会让图签上标的 1:500 变成假的
  · 原点置中（图框左下角放在 -pw/2-12, -ph/2-12），逐株放样坐标读起来才不是天文数字

    python dxf_sheet_scaffold.py --demo 图.dxf        # 干净的示例图，检查应当全绿
    python dxf_sheet_scaffold.py --demo-bad 坏图.dxf  # 故意越框＋压图签，验门会响
    python dxf_sheet_scaffold.py --preview 图.dxf 图.png   # 不装 AutoCAD 把 DXF 渲成 PNG

接到自己项目上：from dxf_sheet_scaffold import Sheet；一张图一个 Sheet 实例；
数据（红线、树位、苗木表）从你自己的数据层 JSON 读，别在这个文件里写死数字——
派生量一律回数据层算，否则点名门抓到的第一个分歧就是这张图。
"""
import argparse
import sys
from pathlib import Path

PAPER = {"A4": (297.0, 210.0), "A3": (420.0, 297.0), "A2": (594.0, 420.0), "A1": (841.0, 594.0)}
# 图层名一旦定了，方案书图例与点名门都认它；改名要图纸、图例、门三处一起改。
LAYERS = [("FRAME", 7, "CONTINUOUS", "图框、图签、比例尺"),
          ("REDLINE", 1, "CONTINUOUS", "用地红线"),
          ("ROAD", 8, "CONTINUOUS", "道路与铺装"),
          ("PLANT", 3, "CONTINUOUS", "乔木／灌木符号与林缘线"),
          ("LABEL", 2, "CONTINUOUS", "标注文字"),
          ("TABLE", 4, "CONTINUOUS", "表格"),
          ("ANNO", 6, "DASHED", "分析线、引出线")]
TITLE_FIELDS = ("图号", "图名", "比例", "图幅", "设计", "审核")     # 日期留空给手写
MIN_TEXT_H = 1.8          # 图面最小字高 mm：A1 打印后约 1.8mm，再小现场读不出来


def rect(r):
    """(x1,y1,x2,y2) → 四个角，给 add_lwpolyline 用。"""
    return [(r[0], r[1]), (r[2], r[1]), (r[2], r[3]), (r[0], r[3])]


class Sheet:
    def __init__(self, code, name, paper="A2", denom=500, line_cap=1400):
        import ezdxf
        self.ezdxf, self.code, self.name = ezdxf, code, name
        self.paper, self.denom, self.cap = paper, denom, line_cap
        self.mm_per_m = 1000.0 / denom                     # 实地一米画多长
        self.pw, self.ph = PAPER[paper]
        self.doc = ezdxf.new("R2010", setup=True)          # setup=True 才有 DASHED 等线型
        self.doc.header["$INSUNITS"] = 4                   # 毫米
        if "CN" not in self.doc.styles:
            # 不建中文文字样式，AutoCAD 里汉字全是问号；simhei.ttf 是 Windows 自带
            self.doc.styles.add("CN", font="simhei.ttf")
        self.ms = self.doc.modelspace()
        for n, c, lt, _u in LAYERS:
            if n not in self.doc.layers:
                self.doc.layers.add(n, color=c, linetype=lt)
        self.x0, self.y0 = -self.pw / 2 - 12, -self.ph / 2 - 12
        self.pad = 25.0 if self.pw >= 594 else 10.0        # 大幅面左侧留装订边
        self.fb = (self.x0 + self.pad, self.y0 + self.pad,
                   self.x0 + self.pw - 10.0, self.y0 + self.ph - 10.0)
        self.tbw, self.tbh = 180.0, 56.0
        self.tb = (self.fb[2] - self.tbw, self.fb[1], self.fb[2], self.fb[1] + self.tbh)
        self.warn, self.errs, self._dw = [], [], []
        self._frame()
        # 图框与图签自己的实体排在最前面：审计时按序号豁免，否则每张图都会
        # 被自己判成「压图签」「越框」。
        self._tb_n = len(self.ms)

    # ── 版面 ────────────────────────────────────────────────────────────────
    def _frame(self):
        m = self.ms
        m.add_lwpolyline([(self.x0, self.y0), (self.x0 + self.pw, self.y0),
                          (self.x0 + self.pw, self.y0 + self.ph), (self.x0, self.y0 + self.ph)],
                         close=True, dxfattribs={"layer": "FRAME"})
        m.add_lwpolyline(rect(self.fb), close=True, dxfattribs={"layer": "FRAME"})
        m.add_lwpolyline(rect(self.tb), close=True, dxfattribs={"layer": "FRAME"})
        rows = [(0.00, 3, "图号", self.code), (0.36, 3, "图名", self.name),
                (0.72, 3, "比例", "1:%d" % self.denom), (0.72, 92, "图幅", self.paper),
                (1.08, 3, "设计", ""), (1.44, 3, "审核", ""), (1.80, 3, "日期", "")]
        for dy, dx, k, v in rows:
            y = self.tb[1] + dy * self.tbh / 2.2 + 3.2
            m.add_line((self.tb[0], y), (self.tb[2], y), dxfattribs={"layer": "FRAME"})
            self.text(self.tb[0] + dx, y, k, 2.6)
            if v:
                self.text(self.tb[0] + dx + 15, y, v, 3.2 if k == "图名" else 2.6)

    def text(self, x, y, s, h=2.5, layer="LABEL"):
        if h < MIN_TEXT_H:
            self._dw.append("字高 %.1fmm 低于下限 %.1fmm｜%s" % (h, MIN_TEXT_H, str(s)[:18]))
        return self.ms.add_text(str(s), dxfattribs={"layer": layer, "height": h, "style": "CN",
                                                    "insert": (x, y)})

    def pline(self, pts, layer="REDLINE", close=True, width=0.0):
        e = self.ms.add_lwpolyline(pts, dxfattribs={"layer": layer})
        if close:
            e.close(True)
        if width:
            e.dxf.const_width = width
        return e

    def circle(self, x, y, r, layer="PLANT"):
        return self.ms.add_circle((x, y), r, dxfattribs={"layer": layer})

    def fit(self, pts_m, box):
        """实地米坐标 → 图面 mm，等比塞进 box=(x1,y1,x2,y2)；返回（点, 系数 k）。
           k 必须接近 1，否则这张图的实际比例就不是图签写的那个分母。"""
        xs = [p[0] for p in pts_m]
        ys = [p[1] for p in pts_m]
        w = max(max(xs) - min(xs), 1e-6)
        h = max(max(ys) - min(ys), 1e-6)
        k = min((box[2] - box[0]) / (w * self.mm_per_m), (box[3] - box[1]) / (h * self.mm_per_m))
        ox = (box[0] + box[2]) / 2.0 - (min(xs) + max(xs)) / 2.0 * self.mm_per_m * k
        oy = (box[1] + box[3]) / 2.0 - (min(ys) + max(ys)) / 2.0 * self.mm_per_m * k
        out = [(x * self.mm_per_m * k + ox, y * self.mm_per_m * k + oy) for x, y in pts_m]
        if k < 0.995:
            self._dw.append("图面比声明的比例小｜缩放系数 k=%.3f，1:%d 是虚的——"
                             "要么换大幅面，要么把图签比例改成实际算出的 1:%d"
                             % (k, self.denom, int(round(self.denom / k))))
        return out, k

    def north(self, x, y, r=12.0):
        m = self.ms
        m.add_circle((x, y), r, dxfattribs={"layer": "FRAME"})
        m.add_lwpolyline([(x, y + r * 0.86), (x - r * 0.30, y), (x + r * 0.30, y)],
                         close=True, dxfattribs={"layer": "FRAME"})
        self.text(x - 1.7, y + r + 1.5, "N", 3.4)

    def scalebar(self, x, y, metres=(0, 5, 10, 20), unit=5.0):
        """直线比例尺：每格的图面长度由 mm_per_m 真算，不照抄别人的分划。"""
        m = self.ms
        total = max(metres) * self.mm_per_m
        m.add_line((x, y), (x + total, y), dxfattribs={"layer": "FRAME"})
        for i, v in enumerate(metres):
            px = x + v * self.mm_per_m
            m.add_line((px, y), (px, y + 3.0), dxfattribs={"layer": "FRAME"})
            self.text(px - 2.0, y + 5.4, str(int(v)), 2.4)
        for j, k in enumerate(range(int(metres[0]), int(max(metres)), int(unit))):
            if j % 2 == 0:
                a, b = x + k * self.mm_per_m, x + (k + unit) * self.mm_per_m
                h = m.add_hatch(dxfattribs={"layer": "FRAME"})
                h.paths.add_polyline_path([(a, y), (b, y), (b, y + 3.0), (a, y + 3.0)],
                                          is_closed=True)
                h.set_pattern_fill("SOLID", color=7)
        self.text(x, y - 6.0, "米（1:%d）" % self.denom, 2.4)

    def table(self, x, y, rows, widths, h=2.4, rh=5.0, layer="TABLE"):
        """rows[0] 当表头；widths 是每列 mm 宽，越框检查会看它画到哪儿。"""
        m = self.ms
        cy = y
        for row in rows:
            cx = x
            for ci, cell in enumerate(row):
                w = widths[ci]
                m.add_lwpolyline([(cx, cy), (cx + w, cy), (cx + w, cy - rh), (cx, cy - rh)],
                                 close=True, dxfattribs={"layer": layer})
                self.text(cx + 1.2, cy - rh / 2.0, cell, h, "LABEL")
                cx += w
            cy -= rh
        return cy

    # ── 检查 ────────────────────────────────────────────────────────────────
    def _pts(self, e):
        t = e.dxftype()
        try:
            if t in ("LWPOLYLINE", "POLYLINE"):
                return [(p[0], p[1]) for p in e.get_points("xy")]
            if t == "LINE":
                return [tuple(e.dxf.start), tuple(e.dxf.end)]
            if t == "CIRCLE":
                c, r = tuple(e.dxf.center), e.dxf.radius
                return [(c[0] - r, c[1] - r), (c[0] + r, c[1] + r)]
            if t == "TEXT":
                p = tuple(e.dxf.insert)
                # 文字宽度按 0.75×字高×字数估（估够用来抓明显越框；要精确用 ezdxf 的文本测量）
                return [p, (p[0] + len(e.dxf.text) * e.dxf.height * 0.75,
                             p[1] + e.dxf.height)]
        except Exception:
            return []
        return []

    def check(self, site_m=None):
        """→ (warns, errs)。这张图自己有没有画坏，跑一遍就知道。"""
        self.errs = []
        self.warn = list(self._dw)      # 画的时候攒下的（字高、缩放系数）
        f = self.fb
        if site_m:
            need = (site_m[0] * self.mm_per_m, site_m[1] * self.mm_per_m)
            avail = (f[2] - f[0] - self.tbw - 24, f[3] - f[1] - 34)
            if need[0] > avail[0] or need[1] > avail[1]:
                self.errs.append("比例装不下｜1:%d 下场地要 %.0f×%.0fmm，画区只有 %.0f×%.0fmm"
                                 "——换大幅面，别靠打印缩放" % (self.denom, need[0], need[1],
                                                              avail[0], avail[1]))
        for i, e in enumerate(self.ms):
            if i < self._tb_n:                 # 图框与图签自己不计
                continue
            for p in self._pts(e):
                if not (f[0] - 0.5 <= p[0] <= f[2] + 0.5 and f[1] - 0.5 <= p[1] <= f[3] + 0.5):
                    self.errs.append("越框｜%s (%.0f,%.0f) 出了内框（图幅 %s）"
                                     % (e.dxftype(), p[0], p[1], self.paper))
                    break
                if self.tb[0] < p[0] < self.tb[2] and self.tb[1] < p[1] < self.tb[3]:
                    self.errs.append("压图签｜%s (%.0f,%.0f) 落在标题栏里" % (e.dxftype(), p[0], p[1]))
                    break
        n = len(self.ms)
        if n > self.cap:
            self.errs.append("线数超预算｜本图 %d 个实体，预算 %d（版面会开始互相压字）" % (n, self.cap))
        elif n > self.cap * 0.9:
            self.warn.append("线数接近预算｜%d／%d" % (n, self.cap))
        txt = " ".join(e.dxf.text for e in self.ms.query("TEXT"))
        for k in TITLE_FIELDS:
            if k not in txt:
                self.errs.append("图签缺项｜没有「%s」" % k)
        for need in ("N", "米"):
            if need not in txt:
                self.warn.append("制图要素｜图面上找不到「%s」——指北针／直线比例尺是任务书点名的" % need)
        if "CN" not in self.doc.styles:
            self.errs.append("文字样式｜没有 CN 样式，AutoCAD 里汉字会变问号")
        return self.warn, self.errs

    def save(self, path):
        self.doc.saveas(str(path))
        return Path(path)


def _demo_sheet(bad=False):
    s = Sheet("PL-01", "种植设计总平面图", paper="A2", denom=500)
    rect = [(0, 0), (60, 0), (60, 45), (0, 45)]
    box = (s.fb[0] + 18, s.fb[1] + 44, s.tb[0] - 16, s.fb[3] - 16)
    pts, _k = s.fit(rect, box)
    s.pline(pts, "REDLINE")
    for i in range(2):
        cx, cy = box[0] + 30 + 40 * i, box[1] + 60
        s.circle(cx, cy, 3.4)
        s.text(cx + 4.2, cy, "ZM-0%d" % (i + 1), 2.2)
    s.north(s.tb[0] - 26, s.fb[3] - 26)
    s.scalebar(s.fb[0] + 20, s.fb[1] + 16, metres=(0, 5, 10, 20), unit=5)
    s.table(s.tb[0] - 96, s.fb[1] + s.tbh + 52,
            [["编号", "树种", "规格"], ["ZM-01", "香樟", "H6.0/Ø12"]], [24, 30, 40])
    if bad:
        s.circle(s.fb[2] + 34, s.fb[1] + 20, 6.0)          # 故意越框
        s.circle(s.tb[0] + 60, s.tb[1] + 20, 8.0)          # 故意压图签
        s.text(s.fb[0] + 22, s.fb[1] + 70, "这行字太小", 1.2)   # 故意低于字高下限
    return s, (60.0, 45.0)


def preview(dxf, png, dpi=170):
    """不装 AutoCAD 也能把 DXF 渲成 PNG——验收要的是"眼睛看过这张图"，不是"脚本说绿了"。"""
    import ezdxf
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    doc = ezdxf.readfile(str(dxf))
    fig = plt.figure(figsize=(11.7, 8.3), facecolor="#1b1b1b")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    Frontend(RenderContext(doc), MatplotlibBackend(ax)).draw_layout(doc.modelspace(), finalize=True)
    fig.savefig(str(png), dpi=dpi, facecolor="#1b1b1b")
    plt.close(fig)
    return Path(png)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--demo", metavar="DXF")
    ap.add_argument("--demo-bad", metavar="DXF")
    ap.add_argument("--preview", nargs=2, metavar=("DXF", "PNG"))
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if a.preview:
        print(preview(*a.preview))
        return 0
    which = a.demo_bad or a.demo
    if not which:
        print(__doc__)
        return 0
    s, site = _demo_sheet(bad=bool(a.demo_bad))
    w, e = s.check(site)
    s.save(which)
    print("%s：图幅 %s、1:%d（实地 1m＝图面 %.2fmm）、%d 个实体" % (
        which, s.paper, s.denom, s.mm_per_m, len(s.ms)))
    for x in w:
        print("  · " + x)
    for x in e:
        print("  ✗ " + x)
    print("检查：%d 条提醒／%d 条红字%s" % (len(w), len(e),
          "（坏示例就该响）" if a.demo_bad else "（示例应当全绿）"))
    return 1 if (e and not a.demo_bad) else 0


if __name__ == "__main__":
    sys.exit(main())
