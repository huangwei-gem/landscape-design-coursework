# -*- coding: utf-8 -*-
"""造一个故意带病的最小假项目（11 处缺陷，用来验 scripts/ 里那两支审计脚本真的抓得住缺陷。

用途有两个：一是本地自检（跑完 make_fixtures.py 再跑两支审计，红字必须逐条命中下面
PLANTED 里列的那几条，抓不到就是脚本自己瞎了）；二是给 evals 当输入，使评分可以判分。
所有图都是纯色 PNG，几十 KB，可以随 skill 分发。

    python make_fixtures.py            # 写到 fixture/
"""
import csv
import re
import sys
from pathlib import Path

OUT = Path(__file__).resolve().parent / "fixture"

# 故意埋的缺陷（审计脚本必须抓到这些，且不该抓出别的硬伤）
PLANTED = [
    "手册要求｜图5 在手册里有图目，对照表里没有承载文件",
    "承载文件不存在｜图2 → 图纸/PL-02_地形分析图.png",
    "手册图名对不上｜对照表写「苗木统计表」",
    "图纸未入册｜PL-06《苗木表》在方案书里没有以图注出现",
    "插图落点｜PL-04 贴在「一、总体设计／1.2 地形与竖向」，这一节正文从头到尾没提过它",
    "印刷分辨率不够｜PL-05",
    "正文引用了 PL-06，但全篇没有以它为图号的图注",
    "出过的图没进报告｜PL-07_未用到的图",
    "对照表多列｜图9 手册里根本没有这个图号",
    "图文不符｜图注写 PL-04，实际嵌进来的文件是 PL-03",
    "图纸目录与磁盘不符｜PL-06《苗木表》登记了",
]


def png(path, w, h, rgb):
    from PIL import Image
    Image.new("RGB", (w, h), rgb).save(path)


def manual(path):
    """指导手册：正文一张 图1（配图），文末一张图版表，图注行与图片行交替（真手册是图在上、号在下）。"""
    import docx
    d = docx.Document()
    d.add_heading("2026 植物景观规划与设计 指导手册", 0)
    d.add_paragraph("四、场地范围")
    d.add_picture(str(OUT / "手册原图/手册_图1_场地位置与范围示意图.png"), width=docx.shared.Cm(6))
    d.add_paragraph("图1 场地位置与范围示意图")
    t = d.add_table(rows=0, cols=1)

    def row(txt=None, img=None):
        r = t.add_row().cells[0]
        if img:
            p = r.paragraphs[0]
            run = p.add_run()
            run.add_picture(str(img), width=docx.shared.Cm(4))
        else:
            r.text = txt or ""
    row(img=OUT / "手册原图/手册_图2.png"); row("图2 场地地形分析图")
    row(img=OUT / "手册原图/手册_图4-1.png"); row("图4-1 苗木表")
    row(img=OUT / "手册原图/手册_图4-2.png"); row("图4-2 苗木形态参考图")
    row(img=OUT / "手册原图/手册_图5.png"); row("图5 种植设计总平面图")
    row("注：以上图例仅作参考，图6 为评分表")     # 行首不是图N，不该被当成图目
    d.save(str(path))


def book(path):
    """方案书：三节正文 + 四张图，故意让 PL-04 贴错节、PL-05 印得太小、PL-06 没贴。"""
    import docx
    d = docx.Document()
    d.add_heading("樟荫书院种植设计方案书", 0)
    d.add_paragraph("一、总体设计")
    d.add_paragraph("1.1 设计立意")
    d.add_paragraph("立意讲环与区，PL-01 是总平面，PL-03 是结构图，两张都在本节。")
    d.add_picture(str(OUT / "图纸/PL-01_种植设计总平面图.png"), width=docx.shared.Cm(15))
    d.add_paragraph("图1　PL-01 种植设计总平面图（1:500）")
    d.add_picture(str(OUT / "图纸/PL-03_植物景观结构图.png"), width=docx.shared.Cm(15))
    d.add_paragraph("图2　PL-03 植物景观结构图")
    d.add_paragraph("1.2 地形与竖向")
    d.add_paragraph("竖向按 PL-02 的等高展开，土方形平衡见 1.2 表格。")
    d.add_picture(str(OUT / "图纸/PL-02_地形竖向分析图.png"), width=docx.shared.Cm(15))
    d.add_paragraph("图3　PL-02 地形竖向分析图")
    # 故意埋两个缺陷：① 落点错（说它的是 1.3）；② 贴错文件（图注写 PL-04，嵌的却是 PL-03 的图）
    d.add_picture(str(OUT / "图纸/PL-03_植物景观结构图.png"), width=docx.shared.Cm(15))
    d.add_paragraph("图4　PL-04 林缘线放线图（放样数据见 1.3）")
    d.add_paragraph("1.3 林缘线与放线")
    d.add_paragraph("林缘线逐圈见 PL-04，放线坐标见 PL-04 附表。")
    d.add_paragraph("二、苗木")
    d.add_paragraph("2.1 苗木规格")
    d.add_paragraph("规格按 PL-05 的形态谱给，逐株编号另见 PL-06 的坐标表。")
    d.add_picture(str(OUT / "图纸/PL-05_树种形态谱.png"), width=docx.shared.Cm(3.2))
    d.add_paragraph("图5　PL-05 树种形态谱")                       # ← 3.2cm 宽、600px：约 480dpi 够；用 240px 的图才不够
    d.save(str(path))


def registry(path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["图号", "图名"])
        for c, n in [("PL-01", "种植设计总平面图"), ("PL-02", "地形竖向分析图"),
                     ("PL-03", "植物景观结构图"), ("PL-04", "林缘线放线图"),
                     ("PL-05", "树种形态谱"), ("PL-06", "苗木表")]:
            w.writerow([c, n])


def coverage(path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["手册图号", "手册图名", "承载图纸", "承载文件"])
        w.writerow(["图1", "场地位置与范围示意图", "PL-01", "图纸/PL-01_种植设计总平面图.png"])
        w.writerow(["图2", "地形分析图", "PL-02", "图纸/PL-02_地形分析图.png"])        # 故意指向不存在的文件名
        w.writerow(["图4-1", "苗木统计表", "PL-06", "图纸/PL-05_树种形态谱.png"])  # 手册里没这名字
        w.writerow(["图4-2", "苗木形态参考图", "PL-05", "图纸/PL-05_树种形态谱.png"])
        w.writerow(["图9", "不存在的图目", "PL-07", "图纸/PL-01_种植设计总平面图.png"])  # 手册没有→提醒
        # 图5 故意不写——手册要求了，对照表漏了


def sheets(path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        f.write(
            "[\n"
            " {\"name\": \"保留树株数\", \"pats\": {\"图纸\": \"(?P<n>\\\\d+)株保留\", "
            "\"正文\": \"保留(?P<n>\\\\d+)株\", \"读我\": \"保留\\\\s*(\\\\d+)株\"}, \"exp\": {\"n\": 24}},\n"
            " {\"name\": \"常绿比\", \"pats\": {\"图纸\": \"常绿(?P<p>\\\\d+)%\", \"正文\": \"(?P<p>\\\\d+)%常绿\"}, "
            "\"exp\": {\"p\": 70}}\n"
            "]\n")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    (OUT / "图纸").mkdir(parents=True, exist_ok=True)
    (OUT / "手册原图").mkdir(parents=True, exist_ok=True)
    manual_pngs = {"手册_图1_场地位置与范围示意图.png": (900, 620),
                   "手册_图2.png": (900, 620), "手册_图4-1.png": (800, 800),
                   "手册_图4-2.png": (800, 800), "手册_图5.png": (900, 620)}
    for i, (n, (w, h)) in enumerate(manual_pngs.items()):
        png(OUT / "手册原图" / n, w, h, (60 + 25 * i, 90, 120))
    sheets_plan = {"PL-01_种植设计总平面图.png": (1600, 1100),
                   "PL-02_地形竖向分析图.png": (1600, 1100),
                   "PL-03_植物景观结构图.png": (1600, 1100),
                   "PL-04_林缘线放线图.png": (1600, 1100),
                   "PL-05_树种形态谱.png": (120, 120),      # 故意：印到 3.2cm 只有约 95dpi
                   "PL-07_未用到的图.png": (1200, 800)}     # 故意：磁盘上有、报告里没用
    for i, (n, (w, h)) in enumerate(sheets_plan.items()):
        png(OUT / "图纸" / n, w, h, (200 - 20 * i, 120, 70 + 15 * i))
    manual(OUT / "指导手册.docx")
    book(OUT / "方案书.docx")
    registry(OUT / "图纸目录.csv")
    coverage(OUT / "对照表.csv")
    sheets(OUT / "gates.json")
    print("fixture 写在 %s" % OUT)
    print("埋的缺陷 %d 条：" % len(PLANTED))
    for p in PLANTED:
        print("  · " + p)


if __name__ == "__main__":
    main()
