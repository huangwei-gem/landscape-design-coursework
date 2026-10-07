# -*- coding: utf-8 -*-
"""环境体检：一次查清这套流程要的开源库、中文字体、可选的 AutoCAD 命令行。

为什么先跑它：后面几支脚本各只用到其中一两个库，"跑到第三步发现少装一个库"最容易
被误判成脚本坏了。字体和 accoreconsole 也一起查——中文字体缺失，DXF 渲成 PNG 时
汉字全是方块；accoreconsole 不在 PATH，出 DWG/PDF 只能走 ezdxf 那条路（也够用）。

用法：
    python check_env.py              # 人读的报告
    python check_env.py --json       # 机器读（给 CI／给别的脚本判断）
    python check_env.py --install    # 只打印那条 pip 命令
"""
import argparse
import importlib
import importlib.metadata as md
import json
import sys

# (导入名, pip 名, 这一支在本流程里干什么)
DEPS = [
    ("ezdxf", "ezdxf",
     "读写 DXF（R2010）：图框、图层、文字、填充、表格；不装 AutoCAD 也能出图出 PDF"),
    ("matplotlib", "matplotlib",
     "PNG 预览图、统计图、逐张目视复验用的核验接触表"),
    ("PIL", "pillow",
     "读图片素尺寸，算「印到纸上的分辨率」；python-docx 认不了的 JPEG 重编码"),
    ("numpy", "numpy",
     "1m 网格面域／掩膜，面积与配比的现算"),
    ("scipy", "scipy",
     "统计检验（Mann-Whitney、Spearman）：判读结论到底有没有对应上，用数说话"),
    ("docx", "python-docx",
     "Word 方案书的生成与读回：章节标题、图注、插图位置、字数"),
    ("pypdfium2", "pypdfium2",
     "读 PDF 页数与逐页渲图，验「合订本到底几页、图有没有被裁一半」"),
]
OPT = [
    ("cv2", "opencv-python", "手绘现状图／照片配准、符号检测（要配准才需要）"),
    ("win32com.client", "pywin32", "驱动本机 AutoCAD（可选，没有也能交付）"),
]
PIP = ("ezdxf matplotlib pillow numpy scipy python-docx pypdfium2 "
       "opencv-python pywin32")
CJK = ["Microsoft YaHei", "SimHei", "SimSun", "Noto Sans CJK SC", "Source Han Sans SC"]


def _probe(mod, pip):
    try:
        importlib.import_module(mod)
    except Exception as e:
        return {"库": pip, "导入名": mod, "状态": "缺", "版本": "", "错误": "%s: %s" % (type(e).__name__, e)}
    try:
        v = md.version(pip)
    except Exception:
        v = "?"
    return {"库": pip, "导入名": mod, "状态": "有", "版本": v, "错误": ""}


def fonts():
    """matplotlib 能不能画中文：查得到中文字体就返回那个名字。"""
    try:
        from matplotlib import font_manager
    except Exception as e:
        return {"可用": [], "查询失败": str(e)}
    ok = []
    for name in CJK:
        try:
            f = font_manager.findfont(font_manager.FontProperties(family=name),
                                      fallback_to_default=False)
            ok.append("%s→%s" % (name, f.split("/")[-1]))
        except Exception:
            continue
    return {"可用": ok, "查询失败": ""}


def accoreconsole():
    """出 DWG/PDF 的可选路径：本机装了 AutoCAD 就有 accoreconsole.exe。"""
    import glob
    import shutil
    p = shutil.which("accoreconsole")
    if p:
        return p
    for pat in (r"C:/Program Files/Autodesk/AutoCAD */accoreconsole.exe",
                r"C:/Program Files (x86)/Autodesk/AutoCAD */accoreconsole.exe"):
        g = sorted(glob.glob(pat))
        if g:
            return g[-1].replace("\\", "/")
    return ""


def collect():
    out = {"python": sys.version.split()[0],
           "python够新": sys.version_info[:2] >= (3, 10),
           "必需": [_probe(m, p) for m, p, _d in DEPS],
           "可选": [_probe(m, p) for m, p, _d in OPT],
           "用途": {m: d for m, _p, d in DEPS + OPT},
           "中文字体": fonts(),
           "accoreconsole": accoreconsole()}
    out["可用"] = all(r["状态"] == "有" for r in out["必需"]) and out["python够新"]
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--install", action="store_true", help="只打印 pip 安装命令")
    a = ap.parse_args()
    if a.install:
        print("python -m pip install %s" % PIP)
        return 0
    r = collect()
    if a.json:
        sys.stdout.reconfigure(encoding="utf-8")
        print(json.dumps(r, ensure_ascii=False, indent=1))
        return 0 if r["可用"] else 1

    w = sys.stdout
    try:
        w.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("Python %s%s" % (r["python"], "" if r["python够新"] else "  ← 太旧，要 3.10+"))
    print("\n必需库")
    for x in r["必需"]:
        print("  %-13s %-4s %-8s %s" % (x["库"], x["状态"], x["版本"],
                                        x["错误"] or r["用途"][x["导入名"]]))
    print("\n可选库")
    for x in r["可选"]:
        print("  %-13s %-4s %-8s %s" % (x["库"], x["状态"], x["版本"],
                                        x["错误"] or r["用途"][x["导入名"]]))
    print("\nmatplotlib 中文字体：%s" % ("、".join(r["中文字体"]["可用"]) or
                                       "一个都没找到——预览图上的汉字会变方块"))
    print("accoreconsole：%s" % (r["accoreconsole"] or "没有（不装 AutoCAD 一样能出图，见 references/dependencies.md）"))
    print("\n%s" % ("结论：环境齐了，可以开跑。" if r["可用"] else
                   "结论：还缺东西，先装——\n  python -m pip install %s" % PIP))
    return 0 if r["可用"] else 1


if __name__ == "__main__":
    sys.exit(main())
