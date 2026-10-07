# -*- coding: utf-8 -*-
"""skill 自检：五支脚本各跑一遍，缺陷必须逐条命中 evals/make_fixtures.py 里 PLANTED 的那些。

    python selfcheck.py
"""
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "evals"))
import make_fixtures                                     # noqa: E402
from make_fixtures import PLANTED                         # 埋的缺陷清单（唯一真源）

F = HERE / "evals" / "fixture"


def run(*cmd):
    p = subprocess.run([sys.executable] + list(cmd), cwd=str(HERE),
                       capture_output=True, shell=False)
    out = p.stdout.decode("utf-8", "replace") + p.stderr.decode("utf-8", "replace")
    return p.returncode, out


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    # 假项目是生成出来的（分发包里只带生成器，不带那堆纯色 PNG），所以每次先重建一遍——
    # 顺带保证「埋的缺陷清单」和 fixture 永远是同一版，不会出现改了 PLANTED 忘了重建。
    make_fixtures.main()
    bad = 0
    rc, o = run("scripts/check_env.py", "--json")
    d = json.loads(o)
    print("check_env    rc=%d 必需 %d 个（齐 %s）中文字体 %d 个 accoreconsole %s"
          % (rc, len(d["必需"]), d["可用"], len(d["中文字体"]["可用"]),
             "有" if d["accoreconsole"] else "无"))
    bad += 0 if d["可用"] else 1

    rc, o = run("scripts/audit_manual_coverage.py", "--manual", str(F / "指导手册.docx"),
                "--manifest", str(F / "对照表.csv"), "--root", str(F), "--json",
                str(HERE / "selfcheck_mc.json"))
    mc = json.loads((HERE / "selfcheck_mc.json").read_text(encoding="utf-8"))
    hits = [q for q in PLANTED if any(q.split("｜")[0] in x and q.split("｜")[-1][:10] in x
                                      for x in mc["硬伤"] + mc["提醒"])]
    print("手册图目     rc=%d 图目 %d 条 硬伤 %d 提醒 %d" % (rc, len(mc["手册图目"]),
                                                          len(mc["硬伤"]), len(mc["提醒"])))
    for q in mc["硬伤"]:
        print("   ✗ " + q)

    rc, o = run("scripts/audit_doc_figs.py", "--doc", str(F / "方案书.docx"),
                "--figdir", str(F / "图纸"), "--registry", str(F / "图纸目录.csv"),
                "--json", str(HERE / "selfcheck_df.json"))
    df = json.loads((HERE / "selfcheck_df.json").read_text(encoding="utf-8"))
    print("方案书插图   rc=%d 插图 %d 图注 %d 硬伤 %d 提醒 %d" % (rc, df["插图"], df["图注"],
                                                             len(df["硬伤"]), len(df["提醒"])))
    for q in df["硬伤"]:
        print("   ✗ " + q)

    allq = mc["硬伤"] + mc["提醒"] + df["硬伤"] + df["提醒"]

    def caught(p):
        """埋的缺陷按「｜／→」切成片段，只要有一条红字把片段全含住就算抓到。"""
        frags = [f.strip() for f in re.split(r"[｜→]", p) if len(f.strip()) >= 4]
        return any(all(f in q for f in frags) for q in allq)

    miss = [p for p in PLANTED if not caught(p)]
    print("埋的 %d 条缺陷被抓到 %d 条" % (len(PLANTED), len(PLANTED) - len(miss)))
    for m in miss:
        print("   ✗ 漏抓｜" + m)
    bad += len(miss)

    rc, o = run("scripts/gates_scaffold.py", "--demo")
    ok = "0 条红字" in o and "3/3 道门响过" in o
    print("三端点名     rc=%d %s" % (rc, "全绿＋3/3 道门响过" if ok else "输出不符预期"))
    bad += 0 if ok else 1

    rc, o = run("scripts/dxf_sheet_scaffold.py", "--demo", str(HERE / "selfcheck_good.dxf"))
    g = "0 条红字" in o
    rc2, o2 = run("scripts/dxf_sheet_scaffold.py", "--demo-bad", str(HERE / "selfcheck_bad.dxf"))
    b = "越框" in o2 and "压图签" in o2
    rc3, o3 = run("scripts/dxf_sheet_scaffold.py", "--preview",
                  str(HERE / "selfcheck_good.dxf"), str(HERE / "selfcheck.png"))
    pv = (HERE / "selfcheck.png").is_file()
    print("DXF 脚手架   干净图 rc=%d 全绿=%s ／ 坏图 rc=%d 会响=%s ／ 渲 PNG rc=%d 出图=%s"
          % (rc, g, rc2, b, rc3, pv))
    bad += (0 if g else 1) + (0 if b else 1) + (0 if pv else 1)

    for f in ("selfcheck_mc.json", "selfcheck_df.json", "selfcheck_good.dxf",
              "selfcheck_bad.dxf", "selfcheck.png"):
        (HERE / f).unlink(missing_ok=True)
    print("\n%s" % ("自检全过：五支脚本可用，埋的缺陷零漏抓。" if not bad
                   else "自检有 %d 处不对，别把 skill 发出去。" % bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
