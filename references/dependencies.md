# 开源依赖清单：装什么、干什么、坑在哪

一条命令装齐（Python 3.10+，Windows/macOS/Linux 都行）：

```bash
python -m pip install ezdxf matplotlib pillow numpy scipy python-docx pypdfium2 opencv-python pywin32
```

先跑 `python scripts/check_env.py` 核对，它会连中文字体和 accoreconsole 一起查。

| 库 | 版本实测 | 在这套流程里干什么 | 已知的坑 |
|---|---|---|---|
| **ezdxf** | 1.4.4 | 生成/读取 DXF（R2010）：图框、图层、LWPOLYLINE、HATCH、TEXT/MTEXT、表格。整套图纸体系靠它，**不装 AutoCAD 也能出图** | ① 必须 `ezdxf.new("R2010", setup=True)`，`setup=True` 才带 DASHED 等线型，否则虚线图层画出来是实线；② 不建中文文字样式（`doc.styles.add("CN", font="simhei.ttf")`），AutoCAD 里汉字全变问号；③ 读文本时 MTEXT 的 `dxf.text` 带 `{\C1;…}` 格式码，要用 `e.plain_text()`，不然判据正则匹配不到；④ 它的 matplotlib 渲图插件在 `ezdxf.addons.drawing.matplotlib`，1.x 里还在 |
| **matplotlib** | 3.10.0 | PNG 预览图、统计图、把一批图拼成核验接触表（一次看 20 张比翻 20 个文件快） | 中文要先 `matplotlib.use("Agg")` 再设 `rcParams["font.sans-serif"]=["Microsoft YaHei","SimHei"]`，并 `axes.unicode_minus=False`，否则负号是方块 |
| **Pillow (PIL)** | 12.3.0 | 读图片像素尺寸算「印到纸上的分辨率」；把网络来的 JPEG 重编码 | python-docx 有时解不动某些渐进式 JPEG（报 unknown image format），用 `Image.open(...).convert("RGB").save(..., quality=88)` 重存一份 baseline 就能插 |
| **numpy** | 2.1.3 | 1m 网格掩膜：面积、坡度、可种面、覆盖率全部在网格上现算，不靠目测 | 网格化时先 `np.floor` 到格心，边界那一圈的面积差就是你和同学数字不一样的来源 |
| **scipy** | 1.15.3 | `mannwhitneyu`／`spearmanr`：判读结论到底对应上没有，用检验说话 | 单侧/双侧要写清；p 值报三位小数，别报 "p<0.05" 蒙混 |
| **python-docx** | 1.2.0 | 方案书的生成与**读回**：章节标题、图注、插图位置、字数 | ① 读回时正文与表格要一起拉平（`doc.element.body.iterchildren()`），只读 `doc.paragraphs` 会漏掉表格里的文字；② 内联图的真实打印宽度在 `wp:extent` 的 EMU 里（÷914400＝英寸），别用版心宽度猜 |
| **pypdfium2** | 5.12.1 | 读 PDF 页数、`get_textpage().get_text_range()` 抽文本、`page.render(scale=2).to_pil()` 逐页渲图做目视复验 | 合订 PDF 的页数是唯一硬证据：图纸 19 张但 PDF 18 页，说明有一张没进去 |
| **opencv-python** | 5.0.0.93 | 手绘现状图／照片配准（相似变换）、点符号检测 | 只在需要把照片/手绘对到场地坐标上时才用；配准控制点要独立于你要验的东西，否则是自证 |
| **pywin32** | 312 | 驱动本机 AutoCAD（可选） | COM 经常连不上（实测过装好但不可达）；连不上就走 ezdxf，不影响交付 |

## 可选但有用的外部工具

- **accoreconsole.exe**（随 AutoCAD 装）：`accoreconsole /i 图.dxf /e 脚本.txt` 能校核并转 DWG/PDF。
  坑：它不加载 .pc3 打印样式表，PDF 出来线宽颜色和你屏幕上看到的不一样 —— 所以**出 PDF 走 ezdxf
  渲图**更可控，accoreconsole 只用来证「这个 DXF 真 AutoCAD 打得开」。
- **ffmpeg**：视频段合成与调色（作业要求做视频才需要）。
- **Firecrawl CLI**（`firecrawl search/scrape`）：检索建成案例与抓原图，见 `case-research.md`。
- **OpenStreetMap 的 landuse/natural/leisure 面**：判读结论的第二来源，见 `no-fabrication.md`。

## 版本要不要锁

不锁。上面这些库近几个大版本 API 都稳。真正要锁的是**随机种子**：管线里凡有随机（布点、
采样、撒点检验），把种子写死在数据层脚本顶部，这样同学之间、你和老师看到的数字才复现得上。
