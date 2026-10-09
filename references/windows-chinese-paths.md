# Windows + 中文路径：会咬人的九件事

同学大多在 Windows、路径里带中文（`C:\Users\...\Documents\Loomy Workspace\CAD-建模`）。
下面每一条都是真踩过、且报错信息不会指向真因的坑。

1. **Python 打印中文乱码／崩** — 控制台默认 cp936。每支脚本顶部
   `sys.stdout.reconfigure(encoding="utf-8")`（用 try 包住，重定向到文件时可能没有这个属性）。
   乱码只是难看，`UnicodeEncodeError` 会让脚本半路死掉。

2. **源码里的中文串带英文引号** — `"读图片素尺寸，算"印到纸上的分辨率""` 直接 SyntaxError，
   而且报错位置指向下一行。**中文里的引号一律用「」**。
   顺带：`"...\C1;..."` 这类字符串要写原始串 `r"..."`，否则 `SyntaxWarning: invalid escape '\C'`。

3. **PowerShell 脚本必须带 BOM** — 含中文的 `.ps1` 存成 UTF-8 无 BOM，PowerShell 5 按 ANSI
   读，中文变问号然后语法崩。存 UTF-8 **with BOM**。另外 Git Bash 里传给 `powershell -Command`
   的中文参数会被再编一次，改成写 `.ps1` 再 `-File` 调用。

4. **`EncodedCommand` 失效** — 别用 `powershell -EncodedCommand` 传中文路径，base64 那层
   编码在 cp936 控制台上会不一致。同上，走文件。

5. **`Start-Process` 参数要各自加引号** — `Arguments` 里带空格的路径不加引号会被拆开，
   报"找不到文件"，看着像权限问题。

6. **`rm -rf` 之外的删除** — Git Bash 里删中文文件名要用引号包住整条路径；
   `find -delete` 配合 `-name "*.png"` 比 shell 通配安全（通配不匹配时会把字面量传下去）。

7. **下载** — `curl` 在部分机器上会卡在证书吊销检查：加 `--ssl-no-revoke`。
   GitHub 直连慢/断：走 `gh-proxy.com` 镜像拼 URL，**下完必须校验**（zip 用
   `python -c "import zipfile;zipfile.ZipFile(f).testzip()"`），镜像会返回截断文件。

8. **AutoCAD 相关** — `accoreconsole` 不一定在 PATH 里，去
   `C:/Program Files/Autodesk/AutoCAD */accoreconsole.exe` 找。它**不加载 .pc3 打印样式表**，
   所以 PDF 别走它；它不主动关掉你已打开的 AutoCAD 进程，也不要为了跑脚本去杀进程
   （`acad.exe` 里有同学没存的图）。

9. **Word 占用** — 交付时如果 `方案书.docx` 正在 Word 里开着，`PermissionError` 会来自
   打包脚本而不是 Word。报错时先确认文件没被打开；不要写"重试三次"的循环掩盖它。

## 顺手一条跨平台的

路径一律 `pathlib.Path` 拼，别手写 `"\\"`；`rglob("*")` 前想清楚要不要递归整棵树
（工作目录里有个几百 GB 的归档目录时，全盘扫会把你卡死）。
## 验收门会被 Word 的锁文件顶红（第十件事）

只要有人**开着**方案书看（他自己查、或 Word 崩过一次留下残file），同目录就多出
`~$方案书.docx` 这种拥有者文件。它跟着 `*.docx` 一起被 glob 命中，于是
「详版 docx 不唯一」这条门当场红——**而文档没坏，坏的是发现文件的那行代码**。

所以任何"数交付里有几份 docx/pdf"的地方都要先滤掉 `~$`：

```python
docx = sorted(p for p in DOC.glob("*.docx") if not p.name.startswith("~$"))
```

本项目一次改了四处（打包装配、方案书文本端、插图落点、时效性），因为它们各自 glob 了一遍。
**别去删那个 `~$` 文件**——那是别人正开着的文档，删了可能丢他的编辑；滤掉就行。

## 用 COM 驱动 Word 数页数：别把自己锁死（第十一件事）

一页门、真排版复验都要真的开一次 Word。`DispatchEx("Word.Application")` 起的实例是**隐形**的，
脚本被 `timeout` 杀掉或中途抛错时它不会跟着死，于是：

- 它**锁着那个 docx**——下一次构建报 `PermissionError`，而 `~$` 拥有者文件可能早就没了
  （锁在进程句柄里，不在目录里）；
- 它**毒化后续所有调用**——新的 `DispatchEx` 会挂在 `Documents.Open` 上不动，看着像"Word 坏了"。

三条自保：

1. **一个进程里只开一次 Word**，把要量的几份文档在同一个 `app` 上循环开完再 `Quit()`。
   在同一个脚本里连着调两次"开 Word 量一页"，第二次很容易踩到没退干净的实例。
2. **超时要能留下现场**：`python -u`（stdout 被重定向时是块缓冲，进程被杀等于一行没打印，
   你会误以为是打开卡住而不是导出卡住），并且把每一步单独 print 出来。
3. **要清实例就先证明是自己的**：`Get-Process WINWORD | Select Id,StartTime,MainWindowTitle`——
   启动时间落在你自己那几次调用里、`MainWindowTitle` 为空的才是脚本留下的；
   有标题或时间更早的是用户开着的，**别动**。`taskkill` 不带 `/F` 对卡在 RPC 等待的实例无效，
   强杀后它可能仍挂在进程表里，这时 Word 自动化整体不可用，只能等或让用户重启会话——
   这条没有干净的服务端解法，别指望脚本自己修好。

顺带一条：`Documents.Open(path, ReadOnly=True)` 之后 `ExportAsFixedFormat` 失败时，
`python-docx` 那边写文件的报错和 Word 的报错长得像，先分清是**句柄被别的 Word 占着**
还是**导出目标目录不对**。
