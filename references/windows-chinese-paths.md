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
