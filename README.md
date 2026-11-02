# Word 页眉页脚清理工具

运行 `WordHeaderFooterCleaner.exe`，选择单个文件或文件夹，再选择处理选项。界面处理后会覆盖原文件。

- 支持 `.doc`、`.docx`、`.docm`、`.dotx`、`.dotm`，可递归处理子文件夹。
- 可移除或修改页眉、页脚，支持对齐方式及 `{PAGE}` / `{NUMPAGES}` 自动页码。
- **删除最后一页图片**：按桌面排版引擎重新分页，删除最后一页正文的嵌入式、浮动图片，也处理表格、组合和文本框内的图片。保留正文文字、普通形状以及页眉页脚。空白末页不会回退删除前一页图片。删除图片可能使文档重新排版。

`.doc` 的全部操作，以及任意格式的末页图片删除，需要 Windows 本机安装 Microsoft Word，或支持 COM 自动化的 WPS 文字。优先使用 Word，无法启动时尝试 WPS；已在 Microsoft Word 上验证，WPS 的具体版本需支持相应自动化接口。`.doc` 保存后仍为原格式。原有新版 Word 页眉页脚处理不依赖 Office。

新功能先处理临时副本，保存、关闭成功后才替换原文件；无末页图片或处理失败时不改原文件，原因显示在日志中。

## 源码运行

```powershell
python -m pip install -r requirements.txt
python word_header_footer_gui.py
```

命令行删除末页图片并输出到另一个目录：

```powershell
python remove_word_headers_footers.py "D:\文档" --recursive --mode last_page_images --output-dir "D:\处理结果"
```

不传 `--output-dir` 时覆盖原文件；末页图片模式不能同时传入页眉或页脚内容。

## 验证与打包

```powershell
python -m pip install -r requirements-dev.txt
python -B -m unittest discover -v
$env:WORD_INTEGRATION_TESTS = '1'  # 本机安装 Microsoft Word 时运行真实文档测试
python -B -m unittest discover -v
python -m PyInstaller --noconfirm WordHeaderFooterCleaner.spec
Copy-Item dist/WordHeaderFooterCleaner.exe WordHeaderFooterCleaner.exe
python -m PyInstaller --noconfirm WordHeaderFooterCleanerSetup.spec
Copy-Item dist/WordHeaderFooterCleanerSetup.exe WordHeaderFooterCleanerSetup.exe
```

分页判断采用不受分节页码影响的物理页码；浮动图片按锚点所在页判断，参考 [Word 页码信息](https://learn.microsoft.com/en-us/office/vba/api/word.wdinformation)及 [Shape.Anchor](https://learn.microsoft.com/en-us/office/vba/api/word.shape.anchor)。
