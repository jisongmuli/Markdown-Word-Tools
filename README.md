# Markdown Word Tools

本地将 UTF-8 Markdown 转为 Word 文档。支持标题、段落、粗体、行内代码、链接、列表、代码块、表格、引用和分隔线；包含中文字体设置、A4 页面、代码底纹和页码。

```sh
python -m pip install -r requirements.txt
python md2docx.py input.md output.docx
```

需要 Python 3.10+。目标机器若没有微软雅黑或 Consolas，Word 会选择替代字体；可修改脚本顶部的字体常量。复杂 HTML、图片与完整 CommonMark 语法不在当前实现范围内。

不包含任何私人记录或原始文档。[MIT](LICENSE)
