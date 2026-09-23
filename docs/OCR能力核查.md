# OCR 能力核查（2026-09-23）

> 这份文档回答一个问题：**这台机器上到底有没有 OCR，插件该不该做 OCR？**
> 结论的证据是命令与原始输出，不是「应该有」或「大概不行」。

## 一、结论（先说结果）

| 问题 | 结论 |
|---|---|
| 本机有没有 OCR 引擎 | **有**：Windows 内置 `Windows.Media.Ocr`，已装识别语言**只有 `zh-Hans-CN`**（简体中文） |
| 它能不能真的读出字 | **能**：实测读出了图片里的英文与中文（原始输出见下） |
| 有没有命令行 OCR 工具 | **没有**：`tesseract` / `gocr` / `ocrad` / `ocrmypdf` / ImageMagick / Poppler 全部未安装；Python 3.14.7 在，但没有 `pytesseract`/`PIL`/`easyocr`/`paddleocr`/`fitz`/`cv2`/`pdf2image` |
| 插件能不能用它 | **不能（在零依赖前提下）**：Node 访问 WinRT 需要原生 addon；而且 OCR 只吃**位图**，插件没有 PDF 光栅化器 |
| 因此 | 插件**不提供 OCR 工具**，`lib/capabilities.js` 里 `pdf.ocr` 保持 `not_implemented`，并在 README 与任务清单里如实写明 |

一句话：**能力在本机存在，但不在插件能触及的运行时里**；扫描件要 OCR，目前得走外部手工路径（见第四节）。

## 二、证据：Windows 内置 OCR 真的能读

生成一张含中英文的图片，再用系统 OCR 引擎识别（探针脚本跑完即删，命令如下）：

```powershell
# 1) 生成图片：System.Drawing（GDI+）画三行字，含中文
#    注意：Windows PowerShell 5.1 会把「无 BOM 的 UTF-8 脚本」当 ANSI 读，
#    中文字面量会先烂掉（第一次实测就是这样，OCR 只是忠实读出了画错的字）。
#    所以中文字符串用码点拼：0x529E 0x516C 0x63D2 0x4EF6 0x80FD 0x529B 0x6838 0x67E5
# 2) 识别：StorageFile → BitmapDecoder → OcrEngine.RecognizeAsync
powershell -NoProfile -ExecutionPolicy Bypass -File probe-ocr-run.ps1
```

原始输出（关键行）：

```
IMAGE_WRITTEN …\ocr-sample.png (14087 bytes)
OCR_LANG zh-Hans-CN
OCR_LINE HELLO OCR 2026
OCR_LINE 办 公 插 件 能 力 核 查
OCR_LINE tesseract missing
```

- 英文行逐字符正确；
- 中文行内容正确（引擎会在汉字之间插空格，这是它的正常行为）；
- 识别语言取自用户配置：`TryCreateFromUserProfileLanguages()` → `zh-Hans-CN`。

识别语言清单（只装了这一个）：

```
PS_VERSION 5.1.26100.7627
WINRT_OCR_LANGS 1
WINRT_OCR_LANG zh-Hans-CN
WINRT_OCR_ENGINE created:zh-Hans-CN
```

## 三、证据：命令行工具与光栅化器

```
tesseract → 未安装      gocr → 未安装        ocrad → 未安装
ocrmypdf  → 未安装      magick → 未安装      pdftoppm → 未安装
python → C:\Program Files\PyManager\python.exe（3.14.7）
PYTHON_OCR_MODULES []        # pytesseract / PIL / easyocr / paddleocr / fitz / cv2 / pdf2image 都没有
宿主依赖里与图像有关的只有：sharp   # 只能解码图片，不带 PDF 光栅化（libvips 预编译版不含 poppler/gs）
```

**没有光栅化器是第二个硬障碍**：OCR 引擎只接受位图，而扫描件通常是 PDF。
把 PDF 变成位图需要 Poppler / Ghostscript / PDFium 之类的渲染器，本机都没有。
宿主自带的 LibreOffice 引擎原则上能「打开 PDF（Draw 导入）再导出 PNG」，
但那条路要经过宿主的引擎接口，属于**另一个组件**的能力，插件不假装拥有它；
真要做得单独验证，不在这一版里。

## 四、今天要做 OCR 的话，手工路径是什么

1. 把扫描页导出成 PNG/JPEG（用 Acrobat、WPS PDF、或任何能另存为图片的阅读器）；
2. 用系统 OCR 识别（本节第二段的脚本就是最小实现；语言取 `Windows.Media.Ocr` 里已安装的识别语言）；
3. 把识别结果作为**纯文本**交给插件的文档工具（插件擅长的是结构化读写，不是识别）。

换句话说：插件这一版的能力边界是「**已经有文字的** Office/PDF 文档」，
而不是「图片里的文字」。

## 五、复现命令

```powershell
# 有 OCR 引擎吗、装了哪些语言
powershell -NoProfile -ExecutionPolicy Bypass -File probe-ocr.ps1

# 真的读一张图（含中英文）
powershell -NoProfile -ExecutionPolicy Bypass -File probe-ocr-run.ps1

# 命令行 OCR 工具与 Python 包
foreach ($n in @('tesseract','gocr','ocrad','ocrmypdf','magick','pdftoppm')) { Get-Command $n -ErrorAction SilentlyContinue }
python -c "import importlib.util as u; print([m for m in ['pytesseract','PIL','easyocr','paddleocr','fitz','cv2','pdf2image'] if u.find_spec(m)])"
```

> 探针脚本属于一次性核查工具，已按工作区清洁约定删除；需要复现时按上面两段命令重写即可
> （第二段的核心是 `StorageFile → BitmapDecoder → OcrEngine.RecognizeAsync` 这条链）。
