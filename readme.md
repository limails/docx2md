# docx2md — Word 转 Markdown 工具包

本目录用于把 Word 文档（`.doc` / `.docx`）转换成高质量 Markdown，
保证图片不丢、链接不断、文件可移植。

## 文件清单

| 文件 | 说明 |
|---|---|
| **`docx2md.py`** | 主程序。零第三方 Python 依赖，用法见[附录](#附录docx2mdpy) |
| **`SPEC_docx2md.md`** | 规格说明书。可据此重写出功能等价的程序 |
| `readme.md` | 本文。pandoc 用法、实测踩坑、跨机器部署 |
| `pandoc_ref\reference.docx` | pandoc 默认样式模板，用于自定义导出 Word 的字体字号 |

## 快速开始

```powershell
python docx2md.py 论文.docx
python docx2md.py D:\docs\          # 递归处理整个目录
```

产物与源文件同目录：`<原名>.md` + `<原名>.media\`（图片）。

首次使用需先装 pandoc，见[安装与卸载](#安装与卸载)。

---

# pandoc 使用说明

> 本文所有结论均在本机（Windows / pandoc 3.11）针对实际文件 `Application of the MEGASYSTEM-C cementless prosthesis system...docx`（学术论文，约 51 KB 正文、11 张内嵌图、1 个跨页表格）实测得出，非转述官方文档。

## 目录

- [安装与卸载](#安装与卸载)
- [换一台机器用（安装位置可能不同）](#换一台机器用安装位置可能不同)
- [一、Word → Markdown](#一word--markdown)
- [二、输出格式怎么选](#二输出格式怎么选)
- [三、Markdown → Word / HTML](#三markdown--word--html)
- [四、自定义 Word 样式](#四自定义-word-样式)
- [五、批量转换](#五批量转换)
- [六、常用选项速查](#六常用选项速查)
- [七、实测踩到的坑](#七实测踩到的坑)
- [八、pandoc 做不到的事](#八pandoc-做不到的事)
- [九、本项目推荐流程](#九本项目推荐流程)
- [附录：docx2md.py](#附录docx2mdpy)

---

## 安装与卸载

### 安装

```
C:\Users\lysjc\AppData\Local\Pandoc\pandoc.exe
```

```powershell
winget install --id JohnMacFarlane.Pandoc -e
```

版本 3.11，占用约 223 MB。安装后已加入用户 PATH，新终端可直接调用 `pandoc`。

> ⚠️ 装完必须**新开一个终端窗口**。已经开着的窗口（包括 IDE 内置终端）
> 进程内的 PATH 还是安装前的旧快照，直接跑会报「找不到 pandoc」。

### 换一台机器用（安装位置可能不同）

`docx2md.py` 不假定 pandoc 装在哪，按下面顺序找，命中即用：

| 优先级 | 途径 | 用法 |
|---|---|---|
| 1 | 命令行指定 | `python docx2md.py 论文.docx --pandoc D:\tools\pandoc.exe` |
| 2 | 环境变量 | `[Environment]::SetEnvironmentVariable('PANDOC_EXE','D:\tools\pandoc.exe','User')` |
| 3 | PATH | 装完新开终端才生效 |
| 4 | 注册表 | 兜住「装了但当前进程 PATH 是旧的」 |
| 5 | 常见安装位置 | winget / choco / scoop / MSI 共 8 处 |
| 6 | 有限广搜 | 上述目录树下递归找 `pandoc.exe` |

覆盖的安装位置：

```
%LOCALAPPDATA%\Pandoc\pandoc.exe                    winget per-user
%LOCALAPPDATA%\Programs\Pandoc\pandoc.exe
%ProgramFiles%\Pandoc\pandoc.exe                    机器级 / MSI
%ProgramFiles(x86)%\Pandoc\pandoc.exe
%USERPROFILE%\scoop\apps\pandoc\current\pandoc.exe  scoop 用户级
%ProgramData%\scoop\apps\pandoc\current\pandoc.exe  scoop 全局
%ProgramData%\chocolatey\bin\pandoc.exe             choco
%ProgramData%\chocolatey\lib\pandoc\tools\pandoc.exe
```

程序启动时会打印找到的路径和版本，便于确认定位是否正确：

```
pandoc : C:\Users\lysjc\AppData\Local\Pandoc\pandoc.exe
版本   : pandoc 3.11
```

全部未命中时会列出所有尝试过的路径，并给出三种解决办法。

> `PANDOC_EXE` 指向目录也行，程序会自动补 `pandoc.exe`。
> 显式指定的路径无效时**直接报错**，不会静默回退 —— 避免以为指定生效了。

### 其他

PowerShell 里定义个别名省事：

```powershell
Set-Alias pd "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
```

### 卸载

```powershell
winget uninstall --id JohnMacFarlane.Pandoc
```

winget 会同时移除安装目录和用户 PATH 里的条目，不需要手动清理。

手动卸载（winget 不可用时）：

```powershell
# 1. 删程序目录
Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Pandoc"

# 2. 清理用户 PATH 里的条目
$p = [Environment]::GetEnvironmentVariable('Path','User')
$p = ($p -split ';' | Where-Object { $_ -notmatch 'Pandoc' }) -join ';'
[Environment]::SetEnvironmentVariable('Path', $p, 'User')
```

### 卸载前需要知道

| 影响 | 说明 |
|---|---|
| **`docx2md.py` 会失效** | 它依赖 pandoc 完成全部转换，卸载后运行会报「找不到 pandoc」并提示安装命令。只是暂时不用的话建议先别删 |
| **已转换的文件不受影响** | 产出的 `.md` 和 `.media\` 里的图片都是独立文件，不依赖 pandoc 继续存在 |
| **本文档会变过时** | 所有命令与参数说明都基于 pandoc。重装后无需修改本文 |
| **`pandoc_ref\reference.docx` 可重生成** | 它是 pandoc 导出的默认模板，删除无影响，重建命令见[第四章](#四自定义-word-样式) |

重装后即完全恢复，`SPEC_docx2md.md` 第 7 章记录了依赖清单。

---

## 一、Word → Markdown

最常用的操作：

```powershell
pandoc "论文.docx" -t gfm --wrap=none --extract-media=media -o 论文.md
```

### 四个参数缺一不可

| 参数 | 不加会怎样（实测） |
|---|---|
| `-t gfm` | 用 `markdown` 或 `commonmark` 时**表格降级为原始 HTML**，管道表消失 |
| `--wrap=none` | 默认每 72 字符硬折行，长段落被切碎，diff 和全文检索体验极差 |
| `--extract-media=media` | **不生成图片文件**，只写出指向它们的引用 → 全部断链 |
| `-o` | 不指定则直接输出到 stdout |

### 关于 `--extract-media`

这是最容易踩的坑。不加这个参数时，pandoc 会输出这样的内容：

```html
<img src="media/image1.jpeg" style="width:2.28678in;height:3.04747in" />
```

引用路径有了，但 **pandoc 不会创建 `media` 目录，也不会写入任何图片文件**。
实测运行后 `media` 目录压根不存在 → Markdown 里 11 个图片引用全部指向不存在的文件。

加上 `--extract-media=DIR` 后，pandoc 才会真正把图片写进 `DIR`，
并自动把引用路径调整为相对于输出文件的相对路径。

---

## 二、输出格式怎么选

`-t` / `--to` 决定输出格式。对同一份 docx 的实测结果：

| 格式 | 体积 | 表格输出 | 说明 |
|---|---|---|---|
| `gfm` | 50786 B | Markdown 管道表 | **默认首选**，GitHub / Obsidian / 通用编辑 |
| `commonmark_x` | 50774 B | Markdown 管道表 | 需要更严格 CommonMark 兼容时用 |
| `commonmark` | 53255 B | 降级为原始 HTML `<table>` | 不推荐 |
| `markdown` | 50865 B | 降级为原始 HTML `<table>` | 不推荐 |

统计方式：统计输出中 `|` 的数量。`gfm` 有 115 个，`commonmark` 为 0 个。

`commonmark` / `markdown` 下表格会变成这样：

```html
<thead>
<tr>
<th style="text-align: left;">Characteristic</th>
<th style="text-align: left;">No. (%) of Patients*</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;">Osteosarcoma</td>
```

内容没丢，但不再是 Markdown 表格，无法用 Markdown 表格语法编辑。

---

## 三、Markdown → Word / HTML

```powershell
# 转 Word 并生成目录
pandoc 论文.md -o 论文.docx --toc --toc-depth=3

# 生成独立完整 HTML（含 <style>，可直接浏览器打开）
pandoc 论文.md -s -o 论文.html

# 只输出正文片段（嵌入现有页面用）
pandoc 论文.md -t html -o 片段.html
```

- `-s` / `--standalone`：输出带完整 HTML 骨架和默认样式的独立页面
- `--toc` / `--toc-depth=N`：给 docx 生成目录，实测正常

### ⚠️ md → pdf 会失败

pandoc 本身**不包含任何排版引擎**。生成 PDF 必须依赖外部引擎之一：

```
pdflatex / lualatex / xelatex / latexmk / tectonic / wkhtmltopdf
weasyprint / prince / pagedjs-cli / context / pdfroff
```

本机目前这些**都没有安装**，所以 `pandoc x.md -o x.pdf` 会直接报错。
安装 pandoc 不会附带它们，需要单独安装并配置。

如果确实需要 md → pdf，可选：
- **Typst**（`winget install Typst.Typst`）：安装最简单，单个二进制，无需 TeX 发行版
- **Tectonic**（`cargo install tectonic` 或下载）：自动下载所需宏包，比完整 TeX Live 轻量
- **MiKTeX / TeX Live**：体积大，但兼容性最好

---

## 四、自定义 Word 样式

导出 docx 时想改字体、字号、标题样式，**不要手改 OOXML**，用 reference.docx：

```powershell
# 1. 导出默认模板
pandoc -o pandoc_ref/reference.docx --print-default-data-file reference.docx

# 2. 用 Word 打开，改好样式（修改「标题 1」「正文」等样式即可），另存回同一路径

# 3. 转换时引用
pandoc 论文.md --reference-doc=pandoc_ref/reference.docx -o 新论文.docx
```

模板已备好：`F:\work\rtest\pandoc_ref\reference.docx`

---

## 五、批量转换

```powershell
Get-ChildItem *.docx | ForEach-Object {
  pandoc $_.FullName -t gfm --wrap=none --extract-media=media -o "out\$($_.BaseName).md"
}
```

含子目录的递归版本：

```powershell
Get-ChildItem *.docx -Recurse | ForEach-Object {
  $dest = "out\$($_.Directory.Name)\$($_.BaseName).md"
  New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
  pandoc $_.FullName -t gfm --wrap=none --extract-media=media -o $dest
}
```

脚本里建议加 `--fail-if-warnings`，让异常情况显式失败而不是静默产出坏文件。

---

## 六、常用选项速查

| 选项 | 作用 |
|---|---|
| `-f FORMAT` / `--from` | 强制指定输入格式（默认按扩展名推断） |
| `-t FORMAT` / `--to` | 指定输出格式 |
| `-o FILE` | 输出到文件（不指定则输出到 stdout） |
| `-s` / `--standalone` | 输出带完整骨架的独立文档 |
| `--wrap=none\|auto\|COLUMN` | 折行策略，默认 72 列 |
| `--extract-media=DIR` | 导出 docx 内嵌媒体到指定目录 |
| `--toc` / `--toc-depth=N` | 生成目录及目录深度 |
| `--number-sections` | 章节标题自动编号 |
| `--number-offset=N` | 起始编号偏移 |
| `--reference-doc=FILE` | 指定 docx 样式模板 |
| `--track-changes=accept\|reject\|all` | Word 修订痕迹处理，默认 `accept`（丢弃修订） |
| `--fail-if-warnings` | 出现警告即报错，适合脚本 |
| `-p FILE` | md→PDF 时指定页眉/页脚（如 `header.tex`） |
| `--verbose` | 输出详细日志 |

### 常用格式名

```
输入:  docx  markdown  gfm  html  latex  rst  org  textile  t2t  json
输出:  docx  markdown  gfm  html  pdf(需引擎)  latex  rst  epub  odt
```

查看完整列表：

```powershell
pandoc --list-input-formats
pandoc --list-output-formats
```

---

## 七、实测踩到的坑

### 坑 1：图片路径不可移植（最严重）

pandoc 会把图片引用写成**绝对路径 + 混用分隔符**的形式：

```html
<img src="F:\work\rtest\out_pandoc\media/media/image1.jpeg" style="width:2.28678in;height:3.04747in" />
      └─ 本机绝对路径，文件一转给别人就断链
                              └─ \ 与 / 混用
                                          └─ 残留 HTML 样式，Markdown 用不上
```

`.docx` 输入产出这种 HTML 标签；`.doc` 输入则产出 Markdown 引用，
但目标同样是绝对路径：

```
F:\work\rtest\_test_real.media/media/image1.jpeg
```

两种形式实测都是 11 个引用、11 个断链。

**注意：pandoc 对图片的放置位置其实是正确的**（本项目论文 5 张图
每张都紧邻对应图注，Figure 4 的 a、b 两张也正确合并在同一行）。
问题只在路径，不在位置。

**解决办法**：只取文件名重建相对引用
`<原名>.media/imageN.png`，并沿用 pandoc 自己的位置判断。
见 `docx2md.py`。

#### 附带坑：文件名含空格/括号会让 Markdown 链接失效

源文件名常有空格和括号（如 `…study (1).md`），则图片路径也带这些字符。
实测（以 pandoc 这个严格 GFM 实现为准）：

| 写法 | 结果 |
|---|---|
| `![x](a b (1).media/img.png)` | **失效**，被当成普通文本 |
| `![x](<a b (1).media/img.png>)` | 正常，src 保持可读 |
| `![x](a%20b%20%281%29.media/img.png)` | 正常，但 src 是百分号编码 |

用尖括号形式即可，无需百分号编码（Windows 文件名不允许出现 `<` `>`）。

### 坑 2：转义符号

pandoc 会把特殊字符转义以防被误解析为 Markdown 语法：

```markdown
No. (%) of Patients\*      ← 原本是 *
interquartile range \[IQR\]   ← 原本是 [ ]
```

论文正文里方括号很常见（`[IQR]`、`[CI]`），会大量出现。
需要后处理还原。

### 坑 3：Word 段落边框变成 `<hr>`

docx 里用于装饰的分隔线，pandoc 会转成 Markdown 水平线 `-----`，
一条条塞进正文，实测产生了 14 条噪音。

### 坑 4：样式语义丢失

Word 里的「标题 1」样式，转出来仍然是 `**Abstract**`（加粗），
而不是 `# Abstract`（标题）。因为很多论文模板里章节标题就是加粗居中，
pandoc 只能识别真正的 heading style。

**解决办法**：后处理把「短加粗独占行」升级为 `##` 标题。

### 坑 5：上标/特殊符号丢失

docx 表格中的 `†`（上标匕首符号）在 pandoc 输出中会变成 `?`。
（对比：MinerU 从 PDF 解析能正确保留 `†`。）

---

## 八、pandoc 做不到的事

| 事项 | 原因 |
|---|---|
| **读 PDF** | 不支持。报 `Unknown input format 'pdf'`，退出码 21 |
| **写 PDF** | 需外部排版引擎（xelatex / tectonic / typst 等），本机未装 |
| **识别 Word 样式语义** | 加粗 ≠ 标题，需后处理 |
| **图片语义定位** | 浮动图片无法确定插入位置 |
| **扫描件 OCR** | 完全不支持 |

验证命令：

```powershell
pandoc --list-input-formats  | Select-String pdf   # 无输出
pandoc --list-output-formats | Select-String pdf   # pdf
```

### 什么时候该换工具

| 场景 | 推荐工具 |
|---|---|
| 规整的 Word 文档 | **pandoc**（本项目场景） |
| 扫描件 PDF | OCR：PaddleOCR / Tesseract |
| 双栏排版 PDF、公式密集 | MinerU（`--tier standard`） |
| 简单 PDF 取纯文本 | `pdftotext -layout`（poppler） |

---

## 九、本项目推荐流程

### 处理本项目的论文 docx

推荐直接用现成的 `docx2md.py`（封装了 pandoc 调用 + 后处理）：

```powershell
cd F:\work\rtest
python docx2md.py "Application of the MEGASYSTEM-C cementless prosthesis system in reconstruction following lower limb bone tumor resection - a retrospective study (1).docx"
```

等价的裸 pandoc 命令：

```powershell
pandoc "论文.docx" -t gfm --wrap=none --extract-media=out_pandoc\media -o out_pandoc\paper_clean.md
```

### 产物

```
论文.docx  ->  论文.md  +  论文.media\
```

`docx2md.py` 生成的 `论文.md`（51 KB / 345 行 + 11 张图）已通过 pandoc 严格 GFM 解析验证，11 个图片引用全部可解析，0 断链。

### 质量核验记录

与 MinerU（从 PDF 解析）对比的实测数据：

| 检查项 | pandoc (docx) | MinerU (pdf) |
|---|---|---|
| 文字粘连错误（`underwentprimary`） | 0 处 | 1 处 |
| 表格 `†` 上标符号 | 丢失变 `?` | 正确保留 |
| 表格缩进层级 | 丢失 | 保留 |
| 图片提取 | 11 张 | flash/basic 档位 0 张 |

结论：**有 docx 就从 docx 转，不要绕 PDF。**
docx 里本来就带着正确的空格和字符；PDF 只记录每个字的坐标，
逆向重建必然有损。

### 遗留问题

- 原文图注在正文和文末图例表中各出现一次，后处理会让同批图出现两次
- `image9.png` 可能未被引用（期刊 logo）
- 「短加粗独占行升级为标题」是启发式判断，可能把论文副标题误判为标题
  （`docx2md.py` 会在输出里逐条列出所有升级位置供核对，
  误判时可用 `--no-postprocess` 生成原始输出对照）

---

## 附录：docx2md.py

`F:\work\rtest\docx2md.py` —— 把 pandoc 调用与后处理封装成一步。

```powershell
python docx2md.py <文件或目录>              # 目录则递归处理
python docx2md.py <文件> --no-postprocess    # 只做 pandoc 转换，跳过后处理
```

它解决的问题（均在本项目论文上实测复现）：

1. **图片路径不可移植** —— pandoc 写出的可能是
   `F:\work\rtest\...media/media/image1.jpeg` 这种绝对路径 + 反斜杠形式，
   发给别人就断链。程序统一重写为 `<原名>.media/imageN.png` 相对引用
2. **Markdown 链接语法失效** —— 源文件名常含空格和括号，
   `![x](a b (1).media/img.png)` 在 pandoc / GitHub 里会被当成普通文本。
   程序自动改用尖括号形式 `![x](<a b (1).media/img.png>)`（实测可正常解析，
   且 src 保持可读，无需百分号编码）
3. **`<img>` 与 `![](...)` 两种形式并存** —— .docx 路径产出 HTML 标签，
   .doc 路径产出 Markdown 引用，程序统一处理
4. **`--extract-media` 多套一层目录** —— 程序自动展平
5. **覆盖残留** —— 重跑前清空 `<原名>.media\`，不留下上一次的孤儿文件

关于「图片位置」：pandoc 本身就把浮动图片放在正确位置（实测每张都紧邻
对应图注），程序**沿用它的判断，不猜位置、不重排**，所以不存在错位风险。

`.doc` 支持通过 PowerShell 驱动 Word COM 转换（无需 pywin32），
中间 `.docx` 存于系统临时目录并即时删除。

局限：

- 标题升级是启发式判断，输出中会逐条列出供核对
- `.doc` 转换依赖本机安装 Microsoft Word；文件有密码或弹窗时会在
  `--timeout`（默认 300 秒）后放弃并清理残留进程

