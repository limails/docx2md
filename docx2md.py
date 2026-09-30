#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
docx2md —— Word 文档 (.doc / .docx) 转 Markdown

依赖：Python 3.8+ 标准库 + pandoc。.doc 的转换依赖本机安装的 Microsoft Word
      （通过 PowerShell 驱动 COM 调用，不使用 pywin32）。

用法：
    python docx2md.py 论文.docx
    python docx2md.py D:\\docs\\            # 递归处理整个目录

产物（与源文件同目录）：
    论文.docx  ->  论文.md  +  论文.media\\      # 图片目录，覆盖时整体重建
    论文.doc   ->  论文.md  +  论文.media\\      # 中间 .docx 存于系统临时目录，用完即删

后处理 4 步：
    1. 删除 docx 段落边框产生的 <hr> 噪音
    2. 还原 pandoc 的转义符 \\[ \\] \\* \\_ \\# \\`
    3. 就地把 <img src="绝对路径" style="..."> 规范化为 ![](相对路径)
       —— 保留 pandoc 自己的位置判断（实测它是正确的），不猜位置、不重排
    4. 将「短加粗独占行」升级为 ## 标题
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# --------------------------------------------------------------------------
# Windows 控制台中文输出（默认 GBK 会花屏）
# --------------------------------------------------------------------------
if sys.platform == "win32":
    for _s in ("stdout", "stderr"):
        try:
            setattr(sys, _s, reconfigure(encoding="utf-8", errors="replace"))
        except Exception:
            pass


# --------------------------------------------------------------------------
# 正则
# --------------------------------------------------------------------------
# 独占一行的水平线（docx 段落边框被 pandoc 转成 <hr>）
RE_HR = re.compile(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$")
# <img ...> 标签（可能一行多个）
RE_IMG = re.compile(r"<img\b[^>]*/?>", re.IGNORECASE)
RE_ATTR = re.compile(r"([\w:-]+)\s*=\s*\"([^\"]*)\"")
# pandoc 转义
RE_UNESCAPE = re.compile(r"\\([\[\]\*_#`>~|])")
# 图注：Figure 1. / Fig. 2 / Table 3
RE_CAPTION = re.compile(r"^\**\s*(figure|fig\.?|table)\s*(\d+)", re.IGNORECASE)
# 整行加粗（可升级为标题）
RE_BOLD_ONLY = re.compile(r"^\*{1,2}([^*\n]{2,60}?)\*{1,2}:?\s*$")
# Markdown 表格行 / 分隔行，用于跳过表格内内容
RE_TABLE_ROW = re.compile(r"^\s*\|")
RE_TABLE_SEP = re.compile(r"^\s*\|?[\s:|-]+\|[\s:|-]*$")

DOC_EXT = {".doc", ".docx"}


# --------------------------------------------------------------------------
# pandoc 定位
#
# 跨机器使用时安装位置可能不同（winget/choco/scoop/手动解压/MSI），
# 因此按下列顺序查找，命中即返回：
#   1. 命令行 --pandoc 显式指定
#   2. 环境变量 PANDOC_EXE
#   3. PATH（shutil.which）
#   4. Windows 注册表（winget/MSI 装了但当前进程 PATH 未刷新时靠这条兜住）
#   5. 各包管理器的常见安装位置
#   6. 上述位置下的有限广搜
# --------------------------------------------------------------------------
def _registry_pandoc() -> str | None:
    """
    从注册表里找 pandoc 的安装路径。

    解决一类常见故障：用 winget 装完 pandoc 后，当前已经打开的
    PowerShell / IDE 进程内 PATH 还是旧快照，shutil.which 会失败，
    但注册表里已经写入了真实路径。
    """
    if sys.platform != "win32":
        return None
    try:
        import winreg
    except ImportError:
        return None

    subkeys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pandoc.exe"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pandoc.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Pandoc"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Pandoc"),
    ]
    for root, sub in subkeys:
        try:
            with winreg.OpenKey(root, sub) as k:
                for value_name in (None, "Location", "InstallLocation", "Path", ""):
                    try:
                        raw = winreg.QueryValueEx(k, value_name)[0]
                    except FileNotFoundError:
                        continue
                    for cand in str(raw).split(";"):
                        cand = cand.strip().strip('"')
                        if not cand:
                            continue
                        p = Path(cand)
                        if p.is_file() and p.name.lower() == "pandoc.exe":
                            return str(p)
                        if p.is_dir() and (p / "pandoc.exe").is_file():
                            return str(p / "pandoc.exe")
        except OSError:
            continue
    return None


def _candidate_paths() -> list[Path]:
    """各包管理器 / 安装方式下的常见位置。"""
    local = os.environ.get("LOCALAPPDATA", "")
    prog = os.environ.get("ProgramFiles", r"C:\Program Files")
    prog86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    user = os.environ.get("USERPROFILE", "")
    progdata = os.environ.get("ProgramData", r"C:\ProgramData")

    return [
        # winget per-user（本项目实测位置）
        Path(local) / "Pandoc" / "pandoc.exe",
        Path(local) / "Programs" / "Pandoc" / "pandoc.exe",
        # 机器级 / MSI
        Path(prog) / "Pandoc" / "pandoc.exe",
        Path(prog86) / "Pandoc" / "pandoc.exe",
        # scoop
        Path(user) / "scoop" / "apps" / "pandoc" / "current" / "pandoc.exe",
        Path(progdata) / "scoop" / "apps" / "pandoc" / "current" / "pandoc.exe",
        # chocolatey
        Path(progdata) / "chocolatey" / "bin" / "pandoc.exe",
        Path(progdata) / "chocolatey" / "lib" / "pandoc" / "tools" / "pandoc.exe",
    ]


def _sweep_dirs() -> list[Path]:
    """在几个小范围内广搜 pandoc.exe，命中即用。"""
    roots = []
    for env in ("LOCALAPPDATA", "ProgramFiles", "ProgramFiles(x86)", "ProgramData"):
        v = os.environ.get(env)
        if v:
            roots.append(Path(v) / "Pandoc")
    roots += [
        Path(os.environ.get("USERPROFILE", "")) / "scoop" / "apps" / "pandoc",
        Path(os.environ.get("ProgramData", "")) / "scoop" / "apps" / "pandoc",
    ]
    out: list[Path] = []
    for root in roots:
        if not root.is_dir():
            continue
        try:
            # 限制递归深度，避免在大目录上无谓遍历
            for p in root.glob("**/pandoc.exe"):
                if p.is_file():
                    out.append(p)
        except (OSError, PermissionError):
            continue
    return out


def find_pandoc(explicit: str = "") -> str:
    """按优先级查找 pandoc 可执行文件。找不到则抛 RuntimeError。"""
    tried: list[str] = []

    # 1. 命令行显式指定
    if explicit:
        p = Path(explicit).expanduser()
        if p.is_file():
            return str(p)
        if shutil.which(explicit):
            return shutil.which(explicit)
        raise RuntimeError(f"--pandoc 指定的路径不存在：{explicit}")

    # 2. 环境变量
    env = os.environ.get("PANDOC_EXE", "").strip()
    if env:
        p = Path(env).expanduser()
        if p.is_file():
            return str(p)
        if p.is_dir() and (p / "pandoc.exe").is_file():
            return str(p / "pandoc.exe")
        found = shutil.which(env)
        if found:
            return found
        raise RuntimeError(
            f"环境变量 PANDOC_EXE 指向的路径无效：{env}\n"
            "请修正该变量，或删除它以改用自动查找。"
        )

    # 3. PATH
    exe = shutil.which("pandoc")
    if exe:
        return exe

    # 4. 注册表（装了但当前进程 PATH 未刷新的情况）
    reg = _registry_pandoc()
    if reg:
        return reg

    # 5. 常见安装位置
    for c in _candidate_paths():
        tried.append(str(c))
        if c.is_file():
            return str(c)

    # 6. 有限广搜
    for c in _sweep_dirs():
        tried.append(str(c))
        return str(c)

    tried_txt = "\n".join(f"    {t}" for t in tried) or "    （无候选路径）"
    raise RuntimeError(
        "找不到 pandoc。\n\n"
        f"已尝试的位置：\n{tried_txt}\n\n"
        "解决办法（任选其一）：\n"
        "  1. 装到 PATH 上（推荐）：\n"
        "       winget install --id JohnMacFarlane.Pandoc -e\n"
        "     安装后请新开一个终端窗口 —— 当前已打开的窗口 PATH 还是旧的。\n"
        "  2. 指定环境变量（适合非标准位置）：\n"
        "       [Environment]::SetEnvironmentVariable('PANDOC_EXE','D:\\tools\\pandoc.exe','User')\n"
        "  3. 本次运行临时指定：\n"
        "       python docx2md.py 论文.docx --pandoc D:\\tools\\pandoc.exe"
    )


# --------------------------------------------------------------------------
# .doc -> .docx（PowerShell 驱动 Word COM）
# --------------------------------------------------------------------------
PS_CONVERT = r"""
param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
# 若用户本来就开着 Word，则不要退出他的实例
$wasRunning = $false
try { $wasRunning = [bool](Get-Process WINWORD -EA SilentlyContinue) } catch { $wasRunning = $false }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    $doc = $word.Documents.Open($In, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    $doc.SaveAs([ref]$Out, [ref]16)                   # 16 = wdFormatDocumentDefault (.docx)
    $doc.Close([ref]0)                               # 0 = wdDoNotSaveChanges
    $doc = $null
} finally {
    if ($doc -ne $null) { try { $doc.Close([ref]0) } catch {} }
    if (-not $wasRunning) { try { $word.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""

WORD_TIMEOUT = 300  # 秒


def doc_to_docx(src: Path, workdir: Path) -> Path:
    """用 Word 把 .doc 转成 .docx，输出到 workdir。返回生成的 .docx 路径。"""
    dst = workdir / (src.stem + ".docx")
    script = workdir / "_convert.ps1"
    script.write_text(PS_CONVERT, encoding="utf-8-sig")

    cmd = [
        "powershell", "-NoProfile", "-NonInteractive",
        "-ExecutionPolicy", "Bypass",
        "-File", str(script), "-In", str(src), "-Out", str(dst),
    ]
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=WORD_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        _kill_stray_word()
        raise RuntimeError(
            f"Word 转换超时（>{WORD_TIMEOUT}s）。文件可能有密码保护或弹出对话框。"
        )

    if proc.returncode != 0 or not dst.is_file():
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"Word 转换 .doc 失败：{err[:500] or f'退出码 {proc.returncode}'}")

    return dst


def _kill_stray_word() -> None:
    """超时后清理可能残留的无头 WINWORD 进程。"""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            capture_output=True, timeout=30,
        )
    except Exception:
        pass


# --------------------------------------------------------------------------
# pandoc 转换
# --------------------------------------------------------------------------
def run_pandoc(pandoc: str, src_docx: Path, out_md: Path, media_dir: Path) -> None:
    cmd = [
        pandoc, str(src_docx),
        "-f", "docx",
        "-t", "gfm",
        "--wrap=none",
        f"--extract-media={media_dir}",
        "-o", str(out_md),
    ]
    proc = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=900,
    )
    if proc.returncode != 0 or not out_md.is_file():
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"pandoc 转换失败：{err[:500] or f'退出码 {proc.returncode}'}")


def flatten_media(media_dir: Path) -> int:
    """pandoc 的 --extract-media=DIR 会固定多套一层 media/，这里展平成 DIR/。"""
    inner = media_dir / "media"
    if not inner.is_dir():
        return len(list(media_dir.glob("*"))) if media_dir.is_dir() else 0
    n = 0
    for item in inner.iterdir():
        shutil.move(str(item), str(media_dir / item.name))
        n += 1
    try:
        inner.rmdir()
    except OSError:
        pass
    return n


# --------------------------------------------------------------------------
# 后处理
# --------------------------------------------------------------------------
def postprocess(md: str, media_dir_name: str) -> tuple[str, list[str], list[str]]:
    """
    返回 (处理后文本, 图片处理日志, 标题升级日志)
    """
    img_log: list[str] = []
    head_log: list[str] = []
    out_lines: list[str] = []
    img_no = 0

    lines = md.split("\n")
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw

        # --- 步骤 1：删除水平线噪音 ---
        if RE_HR.match(line):
            i += 1
            continue

        # --- 步骤 3：就地规范化图片引用（<img> 标签与已有的 ![]() 两种形式）---
        if RE_IMG.search(line):
            following = lines[i + 1:]
            line, n, notes = _rewrite_images(line, media_dir_name, img_no, following)
            img_no += n
            img_log.extend(notes)
        elif "![" in line:
            following = lines[i + 1:]
            line, n, notes = _normalize_md_images(line, media_dir_name, img_no, following)
            img_no += n
            img_log.extend(notes)

        # --- 步骤 2：还原转义 ---
        line = RE_UNESCAPE.sub(r"\1", line)

        # --- 步骤 4：短加粗独占行 -> 标题 ---
        if not RE_TABLE_ROW.match(line) and not RE_TABLE_SEP.match(line):
            m = RE_BOLD_ONLY.match(line)
            if m:
                body = m.group(1).strip()
                is_caption = bool(RE_CAPTION.match(body))
                ends_sentence = body.endswith((".", "。", "!", "?", ";", ":", ","))
                if body and not is_caption and not ends_sentence:
                    line = f"## {body}"
                    head_log.append(f"第{i + 1}行: **{body}** -> ## {body}")

        out_lines.append(line)
        i += 1

    # 压缩 3 个以上连续空行
    cleaned: list[str] = []
    for line in out_lines:
        if line.strip() == "" and len(cleaned) >= 2 and cleaned[-1].strip() == "" and cleaned[-2].strip() == "":
            continue
        cleaned.append(line)

    return "\n".join(cleaned), img_log, head_log


def _next_caption_alt(following: list[str]) -> str:
    """本行之后最近的非空行若是图注，取 "Figure N" 作为 alt。"""
    nxt = next((l for l in following if l.strip()), "")
    cm = RE_CAPTION.match(nxt.strip())
    return f"{cm.group(1).title()} {cm.group(2)}" if cm else ""


def _rel_dest(raw: str, media_dir_name: str) -> str:
    """从任意形式的路径重建相对引用（只取文件名，免疫绝对路径与反斜杠）。"""
    d = raw.strip()
    if d.startswith("<") and d.endswith(">"):
        d = d[1:-1]
    base = re.split(r"[\\/]", d)[-1].strip()
    if not base:
        return raw
    return md_dest(f"{media_dir_name}/{base}")


def _normalize_md_images(
    line: str, media_dir_name: str, start_no: int, following: list[str]
) -> tuple[str, int, list[str]]:
    """
    规范化行内已存在的 Markdown 图片引用 !alt(dest)。

    Word 转出的 .docx 常让 pandoc 直接输出 markdown 引用，且 dest 是
    F:\\dir\\name.media/media/image1.png 这种绝对路径 + 反斜杠形式，全部断链。
    这里按括号配对扫描（路径里可能含括号），逐个重写为相对引用。
    """
    notes: list[str] = []
    count = 0
    caption_alt = _next_caption_alt(following)
    out: list[str] = []
    i = 0

    while True:
        j = line.find("![", i)
        if j < 0:
            out.append(line[i:])
            break
        k = line.find("]", j)
        if k < 0:
            out.append(line[i:])
            break
        if k + 1 >= len(line) or line[k + 1] != "(":
            out.append(line[i:k + 1])
            i = k + 1
            continue

        depth, m = 0, k + 1
        while m < len(line):
            if line[m] == "(":
                depth += 1
            elif line[m] == ")":
                depth -= 1
                if depth == 0:
                    break
            m += 1
        if m >= len(line):          # 括号不配对，原样保留
            out.append(line[i:])
            break

        alt = line[j + 2:k]
        raw_dest = line[k + 2:m]
        new_dest = _rel_dest(raw_dest, media_dir_name)
        new_alt = alt.strip() or caption_alt or "image"
        out.append(line[i:j])
        out.append(f"![{new_alt}]({new_dest})")
        notes.append(
            f"第{start_no + count + 1}张: {re.split(r'[\\\\/]', raw_dest.strip('<> '))[-1]}"
            f" -> ![{new_alt}]({new_dest})"
        )
        count += 1
        i = m + 1

    return "".join(out), count, notes


def md_dest(rel: str) -> str:
    """
    生成 Markdown 链接目标。

    源文件名常含空格和括号（如 "…study (1).md"），直接写
        ![x](a b (1).media/img.png)
    在严格 GFM/CommonMark 解析器（pandoc、GitHub）里会被当成普通文本，链接失效。
    实测用尖括号包裹即可正常解析，且 src 保持可读，无需百分号编码：
        ![x](<a b (1).media/img.png>)   ->  OK
    Windows 文件名不允许出现 < > ，故无需处理。
    """
    if any(c in rel for c in " ()"):
        return f"<{rel}>"
    return rel


def _rewrite_images(
    line: str, media_dir_name: str, start_no: int, following: list[str]
) -> tuple[str, int, list[str]]:
    """把一行里的所有 <img> 标签改写为 Markdown 图片引用，保留原有位置。"""
    notes: list[str] = []
    count = 0

    # 本行之后最近的非空行，用于推断 alt 文字
    caption_alt = _next_caption_alt(following)

    def repl(m: re.Match) -> str:
        nonlocal count
        attrs = dict(RE_ATTR.findall(m.group(0)))
        src = attrs.get("src", "")
        if not src:
            return m.group(0)

        # 只取文件名，重建相对路径 —— 免疫 pandoc 写出的绝对路径与分隔符混用
        base = re.split(r"[\\/]", src)[-1]
        if not base:
            return m.group(0)

        rel = md_dest(f"{media_dir_name}/{base}")
        alt = caption_alt or "image"
        notes.append(f"第{start_no + count + 1}张: {base} -> ![{alt}]({rel})")
        count += 1
        return f"![{alt}]({rel})"

    new_line = RE_IMG.sub(repl, line)
    return new_line, count, notes


# --------------------------------------------------------------------------
# 单文件转换
# --------------------------------------------------------------------------
class Result:
    __slots__ = ("src", "ok", "error", "images", "heads", "elapsed")

    def __init__(self, src: Path):
        self.src = src
        self.ok = False
        self.error = ""
        self.images: list[str] = []
        self.heads: list[str] = []
        self.elapsed = 0.0


def convert_one(pandoc: str, src: Path, workdir: Path, do_post: bool = True) -> Result:
    res = Result(src)
    t0 = time.time()
    tmpdir: Path | None = None

    try:
        out_md = src.with_suffix(".md")
        media_dir = src.parent / (src.stem + ".media")

        # 中间 .docx（仅 .doc 需要）
        if src.suffix.lower() == ".doc":
            tmpdir = Path(tempfile.mkdtemp(prefix="docx2md_"))
            docx = doc_to_docx(src, tmpdir)
        else:
            docx = src

        # 覆盖前清空，保证不残留上一次的多余文件
        if media_dir.is_dir():
            shutil.rmtree(media_dir, ignore_errors=True)

        run_pandoc(pandoc, docx, out_md, media_dir)
        n_files = flatten_media(media_dir)

        if do_post:
            text = out_md.read_text(encoding="utf-8")
            new_text, img_log, head_log = postprocess(text, media_dir.name)
            out_md.write_text(new_text, encoding="utf-8")
            res.images = img_log
            res.heads = head_log
        else:
            res.images = [f"  (未做后处理，pandoc 原样输出，共 {n_files} 个媒体文件)"]

        res.ok = True

    except Exception as exc:  # noqa: BLE001
        res.error = f"{type(exc).__name__}: {exc}"
    finally:
        if tmpdir is not None:
            shutil.rmtree(tmpdir, ignore_errors=True)
        res.elapsed = time.time() - t0

    return res


# --------------------------------------------------------------------------
# 批量
# --------------------------------------------------------------------------
def collect_targets(target: Path) -> list[Path]:
    if target.is_file():
        if target.suffix.lower() not in DOC_EXT:
            raise RuntimeError(f"不是 Word 文档：{target.name}（仅支持 .doc / .docx）")
        return [target]
    if target.is_dir():
        found = [
            p for p in sorted(target.rglob("*"))
            if p.is_file() and p.suffix.lower() in DOC_EXT
            and not p.name.startswith("~$")          # Word 锁文件
        ]
        return found
    raise RuntimeError(f"路径不存在：{target}")


def _pandoc_version(pandoc: str) -> str:
    """读取 pandoc 版本，用于运行时确认找到的是可用的程序。"""
    try:
        proc = subprocess.run(
            [pandoc, "--version"], capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=30,
        )
        return (proc.stdout or proc.stderr).strip().splitlines()[0]
    except Exception:  # noqa: BLE001
        return "(读取失败)"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Word (.doc/.docx) -> Markdown，调用 pandoc 并做后处理",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("target", help="要转换的 .doc/.docx 文件，或包含它们的目录（递归）")
    ap.add_argument("--no-postprocess", action="store_true",
                    help="只做 pandoc 转换，跳过后处理（用于排查后处理问题）")
    ap.add_argument("--pandoc", default="", metavar="路径",
                    help="指定 pandoc.exe 路径，跳过自动查找。"
                         "也可改用环境变量 PANDOC_EXE，效果相同")
    ap.add_argument("--timeout", type=int, default=WORD_TIMEOUT,
                    help=f"Word COM 转换超时秒数，默认 {WORD_TIMEOUT}")
    args = ap.parse_args()

    try:
        pandoc = find_pandoc(args.pandoc)
    except RuntimeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        return 2

    target = Path(args.target).expanduser()
    try:
        targets = collect_targets(target)
    except RuntimeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        return 2

    if not targets:
        print(f"[提示] {target} 下没有找到 .doc / .docx 文件。")
        return 0

    print(f"pandoc : {pandoc}")
    print(f"版本   : {_pandoc_version(pandoc)}")
    print(f"目标   : {target}")
    print(f"待处理 : {len(targets)} 个文件")
    print("-" * 68)

    results: list[Result] = []
    with tempfile.TemporaryDirectory(prefix="docx2md_work_") as wd:
        workdir = Path(wd)
        for src in targets:
            r = convert_one(pandoc, src, workdir, do_post=not args.no_postprocess)
            results.append(r)

            if r.ok:
                out = src.with_suffix(".md")
                print(f"[成功] {src.name}")
                print(f"       -> {out.name}  ({out.stat().st_size:,} 字节, {r.elapsed:.1f}s)")
                if r.images:
                    print(f"       图片 {len(r.images)} 处（就地规范化，未移动位置）：")
                    for line in r.images:
                        print(f"       {line}")
                if r.heads:
                    print(f"       标题升级 {len(r.heads)} 处（如有误判请用 --no-postprocess 对照）：")
                    for line in r.heads[:5]:
                        print(f"       {line}")
                    if len(r.heads) > 5:
                        print(f"       ...另有 {len(r.heads) - 5} 处")
            else:
                print(f"[失败] {src.name}")
                print(f"       {r.error}")
            print("-" * 68)

    ok = [r for r in results if r.ok]
    bad = [r for r in results if not r.ok]
    print("=" * 68)
    print(f"汇总：成功 {len(ok)} 个，失败 {len(bad)} 个，共 {len(results)} 个")
    if bad:
        print()
        print("失败清单：")
        for r in bad:
            print(f"  - {r.src}")
            print(f"      {r.error}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
