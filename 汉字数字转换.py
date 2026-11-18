# -*- coding: utf-8 -*-
"""
汉字数字 → 阿拉伯数字 批量转换器（Windows 图形界面）

把一个文件夹里所有文本文件中的汉字数字
（一二三四五六七八九十十一十二……）
替换成阿拉伯数字（1 2 3 4 5 6 7 8 9 10 11 12……），也可以同时改文件名。

特点：
  · 转换前自动备份，随时可以一键撤销
  · 智能模式只改“明确的数字”，不会把“一样”改成“1样”
  · 保留原文件编码（UTF-8 / 带 BOM 的 UTF-8 / GBK 等）

命令行用法（给喜欢折腾的人 / 自动化用）：
    python 汉字数字转换.py --convert "D:\\某个文件夹"
"""

import codecs
import json
import os
import queue
import re
import shutil
import sys
import threading
import time
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "汉字数字 → 阿拉伯数字 转换器"
BACKUP_PREFIX = "_汉字数字转换备份_"
MAX_FILE_BYTES = 30 * 1024 * 1024      # 超过 30 MB 的文本文件跳过
MAX_SAMPLES = 400                      # 预览时最多展示多少处修改

# ------------------------------------------------------------------ 数字表

DIGITS = {
    "零": 0, "〇": 0, "○": 0,
    "一": 1, "壹": 1,
    "二": 2, "贰": 2, "两": 2,
    "三": 3, "叁": 3,
    "四": 4, "肆": 4,
    "五": 5, "伍": 5,
    "六": 6, "陆": 6,
    "七": 7, "柒": 7,
    "八": 8, "捌": 8,
    "九": 9, "玖": 9,
}
SMALL_UNITS = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000}
BIG_UNITS = {"万": 10000, "萬": 10000, "亿": 100000000, "億": 100000000}

NUMERAL_CHARS = "".join(DIGITS) + "".join(SMALL_UNITS) + "".join(BIG_UNITS)
DECIMAL_CHARS = "".join(DIGITS)

# 一连串汉字数字（可以带“点几”的小数）
RUN_RE = re.compile(
    "[" + re.escape(NUMERAL_CHARS) + "]+"
    "(?:[点點][" + re.escape(DECIMAL_CHARS) + "]+)?"
)

# 这些字出现在数字前面时，基本可以确定后面是数字
PREFIX_OK = set("第共约近超达满总每从到至前")
# 这些字出现在数字后面时，基本可以确定前面是数字（章节、页码、序号……）
SUFFIX_OK = set("章节条款项集话回卷册部篇页行列组班级层楼栋号期届轮课题步版年月日")
# 整词黑名单：它们是词，不是数字
BLOCK_WORDS = {"万一", "千万", "万万", "一五一十"}

# 文档类文件
DOC_EXTS = {
    ".txt", ".md", ".markdown", ".csv", ".tsv", ".srt", ".ass", ".ssa", ".vtt",
    ".log", ".lrc", ".rst", ".tex", ".html", ".htm", ".xhtml", ".xml", ".json",
    ".yaml", ".yml", ".ini", ".conf", ".properties", ".nfo", ".htm", ".mht",
}
# 代码类文件
CODE_EXTS = {
    ".py", ".pyw", ".js", ".mjs", ".cjs", ".ts", ".css", ".scss", ".less",
    ".java", ".c", ".h", ".cpp", ".hpp", ".cc", ".cs", ".php", ".go", ".rs",
    ".lua", ".rb", ".sh", ".bash", ".bat", ".cmd", ".ps1", ".sql", ".vue",
    ".swift", ".kt", ".m", ".mm", ".r", ".pl", ".asm", ".proto", ".gradle",
    ".toml", ".cfg", ".config", ".sln", ".csproj", ".vcxproj", ".gitignore",
    ".editorconfig", ".htaccess", ".dockerfile",
}


def _is_cjk(ch):
    """是否是汉字（含扩展区）。"""
    if not ch:
        return False
    o = ord(ch)
    return (0x3400 <= o <= 0x4DBF) or (0x4E00 <= o <= 0x9FFF) \
        or (0xF900 <= o <= 0xFAFF) or (0x20000 <= o <= 0x2FA1F)


def _parse_int(s):
    """把纯整数部分的汉字数字转成 int，转不了返回 None。"""
    if not s:
        return None
    has_unit = any(c in SMALL_UNITS or c in BIG_UNITS for c in s)

    if not has_unit:
        # 没有单位：一位一位读，比如 二〇二六 → 2026
        if not all(c in DIGITS for c in s):
            return None
        if len(s) == 1:
            return DIGITS[s]
        return int("".join(str(DIGITS[c]) for c in s))

    total = 0      # 已经算好的部分
    section = 0    # 万以下的一段
    num = 0        # 当前这一个数字
    seen = False
    for c in s:
        if c in DIGITS:
            num = DIGITS[c]
            seen = True
        elif c in SMALL_UNITS:
            unit = SMALL_UNITS[c]
            if num == 0:
                num = 1          # “十二”的“十”前面没有数字，按 1 算
            section += num * unit
            num = 0
            seen = True
        elif c in BIG_UNITS:
            unit = BIG_UNITS[c]
            section += num
            num = 0
            if section == 0:
                section = 1
            if unit == 10000:
                total += section * 10000
            else:
                total = (total + section) * 100000000
            section = 0
            seen = True
        else:
            return None
    if not seen:
        return None
    return total + section + num


def cn_to_number(token):
    """把一段汉字数字转成阿拉伯数字字符串；转不了返回 None。"""
    if not token:
        return None
    if "点" in token or "點" in token:
        sep = "点" if "点" in token else "點"
        left, _, right = token.partition(sep)
        if not right or any(c not in DIGITS for c in right):
            return None
        left_value = _parse_int(left) if left else 0
        if left_value is None:
            return None
        return "%d.%s" % (left_value, "".join(str(DIGITS[c]) for c in right))
    value = _parse_int(token)
    return None if value is None else str(value)


def _should_convert(token, prev, nxt, text, end, aggressive):
    if aggressive:
        return token not in BLOCK_WORDS
    if token in BLOCK_WORDS:
        return False
    if prev == "第":
        return True
    left_ok = prev is None or not _is_cjk(prev) or prev in PREFIX_OK
    if nxt is None or not _is_cjk(nxt):
        right_ok = True
    elif nxt in SUFFIX_OK:
        right_ok = True
    elif text.startswith("单元", end):
        right_ok = True
    else:
        right_ok = False
    return left_ok and right_ok


def convert_text(text, aggressive=False, max_samples=0):
    """
    转换一段文本。
    返回 (新文本, 修改处数, 样例列表[(行号, 原来的整行, 改后的整行)])。
    """
    if not text:
        return text, 0, []

    out = []
    last = 0
    count = 0
    samples = []
    line_no = 1
    scan_pos = 0

    for m in RUN_RE.finditer(text):
        token = m.group(0)
        prev = text[m.start() - 1] if m.start() > 0 else None
        nxt = text[m.end()] if m.end() < len(text) else None
        if not _should_convert(token, prev, nxt, text, m.end(), aggressive):
            continue
        number = cn_to_number(token)
        if number is None:
            continue
        out.append(text[last:m.start()])
        out.append(number)
        last = m.end()
        count += 1

        if max_samples and len(samples) < max_samples:
            line_no += text.count("\n", scan_pos, m.start())
            scan_pos = m.start()
            if not samples or samples[-1][0] != line_no:
                ls = text.rfind("\n", 0, m.start()) + 1
                le = text.find("\n", m.end())
                if le < 0:
                    le = len(text)
                old_line = text[ls:le].strip()
                if len(old_line) > 120:
                    old_line = old_line[:120] + "…"
                new_line, _, _ = convert_text(old_line, aggressive)
                samples.append((line_no, old_line, new_line))

    if count == 0:
        return text, 0, []
    out.append(text[last:])
    return "".join(out), count, samples


# ---------------------------------------------------------------- 编码处理

def decode_bytes(data):
    """按 BOM / UTF-8 / GBK 的顺序识别编码，返回 (文本, 编码名)。"""
    if data.startswith(codecs.BOM_UTF8):
        return data[len(codecs.BOM_UTF8):].decode("utf-8"), "utf-8-sig"
    if data.startswith(codecs.BOM_UTF16_LE) or data.startswith(codecs.BOM_UTF16_BE):
        return data.decode("utf-16"), "utf-16"
    if b"\x00" in data[:4096]:
        raise ValueError("看起来是二进制文件")
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    try:
        return data.decode("gb18030"), "gb18030"
    except UnicodeDecodeError:
        raise ValueError("无法识别文件编码")


# ---------------------------------------------------------------- 文件扫描

def _file_kind(name, mode):
    ext = os.path.splitext(name)[1].lower()
    if mode == "all":
        return True
    if ext in DOC_EXTS:
        return True
    if mode == "text" and ext in CODE_EXTS:
        return True
    return False


def list_files(folder, include_sub, mode, exclude=()):
    """列出需要处理的文件（绝对路径）。"""
    result = []
    exclude = {os.path.normcase(os.path.abspath(p)) for p in exclude}
    if include_sub:
        walker = os.walk(folder)
    else:
        walker = [(folder, [], os.listdir(folder))]

    for root, dirs, files in walker:
        dirs[:] = [d for d in dirs
                   if not d.startswith(BACKUP_PREFIX)
                   and d not in ("$RECYCLE.BIN", "System Volume Information")]
        for name in files:
            path = os.path.join(root, name)
            if os.path.normcase(os.path.abspath(path)) in exclude:
                continue
            if name.startswith("~$"):
                continue
            if not _file_kind(name, mode):
                continue
            if os.path.isfile(path):
                result.append(path)
    result.sort()
    return result


class Plan(object):
    def __init__(self):
        self.renames = []    # {"src", "dst", "rel"}
        self.contents = []   # {"path", "rel", "new_bytes", "enc", "count", "samples"}
        self.skipped = []    # (文件名, 原因)
        self.scanned = 0
        self.folder = ""


TOOL_FILE_NAMES = ("使用说明.txt", "双击启动.bat", "一键转换（不弹界面）.bat",
                   "界面预览.png", "错误日志.txt")


def _self_paths():
    """本程序自己的文件，永远不处理，免得把自己改坏。"""
    out = []

    def add(path):
        try:
            p = os.path.abspath(path)
        except Exception:
            return
        if p and p not in out:
            out.append(p)

    try:
        add(__file__)
    except NameError:
        pass
    try:
        add(sys.argv[0])
    except Exception:
        pass
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        for name in TOOL_FILE_NAMES:
            add(os.path.join(here, name))
    except NameError:
        pass
    return out


def collect_plan(folder, include_sub=False, do_names=True, do_content=True,
                 aggressive=False, mode="doc", exclude=()):
    """扫描文件夹，算出要改什么（不写任何文件）。"""
    plan = Plan()
    plan.folder = folder
    exclude = list(exclude) + _self_paths()
    # 改名要对“所有文件”生效，所以先按全部文件列出，改内容时再按类型筛
    files = list_files(folder, include_sub, "all", exclude)
    plan.scanned = len(files)

    # ---- 文件名 ----
    if do_names:
        taken = {os.path.normcase(os.path.abspath(p)) for p in files}
        for path in files:
            name = os.path.basename(path)
            stem, ext = os.path.splitext(name)
            new_stem, count, _ = convert_text(stem, aggressive)
            if not count or new_stem == stem or not new_stem:
                continue
            candidate = new_stem + ext
            dst = os.path.join(os.path.dirname(path), candidate)
            # 目标重名时自动加序号，绝不覆盖已有文件
            base_stem, base_ext = os.path.splitext(candidate)
            i = 2
            while os.path.normcase(dst) in taken and \
                    os.path.normcase(dst) != os.path.normcase(os.path.abspath(path)):
                candidate = "%s (%d)%s" % (base_stem, i, base_ext)
                dst = os.path.join(os.path.dirname(path), candidate)
                i += 1
            taken.add(os.path.normcase(dst))
            plan.renames.append({
                "src": path,
                "dst": dst,
                "rel": os.path.relpath(path, folder),
            })

    # ---- 文件内容 ----
    if do_content:
        for path in files:
            if not _file_kind(os.path.basename(path), mode):
                continue
            try:
                size = os.path.getsize(path)
            except OSError:
                continue
            if size == 0:
                continue
            if size > MAX_FILE_BYTES:
                plan.skipped.append((os.path.relpath(path, folder), "文件太大，已跳过"))
                continue
            try:
                with open(path, "rb") as f:
                    data = f.read()
            except OSError as exc:
                plan.skipped.append((os.path.relpath(path, folder), "读取失败：%s" % exc))
                continue
            try:
                text, enc = decode_bytes(data)
            except ValueError as exc:
                plan.skipped.append((os.path.relpath(path, folder), str(exc)))
                continue
            new_text, count, samples = convert_text(text, aggressive, MAX_SAMPLES)
            if not count:
                continue
            try:
                new_bytes = new_text.encode(enc)
            except UnicodeEncodeError:
                new_bytes = new_text.encode("utf-8")
                enc = "utf-8"
            plan.contents.append({
                "path": path,
                "rel": os.path.relpath(path, folder),
                "new_bytes": new_bytes,
                "enc": enc,
                "count": count,
                "samples": samples,
            })
    return plan


def plan_is_empty(plan):
    return not plan.renames and not plan.contents


# ---------------------------------------------------------------- 执行修改

def apply_plan(plan, log=None):
    """备份后把计划写到磁盘上，返回备份目录。"""
    folder = plan.folder
    stamp = time.strftime("%Y%m%d_%H%M%S")
    backup = os.path.join(folder, BACKUP_PREFIX + stamp)
    n = 1
    while os.path.exists(backup):
        n += 1
        backup = os.path.join(folder, BACKUP_PREFIX + stamp + "_%d" % n)
    os.makedirs(backup)

    manifest = {
        "created": stamp,
        "folder": folder,
        "renames": [],
        "files": [],
    }

    # 1) 备份
    backed_up = set()
    for item in plan.contents:
        dst = os.path.join(backup, item["rel"])
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(item["path"], dst)
        backed_up.add(item["rel"])
        manifest["files"].append(item["rel"])
    for item in plan.renames:
        if item["rel"] in backed_up:
            continue
        dst = os.path.join(backup, item["rel"])
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(item["src"], dst)
        backed_up.add(item["rel"])
        manifest["files"].append(item["rel"])

    # 2) 写内容
    for item in plan.contents:
        with open(item["path"], "wb") as f:
            f.write(item["new_bytes"])
        if log:
            log("  已修改内容：%s（%d 处）" % (item["rel"], item["count"]))

    # 3) 改名（从深到浅，避免父目录先改）
    for item in sorted(plan.renames, key=lambda x: x["src"].count(os.sep), reverse=True):
        try:
            if os.path.exists(item["dst"]):
                raise OSError("目标已存在：%s" % os.path.basename(item["dst"]))
            os.replace(item["src"], item["dst"])
            manifest["renames"].append([item["rel"], os.path.relpath(item["dst"], folder)])
            if log:
                log("  已改名：%s → %s" % (os.path.basename(item["src"]),
                                          os.path.basename(item["dst"])))
        except OSError as exc:
            if log:
                log("  [跳过] 改名失败 %s：%s" % (os.path.basename(item["src"]), exc))

    with open(os.path.join(backup, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    with open(os.path.join(backup, "备份说明.txt"), "w", encoding="utf-8") as f:
        f.write("这是“汉字数字 → 阿拉伯数字 转换器”自动生成的备份。\n")
        f.write("不要手动删除这个文件夹，否则就无法撤销了。\n")
        f.write("转换时间：%s\n" % stamp)
        f.write("原文件夹：%s\n" % folder)
        f.write("改名的文件：%d 个\n" % len(manifest["renames"]))
        f.write("改内容的文件：%d 个\n" % len(manifest["files"]))
    return backup


def find_backups(folder):
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    result = []
    for name in names:
        path = os.path.join(folder, name)
        if name.startswith(BACKUP_PREFIX) and os.path.isdir(path):
            result.append(path)
    result.sort(reverse=True)
    return result


def undo_backup(backup_dir, log=None):
    """把一次转换还原回去。"""
    manifest_path = os.path.join(backup_dir, "manifest.json")
    if not os.path.isfile(manifest_path):
        raise ValueError("这个备份文件夹里没有 manifest.json，无法自动撤销")
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    folder = manifest["folder"]
    count = 0

    # 先把名字改回去（被改名的文件内容可能也改过，所以先恢复名字）
    for old_rel, new_rel in manifest.get("renames", []):
        old_path = os.path.join(folder, old_rel)
        new_path = os.path.join(folder, new_rel)
        if os.path.exists(new_path):
            if os.path.exists(old_path):
                if log:
                    log("  [跳过] %s 已存在，没有还原" % old_rel)
                continue
            os.replace(new_path, old_path)
            count += 1

    # 再按备份把内容写回去
    for rel in manifest.get("files", []):
        src = os.path.join(backup_dir, rel)
        dst = os.path.join(folder, rel)
        if os.path.isfile(src):
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            count += 1

    if log:
        log("已还原 %d 项。" % count)
    shutil.rmtree(backup_dir, ignore_errors=True)
    return count


# ---------------------------------------------------------------- 界面

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    if sys.argv and sys.argv[0]:
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.abspath(os.getcwd())


class ConverterApp(object):
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.q = queue.Queue()
        self.plan = None
        self.plan_key = None
        self.last_backup = None

        root.title(APP_TITLE)
        root.geometry("940x700")
        root.minsize(780, 560)

        try:
            root.iconname("汉字数字转换")
        except Exception:
            pass

        ui_font = ("Microsoft YaHei UI", 10)
        root.option_add("*Font", ui_font)

        pad = {"padx": 10, "pady": 4}

        # ---------- 顶部说明 ----------
        head = tk.Frame(root)
        head.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(head, text="汉字数字 → 阿拉伯数字", font=("Microsoft YaHei UI", 15, "bold"))\
            .pack(anchor="w")
        tk.Label(head,
                 text="例：第一章 → 第1章　　十二、 → 12、　　（一） → (1)",
                 fg="#555555").pack(anchor="w")
        tk.Label(head,
                 text="像「一样、一直、一个人、十二生肖」这样的词不会被误改",
                 fg="#777777").pack(anchor="w")

        # ---------- 文件夹 ----------
        box = tk.LabelFrame(root, text=" 要处理的文件夹 ")
        box.pack(fill="x", padx=12, pady=8)
        row = tk.Frame(box)
        row.pack(fill="x", padx=8, pady=6)
        self.folder_var = tk.StringVar(value=app_dir())
        self.entry = tk.Entry(row, textvariable=self.folder_var, font=ui_font)
        self.entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(row, text="浏览…", width=8, command=self.choose_folder).pack(side="left", padx=(8, 0))

        # ---------- 选项 ----------
        opts = tk.LabelFrame(root, text=" 选项 ")
        opts.pack(fill="x", padx=12)
        line1 = tk.Frame(opts)
        line1.pack(fill="x", padx=8, pady=(6, 2))
        self.sub_var = tk.BooleanVar(value=False)
        self.names_var = tk.BooleanVar(value=True)
        self.content_var = tk.BooleanVar(value=True)
        self.kind_var = tk.StringVar(value="文档类（txt / md / csv / srt…）")
        tk.Checkbutton(line1, text="包含子文件夹", variable=self.sub_var)\
            .pack(side="left", padx=(0, 18))
        tk.Checkbutton(line1, text="修改文件名", variable=self.names_var)\
            .pack(side="left", padx=(0, 18))
        tk.Checkbutton(line1, text="修改文件内容", variable=self.content_var)\
            .pack(side="left")

        line2 = tk.Frame(opts)
        line2.pack(fill="x", padx=8, pady=(2, 2))
        self.mode_var = tk.StringVar(value="smart")
        tk.Label(line2, text="转换方式：").pack(side="left")
        tk.Radiobutton(line2, text="智能（推荐：只改明确的数字）",
                       variable=self.mode_var, value="smart").pack(side="left", padx=(0, 14))
        tk.Radiobutton(line2, text="全部（“一样”也会变成“1样”，慎用）",
                       variable=self.mode_var, value="all").pack(side="left")

        line3 = tk.Frame(opts)
        line3.pack(fill="x", padx=8, pady=(2, 6))
        tk.Label(line3, text="处理范围（只影响文件内容）：").pack(side="left")
        kind = ttk.Combobox(line3, textvariable=self.kind_var, width=34, state="readonly",
                            values=["文档类（txt / md / csv / srt…）",
                                    "文档 + 代码文件",
                                    "所有文本文件"])
        kind.pack(side="left")

        # ---------- 按钮 ----------
        btns = tk.Frame(root)
        btns.pack(fill="x", padx=12, pady=10)
        self.btn_preview = tk.Button(btns, text="预览会改什么", width=16,
                                     font=("Microsoft YaHei UI", 11), command=self.do_preview)
        self.btn_preview.pack(side="left")
        self.btn_run = tk.Button(btns, text="开始转换", width=16,
                                 font=("Microsoft YaHei UI", 11, "bold"),
                                 bg="#2d7ff9", fg="white",
                                 activebackground="#1b64d6", activeforeground="white",
                                 command=self.do_convert)
        self.btn_run.pack(side="left", padx=10)
        self.btn_undo = tk.Button(btns, text="撤销上次转换", width=14,
                                  font=("Microsoft YaHei UI", 11), command=self.do_undo)
        self.btn_undo.pack(side="left")
        self.btn_open = tk.Button(btns, text="打开文件夹", width=12,
                                  command=self.open_folder).pack(side="right")

        # ---------- 日志 ----------
        logbox = tk.LabelFrame(root, text=" 预览 / 运行记录 ")
        logbox.pack(fill="both", expand=True, padx=12)
        wrap = tk.Frame(logbox)
        wrap.pack(fill="both", expand=True, padx=6, pady=6)
        self.text = tk.Text(wrap, wrap="word", font=("Microsoft YaHei UI", 10),
                            bg="#fbfbfb", relief="flat", height=14)
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self.text.tag_configure("title", font=("Microsoft YaHei UI", 10, "bold"),
                                foreground="#1a4a8a")
        self.text.tag_configure("ok", foreground="#1a7f37")
        self.text.tag_configure("warn", foreground="#b35900")
        self.text.tag_configure("err", foreground="#c62828")
        self.text.tag_configure("dim", foreground="#777777")
        self.text.tag_configure("old", foreground="#c62828", overstrike=True)
        self.text.tag_configure("new", foreground="#1a7f37")
        self.text.configure(state="disabled")

        # ---------- 状态栏 ----------
        bar = tk.Frame(root)
        bar.pack(fill="x", padx=12, pady=(6, 10))
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(bar, textvariable=self.status_var, anchor="w").pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="determinate", length=260)
        self.progress.pack(side="right")

        self.log("把本工具放进要处理的文件夹，或者点“浏览…”选一个文件夹。\n", "dim")
        self.log("建议先点「预览会改什么」，确认没问题再点「开始转换」。\n", "dim")
        self.log("转换前会自动备份，随时可以点「撤销上次转换」还原。\n", "dim")
        self.root.after(80, self._poll)
        self.root.after(200, self._quick_scan)

    # ---------------- 小工具 ----------------
    def log(self, text, tag=None):
        self.text.configure(state="normal")
        self.text.insert("end", text if text.endswith("\n") else text + "\n", tag)
        self.text.see("end")
        self.text.configure(state="disabled")

    def set_status(self, text):
        self.status_var.set(text)

    def set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for b in (self.btn_preview, self.btn_run, self.btn_undo):
            b.configure(state=state)

    def choose_folder(self):
        d = filedialog.askdirectory(initialdir=self.folder_var.get() or app_dir(),
                                    title="选择要处理的文件夹")
        if d:
            self.folder_var.set(os.path.normpath(d))
            self.plan = None
            self._quick_scan()

    def open_folder(self):
        folder = self.folder_var.get().strip()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("提示", "这个文件夹不存在。")

    def _options(self):
        kind_map = {
            "文档类（txt / md / csv / srt…）": "doc",
            "文档 + 代码文件": "text",
            "所有文本文件": "all",
        }
        return {
            "folder": self.folder_var.get().strip(),
            "include_sub": self.sub_var.get(),
            "do_names": self.names_var.get(),
            "do_content": self.content_var.get(),
            "aggressive": self.mode_var.get() == "all",
            "mode": kind_map.get(self.kind_var.get(), "doc"),
        }

    def _key(self, opt):
        return (opt["folder"], opt["include_sub"], opt["do_names"],
                opt["do_content"], opt["aggressive"], opt["mode"])

    def _valid_folder(self, opt):
        folder = opt["folder"]
        if not folder:
            messagebox.showwarning("提示", "请先选择文件夹。")
            return False
        if not os.path.isdir(folder):
            messagebox.showwarning("提示", "文件夹不存在：\n%s" % folder)
            return False
        if os.path.splitdrive(os.path.abspath(folder))[1] in ("\\", "/"):
            messagebox.showwarning("提示", "请不要直接选整个磁盘，选一个具体文件夹。")
            return False
        if not opt["do_names"] and not opt["do_content"]:
            messagebox.showwarning("提示", "请至少勾选“修改文件名”或“修改文件内容”中的一项。")
            return False
        return True

    def _quick_scan(self):
        opt = self._options()
        folder = opt["folder"]
        if not os.path.isdir(folder):
            return
        try:
            files = list_files(folder, opt["include_sub"],
                               opt["mode"] if opt["do_content"] else "all",
                               exclude=self._exclude())
        except OSError:
            return
        self.log("\n当前文件夹：%s\n找到 %d 个可处理的文件。\n" % (folder, len(files)), "dim")

    def _exclude(self):
        here = app_dir()
        return [os.path.join(here, os.path.basename(__file__))]

    # ---------------- 任务线程 ----------------
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                self._handle(kind, payload)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _post(self, kind, payload=None):
        self.q.put((kind, payload))

    def _handle(self, kind, payload):
        if kind == "log":
            text, tag = payload
            self.log(text, tag)
        elif kind == "status":
            self.set_status(payload)
        elif kind == "progress":
            cur, total = payload
            if total:
                self.progress.configure(maximum=total, value=cur)
        elif kind == "plan":
            mode, plan, opt = payload
            self.set_busy(False)
            self.progress.configure(value=0)
            if mode == "preview":
                self._show_preview(plan)
                self.set_status("预览完成")
            else:
                self._after_scan_for_apply(plan, opt)
        elif kind == "applied":
            summary = payload
            self.set_busy(False)
            self.progress.configure(value=0)
            self.plan = None
            self.last_backup = summary["backup"]
            self.set_status("转换完成")
            self.log("\n✔ 完成：改名 %d 个，改内容 %d 个文件（共 %d 处）。\n" %
                     (summary["renamed"], summary["files"], summary["changes"]), "ok")
            self.log("备份在：%s\n" % summary["backup"], "dim")
            self.log("如果改错了，点「撤销上次转换」就能全部还原。\n", "dim")
            messagebox.showinfo("完成",
                                "转换完成！\n\n改名 %d 个\n改内容 %d 个文件（共 %d 处）\n\n"
                                "已自动备份，可点「撤销上次转换」还原。"
                                % (summary["renamed"], summary["files"], summary["changes"]))
        elif kind == "undone":
            self.set_busy(False)
            self.set_status("已撤销")
        elif kind == "error":
            self.set_busy(False)
            self.progress.configure(value=0)
            self.set_status("出错")
            self.log("\n✘ 出错：%s\n" % payload, "err")
            messagebox.showerror("出错", payload)

    def _start(self, target, *args):
        self.set_busy(True)
        t = threading.Thread(target=target, args=args, daemon=True)
        t.start()

    def do_preview(self):
        if self.busy:
            return
        opt = self._options()
        if not self._valid_folder(opt):
            return
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self.set_status("正在检查…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._start(self._scan_job, "preview", opt)

    def do_convert(self):
        if self.busy:
            return
        opt = self._options()
        if not self._valid_folder(opt):
            return
        self.set_status("正在检查…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._start(self._scan_job, "apply", opt)

    def _scan_job(self, mode, opt):
        try:
            plan = collect_plan(
                opt["folder"],
                include_sub=opt["include_sub"],
                do_names=opt["do_names"],
                do_content=opt["do_content"],
                aggressive=opt["aggressive"],
                mode=opt["mode"],
                exclude=self._exclude(),
            )
            self._post("plan", (mode, plan, opt))
        except Exception as exc:
            self._post("error", "%s\n%s" % (exc, traceback.format_exc(limit=3)))

    def _show_preview(self, plan):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

        if plan_is_empty(plan):
            self.log("没有发现需要修改的内容。\n", "warn")
            self.log("如果确认文件里确实有汉字数字，可以试试把「智能」换成「全部」，"
                     "或者换一下「处理范围」。\n", "dim")
            return

        self.log("扫描了 %d 个文件\n" % plan.scanned, "dim")
        if plan.renames:
            self.log("\n【文件名】要改 %d 个\n" % len(plan.renames), "title")
            limit = 60
            for item in plan.renames[:limit]:
                self.log("    %s  →  %s" % (os.path.basename(item["src"]),
                                            os.path.basename(item["dst"])), None)
            if len(plan.renames) > limit:
                self.log("    …还有 %d 个" % (len(plan.renames) - limit), "dim")
            self.log("")

        shown = 0
        if plan.contents:
            self.log("\n【文件内容】要改 %d 个文件" % len(plan.contents), "title")
            total = sum(i["count"] for i in plan.contents)
            self.log("（共 %d 处）\n" % total, "dim")
            for item in plan.contents:
                self.log("\n  ● %s（%s，%d 处）" % (item["rel"], item["enc"], item["count"]), "title")
                for line_no, old, new in item["samples"]:
                    if shown >= MAX_SAMPLES:
                        break
                    shown += 1
                    self.log("    第 %d 行：" % line_no, "dim")
                    self.log("      " + old, "old")
                    self.log("   →  " + new, "new")
                if shown >= MAX_SAMPLES:
                    self.log("\n  …预览只显示前 %d 处" % MAX_SAMPLES, "dim")
                    break
        if plan.skipped:
            self.log("\n\n【跳过】%d 个文件" % len(plan.skipped), "warn")
            for name, why in plan.skipped[:10]:
                self.log("\n  %s：%s" % (name, why), "dim")
            if len(plan.skipped) > 10:
                self.log("\n  …", "dim")
        self.log("\n\n确认没问题，就点「开始转换」。\n", "ok")

    def _after_scan_for_apply(self, plan, opt):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        if plan_is_empty(plan):
            self.log("\n没有发现需要修改的内容。\n", "warn")
            self.set_status("没有需要修改的内容")
            messagebox.showinfo("没有变化", "这个文件夹里没有发现需要修改的汉字数字。")
            return

        total = sum(i["count"] for i in plan.contents)
        msg = ("即将修改：\n\n· 文件名 %d 个\n· 文件内容 %d 个（共 %d 处）\n\n"
               "会先把原文件备份到这个文件夹里，随时可以撤销。\n\n确定开始吗？"
               % (len(plan.renames), len(plan.contents), total))
        if not messagebox.askyesno("确认转换", msg):
            self.set_status("已取消")
            return
        self.log("\n开始转换…\n", "title")
        self.set_busy(True)
        self._start(self._apply_job, plan)

    def _apply_job(self, plan):
        try:
            def cb(line):
                self._post("log", (line, None))
            backup = apply_plan(plan, log=cb)
            self._post("applied", {
                "backup": backup,
                "renamed": len(plan.renames),
                "files": len(plan.contents),
                "changes": sum(i["count"] for i in plan.contents),
            })
        except Exception as exc:
            self._post("error", "%s\n%s" % (exc, traceback.format_exc(limit=3)))

    def do_undo(self):
        if self.busy:
            return
        folder = self.folder_var.get().strip()
        if not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择要撤销的文件夹。")
            return
        backups = find_backups(folder)
        if not backups:
            messagebox.showinfo("没有备份", "这个文件夹里没有找到转换备份，没法撤销。")
            return
        latest = backups[0]
        if not messagebox.askyesno(
                "撤销上次转换",
                "要把这个文件夹还原到上一次转换之前吗？\n\n备份：%s\n\n"
                "（会还原被改的文件名和文件内容）" % os.path.basename(latest)):
            return
        self.set_busy(True)
        self.set_status("正在撤销…")

        def job():
            try:
                def cb(line):
                    self._post("log", (line, None))
                n = undo_backup(latest, log=cb)
                self._post("log", ("✔ 已撤销，还原 %d 项。\n" % n, "ok"))
                self._post("undone", None)
            except Exception as exc:
                self._post("error", "%s\n%s" % (exc, traceback.format_exc(limit=3)))

        self._start(job)


# ---------------------------------------------------------------- 命令行

def cli_convert(argv):
    folder = None
    aggressive = "--all" in argv
    do_names = "--no-rename" not in argv
    do_content = "--no-content" not in argv
    include_sub = "--sub" in argv
    quiet = "--quiet" in argv
    mode = "all" if "--any-file" in argv else "doc"
    for i, a in enumerate(argv):
        if a == "--convert" and i + 1 < len(argv):
            folder = argv[i + 1]
    if not folder:
        print("用法: python 汉字数字转换.py --convert <文件夹> [--all] [--any-file] "
              "[--no-rename] [--no-content] [--sub] [--quiet]")
        return 2
    if not os.path.isdir(folder):
        print("文件夹不存在: %s" % folder)
        return 2
    plan = collect_plan(folder, include_sub=include_sub, do_names=do_names,
                        do_content=do_content, aggressive=aggressive, mode=mode)
    if plan_is_empty(plan):
        print("没有发现需要修改的内容。")
        return 0
    if not quiet:
        for item in plan.renames:
            print("[改名] %s -> %s" % (item["rel"], os.path.basename(item["dst"])))
        for item in plan.contents:
            print("[内容] %s (%d 处)" % (item["rel"], item["count"]))
            for line_no, old, new in item["samples"][:5]:
                print("       第%d行: %s  =>  %s" % (line_no, old, new))
    backup = apply_plan(plan, log=(print if not quiet else None))
    print("完成。备份: %s" % backup)
    return 0


def gui_main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("vista")
    except Exception:
        pass
    app = ConverterApp(root)

    def on_error(exc_type, exc_value, exc_tb):
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        try:
            with open(os.path.join(app_dir(), "错误日志.txt"), "a", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n" + detail + "\n")
        except Exception:
            pass
        try:
            messagebox.showerror("程序出错", detail[-1500:])
        except Exception:
            pass

    root.report_callback_exception = on_error
    sys.excepthook = on_error
    root.mainloop()


def main():
    args = sys.argv[1:]
    if "--convert" in args:
        sys.exit(cli_convert(args))
    gui_main()


if __name__ == "__main__":
    main()
