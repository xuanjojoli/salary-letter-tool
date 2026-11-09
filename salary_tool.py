#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量信件生成工具（调薪信 / 晋升信 / 录用信 通用）
=====================================================

功能：
  1. 读取 Excel 数据源，逐行遍历员工数据；
  2. 根据「模板编号」列自动选择对应 .docx 模板；
  3. 把整行字段填充进 Word 模板（占位符写法 {{字段名}}，字段名与 Excel 表头一致）；
  4. 每个员工导出一个 PDF，按可配置的命名规则命名；
  5. 全部 PDF 输出到同一个文件夹；
  6. 支持「场景」保存 / 切换：换一套模板 + 换一份 Excel + 改一下命名规则即可复用于其他信件，
     无需修改任何代码。

依赖：python-docx / docxtpl / openpyxl / pywin32（Word 转 PDF）
Python 版本：3.11+
"""

import os
import re
import sys
import json
import queue
import threading
import tempfile
import traceback
from datetime import datetime, date

import openpyxl
from docxtpl import DocxTemplate

try:
    import pythoncom
    import win32com.client
    HAVE_WORD = True
except Exception:
    HAVE_WORD = False

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
WD_FORMAT_PDF = 17            # Word 的 wdFormatPDF 文件格式常量
ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|]')   # Windows 文件名非法字符

DEFAULT_CONFIG = {
    "name": "调薪信件",
    "excel_file": "",
    "template_dir": "",
    "output_dir": "",
    "sheet_name": "",
    "template_number_column": "模板编号",
    "filename_pattern": "{员工编号}_{员工姓名}_2026 salary review letter",
    "date_format": "%Y-%m-%d",
}


# ---------------------------------------------------------------------------
# 通用工具函数
# ---------------------------------------------------------------------------
def safe_str(x):
    """None -> 空字符串，其余转字符串并去首尾空格。"""
    return "" if x is None else str(x).strip()


def normalize_value(v, date_format="%Y-%m-%d"):
    """把 Excel 单元格值转成适合渲染进 Word 的字符串。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime(date_format)
    if isinstance(v, date):
        return v.strftime(date_format)
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, float):
        # 12345.0 这种整型被 openpyxl 读成 float，去掉多余 .0
        if v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def normalize_number(x):
    """把「模板编号」归一化成字符串，例如 1 / "1" / "01" / 1.0 / "模板1" 都归一到 "1"。"""
    s = safe_str(x)
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        m = re.search(r'(\d+)', s)
        return m.group(1) if m else s


def sanitize_filename(name):
    """去掉 Windows 文件名里的非法字符。"""
    return ILLEGAL_CHARS.sub("_", name).strip()


def guess_template_number_column(headers):
    """从表头里猜测「模板编号」列名，优先找「模板」，避免误选「员工编号」。"""
    for h in headers:
        if h == "模板编号":
            return h
    for h in headers:
        if h and "模板" in h:
            return h
    for h in headers:
        if h and "template" in h.lower():
            return h
    for h in headers:
        if h and "编号" in h:
            return h
    return "模板编号"


# ---------------------------------------------------------------------------
# Excel 读取
# ---------------------------------------------------------------------------
def read_excel(excel_path, sheet_name=None):
    """读取 Excel，返回 (表头列表, 数据行列表, 实际工作表名)。"""
    if not os.path.isfile(excel_path):
        raise ValueError(f"Excel 文件不存在：{excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 是空的（没有任何数据行）")

    headers = [safe_str(h) for h in rows[0]]

    # 检查重复表头（重复会导致字段互相覆盖，属于数据错误，尽早暴露）
    seen = {}
    for h in headers:
        if not h:
            continue
        if h in seen:
            raise ValueError(f"Excel 表头「{h}」重复出现，请先修改 Excel 保证表头唯一")
        seen[h] = True

    data = []
    for r in rows[1:]:
        if r is None or all(c is None or safe_str(c) == "" for c in r):
            continue  # 跳过空行
        record = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            record[h] = r[i] if i < len(r) else None
        data.append(record)

    return headers, data, ws.title


def list_sheets(excel_path):
    """列出 Excel 里所有工作表名。"""
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    return wb.sheetnames


# ---------------------------------------------------------------------------
# 模板匹配
# ---------------------------------------------------------------------------
def build_template_map(template_dir):
    """扫描模板文件夹，按文件名里的数字建立 {编号字符串: 文件名} 映射。"""
    if not os.path.isdir(template_dir):
        raise ValueError(f"模板文件夹不存在：{template_dir}")

    tmap = {}
    docx_files = []
    for f in sorted(os.listdir(template_dir)):
        if not f.lower().endswith(".docx") or f.startswith("~$"):
            continue  # 跳过临时文件（Word 打开时会生成 ~$xxx.docx）
        m = re.search(r'(\d+)', os.path.basename(f))
        if m:
            docx_files.append((f, normalize_number(m.group(1))))
        else:
            docx_files.append((f, None))  # 文件名里没有数字，预览时提示

    if not docx_files:
        raise ValueError(f"模板文件夹里没有找到任何 .docx 文件：{template_dir}")

    for f, num in docx_files:
        if num is None:
            continue
        if num in tmap:
            raise ValueError(
                f"模板编号「{num}」对应了多个文件：{tmap[num]} 和 {f}，请检查模板文件名")
        tmap[num] = f

    return tmap, [f for f, _ in docx_files]


def extract_template_variables(template_path):
    """解析模板里用到的所有 {{变量名}}，失败返回 None。"""
    try:
        doc = DocxTemplate(template_path)
        doc.render({})  # 用空上下文渲染，所有变量都会成为 undeclared
        return set(doc.get_undeclared_template_variables())
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 文件名生成
# ---------------------------------------------------------------------------
def resolve_filename(pattern, record, date_format):
    """把命名规则里的 {字段名} 替换为该行对应值，并清理非法字符。"""
    def repl(m):
        key = m.group(1)
        return normalize_value(record.get(key, ""), date_format)

    name = re.sub(r'\{([^{}]+)\}', repl, pattern)
    return sanitize_filename(name)


# ---------------------------------------------------------------------------
# 核心生成流程
# ---------------------------------------------------------------------------
class Config:
    """一个「场景」的配置，可序列化到 JSON。"""
    def __init__(self, **kwargs):
        d = dict(DEFAULT_CONFIG)
        d.update(kwargs)
        for k, v in d.items():
            setattr(self, k, v)

    def to_dict(self):
        return {k: getattr(self, k) for k in DEFAULT_CONFIG}

    def save(self, path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path):
        with open(path, "r", encoding="utf-8") as f:
            return cls(**json.load(f))


def get_config_dir():
    """配置文件目录：打包成 exe 后放在 exe 同级目录，方便携带和手工编辑。"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    d = os.path.join(base, "configs")
    os.makedirs(d, exist_ok=True)
    return d


def word_error_hint(e):
    """把 Word 转 PDF 时的常见 COM 报错翻译成可操作提示。"""
    s = str(e)
    if ("命令失败" in s or "Command failed" in s.lower()
            or "发生意外" in s or "-2146824090" in s or "-2147352567" in s):
        return ("\n    常见原因：\n"
                "    ① Word 未登录/未激活（Office 365 需手动打开一次 Word 并登录账号激活）；\n"
                "    ② Word 首次启动有登录/授权弹窗，请先手动打开 Word 完成初始化；\n"
                "    ③ 文档处于「受保护的视图」（从网上下载）：对压缩包/文件夹右键→属性→勾选「解除锁定」；\n"
                "    ④ 请先在 Word 里手动「另存为 PDF」测试能否成功。")
    return ""


def convert_to_pdf_batch(items, log):
    """用一次 Word 会话把多个 docx 转成 PDF（items: [(docx路径, pdf路径), ...]）。

    返回成功生成的 PDF 路径列表。
    注意：Word 保存含空格的 PDF 文件名时可能把空格编码成 %20，
    所以先存成无空格的临时名，再用 os.replace 改名到最终名。
    """
    if not HAVE_WORD:
        raise RuntimeError("未检测到 pywin32，无法调用 Word 转 PDF。请安装 pywin32：pip install pywin32")
    if not items:
        return []

    pythoncom.CoInitialize()
    word = None
    succeeded = []
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:
            raise RuntimeError(
                "无法启动 Microsoft Word，请确认本机已安装 MS Office Word。\n"
                f"原始错误：{e}") from e

        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass
        try:
            # 强制禁用宏，避免「安全警告：宏已被禁用」之类弹窗卡住自动化
            word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass

        for idx, (docx_path, pdf_path) in enumerate(items):
            # 无空格的临时名，避免 Word 把空格写成 %20
            tmp_pdf = os.path.join(os.path.dirname(pdf_path),
                                   f"__tmp_pdf_{os.getpid()}_{idx}.pdf")
            doc = None
            try:
                doc = word.Documents.Open(docx_path, ReadOnly=True)
                # 优先 ExportAsFixedFormat（内置 PDF 引擎，比 SaveAs 更稳），失败再退回 SaveAs
                try:
                    doc.ExportAsFixedFormat(OutputFileName=tmp_pdf, ExportFormat=WD_FORMAT_PDF)
                except Exception:
                    doc.SaveAs(tmp_pdf, FileFormat=WD_FORMAT_PDF)
                doc.Close(False)
                doc = None
                if not os.path.isfile(tmp_pdf):
                    raise RuntimeError("Word 未生成 PDF 文件")
                os.replace(tmp_pdf, pdf_path)
                succeeded.append(pdf_path)
            except Exception as e:
                log(f"  转换失败 [{os.path.basename(pdf_path)}]：{e}{word_error_hint(e)}")
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                try:
                    if os.path.isfile(tmp_pdf):
                        os.remove(tmp_pdf)
                except Exception:
                    pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return succeeded


def generate_all(cfg, log=None, progress=None):
    """
    完整流程。log(msg) / progress(cur, total) 为可选回调（线程安全由调用方保证）。
    返回结果字典：{"ok": 数量, "fail": 数量, "warn": 数量, "output": [pdf路径...]}
    """
    log = log or (lambda m: None)
    progress = progress or (lambda c, t: None)

    # ---- 1. 校验输入 ----
    if not cfg.excel_file:
        raise ValueError("请先选择 Excel 数据源文件")
    if not cfg.template_dir:
        raise ValueError("请先选择 Word 模板文件夹")
    if not cfg.output_dir:
        raise ValueError("请先选择输出文件夹")

    os.makedirs(cfg.output_dir, exist_ok=True)

    # ---- 2. 读 Excel ----
    log("正在读取 Excel ...")
    headers, records, sheet = read_excel(cfg.excel_file, cfg.sheet_name or None)
    log(f"  读取到工作表「{sheet}」，共 {len(records)} 行有效员工数据")
    log(f"  表头列：{'、'.join(headers)}")

    if not records:
        raise ValueError("Excel 里没有有效员工数据行")

    # ---- 3. 模板映射 ----
    log("正在扫描模板文件夹 ...")
    tmap, docx_files = build_template_map(cfg.template_dir)
    log(f"  发现 {len(docx_files)} 份模板文件")
    for k in sorted(tmap, key=lambda x: int(x) if x.isdigit() else 10**9):
        log(f"    编号 {k} -> {tmap[k]}")

    # ---- 4. 校验「模板编号」列 ----
    tcol = cfg.template_number_column
    if not tcol or tcol not in headers:
        raise ValueError(
            f"Excel 中找不到「模板编号」列「{tcol}」。\n实际表头：{'、'.join(headers)}\n"
            f"请在界面里把「模板编号列名」改成正确表头。")

    log(f"  模板编号列：「{tcol}」")

    # ---- 5. 校验命名规则里的字段都存在 ----
    pattern_fields = set(re.findall(r'\{([^{}]+)\}', cfg.filename_pattern))
    missing_fields = [f for f in pattern_fields if f not in headers]
    if missing_fields:
        raise ValueError(
            f"命名规则里用到的字段在 Excel 表头中不存在：{'、'.join(missing_fields)}\n"
            f"请检查命名规则或 Excel 表头。")

    # ---- 6. 逐行渲染 ----
    log("正在逐行填充模板 ...")
    tmp_dir = tempfile.mkdtemp(prefix="salary_tool_")
    ok, fail, warn = 0, 0, 0
    pdf_jobs = []          # [(临时docx, 最终pdf路径)]
    used_names = set()     # 用于检测重名

    try:
        total = len(records)
        for idx, rec in enumerate(records):
            progress(idx + 1, total)
            tkey = normalize_number(rec.get(tcol))
            name = normalize_value(rec.get("员工姓名", ""), cfg.date_format)
            emp_id = normalize_value(rec.get("员工编号", ""), cfg.date_format)

            if tkey not in tmap:
                warn += 1
                log(f"  [跳过] 第 {idx + 2} 行：模板编号「{tkey}」没有对应模板文件")
                continue

            tpl_path = os.path.join(cfg.template_dir, tmap[tkey])
            # 渲染
            try:
                doc = DocxTemplate(tpl_path)
                context = {h: normalize_value(rec.get(h), cfg.date_format) for h in headers}
                doc.render(context)
                tmp_docx = os.path.join(tmp_dir, f"row_{idx:04d}.docx")
                doc.save(tmp_docx)
            except Exception as e:
                fail += 1
                log(f"  [填充失败] 第 {idx + 2} 行（{emp_id} {name}）：{e}")
                continue

            # 命名（含去重）
            pdf_name = resolve_filename(cfg.filename_pattern, rec, cfg.date_format)
            if not pdf_name:
                pdf_name = f"row_{idx:04d}"
                warn += 1
            if pdf_name in used_names:
                n = 2
                base, ext = os.path.splitext(pdf_name)
                while f"{base}_{n}{ext}" in used_names:
                    n += 1
                warn += 1
                log(f"  [提示] 第 {idx + 2} 行 PDF 重名，改为 {base}_{n}{ext}")
                pdf_name = f"{base}_{n}{ext}"
            used_names.add(pdf_name)

            pdf_jobs.append((tmp_docx, os.path.join(cfg.output_dir, pdf_name + ".pdf")))

        progress(total, total)

        # ---- 7. Word 转 PDF ----
        ok_paths = []
        if pdf_jobs:
            log(f"正在用 Word 把 {len(pdf_jobs)} 份文档转成 PDF ...")
            ok_paths = convert_to_pdf_batch(pdf_jobs, log)
            fail += (len(pdf_jobs) - len(ok_paths))
        ok = len(ok_paths)

        log("")
        log(f"========== 完成 ==========")
        log(f"  成功导出：{ok} 份 PDF")
        log(f"  失败：{fail} 行")
        log(f"  警告/跳过：{warn} 行")
        log(f"  输出文件夹：{cfg.output_dir}")
        return {"ok": ok, "fail": fail, "warn": warn, "output": ok_paths}
    finally:
        # 清理临时 docx
        try:
            for f in os.listdir(tmp_dir):
                os.remove(os.path.join(tmp_dir, f))
            os.rmdir(tmp_dir)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# GUI（tkinter）
# ---------------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("批量信件生成工具")
        self.root.geometry("880x720")
        self.root.minsize(760, 640)

        self.msg_queue = queue.Queue()
        self.busy = False

        # 场景列表
        self.config_dir = get_config_dir()
        self.scenarios = self._load_scenario_list()

        # 变量
        self.var_scenario = tk.StringVar()
        self.var_excel = tk.StringVar()
        self.var_template = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_tcol = tk.StringVar()
        self.var_pattern = tk.StringVar()
        self.var_datefmt = tk.StringVar()

        self._build_ui()
        self._apply_default()

        if not HAVE_WORD:
            self.log("⚠ 未检测到 pywin32，无法调用 Word 转 PDF，请先执行：pip install pywin32", "warn")

        self.root.after(100, self._poll_queue)

    # ---- 场景读写 ----
    def _load_scenario_list(self):
        names = []
        for f in sorted(os.listdir(self.config_dir)):
            if f.endswith(".json"):
                names.append(f[:-5])
        return names

    def _scenario_path(self, name):
        return os.path.join(self.config_dir, name + ".json")

    def _load_scenario(self, name):
        path = self._scenario_path(name)
        if os.path.isfile(path):
            cfg = Config.load(path)
            self._set_from_config(cfg)
            return cfg
        return None

    def _set_from_config(self, cfg):
        self.var_excel.set(cfg.excel_file)
        self.var_template.set(cfg.template_dir)
        self.var_output.set(cfg.output_dir)
        self.var_sheet.set(cfg.sheet_name)
        self.var_tcol.set(cfg.template_number_column)
        self.var_pattern.set(cfg.filename_pattern)
        self.var_datefmt.set(cfg.date_format)

    def _current_config(self):
        return Config(
            name=self.var_scenario.get() or "未命名",
            excel_file=self.var_excel.get().strip(),
            template_dir=self.var_template.get().strip(),
            output_dir=self.var_output.get().strip(),
            sheet_name=self.var_sheet.get().strip(),
            template_number_column=self.var_tcol.get().strip(),
            filename_pattern=self.var_pattern.get().strip(),
            date_format=self.var_datefmt.get().strip() or "%Y-%m-%d",
        )

    # ---- UI ----
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # 场景栏
        frm_scene = ttk.LabelFrame(self.root, text="场景（可保存多套，用于调薪信/晋升信/录用信等）")
        frm_scene.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(frm_scene, text="当前场景：").pack(side="left", **pad)
        self.cb_scene = ttk.Combobox(frm_scene, textvariable=self.var_scenario,
                                     values=self.scenarios, width=20)
        self.cb_scene.pack(side="left", **pad)
        self.cb_scene.bind("<<ComboboxSelected>>", lambda e: self._on_scene_selected())
        ttk.Button(frm_scene, text="加载", command=self._on_scene_selected).pack(side="left", **pad)
        ttk.Button(frm_scene, text="保存", command=self._on_save_scene).pack(side="left", **pad)
        ttk.Button(frm_scene, text="另存为", command=self._on_save_scene_as).pack(side="left", **pad)
        ttk.Button(frm_scene, text="删除", command=self._on_delete_scene).pack(side="left", **pad)

        # 文件区
        frm_file = ttk.LabelFrame(self.root, text="文件路径")
        frm_file.pack(fill="x", padx=10, pady=4)
        self._file_row(frm_file, 0, "Excel 数据源：", self.var_excel, self._pick_excel, "file")
        self._file_row(frm_file, 1, "Word 模板文件夹：", self.var_template, self._pick_template, "dir")
        self._file_row(frm_file, 2, "输出文件夹：", self.var_output, self._pick_output, "dir")

        # 工作表 + 模板编号列
        frm_opt = ttk.LabelFrame(self.root, text="选项")
        frm_opt.pack(fill="x", padx=10, pady=4)
        ttk.Label(frm_opt, text="工作表：").grid(row=0, column=0, sticky="e", **pad)
        self.cb_sheet = ttk.Combobox(frm_opt, textvariable=self.var_sheet, width=22, state="readonly")
        self.cb_sheet.grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(frm_opt, text="模板编号列名：").grid(row=1, column=0, sticky="e", **pad)
        self.cb_tcol = ttk.Combobox(frm_opt, textvariable=self.var_tcol, width=22)
        self.cb_tcol.grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(frm_opt, text="日期格式：").grid(row=0, column=2, sticky="e", **pad)
        ttk.Entry(frm_opt, textvariable=self.var_datefmt, width=18).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(frm_opt, text="（Excel 日期按此格式化，默认 %Y-%m-%d）").grid(
            row=0, column=4, sticky="w", **pad)

        # 命名规则
        frm_name = ttk.LabelFrame(self.root, text="PDF 命名规则（用 {字段名} 引用 Excel 列）")
        frm_name.pack(fill="x", padx=10, pady=4)
        ttk.Entry(frm_name, textvariable=self.var_pattern).pack(fill="x", padx=8, pady=6)

        # 按钮区
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(fill="x", padx=10, pady=6)
        self.btn_preview = ttk.Button(frm_btn, text="预览校验", command=self._on_preview)
        self.btn_preview.pack(side="left", **pad)
        self.btn_open = ttk.Button(frm_btn, text="打开输出文件夹", command=self._on_open_output)
        self.btn_open.pack(side="left", **pad)
        self.btn_run = ttk.Button(frm_btn, text="开始生成", command=self._on_run)
        self.btn_run.pack(side="right", **pad)

        self.progress = ttk.Progressbar(frm_btn, mode="determinate", length=260)
        self.progress.pack(side="right", padx=8)

        # 日志区
        frm_log = ttk.LabelFrame(self.root, text="运行日志")
        frm_log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.txt_log = tk.Text(frm_log, height=14, wrap="word", state="disabled",
                               font=("Consolas", 9))
        self.txt_log.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        sb = ttk.Scrollbar(frm_log, command=self.txt_log.yview)
        sb.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=sb.set)
        self.txt_log.tag_config("error", foreground="#c0392b")
        self.txt_log.tag_config("ok", foreground="#1e8449")
        self.txt_log.tag_config("warn", foreground="#b9770e")

    def _file_row(self, parent, row, label, var, cmd, kind):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=8, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(parent, text="选择...", command=cmd).grid(row=row, column=2, padx=8, pady=4)
        parent.columnconfigure(1, weight=1)

    def _apply_default(self):
        self.var_pattern.set(DEFAULT_CONFIG["filename_pattern"])
        self.var_tcol.set(DEFAULT_CONFIG["template_number_column"])
        self.var_datefmt.set(DEFAULT_CONFIG["date_format"])

    # ---- 文件选择 ----
    def _pick_excel(self):
        p = filedialog.askopenfilename(
            title="选择 Excel 数据源",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if p:
            self.var_excel.set(p)
            self._load_sheet_and_columns(p)

    def _pick_template(self):
        p = filedialog.askdirectory(title="选择 Word 模板文件夹")
        if p:
            self.var_template.set(p)

    def _pick_output(self):
        p = filedialog.askdirectory(title="选择 PDF 输出文件夹")
        if p:
            self.var_output.set(p)

    def _load_sheet_and_columns(self, excel_path):
        """选中 Excel 后自动读表头，填充工作表下拉框和「模板编号列名」候选。"""
        try:
            sheets = list_sheets(excel_path)
            self.cb_sheet["values"] = sheets
            if sheets:
                self.var_sheet.set(sheets[0])
            headers, _, _ = read_excel(excel_path, sheets[0] if sheets else None)
            self.cb_tcol["values"] = headers
            guessed = guess_template_number_column(headers)
            if guessed in headers:
                self.var_tcol.set(guessed)
            self.log(f"已读取 Excel：{len(sheets)} 个工作表，{len(headers)} 列", "info")
        except Exception as e:
            self.log(f"读取 Excel 失败：{e}", "error")

    # ---- 场景操作 ----
    def _on_scene_selected(self):
        name = self.var_scenario.get().strip()
        if not name:
            return
        cfg = self._load_scenario(name)
        if cfg is None:
            self.log(f"场景「{name}」不存在", "warn")
        else:
            self.log(f"已加载场景「{name}」", "ok")

    def _refresh_scenes(self):
        self.scenarios = self._load_scenario_list()
        self.cb_scene["values"] = self.scenarios

    def _on_save_scene(self):
        name = self.var_scenario.get().strip()
        if not name:
            self._on_save_scene_as()
            return
        self._current_config().save(self._scenario_path(name))
        self._refresh_scenes()
        self.log(f"场景「{name}」已保存", "ok")

    def _on_save_scene_as(self):
        name = self.var_scenario.get().strip() or "新场景"
        # 用 simpledialog 输入名字
        from tkinter import simpledialog
        name = simpledialog.askstring("另存为场景", "输入场景名称：", initialvalue=name)
        if not name:
            return
        name = sanitize_filename(name).strip()
        if not name:
            return
        self.var_scenario.set(name)
        self._current_config().save(self._scenario_path(name))
        self._refresh_scenes()
        self.log(f"场景「{name}」已保存", "ok")

    def _on_delete_scene(self):
        name = self.var_scenario.get().strip()
        if not name:
            return
        path = self._scenario_path(name)
        if not os.path.isfile(path):
            self.log(f"场景「{name}」不存在", "warn")
            return
        if messagebox.askyesno("删除场景", f"确定删除场景「{name}」吗？"):
            os.remove(path)
            self.var_scenario.set("")
            self._refresh_scenes()
            self.log(f"场景「{name}」已删除", "warn")

    # ---- 预览校验 ----
    def _on_preview(self):
        if self.busy:
            return
        cfg = self._current_config()
        self.log("========== 预览校验 ==========", "info")
        try:
            if not cfg.excel_file:
                raise ValueError("请先选择 Excel 数据源")
            headers, records, sheet = read_excel(cfg.excel_file, cfg.sheet_name or None)
            self.log(f"Excel「{sheet}」：{len(records)} 行有效数据，表头 {len(headers)} 列")

            if not cfg.template_dir:
                raise ValueError("请先选择模板文件夹")
            tmap, docx_files = build_template_map(cfg.template_dir)
            self.log(f"模板：{len(docx_files)} 份文件")
            for k in sorted(tmap, key=lambda x: int(x) if x.isdigit() else 10**9):
                self.log(f"  编号 {k} -> {tmap[k]}")

            # 模板编号列是否存在
            tcol = cfg.template_number_column
            self.log(f"  使用模板编号列：「{tcol}」", "info")
            if tcol not in headers:
                self.log(f"⚠ 「模板编号列」{tcol} 在表头中不存在（实际：{'、'.join(headers)}）", "warn")
            else:
                used = {normalize_number(r.get(tcol)) for r in records}
                missing = sorted([u for u in used if u not in tmap], key=lambda x: x)
                if missing:
                    self.log(f"⚠ 数据里用到的模板编号在模板里缺失：{missing}", "warn")
                else:
                    self.log("✓ 所有模板编号都有对应文件", "ok")

            # 命名规则字段
            pf = set(re.findall(r'\{([^{}]+)\}', cfg.filename_pattern))
            missing = [f for f in pf if f not in headers]
            if missing:
                self.log(f"⚠ 命名规则用到的字段在表头中不存在：{missing}", "warn")
            else:
                self.log(f"✓ 命名规则字段齐全：{sorted(pf)}", "ok")

            # 每个模板的占位符 vs 表头
            for f in docx_files:
                vars_ = extract_template_variables(os.path.join(cfg.template_dir, f))
                if vars_ is None:
                    self.log(f"⚠ 模板 {f} 无法解析（可能是复杂模板），跳过校验", "warn")
                    continue
                in_tpl_not_excel = [v for v in sorted(vars_) if v not in headers]
                if in_tpl_not_excel:
                    self.log(f"⚠ 模板 {f} 用到但 Excel 表头缺失的字段：{in_tpl_not_excel}", "warn")
                else:
                    self.log(f"✓ 模板 {f} 字段与表头匹配（{len(vars_)} 个占位符）", "ok")
            self.log("========== 校验结束 ==========", "info")
        except Exception as e:
            self.log(f"预览校验出错：{e}", "error")

    # ---- 开始生成 ----
    def _on_run(self):
        if self.busy:
            return
        cfg = self._current_config()
        if not cfg.excel_file or not cfg.template_dir or not cfg.output_dir:
            messagebox.showwarning("提示", "请先选择 Excel 数据源、Word 模板文件夹、输出文件夹")
            return
        self._set_busy(True)
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        self.log("========== 开始生成 ==========", "info")
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _worker(self, cfg):
        def log(msg, tag="info"):
            self.msg_queue.put(("log", (msg, tag)))
        def progress(cur, total):
            self.msg_queue.put(("progress", (cur, total)))
        try:
            generate_all(cfg, log=log, progress=progress)
            self.msg_queue.put(("log", ("全部流程执行完毕。", "ok")))
        except Exception as e:
            self.msg_queue.put(("log", (f"流程中止：{e}", "error")))
            self.msg_queue.put(("log", (traceback.format_exc(), "error")))
        finally:
            self.msg_queue.put(("done", None))

    def _set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for b in (self.btn_run, self.btn_preview, self.btn_open):
            b.configure(state=state)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    msg, tag = payload
                    self.log(msg, tag)
                elif kind == "progress":
                    cur, total = payload
                    if total:
                        self.progress["maximum"] = total
                        self.progress["value"] = cur
                    else:
                        self.progress["value"] = 0
                elif kind == "done":
                    self._set_busy(False)
                    self.progress["value"] = self.progress["maximum"]
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---- 日志 ----
    def log(self, msg, tag="info"):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _on_open_output(self):
        d = self.var_output.get().strip()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo("提示", "输出文件夹不存在或未选择")


def main():
    root = tk.Tk()
    try:
        # 让界面在高 DPI 下更清晰
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
