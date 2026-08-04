# -*- coding: utf-8 -*-
"""
本地 Outlook PST 邮件搜索工具
=================================
功能：
1. 选择本地 .pst 文件，直接解析（不依赖 Outlook 运行）
2. 按 主题关键字 / 发件人 / 收件人 / 日期区间 进行组合搜索
3. 生成 Excel 报表（含匹配率统计）
4. 将匹配到的邮件另存为 .eml 文件，放入指定文件夹

依赖：
    pip install libpff-python-windows openpyxl

作者：Claude 生成，供内部使用
"""

import os
import sys
import threading
import traceback
import datetime
from email.message import EmailMessage
from email.utils import format_datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pypff
except ImportError:
    pypff = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    openpyxl = None


# ----------------------------------------------------------------------
# 核心：PST 解析与搜索逻辑
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# 主题标准化：去除"回复:"/"转发:"/"RE:"/"[External]"等前缀，
# 让同一个邮件会话（线程）能被识别为"同一主题"，而不是因为回复/转发
# 加了前缀就被当成不同的主题。规则取自用户提供的《邮件检索设置》表。
# ----------------------------------------------------------------------

_RAW_SUBJECT_NOISE_PREFIXES = [
    '回复:', '回复：', '回复_', '回复-', '回复 ', '回复\u3000',
    '转发:', '转发：', '转发_', '转发-', '转发 ', '转发\u3000',
    'RE:', 'RE：', 'RE_', 'RE-', 'RE ',
    'Re:', 'Re：', 'Re_', 'Re-', 'Re ',
    'FW:', 'FW：', 'FW_', 'FW-', 'FW ',
    'Fw:', 'Fw：', 'Fw_', 'Fw-', 'Fw ',
    '以此为准',
    'Recall:', 'Recall：', 'Recall_', 'Recall-', 'Recall ',
    'Update_', 'Update:', 'Update：', 'Update-', 'Update ',
    'Updated_', 'Updated:', 'Updated：', 'Updated-', 'Updated ',
    '答复_', '答复:', '答复：', '答复-', '答复 ', '答复\u3000',
    '更新_', '更新:', '更新：', '更新-', '更新 ', '更新\u3000',
    '[External]',
]
# 按长度从长到短排序、去重，避免短前缀抢先误匹配
SUBJECT_NOISE_PREFIXES = sorted(set(_RAW_SUBJECT_NOISE_PREFIXES), key=len, reverse=True)


def normalize_subject(subject):
    """
    反复剥除主题开头的 回复/转发/RE/FW/[External] 等标记，
    直到剥不动为止，得到"核心主题"，用于识别同一封邮件的会话线程。
    例如：
        "[External] Re: 转发: SABIC FUJIAN HDPE- 配管范围"
        -> "SABIC FUJIAN HDPE- 配管范围"
    """
    s = (subject or "").strip()
    for _ in range(15):  # 最多剥15层前缀，防止极端情况死循环
        changed = False
        for prefix in SUBJECT_NOISE_PREFIXES:
            plen = len(prefix)
            if plen and len(s) >= plen and s[:plen].casefold() == prefix.casefold():
                s = s[plen:].lstrip()
                changed = True
                break
        if not changed:
            break
    return s or (subject or "").strip()


class SearchCriteria:
    """搜索条件"""
    def __init__(self, subject_kw="", from_kw="", to_kw="",
                 date_start=None, date_end=None):
        self.subject_kw = subject_kw.strip().lower()
        self.from_kw = from_kw.strip().lower()
        self.to_kw = to_kw.strip().lower()
        self.date_start = date_start  # datetime.date 或 None
        self.date_end = date_end      # datetime.date 或 None

    def has_any(self):
        return bool(self.subject_kw or self.from_kw or self.to_kw
                    or self.date_start or self.date_end)


class MailRecord:
    """一封邮件的提取结果，用于报表和导出"""
    def __init__(self, folder_path, message):
        self.folder_path = folder_path
        self.subject = safe_str(getattr(message, "subject", None))
        self.normalized_subject = normalize_subject(self.subject)
        self.sender_name = safe_str(getattr(message, "sender_name", None))
        self.display_to = safe_str(getattr(message, "display_to", None))
        self.display_cc = safe_str(getattr(message, "display_cc", None))
        self.date = get_message_date(message)
        self.num_attachments = safe_int(getattr(message, "number_of_attachments", 0))
        self.plain_body = safe_str(getattr(message, "plain_text_body", None))
        self.html_body = safe_str(getattr(message, "html_body", None))
        self.transport_headers = safe_str(getattr(message, "transport_headers", None))

    def matches(self, criteria: SearchCriteria):
        if criteria.subject_kw and criteria.subject_kw not in self.subject.lower() \
                and criteria.subject_kw not in self.normalized_subject.lower():
            return False
        if criteria.from_kw and criteria.from_kw not in self.sender_name.lower() \
                and criteria.from_kw not in self.transport_headers.lower():
            return False
        if criteria.to_kw:
            combined_to = (self.display_to + " " + self.display_cc + " "
                           + self.transport_headers).lower()
            if criteria.to_kw not in combined_to:
                return False
        if criteria.date_start or criteria.date_end:
            if self.date is None:
                return False
            d = self.date.date()
            if criteria.date_start and d < criteria.date_start:
                return False
            if criteria.date_end and d > criteria.date_end:
                return False
        return True

    def to_eml_bytes(self):
        """把邮件内容组装成标准 .eml（可用 Outlook / 其他邮件客户端直接打开）"""
        msg = EmailMessage()
        msg["Subject"] = self.subject or "(无主题)"
        msg["From"] = self.sender_name or "(未知发件人)"
        msg["To"] = self.display_to or "(未知收件人)"
        if self.display_cc:
            msg["Cc"] = self.display_cc
        if self.date:
            try:
                msg["Date"] = format_datetime(self.date)
            except Exception:
                pass
        body = self.plain_body
        if not body and self.html_body:
            body = "(仅有HTML正文，请见下方 html 部分)"
        msg.set_content(body or "(无正文)")
        if self.html_body:
            try:
                msg.add_alternative(self.html_body, subtype="html")
            except Exception:
                pass
        return bytes(msg)


def resource_path(relative_path):
    """兼容直接运行 .py 和打包成 exe 后（PyInstaller）两种情况下查找资源文件"""
    try:
        base_path = sys._MEIPASS  # PyInstaller 打包后解压的临时目录
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def win_long_path(path):
    """
    给绝对路径加上 Windows 长路径前缀 \\\\?\\，绕过经典的 260 字符路径长度限制。
    常见于 OneDrive 同步目录 + 长中文主题文件夹名 + 长文件名 叠加导致超限的场景。
    非 Windows 系统直接原样返回。
    """
    if os.name != "nt":
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):  # UNC 网络路径
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def robust_makedirs(path):
    """创建目录，自动处理超长路径"""
    os.makedirs(win_long_path(path), exist_ok=True)


def robust_open(path, mode):
    """打开文件，自动处理超长路径"""
    return open(win_long_path(path), mode)


def safe_str(v):
    if v is None:
        return ""
    if isinstance(v, bytes):
        try:
            return v.decode("utf-8", errors="ignore")
        except Exception:
            return ""
    return str(v)


def safe_int(v):
    try:
        return int(v)
    except Exception:
        return 0


def get_message_date(message):
    """优先取送达时间，其次取发送时间"""
    for attr in ("delivery_time", "client_submit_time", "creation_time"):
        try:
            v = getattr(message, attr, None)
        except Exception:
            v = None
        if v:
            return v
    return None


def sanitize_filename(name, max_len=80):
    name = name or "无主题"
    bad = '\\/:*?"<>|\n\r\t'
    for ch in bad:
        name = name.replace(ch, "_")
    name = name.strip()
    if not name:
        name = "无主题"
    return name[:max_len]


def iter_pst_messages(pst_path, progress_cb=None):
    """
    遍历 PST 文件中的所有邮件。
    progress_cb(scanned_count) 用于回调更新进度。
    yields (folder_path_str, message)
    """
    if pypff is None:
        raise RuntimeError(
            "未安装 pypff 库，请先执行: pip install libpff-python-windows")

    pst_path = os.path.normpath(os.path.abspath(pst_path))
    pst_file = pypff.file()
    pst_file.open(pst_path)
    root = pst_file.get_root_folder()

    scanned = [0]

    def walk(folder, path_parts):
        # 当前文件夹下的邮件
        for message in folder.sub_messages:
            scanned[0] += 1
            if progress_cb and scanned[0] % 25 == 0:
                progress_cb(scanned[0])
            yield ("/".join(path_parts) or "根目录", message)
        # 递归子文件夹
        for sub in folder.sub_folders:
            sub_name = safe_str(getattr(sub, "name", None)) or "未命名文件夹"
            yield from walk(sub, path_parts + [sub_name])

    try:
        yield from walk(root, [])
    finally:
        if progress_cb:
            progress_cb(scanned[0])
        pst_file.close()


def run_search(pst_path, criteria: SearchCriteria, status_cb=None):
    """
    执行搜索，返回 (matched_records, total_scanned)
    status_cb(text) 用于更新界面状态文字
    """
    matched = []
    total = 0

    def progress(n):
        if status_cb:
            status_cb(f"正在扫描第 {n} 封邮件 ...")

    for folder_path, message in iter_pst_messages(pst_path, progress_cb=progress):
        total += 1
        try:
            record = MailRecord(folder_path, message)
        except Exception:
            continue
        if not criteria.has_any() or record.matches(criteria):
            matched.append(record)

    return matched, total


def export_results(matched, total, output_dir, export_eml=True, status_cb=None):
    """
    生成 Excel 报表 + 导出匹配邮件为 .eml 文件。
    同一个"标准化主题"（去掉 回复/转发/RE/FW/[External] 等前缀后相同）
    的邮件会被归入同一个子文件夹，而不是各自分散。
    返回报表文件路径
    """
    if openpyxl is None:
        raise RuntimeError("未安装 openpyxl 库，请先执行: pip install openpyxl")

    robust_makedirs(output_dir)
    eml_dir = os.path.join(output_dir, "匹配邮件")
    if export_eml:
        robust_makedirs(eml_dir)

    # 按标准化主题分组（同一会话线程的邮件归为一组）
    groups = {}          # normalized_subject -> [MailRecord, ...]
    group_order = []     # 保持首次出现的顺序
    for rec in matched:
        key = rec.normalized_subject or "(无主题)"
        if key not in groups:
            groups[key] = []
            group_order.append(key)
        groups[key].append(rec)

    # 组内按时间排序，方便看出"最早-最新"
    for key in groups:
        groups[key].sort(key=lambda r: (r.date is None, r.date))

    # 为每个分组生成一个安全的文件夹名（加序号前缀防止重名/过长）
    group_folder_names = {}
    for g_idx, key in enumerate(group_order, start=1):
        group_folder_names[key] = f"{g_idx:03d}_{sanitize_filename(key, 50)}"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "搜索结果明细"

    headers = ["序号", "主题", "标准化主题(线程)", "发件人", "收件人", "抄送",
               "日期", "所在PST文件夹", "附件数", "导出路径"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    for idx, rec in enumerate(matched, start=1):
        if status_cb and idx % 20 == 0:
            status_cb(f"正在导出第 {idx}/{len(matched)} 封匹配邮件 ...")

        export_rel_path = ""
        if export_eml:
            key = rec.normalized_subject or "(无主题)"
            sub_folder = group_folder_names[key]
            sub_folder_abs = os.path.join(eml_dir, sub_folder)
            try:
                robust_makedirs(sub_folder_abs)
                date_prefix = rec.date.strftime("%Y%m%d_%H%M%S") if rec.date else "无日期"
                eml_filename = f"{date_prefix}_{sanitize_filename(rec.sender_name, 24)}.eml"
                eml_abs_path = os.path.join(sub_folder_abs, eml_filename)
                with robust_open(eml_abs_path, "wb") as f:
                    f.write(rec.to_eml_bytes())
                export_rel_path = os.path.join("匹配邮件", sub_folder, eml_filename)
            except Exception as e:
                export_rel_path = f"(导出失败：{str(e)[:120]})"

        ws.append([
            idx,
            rec.subject,
            rec.normalized_subject,
            rec.sender_name,
            rec.display_to,
            rec.display_cc,
            rec.date.strftime("%Y-%m-%d %H:%M:%S") if rec.date else "",
            rec.folder_path,
            rec.num_attachments,
            export_rel_path,
        ])

    for col, width in zip("ABCDEFGHIJ", [6, 36, 36, 22, 26, 22, 20, 22, 8, 50]):
        ws.column_dimensions[col].width = width

    # 按主题分组的汇总表：同一线程一行，含最早/最新时间和涉及邮件数
    ws_group = wb.create_sheet("主题分组汇总")
    ws_group.append(["序号", "主题(标准化后)", "涉及邮件数", "最早邮件时间",
                      "最新邮件时间", "涉及发件人", "对应文件夹", "快速访问"])
    for cell in ws_group[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    for g_idx, key in enumerate(group_order, start=1):
        recs = groups[key]
        dated = [r.date for r in recs if r.date]
        earliest = min(dated).strftime("%Y-%m-%d %H:%M:%S") if dated else ""
        latest = max(dated).strftime("%Y-%m-%d %H:%M:%S") if dated else ""
        senders = "; ".join(sorted({r.sender_name for r in recs if r.sender_name}))
        folder_name = group_folder_names[key] if export_eml else ""
        ws_group.append([
            g_idx, key, len(recs), earliest, latest, senders, folder_name, "",
        ])
        if export_eml:
            sub_folder_abs = os.path.abspath(os.path.join(eml_dir, folder_name))
            link_cell = ws_group.cell(row=ws_group.max_row, column=8)
            link_cell.value = "打开文件夹"
            link_cell.hyperlink = "file:///" + sub_folder_abs.replace("\\", "/")
            link_cell.font = Font(color="0563C1", underline="single")
    for col, width in zip("ABCDEFGH", [6, 40, 10, 20, 20, 35, 30, 14]):
        ws_group.column_dimensions[col].width = width

    # 统计汇总 sheet
    ws2 = wb.create_sheet("统计汇总")
    match_rate = (len(matched) / total * 100) if total else 0
    summary_rows = [
        ("生成时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("PST 中邮件总数", total),
        ("匹配邮件数", len(matched)),
        ("匹配率", f"{match_rate:.2f}%"),
        ("识别出的主题(线程)数", len(group_order)),
    ]
    for r in summary_rows:
        ws2.append(r)
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 30

    report_path = os.path.join(
        output_dir, f"邮件搜索报表_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(win_long_path(report_path))
    return report_path, match_rate, len(group_order)


# ----------------------------------------------------------------------
# 图形界面
# ----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook PST 邮件搜索工具")
        self.geometry("640x600")
        self.resizable(False, False)

        self.pst_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.date_start_var = tk.StringVar()
        self.date_end_var = tk.StringVar()
        self.export_eml_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就绪")

        self.logo_img = None
        try:
            self.logo_img = tk.PhotoImage(file=resource_path(os.path.join("assets", "logo.png")))
        except Exception:
            self.logo_img = None  # 找不到logo文件时静默跳过，不影响软件其他功能

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_header = ttk.Frame(self)
        frm_header.pack(fill="x", padx=14, pady=(12, 2))
        if self.logo_img is not None:
            ttk.Label(frm_header, image=self.logo_img).pack(side="left")
        ttk.Label(frm_header, text="Outlook PST 邮件搜索工具",
                  font=("Microsoft YaHei UI", 12, "bold")).pack(side="left", padx=10)

        frm_file = ttk.LabelFrame(self, text="第一步：选择 PST 文件与输出文件夹")
        frm_file.pack(fill="x", **pad)

        ttk.Label(frm_file, text="PST 文件：").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_file, textvariable=self.pst_path_var, width=55).grid(row=0, column=1, pady=6)
        ttk.Button(frm_file, text="浏览...", command=self.choose_pst).grid(row=0, column=2, padx=6)

        ttk.Label(frm_file, text="结果输出到：").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_file, textvariable=self.output_dir_var, width=55).grid(row=1, column=1, pady=6)
        ttk.Button(frm_file, text="浏览...", command=self.choose_output).grid(row=1, column=2, padx=6)

        note = ("提示：若 PST 文件正在 Outlook 中打开，Windows 会锁定该文件导致读取失败。\n"
                "请先关闭 Outlook，或先复制一份 PST 文件后再选择该副本进行搜索。")
        ttk.Label(frm_file, text=note, foreground="#a15c00", justify="left").grid(
            row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))

        frm_search = ttk.LabelFrame(self, text="第二步：设置搜索条件（可任意组合，留空则不作为条件）")
        frm_search.pack(fill="x", **pad)

        ttk.Label(frm_search, text="主题关键字：").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_search, textvariable=self.subject_var, width=40).grid(row=0, column=1, columnspan=2, sticky="w", pady=6)

        ttk.Label(frm_search, text="发件人包含：").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_search, textvariable=self.from_var, width=40).grid(row=1, column=1, columnspan=2, sticky="w", pady=6)

        ttk.Label(frm_search, text="收件人/抄送包含：").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_search, textvariable=self.to_var, width=40).grid(row=2, column=1, columnspan=2, sticky="w", pady=6)

        ttk.Label(frm_search, text="日期范围：").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm_search, textvariable=self.date_start_var, width=16).grid(row=3, column=1, sticky="w", pady=6)
        ttk.Label(frm_search, text="至").grid(row=3, column=1, padx=(120, 0), sticky="w")
        ttk.Entry(frm_search, textvariable=self.date_end_var, width=16).grid(row=3, column=2, sticky="w", pady=6)
        ttk.Label(frm_search, text="格式：YYYY-MM-DD，例如 2026-01-01（可只填一个）",
                  foreground="#666666").grid(row=4, column=0, columnspan=3, sticky="w", padx=6)

        ttk.Checkbutton(frm_search, text="将匹配到的邮件另存为 .eml 文件（放入输出文件夹）",
                         variable=self.export_eml_var).grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=(8, 6))

        self.search_btn = ttk.Button(self, text="开始搜索并生成报表", command=self.start_search)
        self.search_btn.pack(pady=10)

        frm_status = ttk.LabelFrame(self, text="进度与结果")
        frm_status.pack(fill="both", expand=True, **pad)

        self.progress = ttk.Progressbar(frm_status, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=10)

        ttk.Label(frm_status, textvariable=self.status_var, wraplength=580, justify="left").pack(
            anchor="w", padx=10, pady=4)

        self.result_text = tk.Text(frm_status, height=8, wrap="word")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=6)
        self.result_text.configure(state="disabled")

    # ---------------- 交互逻辑 ----------------

    def choose_pst(self):
        path = filedialog.askopenfilename(
            title="选择 PST 文件", filetypes=[("Outlook 数据文件", "*.pst"), ("所有文件", "*.*")])
        if path:
            path = os.path.normpath(path)
            self.pst_path_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.join(os.path.dirname(path), "搜索结果"))

    def choose_output(self):
        path = filedialog.askdirectory(title="选择结果输出文件夹")
        if path:
            self.output_dir_var.set(os.path.normpath(path))

    def parse_date(self, s):
        s = s.strip()
        if not s:
            return None
        try:
            return datetime.datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError(f"日期格式不正确：{s}，请使用 YYYY-MM-DD 格式")

    def set_status(self, text):
        self.status_var.set(text)
        self.update_idletasks()

    def append_result(self, text):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", text + "\n")
        self.result_text.configure(state="disabled")
        self.result_text.see("end")

    def start_search(self):
        pst_path = self.pst_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not pst_path or not os.path.isfile(pst_path):
            messagebox.showerror("错误", "请先选择有效的 PST 文件")
            return
        if not output_dir:
            messagebox.showerror("错误", "请选择结果输出文件夹")
            return
        if pypff is None:
            messagebox.showerror(
                "缺少依赖", "未检测到 pypff 库。\n请先执行：pip install libpff-python-windows")
            return
        if openpyxl is None:
            messagebox.showerror("缺少依赖", "未检测到 openpyxl 库。\n请先执行：pip install openpyxl")
            return

        try:
            date_start = self.parse_date(self.date_start_var.get())
            date_end = self.parse_date(self.date_end_var.get())
        except ValueError as e:
            messagebox.showerror("日期格式错误", str(e))
            return

        criteria = SearchCriteria(
            subject_kw=self.subject_var.get(),
            from_kw=self.from_var.get(),
            to_kw=self.to_var.get(),
            date_start=date_start,
            date_end=date_end,
        )

        self.search_btn.configure(state="disabled")
        self.progress.start(12)
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.configure(state="disabled")

        thread = threading.Thread(
            target=self._run_search_thread,
            args=(pst_path, output_dir, criteria, self.export_eml_var.get()),
            daemon=True,
        )
        thread.start()

    def _run_search_thread(self, pst_path, output_dir, criteria, export_eml):
        try:
            self.set_status("正在打开并扫描 PST 文件，请稍候（大文件可能需要几分钟）...")
            matched, total = run_search(pst_path, criteria, status_cb=self.set_status)

            self.set_status(f"扫描完成，共 {total} 封邮件，匹配 {len(matched)} 封。正在生成报表...")
            report_path, match_rate, thread_count = export_results(
                matched, total, output_dir, export_eml=export_eml, status_cb=self.set_status)

            self.set_status("完成！")
            self.append_result(f"PST 文件：{pst_path}")
            self.append_result(f"扫描邮件总数：{total}")
            self.append_result(f"匹配邮件数量：{len(matched)}")
            self.append_result(f"匹配率：{match_rate:.2f}%")
            self.append_result(f"识别出的主题(线程)数：{thread_count}（已按标准化主题合并 回复/转发/RE/FW 等变体）")
            self.append_result(f"报表文件：{report_path}")
            if export_eml:
                self.append_result(f"匹配邮件已导出到：{os.path.join(output_dir, '匹配邮件')}（按主题分了子文件夹）")

            messagebox.showinfo(
                "搜索完成",
                f"共扫描 {total} 封邮件，匹配 {len(matched)} 封（合并为 {thread_count} 个主题线程），"
                f"匹配率 {match_rate:.2f}%。\n"
                f"报表与邮件已保存到：\n{output_dir}")
        except Exception as e:
            traceback.print_exc()
            self.set_status("出现错误")
            messagebox.showerror("执行出错", f"{e}\n\n请确认 PST 文件未被 Outlook 占用，且已安装所需依赖库。")
        finally:
            self.progress.stop()
            self.search_btn.configure(state="normal")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
