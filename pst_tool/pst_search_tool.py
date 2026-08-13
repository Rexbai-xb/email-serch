
# -*- coding: utf-8 -*-
"""
本地 Outlook PST 邮件搜索工具 v3
依赖：pip install libpff-python-windows openpyxl
"""

import os, sys, threading, traceback, datetime, re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pypff
except ImportError:
    pypff = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

# ── 主题标准化前缀列表（来自检索设置表）──────────────────────────────
_NOISE = [
    '回复:', '回复：', '回复_', '回复-', '回复 ', '回复\u3000',
    '转发:', '转发：', '转发_', '转发-', '转发 ', '转发\u3000',
    'RE:', 'RE：', 'RE_', 'RE-', 'RE ',
    'Re:', 'Re：', 'Re_', 'Re-', 'Re ',
    'FW:', 'FW：', 'FW_', 'FW-', 'FW ',
    'Fw:', 'Fw：', 'Fw_', 'Fw-', 'Fw ',
    '答复:', '答复：', '答复_', '答复-', '答复 ', '答复\u3000',
    '更新:', '更新：', '更新_', '更新-', '更新 ', '更新\u3000',
    'Recall:', 'Recall：', 'Recall_', 'Recall-', 'Recall ',
    'Update:', 'Update：', 'Update_', 'Update-', 'Update ',
    'Updated:', 'Updated：', 'Updated_', 'Updated-', 'Updated ',
    '[External]', '以此为准',
]
NOISE_PREFIXES = sorted(set(_NOISE), key=len, reverse=True)


def normalize_subject(s):
    s = (s or "").strip()
    for _ in range(20):
        changed = False
        for p in NOISE_PREFIXES:
            if len(s) >= len(p) and s[:len(p)].casefold() == p.casefold():
                s = s[len(p):].lstrip()
                changed = True
                break
        if not changed:
            break
    return s.strip()


# ── 路径工具 ──────────────────────────────────────────────────────────
def resource_path(rel):
    try:
        base = sys._MEIPASS
    except AttributeError:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)


def win_long(path):
    if os.name != "nt":
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def mkdirs(path):
    os.makedirs(win_long(path), exist_ok=True)


def open_file(path, mode):
    return open(win_long(path), mode)


# ── 文本工具 ──────────────────────────────────────────────────────────
def safe_str(v):
    if v is None:
        return ""
    if isinstance(v, bytes):
        # 尝试多种编码，修复乱码问题
        for enc in ("utf-8", "gbk", "gb2312", "big5", "latin-1"):
            try:
                return v.decode(enc)
            except Exception:
                pass
        return v.decode("utf-8", errors="replace")
    return str(v)


def safe_int(v):
    try:
        return int(v)
    except Exception:
        return 0


def sanitize_filename(name, maxlen=60):
    name = (name or "无主题").strip()
    for ch in '\\/:*?"<>|\n\r\t':
        name = name.replace(ch, "_")
    return name[:maxlen].strip() or "无主题"


def get_date(message):
    for attr in ("delivery_time", "client_submit_time", "creation_time"):
        try:
            v = getattr(message, attr, None)
            if v:
                return v
        except Exception:
            pass
    return None


# ── 邮件记录 ──────────────────────────────────────────────────────────
class MailRecord:
    def __init__(self, folder_path, message):
        self.folder_path = folder_path
        self.subject      = safe_str(getattr(message, "subject", None))
        self.norm_subject = normalize_subject(self.subject)
        self.sender       = safe_str(getattr(message, "sender_name", None))
        self.display_to   = safe_str(getattr(message, "display_to", None))
        self.display_cc   = safe_str(getattr(message, "display_cc", None))
        self.date         = get_date(message)
        self.n_attach     = safe_int(getattr(message, "number_of_attachments", 0))
        self.headers      = safe_str(getattr(message, "transport_headers", None))

        # 修复乱码：优先用 HTML body（格式保真），其次 plain text
        plain = getattr(message, "plain_text_body", None)
        html  = getattr(message, "html_body", None)
        self.plain_body = safe_str(plain)
        self.html_body  = safe_str(html)

    def matches(self, c):
        if c.subject_kw and c.subject_kw not in self.subject.lower() \
                and c.subject_kw not in self.norm_subject.lower():
            return False
        if c.from_kw and c.from_kw not in self.sender.lower() \
                and c.from_kw not in self.headers.lower():
            return False
        if c.to_kw:
            pool = (self.display_to + " " + self.display_cc + " " + self.headers).lower()
            if c.to_kw not in pool:
                return False
        if c.date_start or c.date_end:
            if self.date is None:
                return False
            d = self.date.date()
            if c.date_start and d < c.date_start:
                return False
            if c.date_end and d > c.date_end:
                return False
        return True

    def to_msg_bytes(self):
        """
        导出为 .msg 格式（Outlook 可直接打开，不会乱码）。
        实际上我们用带正确 Content-Type 和 charset 声明的 .eml，
        强制指定 UTF-8 编码，解决乱码问题。
        """
        lines = []
        def h(k, v):
            if v:
                lines.append(f"{k}: {v}")

        h("MIME-Version", "1.0")
        h("Subject", self.subject or "(无主题)")
        h("From", self.sender or "")
        h("To", self.display_to or "")
        if self.display_cc:
            h("Cc", self.display_cc)
        if self.date:
            try:
                from email.utils import format_datetime
                h("Date", format_datetime(self.date))
            except Exception:
                pass

        if self.html_body:
            lines.append('Content-Type: text/html; charset="utf-8"')
            lines.append("Content-Transfer-Encoding: quoted-printable")
            lines.append("")
            import quopri, io
            encoded = quopri.encodestring(self.html_body.encode("utf-8")).decode("ascii")
            lines.append(encoded)
        else:
            lines.append('Content-Type: text/plain; charset="utf-8"')
            lines.append("Content-Transfer-Encoding: quoted-printable")
            lines.append("")
            import quopri
            encoded = quopri.encodestring(
                (self.plain_body or "(无正文)").encode("utf-8")).decode("ascii")
            lines.append(encoded)

        return "\r\n".join(lines).encode("utf-8")


# ── 搜索条件 ──────────────────────────────────────────────────────────
class SearchCriteria:
    def __init__(self, subject_kw="", from_kw="", to_kw="",
                 date_start=None, date_end=None, folder_filter=""):
        self.subject_kw   = subject_kw.strip().lower()
        self.from_kw      = from_kw.strip().lower()
        self.to_kw        = to_kw.strip().lower()
        self.date_start   = date_start
        self.date_end     = date_end
        self.folder_filter = folder_filter.strip()   # PST 内子文件夹名

    def has_any(self):
        return bool(self.subject_kw or self.from_kw or self.to_kw
                    or self.date_start or self.date_end)


# ── PST 遍历 ──────────────────────────────────────────────────────────
def iter_pst(pst_path, folder_filter="", progress_cb=None):
    if pypff is None:
        raise RuntimeError("未安装 pypff 库，请执行：pip install libpff-python-windows")
    pst_path = os.path.normpath(os.path.abspath(pst_path))
    pf = pypff.file()
    pf.open(pst_path)
    root = pf.get_root_folder()
    scanned = [0]

    def walk(folder, parts):
        for msg in folder.sub_messages:
            scanned[0] += 1
            if progress_cb and scanned[0] % 25 == 0:
                progress_cb(scanned[0])
            yield ("/".join(parts) or "根目录", msg)
        for sub in folder.sub_folders:
            name = safe_str(getattr(sub, "name", None)) or "未命名"
            yield from walk(sub, parts + [name])

    def walk_filtered(folder, parts):
        """只遍历名称匹配 folder_filter 的子文件夹"""
        for sub in folder.sub_folders:
            name = safe_str(getattr(sub, "name", None)) or "未命名"
            if name == folder_filter:
                yield from walk(sub, parts + [name])
            else:
                yield from walk_filtered(sub, parts + [name])

    try:
        if folder_filter:
            yield from walk_filtered(root, [])
        else:
            yield from walk(root, [])
    finally:
        if progress_cb:
            progress_cb(scanned[0])
        pf.close()


def run_search(pst_path, criteria, status_cb=None):
    matched, total = [], 0

    def prog(n):
        if status_cb:
            status_cb(f"正在扫描第 {n} 封邮件 ...")

    for fp, msg in iter_pst(pst_path, criteria.folder_filter, prog):
        total += 1
        try:
            rec = MailRecord(fp, msg)
        except Exception:
            continue
        if not criteria.has_any() or rec.matches(criteria):
            matched.append(rec)

    return matched, total


# ── 导出 ──────────────────────────────────────────────────────────────
def export_results(matched, total, output_dir, export_eml=True, status_cb=None):
    if openpyxl is None:
        raise RuntimeError("未安装 openpyxl，请执行：pip install openpyxl")

    mkdirs(output_dir)
    eml_dir = os.path.join(output_dir, "匹配邮件")
    if export_eml:
        mkdirs(eml_dir)

    # 按标准化主题分组
    groups, group_order = {}, []
    for rec in matched:
        key = rec.norm_subject or "(无主题)"
        if key not in groups:
            groups[key] = []
            group_order.append(key)
        groups[key].append(rec)

    for key in groups:
        groups[key].sort(key=lambda r: (r.date is None, r.date))

    # 文件夹名 = 主题本身（不加序号）
    folder_names = {}
    used = {}
    for key in group_order:
        base = sanitize_filename(key, 60)
        cnt = used.get(base, 0)
        used[base] = cnt + 1
        folder_names[key] = base if cnt == 0 else f"{base}_{cnt}"

    # ── Excel ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # 【Sheet1】汇总表（参照截图：序号/主题/接收时间/最新邮件时间/邮件发件人）
    ws = wb.active
    ws.title = "检索主题邮件"
    hdr = ["序号", "主题", "接收时间", "最新邮件时间", "邮件发件人"]
    ws.append(hdr)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    for g_idx, key in enumerate(group_order, 1):
        recs = groups[key]
        dated = [r.date for r in recs if r.date]
        earliest = min(dated).strftime("%Y-%m-%d %H:%M") if dated else ""
        latest   = max(dated).strftime("%Y-%m-%d %H:%M") if dated else ""
        senders  = "; ".join(sorted({r.sender for r in recs if r.sender}))

        ws.append([g_idx, key, earliest, latest, senders])
        row = ws.max_row

        # 主题列加超链接（直接打开文件夹，不弹确认）
        if export_eml:
            folder_abs = os.path.abspath(os.path.join(eml_dir, folder_names[key]))
            link_url = "file:///" + folder_abs.replace("\\", "/")
            subj_cell = ws.cell(row=row, column=2)
            subj_cell.hyperlink = link_url
            subj_cell.font = Font(color="0563C1", underline="single")

    for col, w in zip("ABCDE", [6, 44, 18, 18, 30]):
        ws.column_dimensions[col].width = w

    # 【Sheet2】明细表
    ws2 = wb.create_sheet("邮件明细")
    ws2.append(["序号", "主题", "发件人", "收件人", "日期", "所在PST文件夹", "附件数"])
    for cell in ws2[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    for idx, rec in enumerate(matched, 1):
        if status_cb and idx % 20 == 0:
            status_cb(f"正在导出第 {idx}/{len(matched)} 封邮件 ...")

        if export_eml:
            key = rec.norm_subject or "(无主题)"
            sub_abs = os.path.join(eml_dir, folder_names[key])
            try:
                mkdirs(sub_abs)
                # 文件名：时间_主题.eml（参照截图）
                ts = rec.date.strftime("%Y%m%d_%H%M%S") if rec.date else "无日期"
                fname = f"{ts}_{sanitize_filename(rec.norm_subject, 40)}.eml"
                fpath = os.path.join(sub_abs, fname)
                with open_file(fpath, "wb") as f:
                    f.write(rec.to_msg_bytes())
            except Exception as e:
                pass  # 不中断，继续下一封

        ws2.append([
            idx, rec.subject, rec.sender, rec.display_to,
            rec.date.strftime("%Y-%m-%d %H:%M:%S") if rec.date else "",
            rec.folder_path, rec.n_attach,
        ])

    for col, w in zip("ABCDEFG", [6, 36, 22, 26, 20, 22, 8]):
        ws2.column_dimensions[col].width = w

    # 【Sheet3】统计
    ws3 = wb.create_sheet("统计汇总")
    rate = (len(matched) / total * 100) if total else 0
    for r in [
        ("生成时间",   datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("PST邮件总数", total),
        ("匹配邮件数",  len(matched)),
        ("匹配率",      f"{rate:.2f}%"),
        ("识别主题数",  len(group_order)),
    ]:
        ws3.append(r)
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 28

    report = os.path.join(
        output_dir,
        f"邮件搜索报表_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(win_long(report))
    return report, rate, len(group_order)


# ── 界面 ──────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook PST 邮件搜索工具")
        self.geometry("680x680")
        self.resizable(False, False)

        self.pst_path_var    = tk.StringVar()
        self.folder_var      = tk.StringVar()   # 新增：PST内子文件夹筛选
        self.output_dir_var  = tk.StringVar()
        self.subject_var     = tk.StringVar()
        self.from_var        = tk.StringVar()
        self.to_var          = tk.StringVar()
        self.date_start_var  = tk.StringVar()
        self.date_end_var    = tk.StringVar()
        self.export_eml_var  = tk.BooleanVar(value=True)
        self.status_var      = tk.StringVar(value="就绪")

        self.logo_img = None
        try:
            self.logo_img = tk.PhotoImage(
                file=resource_path(os.path.join("assets", "logo.png")))
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # ── 顶部 logo ──
        frm_hdr = ttk.Frame(self)
        frm_hdr.pack(fill="x", padx=14, pady=(10, 4))
        if self.logo_img:
            ttk.Label(frm_hdr, image=self.logo_img).pack(side="left")
        ttk.Label(frm_hdr, text="Outlook PST 邮件搜索工具",
                  font=("Microsoft YaHei UI", 11, "bold"),
                  foreground="#3c3c3c").pack(side="right", padx=4)

        # ── 第一步：文件选择 ──
        frm_file = ttk.LabelFrame(self, text="第一步：选择 PST 文件与输出文件夹")
        frm_file.pack(fill="x", **pad)

        ttk.Label(frm_file, text="PST 文件：").grid(
            row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_file, textvariable=self.pst_path_var, width=52).grid(
            row=0, column=1, pady=5)
        ttk.Button(frm_file, text="浏览...", command=self.choose_pst).grid(
            row=0, column=2, padx=6)

        # ── 新增：PST 内子文件夹 ──
        ttk.Label(frm_file, text="指定子文件夹：").grid(
            row=1, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_file, textvariable=self.folder_var, width=52).grid(
            row=1, column=1, pady=5, sticky="w")
        ttk.Label(frm_file,
                  text="可选，填写 PST 内子文件夹名（如 HDPE、重要、主送），留空则搜全部",
                  foreground="#666666").grid(
            row=2, column=0, columnspan=3, sticky="w", padx=6)

        ttk.Label(frm_file, text="结果输出到：").grid(
            row=3, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_file, textvariable=self.output_dir_var, width=52).grid(
            row=3, column=1, pady=5)
        ttk.Button(frm_file, text="浏览...", command=self.choose_output).grid(
            row=3, column=2, padx=6)

        ttk.Label(frm_file,
                  text="提示：若 PST 正被 Outlook 打开，请先关闭 Outlook 或使用副本。",
                  foreground="#a15c00").grid(
            row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 5))

        # ── 第二步：搜索条件 ──
        frm_s = ttk.LabelFrame(self, text="第二步：设置搜索条件（可任意组合，留空不限制）")
        frm_s.pack(fill="x", **pad)

        ttk.Label(frm_s, text="主题关键字：").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_s, textvariable=self.subject_var, width=38).grid(
            row=0, column=1, columnspan=2, sticky="w", pady=5)

        ttk.Label(frm_s, text="发件人包含：").grid(row=1, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_s, textvariable=self.from_var, width=38).grid(
            row=1, column=1, columnspan=2, sticky="w", pady=5)

        ttk.Label(frm_s, text="收件人/抄送：").grid(row=2, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_s, textvariable=self.to_var, width=38).grid(
            row=2, column=1, columnspan=2, sticky="w", pady=5)

        ttk.Label(frm_s, text="日期范围：").grid(row=3, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(frm_s, textvariable=self.date_start_var, width=14).grid(
            row=3, column=1, sticky="w", pady=5)
        ttk.Label(frm_s, text="至").grid(row=3, column=1, padx=(112, 0), sticky="w")
        ttk.Entry(frm_s, textvariable=self.date_end_var, width=14).grid(
            row=3, column=2, sticky="w", pady=5)
        ttk.Label(frm_s, text="格式：YYYY-MM-DD（可只填一端）",
                  foreground="#666666").grid(row=4, column=0, columnspan=3, sticky="w", padx=6)

        ttk.Checkbutton(frm_s, text="将匹配邮件另存为 .eml 文件（按主题分子文件夹）",
                        variable=self.export_eml_var).grid(
            row=5, column=0, columnspan=3, sticky="w", padx=6, pady=(6, 4))

        # ── 搜索按钮 ──
        self.search_btn = ttk.Button(self, text="开始搜索并生成报表",
                                     command=self.start_search)
        self.search_btn.pack(pady=8)

        # ── 进度区域 ──
        frm_prog = ttk.LabelFrame(self, text="进度与结果")
        frm_prog.pack(fill="both", expand=True, **pad)

        self.progress = ttk.Progressbar(frm_prog, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=8)

        ttk.Label(frm_prog, textvariable=self.status_var,
                  wraplength=600, justify="left").pack(anchor="w", padx=10, pady=2)

        self.result_text = tk.Text(frm_prog, height=7, wrap="word")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=6)
        self.result_text.configure(state="disabled")

        # ── 版权 ──
        ttk.Label(self,
                  text="Copyright © 2026 CTCI Beijing Co., Ltd.",
                  foreground="#999999",
                  font=("Arial", 8)).pack(side="bottom", pady=(2, 6))

    # ── 交互 ──────────────────────────────────────────────────────────
    def choose_pst(self):
        path = filedialog.askopenfilename(
            title="选择 PST 文件",
            filetypes=[("Outlook 数据文件", "*.pst"), ("所有文件", "*.*")])
        if path:
            path = os.path.normpath(path)
            self.pst_path_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(
                    os.path.join(os.path.dirname(path), "搜索结果"))

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
            raise ValueError(f"日期格式不正确：{s}，请用 YYYY-MM-DD")

    def set_status(self, text):
        self.status_var.set(text)
        self.update_idletasks()

    def append_result(self, text):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", text + "\n")
        self.result_text.configure(state="disabled")
        self.result_text.see("end")

    def start_search(self):
        pst_path   = self.pst_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not pst_path or not os.path.isfile(pst_path):
            messagebox.showerror("错误", "请先选择有效的 PST 文件")
            return
        if not output_dir:
            messagebox.showerror("错误", "请选择结果输出文件夹")
            return
        if pypff is None:
            messagebox.showerror("缺少依赖",
                "未检测到 pypff 库。\n请先执行：pip install libpff-python-windows")
            return
        if openpyxl is None:
            messagebox.showerror("缺少依赖",
                "未检测到 openpyxl 库。\n请先执行：pip install openpyxl")
            return

        try:
            date_start = self.parse_date(self.date_start_var.get())
            date_end   = self.parse_date(self.date_end_var.get())
        except ValueError as e:
            messagebox.showerror("日期格式错误", str(e))
            return

        criteria = SearchCriteria(
            subject_kw    = self.subject_var.get(),
            from_kw       = self.from_var.get(),
            to_kw         = self.to_var.get(),
            date_start    = date_start,
            date_end      = date_end,
            folder_filter = self.folder_var.get(),
        )

        self.search_btn.configure(state="disabled")
        self.progress.start(12)
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.configure(state="disabled")

        threading.Thread(
            target=self._run,
            args=(pst_path, output_dir, criteria, self.export_eml_var.get()),
            daemon=True,
        ).start()

    def _run(self, pst_path, output_dir, criteria, export_eml):
        try:
            self.set_status("正在扫描 PST 文件，请稍候...")
            matched, total = run_search(pst_path, criteria, self.set_status)

            self.set_status(f"扫描完成，共 {total} 封，匹配 {len(matched)} 封，正在生成报表...")
            report, rate, threads = export_results(
                matched, total, output_dir, export_eml, self.set_status)

            self.set_status("完成！")
            self.append_result(f"PST 文件：{pst_path}")
            self.append_result(f"扫描总数：{total}　匹配：{len(matched)}　匹配率：{rate:.2f}%")
            self.append_result(f"识别主题（线程）数：{threads}")
            self.append_result(f"报表：{report}")
            if export_eml:
                self.append_result(
                    f"邮件导出到：{os.path.join(output_dir, '匹配邮件')}")

            messagebox.showinfo("搜索完成",
                f"扫描 {total} 封，匹配 {len(matched)} 封，匹配率 {rate:.2f}%\n"
                f"识别主题数：{threads}\n报表已保存到：\n{output_dir}")
        except Exception as e:
            traceback.print_exc()
            self.set_status("出现错误")
            messagebox.showerror("执行出错",
                f"{e}\n\n请确认 PST 未被 Outlook 占用，且已安装所需依赖库。")
        finally:
            self.progress.stop()
            self.search_btn.configure(state="normal")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
