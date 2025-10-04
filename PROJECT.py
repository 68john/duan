# QL_Diem_HS.py
# -*- coding: utf-8 -*-
import os, sys, csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Biểu đồ
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator  # trục Y hiển thị số nguyên
import numpy as np

# Load .env if present
try:
    from dotenv import load_dotenv, find_dotenv
    def _load_dotenv_robust():
        base_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
        candidates = [
            os.path.join(base_dir, ".env"),
            find_dotenv(usecwd=True),
        ]
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            candidates.insert(1, os.path.join(sys._MEIPASS, ".env"))
        for path in candidates:
            if path and os.path.isfile(path):
                load_dotenv(path, override=False)
                break
    _load_dotenv_robust()
except Exception:
    pass

# (tuỳ chọn) Logo
try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False

# ---- Cấu hình font Unicode cho Matplotlib ----
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI']
matplotlib.rcParams['axes.unicode_minus'] = False

def resource_path(rel_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.abspath(".")
    return os.path.join(base, rel_path)

SUBJECTS = ["toan", "ly", "hoa", "van", "anh", "tin"]

# ================== MẢNG THỐNG KÊ (10 BIN) & MẢNG RAW ==================
# 10 bin: [0,1), [1,2), ..., [8,9), [9,10]
TOAN = [0]*10; LI = [0]*10; HOA = [0]*10
VAN  = [0]*10; ANH = [0]*10; TIN = [0]*10; TB = [0]*10

TOAN_RAW, LI_RAW, HOA_RAW = [], [], []
VAN_RAW, ANH_RAW, TIN_RAW = [], [], []
TB_RAW = []

# ================== HÀM TÍNH TOÁN ==================
def classify(avg: float) -> str:
    if avg >= 8.5: return "Giỏi"
    if avg >= 7.0: return "Khá"
    if avg >= 5.0: return "Trung bình"
    return "Yếu"

def wavg(scores: dict) -> float:
    return sum(scores.values())/len(scores) if scores else 0.0

def parse_score_any(txt: str) -> float:
    if txt is None: return 0.0
    s = str(txt).strip()
    if not s: return 0.0
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        v = float(s)
        if 0 <= v <= 10: return v
    except:
        pass
    return 0.0

def _bucket_index(v: float) -> int:
    try:
        x = float(v)
    except:
        x = 0.0
    if x < 0: x = 0.0
    if x > 10: x = 10.0
    return 9 if x >= 9.0 else int(x)

def update_histograms_from_raw():
    global TOAN, LI, HOA, VAN, ANH, TIN, TB
    TOAN = [0]*10; LI = [0]*10; HOA = [0]*10
    VAN  = [0]*10; ANH = [0]*10; TIN = [0]*10; TB = [0]*10
    for s in TOAN_RAW: TOAN[_bucket_index(s)] += 1
    for s in LI_RAW:   LI[_bucket_index(s)]   += 1
    for s in HOA_RAW:  HOA[_bucket_index(s)]  += 1
    for s in VAN_RAW:  VAN[_bucket_index(s)]  += 1
    for s in ANH_RAW:  ANH[_bucket_index(s)]  += 1
    for s in TIN_RAW:  TIN[_bucket_index(s)]  += 1
    for s in TB_RAW:   TB[_bucket_index(s)]   += 1

# ---------- AI helpers (đặt trước open_qna_window để tránh 'not defined') ----------
def _ai__summarize_records(records, limit_rows=15):
    if not records:
        return "Bảng hiện đang trống."
    ranks = {}
    for s in records:
        xl = (s.get("xep_loai") if isinstance(s, dict) else getattr(s, "xep_loai", "")) or \
             (s.get("xếp loại") if isinstance(s, dict) else "")
        ranks[xl] = ranks.get(xl, 0) + 1
    head = []
    for i, s in enumerate(records[:limit_rows], start=1):
        if isinstance(s, dict):
            name = s.get("ho_ten") or s.get("Họ Tên") or "?"
            clss = s.get("lop") or s.get("Lớp") or "?"
            tb = s.get("diem_tb") or s.get("Điểm TB") or 0.0
        else:
            name = getattr(s, "ho_ten", "?"); clss = getattr(s, "lop", "?"); tb = getattr(s, "diem_tb", 0.0)
        try: tb = float(tb)
        except Exception: tb = 0.0
        xl = (s.get("xep_loai") if isinstance(s, dict) else getattr(s, "xep_loai","?")) or s.get("xếp loại") if isinstance(s, dict) else "?"
        head.append(f"{i}. {name} ({clss}) - TB={tb:.2f} - XL={xl}")
    return "\n".join([
        f"Số HS hiển thị: {len(records)}",
        "Phân bố xếp loại: " + ", ".join([f"{k}={v}" for k,v in ranks.items()]) if ranks else "(không có)",
        "Một số dòng đầu:",
        *head
    ])

def _ai__ask_chatgpt(question: str, context: str = "", temperature: float = 0.2, max_output_tokens: int = 400) -> str:
    try:
        from openai import OpenAI
    except Exception:
        return "Chưa cài 'openai'. Hãy chạy: pip install --upgrade openai python-dotenv"
    import os

    api_key = os.getenv("OPENAI_API_KEY")
    model   = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    org     = os.getenv("OPENAI_ORG", None)  # optional

    if not api_key:
        return "Chưa thiết lập OPENAI_API_KEY. Đặt biến môi trường hoặc file .env."

    candidates = [model, "gpt-4o-mini", "gpt-4o"]
    client = OpenAI(api_key=api_key, organization=org) if org else OpenAI(api_key=api_key)

    system_prompt = (
        "Bạn là trợ lý AI cho phần mềm Quản lý điểm học sinh. "
        "Trả lời ngắn gọn, chính xác, bằng tiếng Việt. "
        "Nếu câu hỏi liên quan đến dữ liệu bảng đang hiển thị, hãy giải thích cách lọc/tìm/xếp hạng trong app."
    )

    last_err = None
    for md in candidates:
        try:
            rsp = client.responses.create(
                model=md,
                input=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user",
                     "content": (question if not context else f"Ngữ cảnh:\n{context}\n\nCâu hỏi: {question}")},
                ],
                max_output_tokens=max_output_tokens,
                temperature=temperature,
            )
            text = getattr(rsp, "output_text", "").strip()
            if text:
                return text
            last_err = "Không nhận được nội dung trả lời."
        except Exception as e:
            low = str(e).lower()
            if ("not found" in low or "does not exist" in low or "permission" in low or "unsupported" in low):
                last_err = f"Model `{md}` không khả dụng, thử model khác..."
                continue
            if "insufficient_quota" in low or "exceeded your current quota" in low:
                return (
                    "API key hiện không còn quota/credit để gọi.\n"
                    "- Billing: https://platform.openai.com/account/billing\n"
                    "- Usage: https://platform.openai.com/usage\n"
                    "- Nếu bạn có nhiều tổ chức (org), hãy đặt OPENAI_ORG cho đúng."
                )
            if "rate" in low and "limit" in low:
                return "Đang chạm rate limit. Hãy thử lại sau vài giây."
            if "invalid_api_key" in low or "401" in low:
                return "API key không hợp lệ hoặc đã bị thu hồi. Tạo key mới và cập nhật OPENAI_API_KEY."
            last_err = f"Lỗi gọi ChatGPT ({md}): {e}"
            break

    return last_err or "Không nhận được nội dung trả lời."

# ================== ỨNG DỤNG GUI ==================
class StudentManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý điểm học sinh (6 môn)")
        self.root.geometry("1200x760")

        # Trạng thái giao diện
        self.dark_mode = False
        self.buttons: list[tuple[tk.Button, str, str, str]] = []  # (btn, bg, fg, active_bg)

        # Dữ liệu
        self.students = []
        self.next_id = 1

        # Theme gốc
        self._setup_theme()

        # UI
        self._build_header()
        self._build_form()
        self._build_buttons()
        self._build_search()
        self._build_table()
        self._build_statusbar()
        self._set_status("Sẵn sàng.")

    # ---------- THEME ----------
    def _setup_theme(self):
        self.primary = "#1e88e5"
        self.primary_dark = "#1565c0"
        self.bg_light = "#eaf2fb"
        self.text_dark = "#1f2937"

        style = ttk.Style()
        try: style.theme_use("clam")
        except: pass

        self.root.configure(bg=self.bg_light)
        style.configure("TFrame", background=self.bg_light)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"),
                        foreground=self.primary, background=self.bg_light)
        style.configure("TLabel", background=self.bg_light, foreground=self.text_dark, font=("Segoe UI", 10))
        style.configure("TLabelframe", background=self.bg_light)
        style.configure("TLabelframe.Label", background=self.bg_light, foreground=self.text_dark,
                        font=("Segoe UI", 11, "bold"))
        style.configure("TEntry", fieldbackground="white", foreground=self.text_dark, background="white")
        style.configure("TCombobox", fieldbackground="white", foreground=self.text_dark, background="white")
        style.configure("Treeview", font=("Segoe UI", 10),
                        background="white", fieldbackground="white", foreground="#000000")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background=self.primary, foreground="white")
        style.map("Treeview.Heading", background=[("active", self.primary_dark)])

    def _apply_dark_palette(self):
        style = ttk.Style()
        bg = "#1e1e1e"; fg = "#f5f5f5"

        self.root.configure(bg=bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabelframe", background=bg)
        style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("Title.TLabel", background=bg, foreground="#90caf9")
        style.configure("TEntry", fieldbackground="#333333", foreground=fg, background="#333333")
        style.configure("TCombobox", fieldbackground="#333333", foreground=fg, background="#333333")
        style.configure("Treeview", background="#2d2d2d", fieldbackground="#2d2d2d", foreground=fg)
        style.configure("Treeview.Heading", background="#3a3a3a", foreground=fg)

        if hasattr(self, "logo_lbl"):
            try: self.logo_lbl.configure(bg=bg)
            except: pass

        if hasattr(self, "tree"):
            self.tree.tag_configure("oddrow",  background="#1f1f1f", foreground=fg)
            self.tree.tag_configure("evenrow", background="#262626", foreground=fg)

        for btn, _bg, _fg, _abg in self.buttons:
            try: btn.config(bg="#333333", fg="#f5f5f5", activebackground="#555555")
            except: pass

    def _apply_light_palette(self):
        self._setup_theme()
        self.root.configure(bg=self.bg_light)
        for w in self.root.winfo_children():
            try:
                if isinstance(w, (tk.Frame, tk.Label, tk.Button)):
                    w.configure(bg=self.bg_light)
            except:
                pass
            try:
                for c in w.winfo_children():
                    if isinstance(c, (tk.Frame, tk.Label)):
                        c.configure(bg=self.bg_light)
            except:
                pass

        if hasattr(self, "logo_lbl"):
            try: self.logo_lbl.configure(bg=self.bg_light)
            except: pass

        if hasattr(self, "tree"):
            self.tree.tag_configure("oddrow",  background="#ffffff", foreground="#000000")
            self.tree.tag_configure("evenrow", background="#e3f2fd", foreground="#000000")

        for btn, bgc, fgc, abg in self.buttons:
            try:
                btn.config(bg=bgc, fg=fgc, activebackground=abg)
            except Exception:
                pass

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self._apply_dark_palette()
        else:
            self._apply_light_palette()
        self._set_status("Đã đổi chế độ giao diện.")

    # ---------- HEADER ----------
    def _build_header(self):
        h = ttk.Frame(self.root); h.pack(fill="x", padx=12, pady=(12,6))
        try:
            path = resource_path("logo_truong.png")
            if os.path.exists(path):
                if _HAS_PIL:
                    img = Image.open(path).convert("RGBA").resize((56,56), Image.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(img)
                else:
                    self.logo_img = tk.PhotoImage(file=path)
                self.logo_lbl = tk.Label(h, image=self.logo_img, bg=self.bg_light)
                self.logo_lbl.pack(side="left", padx=(0,10))
        except Exception:
            pass
        ttk.Label(h, text="Quản lý điểm học sinh (6 môn)", style="Title.TLabel").pack(side="left")

    # ---------- FORM ----------
    def _build_form(self):
        frm = ttk.LabelFrame(self.root, text="Thông tin học sinh", padding=10)
        frm.pack(fill="x", padx=12, pady=6)

        ttk.Label(frm, text="Họ tên:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.ent_name = ttk.Entry(frm, width=28); self.ent_name.grid(row=0, column=1, padx=6, pady=4)

        ttk.Label(frm, text="Lớp:").grid(row=0, column=2, sticky="w", padx=6, pady=4)
        self.ent_class = ttk.Entry(frm, width=12); self.ent_class.grid(row=0, column=3, padx=6, pady=4)

        self.ent_scores = {}
        for i, subj in enumerate(SUBJECTS):
            r = 1 + (i // 3)
            c = (i % 3) * 2
            ttk.Label(frm, text=f"{subj.capitalize()}:").grid(row=r, column=c, sticky="w", padx=6, pady=4)
            e = ttk.Entry(frm, width=6); e.grid(row=r, column=c+1, padx=6, pady=4)
            self.ent_scores[subj] = e

        for col in range(6): frm.grid_columnconfigure(col, weight=1)

    # ---------- ICONS (tùy chọn) ----------
    def _load_icon(self, name):
        p = os.path.join("icons", name)
        try:
            if os.path.exists(p): return tk.PhotoImage(file=p)
        except Exception:
            return None
        return None

    # ---------- BUTTONS ----------
    def _build_buttons(self):
        btnf = ttk.LabelFrame(self.root, text="Chức năng", padding=10)
        btnf.pack(fill="x", padx=12, pady=6)

        ic_add    = self._load_icon("add.png")
        ic_edit   = self._load_icon("edit.png")
        ic_delete = self._load_icon("delete.png")
        ic_save   = self._load_icon("save.png")
        ic_open   = self._load_icon("open.png")
        ic_excel  = self._load_icon("excel.png")
        ic_clear  = self._load_icon("clear.png")
        ic_chart  = self._load_icon("chart.png")

        def mkbtn(text, cmd, bg, icon=None, fallback=""):
            label = f" {text}" if icon else f"{fallback} {text}"
            b = tk.Button(btnf, text=label, image=icon, compound="left",
                    command=cmd, bg=bg, fg="white",
                    activebackground=self.primary_dark,
                    font=("Segoe UI", 10, "bold"), padx=10, pady=5)
            b.pack(side="left", padx=6)
            self.buttons.append((b, bg, "white", self.primary_dark))
            return b

        mkbtn("Thêm", self.add_student, self.primary, ic_add, "➕")
        mkbtn("Sửa", self.edit_student, self.primary, ic_edit, "✏️")
        mkbtn("Xóa", self.delete_student, self.primary, ic_delete, "🗑️")
        mkbtn("Lưu CSV…", self.save_csv, "#1565c0", ic_save, "💾")
        mkbtn("Đọc CSV…", self.load_csv, "#1565c0", ic_open, "📂")
        mkbtn("Xuất Excel…", self.export_excel, "#2e7d32", ic_excel, "📊")
        mkbtn("Xem đồ thị…", self.show_charts, "#425862", ic_chart, "📈")
        mkbtn("Biểu đồ 3 khối", self.show_block_chart, "#455a64", ic_chart, "🏫")  # nút mới
        mkbtn("Hỏi AI (Q&A)", lambda: open_qna_window(self.root, get_subset_callable=self._get_visible_subset), "#6a1b9a", None, "🤖")
        mkbtn("Dark Mode", self.toggle_dark_mode, self.primary, None, "🌙")

        b_clear = tk.Button(btnf, text=(" Xóa form" if ic_clear else "🧹 Xóa form"),
                            image=ic_clear, compound="left",
                            command=self.clear_form, bg="#90caf9", fg="#0b3060",
                            activebackground="#64b5f6",
                            font=("Segoe UI", 10, "bold"), padx=10, pady=5)
        b_clear.pack(side="left", padx=16)
        self.buttons.append((b_clear, "#90caf9", "#0b3060", "#64b5f6"))

    # ---------- SEARCH ----------
    def _build_search(self):
        sf = ttk.LabelFrame(self.root, text="Tìm kiếm", padding=10)
        sf.pack(fill="x", padx=12, pady=6)

        ttk.Label(sf, text="Tên:").pack(side="left", padx=4)
        self.ent_search_name = ttk.Entry(sf, width=14); self.ent_search_name.pack(side="left")

        ttk.Label(sf, text="Lớp:").pack(side="left", padx=4)
        self.ent_search_class = ttk.Entry(sf, width=10); self.ent_search_class.pack(side="left")

        ttk.Label(sf, text="Xếp loại:").pack(side="left", padx=4)
        self.ent_search_rank = ttk.Entry(sf, width=12); self.ent_search_rank.pack(side="left")

        b_adv = tk.Button(sf, text="🔎 Tìm nâng cao",
                          command=self.advanced_search, bg="#00897b", fg="white",
                          activebackground="#00695c",
                          font=("Segoe UI", 10, "bold"), padx=10, pady=4)
        b_adv.pack(side="left", padx=8)
        self.buttons.append((b_adv, "#00897b", "white", "#00695c"))

        ttk.Label(sf, text="Theo:").pack(side="left", padx=(16,0))
        self.cmb_criteria = ttk.Combobox(
            sf, state="readonly", width=12,
            values=["Tên", "Lớp", "ID", "Xếp loại"]
        )
        self.cmb_criteria.current(0); self.cmb_criteria.pack(side="left", padx=6)
        ttk.Label(sf, text="Giá trị:").pack(side="left", padx=(10,0))
        self.ent_search = ttk.Entry(sf, width=26); self.ent_search.pack(side="left", padx=6)

        b_find = tk.Button(sf, text="🔍 Tìm",
                           command=self.search_student, bg=self.primary, fg="white",
                           activebackground=self.primary_dark,
                           font=("Segoe UI", 10, "bold"), padx=10, pady=4)
        b_find.pack(side="left", padx=6)
        self.buttons.append((b_find, self.primary, "white", self.primary_dark))

        b_all = tk.Button(sf, text="📋 Hiển thị tất cả",
                          command=self.refresh_table, bg=self.primary, fg="white",
                          activebackground=self.primary_dark,
                          font=("Segoe UI", 10, "bold"), padx=10, pady=4)
        b_all.pack(side="left", padx=6)
        self.buttons.append((b_all, self.primary, "white", self.primary_dark))

    # ---------- TABLE ----------
    def _build_table(self):
        wrap = ttk.Frame(self.root); wrap.pack(fill="both", expand=True, padx=12, pady=6)
        self.cols = ["stt","id","ho_ten","lop"] + SUBJECTS + ["diem_tb","xep_loai"]
        self.tree = ttk.Treeview(wrap, columns=self.cols, show="headings", height=18)

        header_vi = {
            "stt": "STT", "id": "ID", "ho_ten": "Họ Tên", "lop": "Lớp",
            "toan": "Toán", "ly": "Lý", "hoa": "Hóa", "van": "Văn",
            "anh": "Anh", "tin": "Tin", "diem_tb": "Điểm TB", "xep_loai": "Xếp Loại"
        }
        for key in self.cols: self.tree.heading(key, text=header_vi.get(key, key))

        self.tree.column("stt", width=60, anchor="center")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("ho_ten", width=230, anchor="w")
        self.tree.column("lop", width=110, anchor="center")
        for s in SUBJECTS: self.tree.column(s, width=85, anchor="e")
        self.tree.column("diem_tb", width=95, anchor="e")
        self.tree.column("xep_loai", width=110, anchor="center")

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1); wrap.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    # ---------- STATUS ----------
    def _build_statusbar(self):
        self.status = tk.StringVar()
        bar = ttk.Frame(self.root); bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=4)
    def _set_status(self, msg): self.status.set(msg)

    # ---------- HELPERS ----------
    def clear_form(self):
        self.ent_name.delete(0, tk.END); self.ent_class.delete(0, tk.END)
        for e in self.ent_scores.values(): e.delete(0, tk.END)
        self._set_status("Đã xóa nội dung form.")

    def _on_tree_select(self, _evt):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]
        self.ent_name.delete(0, tk.END); self.ent_name.insert(0, vals[2])
        self.ent_class.delete(0, tk.END); self.ent_class.insert(0, vals[3])
        for i, subj in enumerate(SUBJECTS, start=4):
            self.ent_scores[subj].delete(0, tk.END)
            self.ent_scores[subj].insert(0, vals[i])
        self._set_status(f"Đang chọn ID {vals[1]}.")

    def _collect_scores(self):
        return {s: parse_score_any(e.get()) for s, e in self.ent_scores.items()}

    # ---------- CRUD ----------
    def add_student(self):
        name = self.ent_name.get().strip(); lop = self.ent_class.get().strip()
        if not name or not lop:
            messagebox.showwarning("Thiếu", "Vui lòng nhập Họ tên và Lớp."); return
        scores = self._collect_scores()
        avg = wavg(scores); xl = classify(avg)
        st = {"id": self.next_id, "ho_ten": name, "lop": lop,
              **scores, "diem_tb": round(avg, 2), "xep_loai": xl}
        self.students.append(st); self.next_id += 1
        self.refresh_table(); self._set_status(f"Đã thêm HS {name} ({lop}).")

    def edit_student(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Chọn một học sinh trong bảng để sửa."); return
        sid = int(self.tree.item(sel[0])["values"][1])
        st = next((s for s in self.students if s["id"] == sid), None)
        if not st: return
        name = self.ent_name.get().strip() or st["ho_ten"]
        lop  = self.ent_class.get().strip() or st["lop"]
        new_scores = self._collect_scores()
        for k in SUBJECTS:
            if self.ent_scores[k].get().strip() == "":
                new_scores[k] = st[k]
        avg = wavg(new_scores); xl = classify(avg)
        st.update({"ho_ten": name, "lop": lop, **new_scores,
                   "diem_tb": round(avg, 2), "xep_loai": xl})
        self.refresh_table(); self._set_status(f"Đã sửa ID {sid}.")

    def delete_student(self):
        sel = self.tree.selection()
        if not sel: return
        sid = int(self.tree.item(sel[0])["values"][1])
        self.students = [s for s in self.students if s["id"] != sid]
        self.refresh_table(); self._set_status(f"Đã xóa ID {sid}.")

    # ---------- REFRESH TABLE + CẬP NHẬT HISTOGRAM ----------
    def refresh_table(self, subset=None):
        for i in self.tree.get_children(): self.tree.delete(i)

        global TOAN_RAW, LI_RAW, HOA_RAW, VAN_RAW, ANH_RAW, TIN_RAW, TB_RAW
        TOAN_RAW, LI_RAW, HOA_RAW = [], [], []
        VAN_RAW, ANH_RAW, TIN_RAW = [], [], []
        TB_RAW = []

        self.tree.tag_configure("oddrow",  background="#ffffff", foreground="#000000")
        self.tree.tag_configure("evenrow", background="#e3f2fd", foreground="#000000")

        data = subset if subset is not None else self.students
        data.sort(key=lambda s: (s["lop"], s["ho_ten"]))

        for idx, s in enumerate(data, start=1):
            row = [idx, s["id"], s["ho_ten"], s["lop"]] + \
                  [s[subj] for subj in SUBJECTS] + [f"{s['diem_tb']:.2f}", s["xep_loai"]]
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=row, tags=(tag,))

            TOAN_RAW.append(s["toan"]); LI_RAW.append(s["ly"]); HOA_RAW.append(s["hoa"])
            VAN_RAW.append(s["van"]); ANH_RAW.append(s["anh"]); TIN_RAW.append(s["tin"])
            TB_RAW.append(s["diem_tb"])

        update_histograms_from_raw()

        if self.dark_mode:
            self.tree.tag_configure("oddrow",  background="#1f1f1f", foreground="#f5f5f5")
            self.tree.tag_configure("evenrow", background="#262626", foreground="#f5f5f5")

    # ---------- SEARCH THƯỜNG ----------
    def search_student(self):
        crit = self.cmb_criteria.get(); q = self.ent_search.get().strip()
        if not q:
            self.refresh_table(); self._set_status("Hiển thị tất cả."); return

        if crit == "Tên":
            ql = q.lower(); filtered = [s for s in self.students if ql in s["ho_ten"].lower()]
        elif crit == "Lớp":
            ql = q.lower(); filtered = [s for s in self.students if ql in s["lop"].lower()]
        elif crit == "ID":
            try:
                qid = int(q); filtered = [s for s in self.students if s["id"] == qid]
            except ValueError:
                messagebox.showwarning("ID không hợp lệ", "Nhập số nguyên cho ID."); return
        elif crit == "Xếp loại":
            ql = q.lower(); filtered = [s for s in self.students if ql in s["xep_loai"].lower()]
        else:
            filtered = self.students

        self.refresh_table(filtered)
        self._set_status(f"Tìm theo {crit}='{q}' → {len(filtered)} kết quả.")

    # ---------- TÌM KIẾM NÂNG CAO ----------
    def advanced_search(self):
        qname  = self.ent_search_name.get().strip().lower()
        qclass = self.ent_search_class.get().strip().lower()
        qrank  = self.ent_search_rank.get().strip().lower()

        filtered = []
        for s in self.students:
            ok = True
            if qname  and qname  not in s["ho_ten"].lower(): ok = False
            if qclass and qclass not in s["lop"].lower():    ok = False
            if qrank  and qrank  not in s["xep_loai"].lower(): ok = False
            if ok: filtered.append(s)

        self.refresh_table(filtered)
        self._set_status(f"Tìm nâng cao → {len(filtered)} kết quả.")

    # ---------- CSV ----------
    def save_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not path: return
        header_vi = {
            "id": "ID", "ho_ten": "Họ Tên", "lop": "Lớp",
            "toan": "Toán", "ly": "Lý", "hoa": "Hóa", "van": "Văn",
            "anh": "Anh", "tin": "Tin", "diem_tb": "Điểm TB", "xep_loai": "Xếp Loại"
        }
        cols_en = ["id","ho_ten","lop"] + SUBJECTS + ["diem_tb","xep_loai"]
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL)
                w.writerow([header_vi[c] for c in cols_en])
                for s in self.students: w.writerow([s[c] for c in cols_en])
            messagebox.showinfo("Thành công", f"Đã lưu {len(self.students)} HS vào:\n{path}")
            self._set_status(f"Đã lưu CSV: {path}")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not path: return
        try:
            self.students.clear()
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                sample = f.read(4096); f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=";,")
                    delimiter = dialect.delimiter
                except Exception:
                    delimiter = ";"
                reader = csv.DictReader(f, delimiter=delimiter, quotechar='"')
                vi_to_en = {
                    "id":"id","họ tên":"ho_ten","ho ten":"ho_ten",
                    "lớp":"lop","lop":"lop","toán":"toan","toan":"toan",
                    "lý":"ly","ly":"ly","hóa":"hoa","hoa":"hoa",
                    "văn":"van","van":"van","anh":"anh","tin":"tin",
                    "điểm tb":"diem_tb","diem tb":"diem_tb",
                    "xếp loại":"xep_loai","xep loai":"xep_loai"
                }
                file_fields = reader.fieldnames or []
                field_map = {}
                for h in file_fields:
                    k = (h or "").strip().lower()
                    en = vi_to_en.get(k)
                    if not en and k in ["id","ho_ten","lop","toan","ly","hoa","van","anh","tin","diem_tb","xep_loai"]:
                        en = k
                    if en: field_map[h] = en

                for row in reader:
                    get = lambda en_key: next((row[h] for h, en in field_map.items() if en == en_key), "")
                    sc = {s: parse_score_any(get(s)) for s in SUBJECTS}
                    avg = wavg(sc); xl = classify(avg)
                    id_val = get("id")
                    try: id_int = int(float(id_val)) if id_val != "" else len(self.students)+1
                    except: id_int = len(self.students)+1
                    self.students.append({
                        "id": id_int,
                        "ho_ten": (get("ho_ten") or "").strip(),
                        "lop": (get("lop") or "").strip(),
                        **sc, "diem_tb": round(avg, 2), "xep_loai": xl
                    })
            self.next_id = max((s["id"] for s in self.students), default=0) + 1
            self.refresh_table()
            self._set_status(f"Đã đọc CSV: {path} (delimiter='{delimiter}')")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    # ---------- EXCEL ----------
    def export_excel(self):
        if not self.students:
            messagebox.showinfo("Trống", "Chưa có dữ liệu để xuất."); return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Workbook","*.xlsx")])
        if not path: return

        headers = ["STT","ID","Họ Tên","Lớp","Toán","Lý","Hóa","Văn","Anh","Tin","Điểm TB","Xếp Loại"]
        wb = Workbook(); ws = wb.active; ws.title = "Bảng điểm"

        head_font = Font(bold=True)
        head_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
        zebra_even = PatternFill(start_color="F7FBFF", end_color="F7FBFF", fill_type="solid")
        zebra_odd  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        thin = Side(style="thin", color="90A4AE")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = head_font; cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        data = sorted(self.students, key=lambda s: (s["lop"], s["ho_ten"]))
        for idx, s in enumerate(data, start=1):
            row_vals = [idx, s["id"], s["ho_ten"], s["lop"],
                        s["toan"], s["ly"], s["hoa"], s["van"], s["anh"], s["tin"],
                        float(f"{s['diem_tb']:.2f}"), s["xep_loai"]]
            r = idx + 1
            for c, v in enumerate(row_vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                if c in (1,2,4,11): cell.alignment = Alignment(horizontal="center")
                elif c == 3: cell.alignment = Alignment(horizontal="left")
                else: cell.alignment = Alignment(horizontal="center")
                cell.fill = zebra_even if idx % 2 == 0 else zebra_odd
                cell.border = border

        for column_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = min(max_len + 2, 28)

        try:
            wb.save(path)
            messagebox.showinfo("Thành công", f"Đã xuất Excel:\n{path}")
            self._set_status(f"Đã xuất Excel: {path}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu Excel:\n{e}")

    # ---------- LẤY TẬP DỮ LIỆU ĐANG HIỂN THỊ ----------
    def _get_visible_subset(self):
        id_to_obj = {s["id"]: s for s in self.students}
        subset = []
        for item in self.tree.get_children():
            vals = self.tree.item(item)["values"]
            if not vals:
                continue
            try:
                sid = int(vals[1])
                if sid in id_to_obj:
                    subset.append(id_to_obj[sid])
            except Exception:
                continue
        return subset

    # ---------- BIỂU ĐỒ HISTOGRAM 7 MÔN ----------
    def show_charts(self):
        subset = self._get_visible_subset()
        if not subset:
            messagebox.showinfo("Chưa có dữ liệu", "Bảng đang rỗng. Hãy thêm học sinh hoặc hiển thị dữ liệu trước.")
            return

        def hist10(values):
            bins = [0]*10
            for v in values:
                bins[_bucket_index(v)] += 1
            return [int(x) for x in bins]

        tb_vals   = [s["diem_tb"] for s in subset]
        toan_vals = [s["toan"]    for s in subset]
        ly_vals   = [s["ly"]      for s in subset]
        hoa_vals  = [s["hoa"]     for s in subset]
        van_vals  = [s["van"]     for s in subset]
        anh_vals  = [s["anh"]     for s in subset]
        tin_vals  = [s["tin"]     for s in subset]

        y_tb   = hist10(tb_vals)
        y_toan = hist10(toan_vals)
        y_li   = hist10(ly_vals)
        y_hoa  = hist10(hoa_vals)
        y_van  = hist10(van_vals)
        y_anh  = hist10(anh_vals)
        y_tin  = hist10(tin_vals)

        x_labels = ["0-1","1-2","2-3","3-4","4-5","5-6","6-7","7-8","8-9","9-10"]
        series = [
            ("TRUNG BÌNH MÔN", y_tb,   "blue"),
            ("TOÁN HỌC",       y_toan, "green"),
            ("VẬT LÝ",         y_li,   "red"),
            ("HÓA HỌC",        y_hoa,  "orange"),
            ("NGỮ VĂN",        y_van,  "purple"),
            ("NGOẠI NGỮ",      y_anh,  "gold"),
            ("TIN HỌC",        y_tin,  "grey"),
        ]

        fig, axs = plt.subplots(4, 2, figsize=(12, 10))
        axs = axs.ravel()

        for i, (title, data, color) in enumerate(series):
            bars = axs[i].bar(x_labels, data, color=color, width=0.6)
            axs[i].set_title(title)
            axs[i].set_ylabel("Số lượng")
            axs[i].set_xlabel("Khoảng điểm")
            axs[i].grid(axis="y", linestyle="--", alpha=0.3)
            axs[i].yaxis.set_major_locator(MaxNLocator(integer=True))
            y_max = max(data) if max(data) > 0 else 1
            axs[i].set_ylim(0, y_max * 1.2 + 1)
            y_off = max(0.4, y_max * 0.03)
            for rect, val in zip(bars, data):
                axs[i].text(rect.get_x() + rect.get_width()/2.0,
                            rect.get_height() + y_off,
                            str(int(val)),
                            ha="center", va="bottom", fontsize=9, clip_on=True)

        axs[-1].axis("off")
        fig.suptitle(f"Phân bố điểm theo khoảng – Số học sinh hiển thị: {len(subset)}", y=0.995)
        plt.tight_layout()
        try:
            fig.savefig("bieudo.png", dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.2)
        except Exception:
            pass
        plt.show()

    # ---------- BIỂU ĐỒ 3 KHỐI: 4 SERIES/ MỖI BIN ----------
    def show_block_chart(self):
        subset = self._get_visible_subset()
        if not subset:
            messagebox.showinfo("Chưa có dữ liệu", "Bảng đang rỗng. Hãy thêm học sinh hoặc hiển thị dữ liệu trước.")
            return

        # gom theo khối dựa vào tiền tố lớp
        blocks = {"10": [], "11": [], "12": []}
        for s in subset:
            lop = (s.get("lop") or "").strip()
            tb  = float(s.get("diem_tb", 0.0))
            if lop.startswith("10"):
                blocks["10"].append(tb)
            elif lop.startswith("11"):
                blocks["11"].append(tb)
            elif lop.startswith("12"):
                blocks["12"].append(tb)

        # histogram 10 bin
        def hist10(values):
            bins = [0]*10
            for v in values:
                try: x = float(v)
                except: x = 0.0
                if x < 0: x = 0.0
                if x > 10: x = 10.0
                idx = 9 if x >= 9.0 else int(x)
                bins[idx] += 1
            return [int(x) for x in bins]

        h10 = hist10(blocks["10"])
        h11 = hist10(blocks["11"])
        h12 = hist10(blocks["12"])
        havg = [(h10[i] + h11[i] + h12[i]) / 3.0 for i in range(10)]  # TB theo từng bin

        x_labels = ["0-1","1-2","2-3","3-4","4-5","5-6","6-7","7-8","8-9","9-10"]
        x = np.arange(len(x_labels))
        width = 0.2

        fig, ax = plt.subplots(figsize=(12, 6))
        r10  = ax.bar(x - 1.5*width, h10, width, label="Khối 10")
        r11  = ax.bar(x - 0.5*width, h11, width, label="Khối 11")
        r12  = ax.bar(x + 0.5*width, h12, width, label="Khối 12")
        ravg = ax.bar(x + 1.5*width, havg, width, label="TB 3 khối")

        ax.set_title("Phân bố Điểm TB theo khoảng – 3 khối & Trung bình cộng")
        ax.set_xlabel("Khoảng điểm")
        ax.set_ylabel("Số học sinh")
        ax.set_xticks(x, x_labels)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax.grid(axis="y", linestyle="--", alpha=0.3)
        ax.legend(loc="upper right", ncols=2)

        ymax = max(max(h10 or [0]), max(h11 or [0]), max(h12 or [0]), int(max(havg or [0])))
        ymax = 1 if ymax <= 0 else ymax
        ax.set_ylim(0, ymax*1.25 + 1)

        def annotate(bars):
            for rect in bars:
                h = rect.get_height()
                ax.text(rect.get_x() + rect.get_width()/2,
                        h + max(0.4, ymax*0.03),
                        f"{h:.0f}" if abs(h - round(h)) < 1e-9 else f"{h:.1f}",
                        ha="center", va="bottom", fontsize=9, clip_on=True)
        annotate(r10); annotate(r11); annotate(r12); annotate(ravg)

        plt.tight_layout()
        plt.show()

# ---------- AI Q&A WINDOW (nâng cấp: 4 tab, ngân hàng câu hỏi, Ctrl+Enter, xuất log) ----------
def open_qna_window(root, get_subset_callable=None):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText
    from datetime import datetime
    import os

    # --- Toplevel riêng, KHÔNG đổi theme toàn cục ---
    win = tk.Toplevel(root)
    win.title("Hỏi AI (Q&A)")
    win.geometry("1100x740")

    # ---------- Styles cục bộ cho Q&A (không ảnh hưởng app chính) ----------
    style = ttk.Style(win)  # style gắn với toplevel này
    # khung & tiêu đề
    style.configure("QnA.TLabelframe", padding=10)
    style.configure("QnA.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
    # nút màu (ttk hỗ trợ hạn chế, nhưng đủ để nhấn nhá)
    style.configure("QnA.Primary.TButton",  foreground="white", background="#1e88e5")
    style.map("QnA.Primary.TButton",
              background=[("active", "#1565c0"), ("!active", "#1e88e5")])
    style.configure("QnA.Info.TButton",     foreground="white", background="#0288d1")
    style.map("QnA.Info.TButton",
              background=[("active", "#0277bd"), ("!active", "#0288d1")])
    style.configure("QnA.Secondary.TButton", foreground="white", background="#607d8b")
    style.map("QnA.Secondary.TButton",
              background=[("active", "#546e7a"), ("!active", "#607d8b")])

    # ---------- State / log / history ----------
    log_lines = []
    history_rows = []  # (time, question, answer, idx)

    def log(msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        log_lines.append(line)
        txt_log.configure(state="normal")
        txt_log.insert("end", line + "\n")
        txt_log.see("end")
        txt_log.configure(state="disabled")

    # ---------- Notebook ----------
    nb = ttk.Notebook(win)
    tab_chat = ttk.Labelframe(nb, text="💬 Chat với AI", style="QnA.TLabelframe")
    tab_ctx  = ttk.Labelframe(nb, text="📄 Ngữ cảnh",   style="QnA.TLabelframe")
    tab_cfg  = ttk.Labelframe(nb, text="⚙️ Cài đặt",    style="QnA.TLabelframe")
    tab_log  = ttk.Labelframe(nb, text="🧾 Nhật ký",    style="QnA.TLabelframe")
    nb.add(tab_chat, text="Chat"); nb.add(tab_ctx, text="Ngữ cảnh")
    nb.add(tab_cfg, text="Cài đặt"); nb.add(tab_log, text="Nhật ký")
    nb.pack(fill="both", expand=True, padx=10, pady=10)

    # ================= TAB CHAT =================
    left = ttk.Labelframe(tab_chat, text="Ngân hàng câu hỏi", style="QnA.TLabelframe")
    left.pack(side="left", fill="y", padx=(0,10))

    questions = [
        "Phân tích phân bố xếp loại theo lớp.",
        "Liệt kê top 10 học sinh theo Điểm TB.",
        "Tìm các lớp có Điểm TB trung bình ≥ 8.0.",
        "Nhận xét chênh lệch điểm Toán-Văn.",
        "Lọc HS khối 11 có Điểm TB ≥ 7.5.",
        "Mô tả báo cáo theo khối 10/11/12.",
        "Tìm HS điểm thấp nhất từng môn.",
        "Gợi ý biểu đồ phù hợp dữ liệu này.",
        "Tạo tiêu đề & tóm tắt báo cáo tuần.",
        "Đề xuất tiêu chí xét học bổng."
    ]
    lst = tk.Listbox(left, height=22)
    for q in questions: lst.insert("end", q)
    lst.pack(fill="y", expand=True)

    right = ttk.Frame(tab_chat)
    right.pack(side="left", fill="both", expand=True)

    frm_q = ttk.Labelframe(right, text="Câu hỏi", style="QnA.TLabelframe")
    frm_q.pack(fill="x")
    txt_q = ScrolledText(frm_q, height=5, wrap="word", font=("Segoe UI", 11), relief="flat")
    txt_q.pack(fill="x", pady=(0,6))

    def insert_from_bank(_evt=None):
        sel = lst.curselection()
        if not sel: return
        q = lst.get(sel[0])
        txt_q.insert("end", ("" if txt_q.index("end-1c")=="1.0" else "\n") + q)
        txt_q.focus_set()
    lst.bind("<Double-Button-1>", insert_from_bank)

    ttk.Separator(right, orient="horizontal").pack(fill="x", pady=6)

    frm_a = ttk.Labelframe(right, text="Trả lời", style="QnA.TLabelframe")
    frm_a.pack(fill="both", expand=True)
    txt_a = ScrolledText(frm_a, height=14, wrap="word", font=("Segoe UI", 11), relief="flat")
    txt_a.pack(fill="both", expand=True)

    action = ttk.Frame(right); action.pack(fill="x", pady=(8,0))
    include_ctx_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(action, text="Đưa ngữ cảnh bảng vào câu hỏi",
                    variable=include_ctx_var).pack(side="left")

    def on_copy():
        win.clipboard_clear()
        win.clipboard_append(txt_a.get("1.0","end").strip())
        log("Đã copy trả lời.")

    def on_export_log():
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Text","*.txt")])
        if not path: return
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines))
        messagebox.showinfo("Xuất", f"Đã lưu log:\n{path}")

    btn_ask    = ttk.Button(action, text="🤖 Hỏi AI (Ctrl+Enter)", style="QnA.Primary.TButton")
    btn_copy   = ttk.Button(action, text="📋 Copy",                 style="QnA.Info.TButton", command=on_copy)
    btn_export = ttk.Button(action, text="💾 Xuất log",             style="QnA.Secondary.TButton", command=on_export_log)
    btn_export.pack(side="right", padx=6); btn_copy.pack(side="right", padx=6); btn_ask.pack(side="right", padx=6)

    # Lịch sử Q&A (panel phải)
    history = ttk.Labelframe(tab_chat, text="Lịch sử hỏi đáp", style="QnA.TLabelframe")
    history.pack(side="left", fill="y", padx=(10,0))
    tv = ttk.Treeview(history, columns=("time","q","a","idx"), show="headings", height=25)
    tv.heading("time", text="Thời gian"); tv.heading("q", text="Câu hỏi"); tv.heading("a", text="Tóm tắt trả lời")
    tv.column("time", width=85, anchor="center")
    tv.column("q", width=280, anchor="w"); tv.column("a", width=280, anchor="w")
    tv.heading("idx", text=""); tv.column("idx", width=0, stretch=False)  # ẩn index
    tv.pack(fill="y", expand=True)

    def on_select_history(_evt=None):
        sel = tv.selection()
        if not sel: return
        vals = tv.item(sel[0])["values"]
        if len(vals) < 4: return
        idx = int(vals[3])
        t,q,a,_ = history_rows[idx]
        txt_q.delete("1.0","end"); txt_q.insert("end", q)
        txt_a.delete("1.0","end"); txt_a.insert("end", a)
    tv.bind("<<TreeviewSelect>>", on_select_history)

    # ================= TAB NGỮ CẢNH =================
    txt_ctx = ScrolledText(tab_ctx, wrap="none", font=("Consolas", 10))
    txt_ctx.pack(fill="both", expand=True)
    def refresh_context():
        txt_ctx.delete("1.0","end")
        if callable(get_subset_callable):
            try:
                subset = get_subset_callable() or []
                txt_ctx.insert("end", _ai__summarize_records(subset, limit_rows=50))
            except Exception as e:
                txt_ctx.insert("end", f"(Lỗi lấy ngữ cảnh) {e}")
        else:
            txt_ctx.insert("end", "(Không có hàm truy cập ngữ cảnh)")
    ttk.Button(tab_ctx, text="🔄 Làm mới", style="QnA.Info.TButton", command=refresh_context)\
        .pack(anchor="e", pady=6)
    refresh_context()

    # ================= TAB CÀI ĐẶT =================
    cfg = ttk.Frame(tab_cfg, padding=4); cfg.pack(fill="x")
    ttk.Label(cfg, text="Model (ENV OPENAI_MODEL):").grid(row=0, column=0, sticky="w", pady=6)
    model_var = tk.StringVar(value=os.getenv("OPENAI_MODEL", "gpt-4o-mini"))
    model_box = ttk.Combobox(cfg, textvariable=model_var,
                             values=("gpt-4o-mini","gpt-4o","gpt-5-mini"),
                             width=20, state="readonly")
    model_box.grid(row=0, column=1, sticky="w", padx=8)
    def on_model_sel(*_):
        os.environ["OPENAI_MODEL"] = model_var.get()
        log(f"Chọn model: {model_var.get()}")
    model_box.bind("<<ComboboxSelected>>", on_model_sel)

    ttk.Label(cfg, text="Temperature (0.0 - 2.0):").grid(row=1, column=0, sticky="w", pady=6)
    spn_temp = ttk.Spinbox(cfg, from_=0.0, to=2.0, increment=0.1, width=6)
    spn_temp.insert(0, "0.2"); spn_temp.grid(row=1, column=1, sticky="w", padx=8)

    ttk.Label(cfg, text="Max output tokens:").grid(row=2, column=0, sticky="w", pady=6)
    spn_tokens = ttk.Spinbox(cfg, from_=100, to=4000, increment=50, width=8)
    spn_tokens.insert(0, "400"); spn_tokens.grid(row=2, column=1, sticky="w", padx=8)

    # ================= TAB NHẬT KÝ =================
    txt_log = ScrolledText(tab_log, wrap="word", state="disabled", font=("Segoe UI", 10))
    txt_log.pack(fill="both", expand=True)

    # ---------- GỬI HỎI AI ----------
    def _shorten(s, n=80):
        s = " ".join(s.split())
        return s if len(s) <= n else s[:n-1] + "…"

    def on_ask(_evt=None):
        q = txt_q.get("1.0","end").strip()
        if not q:
            messagebox.showinfo("Thiếu nội dung", "Hãy nhập câu hỏi."); return

        txt_a.delete("1.0","end"); txt_a.insert("end", "⏳ Đang hỏi AI...\n")

        ctx = ""
        if include_ctx_var.get() and callable(get_subset_callable):
            try: ctx = _ai__summarize_records(get_subset_callable() or [], limit_rows=15)
            except: ctx = ""

        try:
            temp = float(spn_temp.get()); toks = int(spn_tokens.get())
        except: temp, toks = 0.2, 400

        ans = _ai__ask_chatgpt(q, context=ctx, temperature=temp, max_output_tokens=toks)
        txt_a.delete("1.0","end"); txt_a.insert("end", ans)

        t = datetime.now().strftime("%H:%M:%S")
        idx = len(history_rows)
        history_rows.append((t, q, ans, idx))
        tv.insert("", "end", values=(t, _shorten(q, 60), _shorten(ans, 60), idx))
        log(f"Hỏi: {q[:50]}... / Đáp: {ans[:50]}...")

    btn_ask.configure(command=on_ask)
    txt_q.bind("<Control-Return>", on_ask)

    log("Khởi động Q&A (ttk thuần, không đổi theme app).")

# ---------- RUN ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = StudentManagerGUI(root)
    root.mainloop()
