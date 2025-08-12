
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import os, re, csv
from db_helper import ensure_db, upsert_many_simple, replace_all, query, lookup_by_bankname

APP_DIR = Path(__file__).parent
DB_PATH = str(APP_DIR / "codebook.db")

COMMON_ENCODINGS = ["utf-8-sig","utf-8","gbk","gb18030","utf-16","utf-16le","utf-16be","latin1"]
COMMON_DELIMS = ["|", "\t", ",", ";", " "]

def sniff_delimiter(text: str):
    try:
        dialect = csv.Sniffer().sniff(text, delimiters="|\t,; ")
        return dialect.delimiter
    except Exception:
        first = next((ln for ln in text.splitlines() if ln.strip()), "")
        best, bestn = None, 1
        for d in COMMON_DELIMS:
            n = len([c for c in first.split(d) if c != ""])
            if n > bestn:
                best, bestn = d, n
        return best or ","

def try_parse_txt(path: str, encoding=None, delimiter=None, header=None):
    data = Path(path).read_bytes()
    encs = [encoding] if encoding else COMMON_ENCODINGS
    for enc in encs:
        try:
            s = data.decode(enc)
        except Exception:
            continue
        s = s.replace("\r\n","\n").replace("\r","\n")
        if s and s[0] == "\ufeff": s = s[1:]
        delim = delimiter or sniff_delimiter(s)
        from io import StringIO
        sio = StringIO(s)
        try:
            df = pd.read_csv(sio, sep=delim, header=(0 if header else None),
                             dtype=str, engine="python", skip_blank_lines=True)\
                    .dropna(axis=1, how="all").dropna(axis=0, how="all")
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            if df.shape[1] == 1:
                col = df.columns[0]
                for d in ["|","\t",",",";"]:
                    parts = df[col].str.split(d, expand=True)
                    if parts.shape[1] >= 2:
                        return parts
                parts = df[col].str.split(r"\s+", expand=True)
                if parts.shape[1] >= 2:
                    return parts
            return df
        except Exception:
            if delim == " ":
                try:
                    sio = StringIO(s)
                    df = pd.read_csv(sio, sep=r"\s+", header=(0 if header else None), engine="python", dtype=str)
                    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
                    return df
                except Exception:
                    pass
            continue
    raise RuntimeError("无法解析 TXT：编码或分隔符不匹配。")

def read_any(path: str, *, txt_encoding=None, txt_delim=None, txt_header=False):
    p = Path(path)
    if p.suffix.lower() in [".xlsx",".xlsm",".xltx",".xltm"]:
        return pd.read_excel(path, engine="openpyxl", dtype=str)
    if p.suffix.lower() == ".xls":
        try:
            return pd.read_excel(path, dtype=str)
        except Exception as e:
            raise RuntimeError("读取 .xls 需安装 xlrd：pip install xlrd") from e
    if p.suffix.lower() == ".csv":
        return pd.read_csv(path, encoding="utf-8", dtype=str)
    if p.suffix.lower() in [".txt",".dat"]:
        return try_parse_txt(path, encoding=txt_encoding, delimiter=txt_delim, header=txt_header)
    raise RuntimeError("无法解析文件，请转存为 CSV/Excel 再试。")

def export_text_xlsx(df: pd.DataFrame, path: str, *, include_header: bool = True):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title="Sheet1"
    start_row = 1
    if include_header:
        for j, col in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=j, value=str(col)); cell.number_format="@"
        start_row = 2
    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        for j, col in enumerate(df.columns, start=1):
            val = "" if pd.isna(row[col]) else str(row[col])
            cell = ws.cell(row=i, column=j, value=val); cell.number_format="@"
    wb.save(path)

# ---------- helpers for picking columns & cleaning ----------
def pick_ibps(df):
    df2 = df.copy()
    if df2.columns.size:
        df2.columns = [str(c).strip() for c in df2.columns]
    code_alias = ["清算行行号","清算行号","联行号","行号","行号代码","清算行行号代码"]
    name_alias = ["清算行名称","清算行名","名称","银行名称","开户行名称"]
    code_col = next((c for c in df2.columns if c in code_alias), None)
    name_col = next((c for c in df2.columns if c in name_alias), None)
    if code_col and name_col:
        use = df2[[code_col, name_col]].copy()
        use.columns = ["code","name"]
    else:
        use = df2.iloc[:, :2].copy(); use.columns = ["code","name"]
    use["code"] = use["code"].astype(str).str.replace(".0","", regex=False)
    use["code"] = use["code"].apply(lambda s: (re.search(r"(\d{12})", re.sub(r"\D","", s)) or re.search(r"(\d{12})", s or "")))
    use["code"] = use["code"].apply(lambda m: (m.group(1) if m else ""))
    use["name"] = use["name"].astype(str).str.strip()
    return use

def pick_cnaps(df):
    df2 = df.copy()
    if df2.columns.size:
        df2.columns = [str(c).strip() for c in df2.columns]
    needed = ["BNKCODE","CLSCODE","CITYCODE","LNAME"]
    if set(needed).issubset(set(df2.columns)):
        use = df2[needed].copy()
    elif df2.shape[1] >= 4:
        use = df2.iloc[:, :4].copy(); use.columns = needed
    elif df2.shape[1] == 1:
        col = df2.columns[0]
        for d in ["|","\t",",",";"]:
            parts = df2[col].str.split(d, expand=True)
            if parts.shape[1] >= 4:
                use = parts.iloc[:, :4]; use.columns = needed; break
        else:
            parts = df2[col].str.split(r"\s+", expand=True)
            if parts.shape[1] >= 4:
                use = parts.iloc[:, :4]; use.columns = needed
            else:
                def regex_extract(s):
                    s = str(s or "")
                    m = re.search(r"(\d{12})", s)
                    tail = (s[m.end():].strip() if m else "")
                    return (m.group(1) if m else ""), "", "", tail.strip(" ,;|\t")
                use = pd.DataFrame([regex_extract(v) for v in df2[col].tolist()], columns=needed)
    else:
        use = df2.copy()
        while use.shape[1] < 4:
            use[use.shape[1]] = ""
        use = use.iloc[:, :4]; use.columns = needed
    use["BNKCODE"] = use["BNKCODE"].astype(str).str.replace(".0","", regex=False)
    use["BNKCODE"] = use["BNKCODE"].apply(lambda s: (re.search(r"(\d{12})", re.sub(r"\D","", s)) or re.search(r"(\d{12})", s or "")))
    use["BNKCODE"] = use["BNKCODE"].apply(lambda m: (m.group(1) if m else ""))
    use["LNAME"] = use["LNAME"].astype(str).str.strip()
    return use

# ---------- reusable UI ----------
class ScrollableTree(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master)
        self.tree = ttk.Treeview(self, show="headings", **kwargs)
        xbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

class AdvancedImport(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("高级导入参数（TXT专用）"); self.resizable(False, False)
        self.encoding = tk.StringVar(value="自动")
        self.delim = tk.StringVar(value="自动")
        self.has_header = tk.BooleanVar(value=False)
        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="编码：").grid(row=0,column=0,sticky="e",padx=6,pady=6)
        enc_cb = ttk.Combobox(frm, textvariable=self.encoding, width=18, state="readonly",
                              values=["自动"]+COMMON_ENCODINGS); enc_cb.grid(row=0,column=1,sticky="w",padx=6,pady=6)
        ttk.Label(frm, text="分隔符：").grid(row=1,column=0,sticky="e",padx=6,pady=6)
        delim_cb = ttk.Combobox(frm, textvariable=self.delim, width=18, state="readonly",
                                values=["自动","|","\\t（制表）",",",";","空格"]); delim_cb.grid(row=1,column=1,sticky="w",padx=6,pady=6)
        ttk.Checkbutton(frm, text="首行为表头", variable=self.has_header).grid(row=2,column=1,sticky="w",padx=6,pady=6)
        btns = ttk.Frame(frm); btns.grid(row=3,column=0,columnspan=2,pady=10)
        ttk.Button(btns, text="确定", command=self.ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=8)
        self.result=None; self.bind("<Return>", lambda e: self.ok())
    def ok(self):
        enc = None if self.encoding.get()=="自动" else self.encoding.get()
        d = self.delim.get()
        if d=="自动": delim=None
        elif d=="\\t（制表）": delim="\\t"
        elif d=="空格": delim=" "
        else: delim=d
        self.result = dict(encoding=enc, delim=delim, header=self.has_header.get()); self.destroy()

# ---------- Library tab ----------
class LibraryTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.table_choice = tk.StringVar(value="ibps")
        self.kw = tk.StringVar()
        self._adv = dict(encoding=None, delim=None, header=False)
        self._build()

    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="维护库：").pack(side="left")
        ttk.Radiobutton(top, text="IBPS（清算）", variable=self.table_choice, value="ibps").pack(side="left")
        ttk.Radiobutton(top, text="CNAPS（大额）", variable=self.table_choice, value="cnaps").pack(side="left", padx=8)
        ttk.Button(top, text="导入（txt/xls/xlsx/csv）", command=self.import_file).pack(side="left", padx=12)
        ttk.Button(top, text="高级导入参数", command=self.adv_params).pack(side="left")
        ttk.Button(top, text="导出库", command=self.export_db).pack(side="left", padx=12)
        ttk.Label(top, text="关键词：").pack(side="left", padx=12)
        ttk.Entry(top, textvariable=self.kw, width=28).pack(side="left")
        ttk.Button(top, text="查询", command=self.search).pack(side="left", padx=6)

        self.stree = ScrollableTree(self); self.stree.pack(fill="both", expand=True, padx=8, pady=6)

    def adv_params(self):
        dlg = AdvancedImport(self); self.wait_window(dlg)
        if dlg.result:
            self._adv = dlg.result
            messagebox.showinfo("已设置", f"编码={self._adv['encoding'] or '自动'}，分隔符={self._adv['delim'] or '自动'}，首行为表头={self._adv['header']}")

    def import_file(self):
        path = filedialog.askopenfilename(filetypes=[("TXT/Excel/CSV","*.txt;*.dat;*.xls;*.xlsx;*.csv"),("所有文件","*.*")])
        if not path: return
        try:
            df = read_any(path, txt_encoding=self._adv['encoding'], txt_delim=self._adv['delim'], txt_header=self._adv['header'])
        except Exception as e:
            messagebox.showerror("失败", f"读取失败：{e}"); return
        rows = []; raw_src = os.path.basename(path)
        if self.table_choice.get()=="cnaps":
            use = pick_cnaps(df)
            for _, r in use.iterrows():
                code = r["BNKCODE"]; name = r["LNAME"]
                if re.fullmatch(r"\d{12}", code or ""):
                    raw = "|".join([str(r[c]) for c in ["BNKCODE","CLSCODE","CITYCODE","LNAME"] if c in use.columns])
                    rows.append((code, name, raw, raw_src))
            table="cnaps"
        else:
            use = pick_ibps(df)
            for _, r in use.iterrows():
                code = r["code"]; name = r["name"]
                if re.fullmatch(r"\d{12}", code or ""):
                    raw = "|".join([str(r[c]) for c in ["code","name"] if c in use.columns])
                    rows.append((code, name, raw, raw_src))
            table="ibps"
        if not rows:
            messagebox.showwarning("提示","未发现有效的12位行号记录（可尝试：高级导入设置编码=GBK、分隔符=|或Tab；或把TXT另存为CSV/Excel）"); return
        if messagebox.askyesno("导入方式", "选择“是”= 全量替换；“否”= 增量合并（按 code upsert）"):
            replace_all(DB_PATH, table, rows)
        else:
            upsert_many_simple(DB_PATH, table, rows)
        messagebox.showinfo("成功", f"导入完成：共 {len(rows)} 条")

    def search(self):
        table = self.table_choice.get()
        rows = query(DB_PATH, table, self.kw.get().strip(), limit=2000)
        df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["code","name"])
        self._load_df(df)

    def export_db(self):
        rows = query(DB_PATH, self.table_choice.get(), "", limit=999999)
        if not rows:
            messagebox.showinfo("提示","当前库为空"); return
        df = pd.DataFrame(rows)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel",".xlsx"),("CSV",".csv")])
        if not path: return
        try:
            if path.lower().endswith(".csv"):
                df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                export_text_xlsx(df, path, include_header=True)
            messagebox.showinfo("成功", f"已导出：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("失败", f"导出失败：{e}")

    def _load_df(self, df):
        tree = self.stree.tree
        tree["columns"] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=220 if col=="name" else 160, anchor="w")
        tree.delete(*tree.get_children())
        for _, row in df.iterrows():
            tree.insert("", "end", values=[str(row.get(c,"")) for c in df.columns])

# ---------- CodePicker ----------
class ScrollableForm(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

class CodePicker(tk.Toplevel):
    def __init__(self, master, default_source="ibps", ibps_only=False):
        super().__init__(master)
        self.title("选择行号（IBPS/CNAPS）" if not ibps_only else "选择银行（IBPS）")
        self.geometry("760x520"); self.resizable(True, True)
        self.ibps_only = ibps_only
        self.source = tk.StringVar(value="ibps")
        self.kw = tk.StringVar()
        self.selected_row = None
        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        if not ibps_only:
            ttk.Label(top, text="来源：").pack(side="left")
            ttk.Radiobutton(top, text="IBPS（清算）", variable=self.source, value="ibps", command=self.search).pack(side="left")
            ttk.Radiobutton(top, text="CNAPS（大额）", variable=self.source, value="cnaps", command=self.search).pack(side="left", padx=8)
            self.source.set(default_source)
        else:
            ttk.Label(top, text="来源：IBPS（清算）").pack(side="left")
            self.source.set("ibps")
        ttk.Label(top, text="关键字：").pack(side="left", padx=8)
        ent = ttk.Entry(top, textvariable=self.kw, width=32); ent.pack(side="left"); ent.bind("<Return>", lambda e: self.search())
        ttk.Button(top, text="查询", command=self.search).pack(side="left", padx=6)

        self.stree = ScrollableTree(self, height=18); self.stree.pack(fill="both", expand=True, padx=8, pady=6)
        tree = self.stree.tree; tree["columns"] = ["code","name"]
        for c, w in [("code",180),("name",480)]:
            tree.heading(c, text=c); tree.column(c, width=w, anchor="w")
        tree.bind("<Double-1>", lambda e: self.pick()); tree.bind("<Return>", lambda e: self.pick())

        btns = ttk.Frame(self, padding=8); btns.pack(fill="x")
        ttk.Button(btns, text="确定", command=self.pick).pack(side="right", padx=6)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")

        self.search()

    def search(self):
        rows = query(DB_PATH, self.source.get(), self.kw.get().strip(), limit=2000)
        tree = self.stree.tree; tree.delete(*tree.get_children())
        for r in rows:
            tree.insert("", "end", values=[r["code"], r["name"]])

    def pick(self):
        tree = self.stree.tree; sel = tree.selection()
        if not sel: messagebox.showinfo("提示","请先选择一条"); return
        vals = tree.item(sel[0], "values")
        self.selected_row = (vals[0], vals[1])  # (code, name)
        self.destroy()

# ---------- Payroll ----------
class PayrollDialog(tk.Toplevel):
    COLS = ["收款人银行名称","收款人卡号","收款人名称","金额"]
    def __init__(self, master, init_values=None):
        super().__init__(master)
        self.title("新增/编辑 - 代发工资"); self.resizable(True, True)
        self.values = {}
        sf = ScrollableForm(self); sf.pack(fill="both", expand=True)
        frm = sf.inner
        self.vars = {}
        # 收款人银行名称 + 选择（IBPS-only）
        ttk.Label(frm, text="收款人银行名称：").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        v_bank = tk.StringVar(value=(init_values.get("收款人银行名称","") if init_values else ""))
        wrap = ttk.Frame(frm)
        ttk.Entry(wrap, textvariable=v_bank, width=36).pack(side="left")
        def choose_name():
            dlg = CodePicker(self, default_source="ibps", ibps_only=True)
            self.wait_window(dlg)
            if getattr(dlg, "selected_row", None):
                _, name = dlg.selected_row
                v_bank.set(name or v_bank.get())
        ttk.Button(wrap, text="选择…", command=choose_name).pack(side="left", padx=6)
        wrap.grid(row=0, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm, text="（可点击“选择…”从 IBPS 库带出银行名称）", foreground="#666").grid(row=0, column=2, sticky="w", padx=6)
        self.vars["收款人银行名称"] = v_bank

        # 卡号、姓名、金额
        labels = ["收款人卡号","收款人名称","金额"]
        for i, col in enumerate(labels, start=1):
            ttk.Label(frm, text=col + "：").grid(row=i, column=0, sticky="e", padx=6, pady=6)
            var = tk.StringVar(value=(init_values.get(col,"") if init_values else ""))
            ttk.Entry(frm, textvariable=var, width=36).grid(row=i, column=1, sticky="w", padx=6, pady=6)
            self.vars[col] = var

        btns = ttk.Frame(self, padding=8); btns.pack(fill="x")
        ttk.Button(btns, text="确定", command=self.ok).pack(side="right", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")
        self.bind("<Return>", lambda e: self.ok())

    def ok(self):
        vals = {k: v.get().strip() for k, v in self.vars.items()}
        probs = []
        if not vals["收款人银行名称"]: probs.append("收款人银行名称 不能为空")
        if not re.fullmatch(r"\d{6,32}", vals["收款人卡号"]): probs.append("收款人卡号 必须为6-32位数字")
        try:
            if float(vals["金额"]) <= 0: probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是有效数字")
        if not vals["收款人名称"]: probs.append("收款人名称 不能为空")
        if probs:
            messagebox.showwarning("校验不通过", "；".join(probs)); return
        self.values = vals; self.destroy()

class PayrollTab(ttk.Frame):
    COLS = ["收款人银行名称","收款人卡号","收款人名称","金额"]
    def __init__(self, master):
        super().__init__(master)
        self.df = pd.DataFrame(columns=self.COLS); self._build()

    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Button(top, text="新增一条", command=self.add_one).pack(side="left")
        ttk.Button(top, text="编辑选中", command=self.edit_one).pack(side="left", padx=6)
        ttk.Button(top, text="删除选中", command=self.delete_selected).pack(side="left", padx=6)
        ttk.Button(top, text="打开代发工资表", command=self.load_file).pack(side="left", padx=12)
        ttk.Button(top, text="校验并导出（无表头）", command=self.validate_export).pack(side="left", padx=6)

        self.stree = ScrollableTree(self, height=18); self.stree.pack(fill="both", expand=True, padx=8, pady=6)
        self._reload()

    def add_one(self):
        dlg = PayrollDialog(self); self.wait_window(dlg)
        if getattr(dlg, "values", None):
            self.df.loc[len(self.df)] = [dlg.values[c] for c in self.COLS]; self._reload()

    def edit_one(self):
        tree = self.stree.tree; sel = tree.selection()
        if not sel: messagebox.showinfo("提示","请先选择一行"); return
        idx = tree.index(sel[0]); init = {c: str(self.df.iloc[idx][c]) for c in self.COLS}
        dlg = PayrollDialog(self, init_values=init); self.wait_window(dlg)
        if getattr(dlg, "values", None):
            for c in self.COLS: self.df.at[self.df.index[idx], c] = dlg.values[c]
            self._reload()

    def delete_selected(self):
        tree = self.stree.tree; sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0]); self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True); self._reload()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV","*.xlsx;*.xls;*.csv")])
        if not path: return
        try:
            df = read_any(path)
            new = pd.DataFrame(columns=self.COLS)
            for c in self.COLS: new[c] = df[c].astype(str) if c in df.columns else ""
            self.df = new[self.COLS]; self._reload()
        except Exception as e:
            messagebox.showerror("失败", f"读取失败：{e}")

    def validate_export(self):
        df = self.df.fillna("")
        probs = []
        if df["收款人银行名称"].str.strip().eq("").any(): probs.append("收款人银行名称 不能为空")
        if (~df["收款人卡号"].astype(str).str.match(r"^\d{6,32}$", na=False)).any(): probs.append("收款人卡号 格式异常")
        if df["收款人名称"].str.strip().eq("").any(): probs.append("收款人名称 不能为空")
        try:
            if (pd.to_numeric(df["金额"], errors="coerce")<=0).any(): probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是可解析数字")
        if probs: messagebox.showwarning("校验结果","；".join(probs))
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel",".xlsx")])
        if not path: return
        export_text_xlsx(self.df, path, include_header=False)
        messagebox.showinfo("成功", "已导出（无表头，文本格式）")

# ---------- Transfer ----------
class TransferDialog(tk.Toplevel):
    COLS = ["收款方账号","收款方户名","金额","转账方式","行别信息类型",
            "收款方银行名称","收款方银行大额支付行号/跨行清算行号","用途","明细标注"]
    def __init__(self, master, init_values=None, default_source="ibps"):
        super().__init__(master)
        self.title("新增/编辑 - 批量转账"); self.minsize(820, 560); self.resizable(True, True)
        self.values = {}; self.default_source = default_source
        self.transfer_mode_map = {"0":"0 - 行内转账（同一银行）","1":"1 - 跨行转账（不同银行）"}
        self.bankinfo_type_map = {"0":"0 - 跨行清算银行信息（IBPS）","1":"1 - 开户支行信息（开户支行名称/大额）"}
        container = ttk.Frame(self, padding=12); container.pack(fill="both", expand=True)
        container.grid_columnconfigure(1, weight=1)
        tips = {
            "转账方式":"0=行内；1=跨行（跨行需填写下方“行号”）",
            "行别信息类型":"0=跨行清算银行信息；1=开户支行信息。行内不填；银联卡号可不填。",
            "行号":"支持手输或点击【选择…】从本地库选；选择后会自动带出银行名称。",
            "可选项":"行内转账时不输入；收款方账号是银联卡号时可以不输入"
        }
        self.vars = {}; row=0
        def add(label, widget, hint_key=None, below_hint_key=None):
            nonlocal row
            ttk.Label(container, text=label+"：").grid(row=row, column=0, sticky="e", padx=6, pady=6)
            w = widget(container)
            if isinstance(w, tuple):
                main, extra = w; main.grid(row=row, column=1, sticky="ew", padx=6, pady=6)
                if extra is not None: extra.grid(row=row, column=2, sticky="w", padx=6, pady=6)
            else:
                w.grid(row=row, column=1, sticky="ew", padx=6, pady=6)
            if hint_key:
                ttk.Label(container, text="（"+tips[hint_key]+"）", foreground="#666").grid(row=row, column=3, sticky="w", padx=6, pady=6)
            row+=1
            if below_hint_key:
                ttk.Label(container, text="（"+tips[below_hint_key]+"）", foreground="#666").grid(row=row, column=1, columnspan=3, sticky="w", padx=6, pady=(0,8))
                row+=1
        v_acct = tk.StringVar(value=(init_values.get("收款方账号","") if init_values else "")); add("收款方账号", lambda p: ttk.Entry(p, textvariable=v_acct)); self.vars["收款方账号"]=v_acct
        v_uname= tk.StringVar(value=(init_values.get("收款方户名","") if init_values else "")); add("收款方户名", lambda p: ttk.Entry(p, textvariable=v_uname)); self.vars["收款方户名"]=v_uname
        v_amt  = tk.StringVar(value=(init_values.get("金额","") if init_values else "")); add("金额", lambda p: ttk.Entry(p, textvariable=v_amt)); self.vars["金额"]=v_amt
        cur_mode = str(init_values.get("转账方式","")) if init_values else "0"
        if cur_mode not in ("0","1"): cur_mode="0"
        v_mode = tk.StringVar(value=self.transfer_mode_map[cur_mode])
        add("转账方式", lambda p: ttk.Combobox(p, textvariable=v_mode, state="readonly", values=[self.transfer_mode_map["0"], self.transfer_mode_map["1"]]), "转账方式"); self.vars["转账方式"]=v_mode
        cur_btype = str(init_values.get("行别信息类型","")) if init_values else "0"
        if cur_btype not in ("0","1"): cur_btype="0"
        v_btype= tk.StringVar(value=self.bankinfo_type_map[cur_btype])
        add("行别信息类型", lambda p: ttk.Combobox(p, textvariable=v_btype, state="readonly", values=[self.bankinfo_type_map["0"], self.bankinfo_type_map["1"]]), "行别信息类型"); self.vars["行别信息类型"]=v_btype
        v_bname= tk.StringVar(value=(init_values.get("收款方银行名称","") if init_values else ""))
        add("收款方银行名称", lambda p: ttk.Entry(p, textvariable=v_bname), "可选项"); self.vars["收款方银行名称"]=v_bname
        v_code = tk.StringVar(value=(init_values.get("收款方银行大额支付行号/跨行清算行号","") if init_values else ""))
        def code_builder(p):
            wrapper = ttk.Frame(p); ttk.Entry(wrapper, textvariable=v_code, width=30).pack(side="left")
            def do_pick():
                dlg = CodePicker(self, default_source=self.default_source)
                self.wait_window(dlg)
                if getattr(dlg, "selected_row", None):
                    code, name = dlg.selected_row
                    v_code.set(code); self.vars["收款方银行名称"].set(name or self.vars["收款方银行名称"].get())
            ttk.Button(wrapper, text="选择…", command=do_pick).pack(side="left", padx=6)
            return (wrapper, None)
        add("收款方银行大额支付行号/跨行清算行号", code_builder, "行号", "可选项"); self.vars["收款方银行大额支付行号/跨行清算行号"]=v_code
        v_use  = tk.StringVar(value=(init_values.get("用途","") if init_values else "")); add("用途", lambda p: ttk.Entry(p, textvariable=v_use)); self.vars["用途"]=v_use
        v_note = tk.StringVar(value=(init_values.get("明细标注","") if init_values else "")); add("明细标注", lambda p: ttk.Entry(p, textvariable=v_note)); self.vars["明细标注"]=v_note
        btns = ttk.Frame(self, padding=8); btns.pack(fill="x")
        ttk.Button(btns, text="确定", command=self.ok).pack(side="right", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")
        self.bind("<Return>", lambda e: self.ok())
    def ok(self):
        inverse_mode = {v:k for k,v in self.transfer_mode_map.items()}
        inverse_btype= {v:k for k,v in self.bankinfo_type_map.items()}
        v={}
        for k, var in self.vars.items():
            val = var.get().strip()
            if k=="转账方式": val = inverse_mode.get(val, "0")
            if k=="行别信息类型": val = inverse_btype.get(val, "0")
            v[k]=val
        probs=[]
        if not v["收款方户名"]: probs.append("收款方户名 不能为空")
        if not re.fullmatch(r"\d{6,32}", v["收款方账号"]): probs.append("收款方账号 必须为6-32位数字")
        try:
            if float(v["金额"]) <= 0: probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是有效数字")
        if v.get("转账方式","")=="1" and not v.get("收款方银行大额支付行号/跨行清算行号",""):
            probs.append("跨行转账时需提供行号")
        if probs: messagebox.showwarning("校验不通过","；".join(probs)); return
        self.values = v; self.destroy()

class TransferTab(ttk.Frame):
    COLS = ["收款方账号","收款方户名","金额","转账方式","行别信息类型",
            "收款方银行名称","收款方银行大额支付行号/跨行清算行号","用途","明细标注"]
    def __init__(self, master):
        super().__init__(master)
        self.df = pd.DataFrame(columns=self.COLS); self.source_choice = tk.StringVar(value="ibps"); self._build()
    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Button(top, text="新增一条", command=self.add_one).pack(side="left")
        ttk.Button(top, text="编辑选中", command=self.edit_one).pack(side="left", padx=6)
        ttk.Button(top, text="删除选中", command=self.delete_selected).pack(side="left", padx=6)
        ttk.Button(top, text="打开批量转账表", command=self.load_file).pack(side="left", padx=12)
        ttk.Radiobutton(top, text="使用 IBPS（清算行号）", variable=self.source_choice, value="ibps").pack(side="left", padx=12)
        ttk.Radiobutton(top, text="使用 CNAPS（大额行号）", variable=self.source_choice, value="cnaps").pack(side="left")
        ttk.Button(top, text="按选择填充行号", command=self.fill_codes).pack(side="left", padx=6)
        ttk.Button(top, text="校验并导出（保留表头）", command=self.validate_export).pack(side="left", padx=6)
        self.stree = ScrollableTree(self, height=18); self.stree.pack(fill="both", expand=True, padx=8, pady=6)
        self._reload()
    def add_one(self):
        dlg = TransferDialog(self, default_source=self.source_choice.get()); self.wait_window(dlg)
        if getattr(dlg, "values", None):
            self.df.loc[len(self.df)] = [dlg.values.get(c,"") for c in self.COLS]; self._reload()
    def edit_one(self):
        tree = self.stree.tree; sel = tree.selection()
        if not sel: messagebox.showinfo("提示","请先选择一行"); return
        idx = tree.index(sel[0]); init = {c: str(self.df.iloc[idx][c]) for c in self.COLS}
        dlg = TransferDialog(self, init_values=init, default_source=self.source_choice.get()); self.wait_window(dlg)
        if getattr(dlg, "values", None):
            for c in self.COLS: self.df.at[self.df.index[idx], c] = dlg.values.get(c,"")
            self._reload()
    def delete_selected(self):
        tree = self.stree.tree; sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0]); self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True); self._reload()
    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV","*.xlsx;*.xls;*.csv")])
        if not path: return
        try:
            df = read_any(path)
            new = pd.DataFrame(columns=self.COLS)
            for c in self.COLS: new[c] = df[c].astype(str) if c in df.columns else ""
            self.df = new[self.COLS]; self._reload()
        except Exception as e:
            messagebox.showerror("失败", f"读取失败：{e}")
    def fill_codes(self):
        lib = self.source_choice.get()
        if "收款方银行名称" not in self.df.columns:
            messagebox.showwarning("提示","缺少 收款方银行名称 列"); return
        filled = 0
        for i, r in self.df.iterrows():
            bank = str(r["收款方银行名称"] or "").strip()
            if not bank: continue
            hit = lookup_by_bankname(DB_PATH, lib, bank)
            if hit and hit.get("code"):
                self.df.at[i, "收款方银行大额支付行号/跨行清算行号"] = hit["code"]
                if not str(self.df.at[i, "收款方银行名称"]).strip():
                    self.df.at[i, "收款方银行名称"] = hit.get("name","")
                filled += 1
        self._reload(); messagebox.showinfo("结果", f"已填充 {filled} 条（来源：{lib.upper()}）")
    def validate_export(self):
        df = self.df.fillna("")
        probs = []
        if df["收款方户名"].str.strip().eq("").any(): probs.append("收款方户名 不能为空")
        if (~df["收款方账号"].astype(str).str.match(r"^\d{6,32}$", na=False)).any(): probs.append("收款方账号 格式异常")
        try:
            if (pd.to_numeric(df["金额"], errors="coerce")<=0).any(): probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是可解析数字")
        cross = df["转账方式"].astype(str).str.strip()=="1"
        need = cross & df["收款方银行大额支付行号/跨行清算行号"].astype(str).str.strip().eq("")
        if need.any(): probs.append("跨行转账时需提供行号（IBPS或CNAPS）")
        if probs: messagebox.showwarning("校验结果","；".join(probs))
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel",".xlsx")])
        if not path: return
        export_text_xlsx(self.df, path, include_header=True)
        messagebox.showinfo("成功","已导出（保留表头，文本格式）")

# ---------- Main App ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("华夏离线批量编辑器 v2.3")
        self.minsize(1200, 760)
        ensure_db(DB_PATH)
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True)
        nb.add(LibraryTab(nb), text="库维护（IBPS/CNAPS 导入）")
        nb.add(PayrollTab(nb), text="代发工资（逐条新增/编辑）")
        nb.add(TransferTab(nb), text="批量转账（逐条新增/编辑）")

if __name__ == "__main__":
    App().mainloop()
