
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import os, re
from db_helper import ensure_db, upsert_many_simple, replace_all, query, lookup_by_bankname

APP_DIR = Path(__file__).parent
DB_PATH = str(APP_DIR / "codebook.db")

# -------- utils --------
def detect_delimiter(sample: str):
    cands = ["|", "\t", ",", ";", " "]
    best, bestn = None, 1
    for d in cands:
        n = len(sample.split(d))
        if n > bestn:
            best, bestn = d, n
    return best or ","

def read_txt(path: str):
    data = Path(path).read_bytes()
    for enc in ["utf-8", "gbk", "gb18030", "utf-16", "latin1"]:
        try:
            s = data.decode(enc)
            first = next((ln for ln in s.splitlines() if ln.strip()), "")
            delim = detect_delimiter(first)
            from io import StringIO
            df = pd.read_csv(StringIO(s), sep=delim, header=None)
            return df
        except Exception:
            continue
    raise RuntimeError("无法解析 TXT，请确认编码或转存为 CSV/Excel。")

def read_any(path: str):
    p = Path(path)
    if p.suffix.lower() in [".xlsx",".xlsm",".xltx",".xltm"]:
        return pd.read_excel(path, engine="openpyxl")
    if p.suffix.lower() == ".xls":
        try:
            return pd.read_excel(path)
        except Exception as e:
            raise RuntimeError("读取 .xls 需安装 xlrd：pip install xlrd") from e
    if p.suffix.lower() == ".csv":
        return pd.read_csv(path, encoding="utf-8")
    if p.suffix.lower() in [".txt", ".dat"]:
        return read_txt(path)
    raise RuntimeError("无法解析文件，请转存为 CSV/Excel 再试。")

def export_text_xlsx(df: pd.DataFrame, path: str):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for j, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=j, value=str(col)); cell.number_format = "@"
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, col in enumerate(df.columns, start=1):
            val = "" if pd.isna(row[col]) else str(row[col])
            cell = ws.cell(row=i, column=j, value=val); cell.number_format = "@"
    wb.save(path)

# ------- Library maintenance -------
class LibraryTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.table_choice = tk.StringVar(value="ibps")
        self.kw = tk.StringVar()
        self._build()

    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="维护库：").pack(side="left")
        ttk.Radiobutton(top, text="IBPS（清算）", variable=self.table_choice, value="ibps").pack(side="left")
        ttk.Radiobutton(top, text="CNAPS（大额）", variable=self.table_choice, value="cnaps").pack(side="left", padx=8)
        ttk.Button(top, text="导入（txt/xls/xlsx/csv）", command=self.import_file).pack(side="left", padx=12)
        ttk.Button(top, text="导出库", command=self.export_db).pack(side="left")
        ttk.Label(top, text="关键词：").pack(side="left", padx=12)
        ttk.Entry(top, textvariable=self.kw, width=28).pack(side="left")
        ttk.Button(top, text="查询", command=self.search).pack(side="left", padx=6)

        self.tree = ttk.Treeview(self, show="headings")
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)

    def import_file(self):
        path = filedialog.askopenfilename(filetypes=[("TXT/Excel/CSV","*.txt;*.dat;*.xls;*.xlsx;*.csv"),("所有文件","*.*")])
        if not path: return
        try:
            df = read_any(path)
        except Exception as e:
            messagebox.showerror("失败", f"读取失败：{e}"); return

        rows = []
        if self.table_choice.get()=="cnaps":
            if df.shape[1] < 4:
                messagebox.showerror("失败", "CNAPS 导入需4列：BNKCODE, CLSCODE, CITYCODE, LNAME"); return
            df = df.iloc[:, :4]; df.columns = ["BNKCODE","CLSCODE","CITYCODE","LNAME"]
            for _, r in df.iterrows():
                code = str(r["BNKCODE"]).strip()
                name = str(r["LNAME"]).strip()
                if not re.fullmatch(r"\d{12}", code or ""): continue
                raw = "|".join([str(r[c]) for c in ["BNKCODE","CLSCODE","CITYCODE","LNAME"]])
                rows.append((code, name, raw, os.path.basename(path)))
            table="cnaps"
        else:
            if df.shape[1] < 2:
                messagebox.showerror("失败", "IBPS 导入需2列：清算行行号, 清算行名称"); return
            df = df.iloc[:, :2]; df.columns = ["code","name"]
            for _, r in df.iterrows():
                code = str(r["code"]).strip()
                name = str(r["name"]).strip()
                if not re.fullmatch(r"\d{12}", code or ""): continue
                raw = "|".join([str(r[c]) for c in ["code","name"]])
                rows.append((code, name, raw, os.path.basename(path)))
            table="ibps"

        if not rows:
            messagebox.showwarning("提示","未发现有效的12位行号记录"); return

        if messagebox.askyesno("导入方式", "选择“是”= 全量替换；“否”= 增量合并（按 code upsert）"):
            replace_all(DB_PATH, table, rows)
        else:
            upsert_many_simple(DB_PATH, table, rows)
        messagebox.showinfo("成功", f"导入完成：共 {len(rows)} 条")

    def search(self):
        table = self.table_choice.get()
        rows = query(DB_PATH, table, self.kw.get().strip(), limit=2000)
        import pandas as pd
        df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["code","name"])
        self._load_df(df)

    def export_db(self):
        from db_helper import query as q
        table = self.table_choice.get()
        rows = q(DB_PATH, table, "", limit=999999)
        if not rows:
            messagebox.showinfo("提示","当前库为空"); return
        import pandas as pd
        df = pd.DataFrame(rows)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel",".xlsx"),("CSV",".csv")])
        if not path: return
        try:
            if path.lower().endswith(".csv"):
                df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                export_text_xlsx(df, path)
            messagebox.showinfo("成功", f"已导出：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("失败", f"导出失败：{e}")

    def _load_df(self, df):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=200 if col=="name" else 160, anchor="w")
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[str(row.get(c,"")) for c in df.columns])

# ------- Dialogs -------
class PayrollDialog(tk.Toplevel):
    COLS = ["收款人银行名称","收款人卡号","收款人名称","金额"]
    def __init__(self, master, init_values=None):
        super().__init__(master)
        self.title("新增/编辑 - 代发工资")
        self.resizable(False, False)
        self.values = {}
        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        self.vars = {}
        for i, col in enumerate(self.COLS):
            ttk.Label(frm, text=col + "：").grid(row=i, column=0, sticky="e", padx=6, pady=6)
            var = tk.StringVar(value=(init_values.get(col,"") if init_values else ""))
            ent = ttk.Entry(frm, textvariable=var, width=36)
            ent.grid(row=i, column=1, sticky="w", padx=6, pady=6)
            self.vars[col] = var
        btns = ttk.Frame(frm); btns.grid(row=len(self.COLS), column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="确定", command=self.ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=8)
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
        self.values = vals
        self.destroy()

class TransferDialog(tk.Toplevel):
    COLS = ["收款方账号","收款方户名","金额","转账方式","行别信息类型",
            "收款方银行名称","收款方银行大额支付行号/跨行清算行号","用途","明细标注"]
    def __init__(self, master, init_values=None):
        super().__init__(master)
        self.title("新增/编辑 - 批量转账")
        self.resizable(False, False)
        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        self.vars = {}
        for i, col in enumerate(self.COLS):
            ttk.Label(frm, text=col + "：").grid(row=i, column=0, sticky="e", padx=6, pady=6)
            var = tk.StringVar(value=(init_values.get(col,"") if init_values else ""))
            ent = ttk.Entry(frm, textvariable=var, width=40)
            ent.grid(row=i, column=1, sticky="w", padx=6, pady=6)
            self.vars[col] = var
        btns = ttk.Frame(frm); btns.grid(row=len(self.COLS), column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="确定", command=self.ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=8)
        self.bind("<Return>", lambda e: self.ok())

    def ok(self):
        v = {k: var.get().strip() for k, var in self.vars.items()}
        probs = []
        if not v["收款方户名"]: probs.append("收款方户名 不能为空")
        if not re.fullmatch(r"\d{6,32}", v["收款方账号"]): probs.append("收款方账号 必须为6-32位数字")
        try:
            if float(v["金额"]) <= 0: probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是有效数字")
        if v.get("转账方式","") == "1" and not v.get("收款方银行大额支付行号/跨行清算行号",""):
            probs.append("跨行转账时需提供行号")
        if probs:
            messagebox.showwarning("校验不通过", "；".join(probs)); return
        self.values = v
        self.destroy()

# ------- Payroll tab -------
class PayrollTab(ttk.Frame):
    COLS = ["收款人银行名称","收款人卡号","收款人名称","金额"]
    def __init__(self, master):
        super().__init__(master)
        self.df = pd.DataFrame(columns=self.COLS)
        self._build()

    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Button(top, text="新增一条", command=self.add_one).pack(side="left")
        ttk.Button(top, text="编辑选中", command=self.edit_one).pack(side="left", padx=6)
        ttk.Button(top, text="删除选中", command=self.delete_selected).pack(side="left", padx=6)
        ttk.Button(top, text="打开代发工资表", command=self.load_file).pack(side="left", padx=12)
        ttk.Button(top, text="校验并导出", command=self.validate_export).pack(side="left", padx=6)

        self.tree = ttk.Treeview(self, show="headings", selectmode="browse")
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)
        self._reload()

    def add_one(self):
        dlg = PayrollDialog(self)
        self.wait_window(dlg)
        if getattr(dlg, "values", None):
            self.df.loc[len(self.df)] = [dlg.values[c] for c in self.COLS]
            self._reload()

    def edit_one(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示","请先选择一行"); return
        idx = self.tree.index(sel[0])
        init = {c: str(self.df.iloc[idx][c]) for c in self.COLS}
        dlg = PayrollDialog(self, init_values=init)
        self.wait_window(dlg)
        if getattr(dlg, "values", None):
            for c in self.COLS:
                self.df.at[self.df.index[idx], c] = dlg.values[c]
            self._reload()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True)
        self._reload()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV","*.xlsx;*.xls;*.csv")])
        if not path: return
        try:
            df = read_any(path)
            new = pd.DataFrame(columns=self.COLS)
            for c in self.COLS:
                new[c] = df[c].astype(str) if c in df.columns else ""
            self.df = new[self.COLS]
            self._reload()
        except Exception as e:
            messagebox.showerror("失败", f"读取失败：{e}")

    def validate_export(self):
        df = self.df.fillna("")
        probs = []
        if df["收款人银行名称"].str.strip().eq("").any(): probs.append("收款人银行名称 不能为空")
        if (~df["收款人卡号"].astype(str).str.match(r"^\d{6,32}$", na=False)).any(): probs.append("收款人卡号 格式异常")
        if df["收款人名称"].str.strip().eq("").any(): probs.append("收款人名称 不能为空")
        try:
            amt_bad = pd.to_numeric(df["金额"], errors="coerce")<=0
            if amt_bad.any(): probs.append("金额 必须大于0")
        except Exception:
            probs.append("金额 不是可解析数字")
        if probs:
            messagebox.showwarning("校验结果", "；".join(probs))
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel",".xlsx")])
        if not path: return
        export_text_xlsx(self.df, path)
        messagebox.showinfo("成功", "已导出（文本格式）")

    def _reload(self):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.COLS
        for col in self.COLS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=180, anchor="w")
        for i, row in self.df.iterrows():
            self.tree.insert("", "end", values=[str(row.get(c,"")) for c in self.COLS])

# ------- Transfer tab -------
class TransferTab(ttk.Frame):
    COLS = ["收款方账号","收款方户名","金额","转账方式","行别信息类型",
            "收款方银行名称","收款方银行大额支付行号/跨行清算行号","用途","明细标注"]
    def __init__(self, master):
        super().__init__(master)
        self.df = pd.DataFrame(columns=self.COLS)
        self.source_choice = tk.StringVar(value="ibps")
        self._build()

    def _build(self):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=8)
        ttk.Button(top, text="新增一条", command=self.add_one).pack(side="left")
        ttk.Button(top, text="编辑选中", command=self.edit_one).pack(side="left", padx=6)
        ttk.Button(top, text="删除选中", command=self.delete_selected).pack(side="left", padx=6)
        ttk.Button(top, text="打开批量转账表", command=self.load_file).pack(side="left", padx=12)
        ttk.Radiobutton(top, text="使用 IBPS（清算行号）", variable=self.source_choice, value="ibps").pack(side="left", padx=12)
        ttk.Radiobutton(top, text="使用 CNAPS（大额行号）", variable=self.source_choice, value="cnaps").pack(side="left")
        ttk.Button(top, text="按选择填充行号", command=self.fill_codes).pack(side="left", padx=6)
        ttk.Button(top, text="校验并导出", command=self.validate_export).pack(side="left", padx=6)

        self.tree = ttk.Treeview(self, show="headings", selectmode="browse")
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)
        self._reload()

    def add_one(self):
        dlg = TransferDialog(self)
        self.wait_window(dlg)
        if getattr(dlg, "values", None):
            self.df.loc[len(self.df)] = [dlg.values.get(c,"") for c in self.COLS]
            self._reload()

    def edit_one(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示","请先选择一行"); return
        idx = self.tree.index(sel[0])
        init = {c: str(self.df.iloc[idx][c]) for c in self.COLS}
        dlg = TransferDialog(self, init_values=init)
        self.wait_window(dlg)
        if getattr(dlg, "values", None):
            for c in self.COLS:
                self.df.at[self.df.index[idx], c] = dlg.values.get(c,"")
            self._reload()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True)
        self._reload()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV","*.xlsx;*.xls;*.csv")])
        if not path: return
        try:
            df = read_any(path)
            new = pd.DataFrame(columns=self.COLS)
            for c in self.COLS:
                new[c] = df[c].astype(str) if c in df.columns else ""
            self.df = new[self.COLS]
            self._reload()
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
                filled += 1
        self._reload()
        messagebox.showinfo("结果", f"已填充 {filled} 条（来源：{lib.upper()}）")

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
        if probs:
            messagebox.showwarning("校验结果","；".join(probs))
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel",".xlsx")])
        if not path: return
        export_text_xlsx(self.df, path)
        messagebox.showinfo("成功","已导出（文本格式）")

    def _reload(self):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.COLS
        for col in self.COLS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=200 if col=="收款方银行大额支付行号/跨行清算行号" else 160, anchor="w")
        for i, row in self.df.iterrows():
            self.tree.insert("", "end", values=[str(row.get(c,"")) for c in self.COLS])

# ------- Main -------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("华夏离线批量编辑器（v2：支持TXT导入 & 新增对话框）")
        self.geometry("1180x740")
        ensure_db(DB_PATH)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        nb.add(LibraryTab(nb), text="库维护（IBPS/CNAPS 导入）")
        nb.add(PayrollTab(nb), text="代发工资（逐条新增/编辑）")
        nb.add(TransferTab(nb), text="批量转账（逐条新增/编辑）")

if __name__ == "__main__":
    App().mainloop()
