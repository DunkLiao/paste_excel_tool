# -*- coding: utf-8 -*-
"""
專案名稱: 信用風險
功能描述: 根據「EXCEL貼值清單」將來源儲存格的值貼到目的儲存格 (VBA PasteExcel 的 Python + ttk 版本)
執行環境: Windows + 已安裝 Microsoft Excel + pywin32 (pip install pywin32)

清單欄位 (工作頁: EXCEL貼值清單，第 1 列為標題，第 2 列起為資料)
    A 來源EXCEL檔案路徑
    B 來源工作頁名稱 (找不到名稱且為數字時，改以第 N 張工作表)
    C 來源儲存格位置 (i.e. A1:B2)
    D 來源範圍是否為UsedRange(Y)
    E 目的EXCEL檔案路徑
    F 目的工作頁名稱
    G 目的儲存格位置 (i.e. A1:B2)
    H 目的儲存格為合併儲存格(Y)
    I 錯誤訊息 (成功時顯示「成功」)

處理規則 (與 VBA 版一致)
    1. 相同檔案只開啟一次，全部列處理完才統一儲存並關閉
    2. 來源檔若同時為目的檔，以可寫入方式開啟
    3. 複製貼值失敗的列，不會讓該列觸發儲存
    4. H = Y 時以 Value 指定貼值；否則以 Copy + PasteSpecial(貼上值)
"""

import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pythoncom
    import win32com.client as win32
except ImportError:  # 非 Windows 或未安裝 pywin32
    pythoncom = None
    win32 = None

# ---------------------------------------------------------------- 常數
DEFAULT_SHEET_LIST = "EXCEL貼值清單"

COL_SRC_FILE = 1    # A
COL_SRC_SHEET = 2   # B
COL_SRC_RANGE = 3   # C
COL_SRC_USED = 4    # D
COL_TGT_FILE = 5    # E
COL_TGT_SHEET = 6   # F
COL_TGT_RANGE = 7   # G
COL_TGT_MERGE = 8   # H
COL_ERR = 9         # I

MSG_OK = "成功"
HDR_SRC_USED = "來源範圍是否為UsedRange(Y)"
HDR_ERR = "錯誤訊息"

XL_UP = -4162
XL_PASTE_VALUES = -4163


# ---------------------------------------------------------------- 工具函式
def to_text(value):
    """將儲存格值轉成去除前後空白的字串 (1.0 -> '1')。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def get_sheet(wb, sheet_name):
    """先以名稱找 (不分大小寫)；找不到且為數字時，以第 N 張工作表索引找。"""
    if not sheet_name:
        return None
    for ws in wb.Worksheets:
        if ws.Name.lower() == sheet_name.lower():
            return ws
    try:
        idx = int(float(sheet_name))
    except ValueError:
        return None
    if 1 <= idx <= wb.Worksheets.Count:
        return wb.Worksheets(idx)
    return None


# ---------------------------------------------------------------- 核心處理
class PasteEngine:
    def __init__(self, list_file, list_sheet, log, progress, cancel_event):
        self.list_file = list_file
        self.list_sheet = list_sheet
        self.log = log                  # log(row, result, message)
        self.progress = progress        # progress(current, total)
        self.cancel = cancel_event
        self.app = None
        self.pool = {}                  # key -> workbook
        self.owned = set()              # 由本程式開啟、需由本程式關閉
        self.targets = set()            # 需以可寫入開啟的檔案 key
        self.dirty = set()              # 有成功貼值、需儲存
        self.dirty_rows = {}            # key -> [列號]
        self.results = {}               # 列號 -> 訊息

    # ---- 活頁簿管理
    def open_workbook(self, path, writable):
        if not path:
            raise ValueError("檔案路徑為空白")
        key = path_key(path)
        if key in self.pool:
            return self.pool[key]
        if not os.path.isfile(path):
            raise FileNotFoundError("檔案不存在: " + path)
        wb = self.app.Workbooks.Open(
            os.path.abspath(path), UpdateLinks=0, ReadOnly=not writable
        )
        self.pool[key] = wb
        self.owned.add(key)
        return wb

    def close_all(self):
        for key in list(self.owned):
            try:
                self.pool[key].Close(SaveChanges=False)
            except Exception:
                pass
        self.pool.clear()
        self.owned.clear()

    # ---- 主流程
    def run(self):
        ok_count = fail_count = 0
        pythoncom.CoInitialize()
        try:
            self.app = win32.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.ScreenUpdating = False

            list_key = path_key(self.list_file)
            self.targets.add(list_key)
            list_wb = self.open_workbook(self.list_file, True)
            try:
                ws = list_wb.Worksheets(self.list_sheet)
            except Exception:
                raise RuntimeError("清單檔中找不到工作頁: " + self.list_sheet)

            ws.Cells(1, COL_SRC_USED).Value = HDR_SRC_USED
            ws.Cells(1, COL_ERR).Value = HDR_ERR

            last_row = ws.Cells(ws.Rows.Count, COL_SRC_FILE).End(XL_UP).Row
            if last_row < 2:
                self.log(0, "略過", "清單無資料")
                return 0, 0

            # 讀取清單 (一次讀入記憶體)
            rows = []
            for i in range(2, last_row + 1):
                rows.append((i, [to_text(ws.Cells(i, c).Value)
                                 for c in range(1, COL_ERR)]))
            ws.Range(ws.Cells(2, COL_ERR), ws.Cells(last_row, COL_ERR)).ClearContents()

            # 先蒐集所有目的檔
            for _, r in rows:
                if r[COL_TGT_FILE - 1]:
                    self.targets.add(path_key(r[COL_TGT_FILE - 1]))

            # 逐列處理
            total = len(rows)
            for n, (i, r) in enumerate(rows, 1):
                if self.cancel.is_set():
                    self.log(i, "中止", "使用者已中止，後續列未處理")
                    break
                self.progress(n, total)
                if not r[COL_SRC_FILE - 1] and not r[COL_TGT_FILE - 1]:
                    continue
                try:
                    msg = self.process_row(i, r)
                except Exception as e:  # 未預期錯誤只影響該列
                    msg = "處理時發生錯誤: " + str(e)
                self.results[i] = msg or MSG_OK
                self.log(i, "成功" if not msg else "失敗", msg or "")

            # 統一儲存有異動的目的檔 (清單檔最後再存)
            for key in list(self.dirty):
                if key == list_key:
                    continue
                try:
                    self.pool[key].Save()
                except Exception as e:
                    err = "儲存目的檔案時發生錯誤: " + str(e)
                    for rn in self.dirty_rows.get(key, []):
                        self.results[rn] = err
                        self.log(rn, "失敗", err)

            # 回寫結果到清單
            for rn, msg in self.results.items():
                ws.Cells(rn, COL_ERR).Value = msg
                if msg == MSG_OK:
                    ok_count += 1
                else:
                    fail_count += 1
            list_wb.Save()
            return ok_count, fail_count
        finally:
            self.close_all()
            try:
                if self.app is not None:
                    self.app.DisplayAlerts = True
                    self.app.ScreenUpdating = True
                    self.app.Quit()
            except Exception:
                pass
            self.app = None
            pythoncom.CoUninitialize()

    def process_row(self, i, r):
        """處理單列，回傳空字串代表成功，否則回傳錯誤訊息。"""
        src_file = r[COL_SRC_FILE - 1]
        src_sheet = r[COL_SRC_SHEET - 1]
        src_addr = r[COL_SRC_RANGE - 1]
        use_used = r[COL_SRC_USED - 1].upper() == "Y"
        tgt_file = r[COL_TGT_FILE - 1]
        tgt_sheet = r[COL_TGT_SHEET - 1]
        tgt_addr = r[COL_TGT_RANGE - 1]
        is_merge = r[COL_TGT_MERGE - 1].upper() == "Y"

        # 開啟檔案
        try:
            src_wb = self.open_workbook(src_file, path_key(src_file) in self.targets)
        except Exception as e:
            return "無法開啟來源檔案: " + str(e)
        try:
            tgt_wb = self.open_workbook(tgt_file, True)
        except Exception as e:
            return "無法開啟目的檔案: " + str(e)

        # 工作表
        src_ws = get_sheet(src_wb, src_sheet)
        if src_ws is None:
            return "來源工作表不存在: " + src_sheet
        tgt_ws = get_sheet(tgt_wb, tgt_sheet)
        if tgt_ws is None:
            return "目的工作表不存在: " + tgt_sheet

        # 範圍
        try:
            src_range = src_ws.UsedRange if use_used else src_ws.Range(src_addr)
        except Exception:
            return "來源範圍無效: " + src_addr
        try:
            tgt_range = tgt_ws.Range(tgt_addr)
        except Exception:
            return "目的範圍無效: " + tgt_addr

        # 複製貼值
        try:
            if is_merge:
                tgt_range.Value = src_range.Value
            else:
                src_range.Copy()
                tgt_range.PasteSpecial(XL_PASTE_VALUES)
                self.app.CutCopyMode = False
        except Exception as e:
            try:
                self.app.CutCopyMode = False
            except Exception:
                pass
            return "複製資料時發生錯誤: " + str(e)

        key = path_key(tgt_file)
        self.dirty.add(key)
        self.dirty_rows.setdefault(key, []).append(i)
        return ""


# ---------------------------------------------------------------- ttk UI
class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.master = master
        self.q = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker = None

        self.var_file = tk.StringVar()
        self.var_sheet = tk.StringVar(value=DEFAULT_SHEET_LIST)
        self.var_status = tk.StringVar(value="請選擇含「EXCEL貼值清單」的 Excel 檔")

        self._build()
        self.pack(fill="both", expand=True)
        self.after(100, self._poll)

    def _build(self):
        self.master.title("EXCEL 貼值處理工具")
        self.master.geometry("900x560")
        self.master.minsize(760, 440)

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        frm = ttk.LabelFrame(self, text="設定", padding=10)
        frm.pack(fill="x")
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="清單檔案:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(frm, textvariable=self.var_file).grid(
            row=0, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(frm, text="瀏覽…", command=self._browse).grid(row=0, column=2, pady=3)

        ttk.Label(frm, text="清單工作頁:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(frm, textvariable=self.var_sheet, width=30).grid(
            row=1, column=1, sticky="w", padx=6, pady=3)

        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=8)
        self.btn_start = ttk.Button(btns, text="開始執行", command=self._start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(btns, text="中止", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        ttk.Button(btns, text="清除記錄", command=self._clear).pack(side="left")
        ttk.Button(btns, text="離開", command=self.master.destroy).pack(side="right")

        self.pb = ttk.Progressbar(self, mode="determinate")
        self.pb.pack(fill="x")
        ttk.Label(self, textvariable=self.var_status).pack(anchor="w", pady=(4, 6))

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True)
        cols = ("row", "result", "msg")
        self.tree = ttk.Treeview(box, columns=cols, show="headings")
        self.tree.heading("row", text="清單列號")
        self.tree.heading("result", text="結果")
        self.tree.heading("msg", text="訊息")
        self.tree.column("row", width=80, anchor="center", stretch=False)
        self.tree.column("result", width=80, anchor="center", stretch=False)
        self.tree.column("msg", width=600, anchor="w")
        self.tree.tag_configure("ok", foreground="#0a7a2f")
        self.tree.tag_configure("fail", foreground="#c00000")
        self.tree.tag_configure("info", foreground="#555555")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ---- 事件
    def _browse(self):
        path = filedialog.askopenfilename(
            title="選擇清單 Excel 檔",
            filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xlsb *.xls"), ("所有檔案", "*.*")])
        if path:
            self.var_file.set(os.path.normpath(path))

    def _clear(self):
        self.tree.delete(*self.tree.get_children())

    def _start(self):
        if win32 is None:
            messagebox.showerror("缺少套件", "請先安裝 pywin32:\n\npip install pywin32\n\n(且需 Windows + Excel)")
            return
        path = self.var_file.get().strip()
        sheet = self.var_sheet.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提醒", "請選擇存在的清單 Excel 檔")
            return
        if not sheet:
            messagebox.showwarning("提醒", "請輸入清單工作頁名稱")
            return

        self._clear()
        self.cancel_event.clear()
        self.pb["value"] = 0
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.var_status.set("處理中…")

        engine = PasteEngine(
            path, sheet,
            log=lambda r, res, m: self.q.put(("log", r, res, m)),
            progress=lambda c, t: self.q.put(("progress", c, t)),
            cancel_event=self.cancel_event)

        def job():
            try:
                ok, fail = engine.run()
                self.q.put(("done", ok, fail))
            except Exception as e:
                self.q.put(("error", str(e), traceback.format_exc()))

        self.worker = threading.Thread(target=job, daemon=True)
        self.worker.start()

    def _stop(self):
        self.cancel_event.set()
        self.var_status.set("中止中，完成目前這一列後停止…")

    def _finish_ui(self):
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    _, row, res, msg = item
                    tag = {"成功": "ok", "失敗": "fail"}.get(res, "info")
                    self.tree.insert("", "end", values=(row, res, msg), tags=(tag,))
                    self.tree.yview_moveto(1)
                elif kind == "progress":
                    _, cur, total = item
                    self.pb["maximum"] = total
                    self.pb["value"] = cur
                    self.var_status.set("處理中… %d / %d" % (cur, total))
                elif kind == "done":
                    _, ok, fail = item
                    self.var_status.set("處理完成: 成功 %d 筆、失敗 %d 筆" % (ok, fail))
                    self._finish_ui()
                    messagebox.showinfo(
                        "處理完成",
                        "成功 %d 筆、失敗 %d 筆。\n\n請查看清單「%s」的「錯誤訊息」欄位。"
                        % (ok, fail, self.var_sheet.get()))
                elif kind == "error":
                    _, msg, tb = item
                    self.var_status.set("發生錯誤")
                    self._finish_ui()
                    self.tree.insert("", "end", values=("-", "失敗", msg), tags=("fail",))
                    messagebox.showerror("程式發生未預期錯誤", msg)
        except queue.Empty:
            pass
        self.after(100, self._poll)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
