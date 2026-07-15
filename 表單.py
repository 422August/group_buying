#!/usr/bin/env python3
# coding: utf-8
"""
FB group-buy order-sheet generator · v2.8 (No Price Version)
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from typing import Dict, Union, Optional, Any
import pandas as pd


# ── 業務邏輯與解析器 ───────────────────────────────────────────
class OrderParser:
    """處理留言解析與數量加總的邏輯類別"""

    RE_NAME_PLUS = re.compile(
        r"^([\u4e00-\u9fa5A-Za-zA-Z\s·\._-]+?)\s+\+\s*([0-9]+)\s*$"
    )
    RE_NAME_LINE = re.compile(r"^[\u4e00-\u9fa5A-Za-zA-Z\s·\._-]+$")
    RE_PURE_NUM = re.compile(r"^\s*\+?\s*([0-9]+)\s*$")
    RE_ANY_LETTER_PLUS = re.compile(r"[A-Za-z]\s*\+\s*\d+")
    RE_TIME = re.compile(r"^\d+\s*(天|日|週|周|小時|分鐘)$")
    NOISE_WORDS = {"回覆", "讚", "翻譯年糕", "·"}

    @classmethod
    def parse(cls, text: str) -> Dict[str, Union[int, str]]:
        buyers: Dict[str, Union[int, str]] = {}
        current_name: Optional[str] = None

        for line in text.splitlines():
            line = line.strip()

            if (
                not line
                or "已編輯" in line
                or line in cls.NOISE_WORDS
                or cls.RE_TIME.match(line)
            ):
                continue

            if current_name:
                cls._handle_qty_line(buyers, current_name, line)
                current_name = None
                continue

            match_plus = cls.RE_NAME_PLUS.match(line)
            if match_plus:
                name, qty = match_plus.group(1).strip(), int(match_plus.group(2))
                cls.merge_value(buyers, name, qty)
                continue

            if cls.RE_NAME_LINE.match(line):
                current_name = line

        return buyers

    @classmethod
    def _handle_qty_line(cls, buyers: Dict[str, Any], name: str, line: str) -> None:
        match_num = cls.RE_PURE_NUM.match(line)
        if match_num and not cls.RE_ANY_LETTER_PLUS.search(line):
            cls.merge_value(buyers, name, int(match_num.group(1)))
        else:
            cls.merge_value(buyers, name, line)

    @staticmethod
    def merge_value(
        target_dict: Dict[str, Any], key: str, new_val: Union[int, str]
    ) -> None:
        old_val = target_dict.get(key)
        if old_val in (None, ""):
            target_dict[key] = new_val
        elif isinstance(old_val, int) and isinstance(new_val, int):
            target_dict[key] = old_val + new_val
        else:
            target_dict[key] = ",".join(
                sorted({str(old_val).strip(), str(new_val).strip()})
            )


# ── 排序鍵 ─────────────────────────────────────────────
def get_sort_key(text: str) -> tuple[int, str]:
    """統一處理姓名與商品名稱的排序邏輯 (英文優先，中文在後)"""
    first_char = text.lstrip()[0]
    if first_char.isascii() and first_char.isalpha():
        return (0, text.casefold())
    return (1, text)


# ── GUI 應用程式 ──────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("團購訂單表產生器")
        self.geometry("980x620")

        self.data: Dict[str, Dict[str, Any]] = defaultdict(dict)
        self.df_display: Optional[pd.DataFrame] = None
        self._build_ui()

    def _build_ui(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=8, pady=4)

        ttk.Label(
            top_frame, text="商品名稱：", font=("Microsoft JhengHei", 20, "bold")
        ).grid(row=0, column=0, sticky="w")
        self.ent_item = ttk.Entry(top_frame, width=60, font=("Consolas", 20))
        self.ent_item.grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(
            top_frame, text="貼文內容：", font=("Microsoft JhengHei", 20, "bold")
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.txt_post = tk.Text(
            top_frame, width=100, height=10, wrap="word", font=("Consolas", 20)
        )
        self.txt_post.grid(row=2, column=0, columnspan=2, sticky="we", pady=(0, 6))

        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, sticky="we")
        ttk.Button(btn_frame, text="新增商品", command=self.on_add).pack(
            side="left", padx=6
        )
        ttk.Button(btn_frame, text="匯出 Excel", command=self.on_export).pack(
            side="left", padx=6
        )

        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree = ttk.Treeview(table_frame, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def on_add(self):
        item_name = self.ent_item.get().strip()

        if not item_name:
            messagebox.showwarning("缺少商品名稱", "請輸入『商品名稱』")
            return

        raw_text = self.txt_post.get("1.0", "end")
        parsed_buyers = OrderParser.parse(raw_text)

        if not parsed_buyers:
            messagebox.showwarning("無符合格式", "找不到任何有效訂購資訊")
            return

        for buyer, val in parsed_buyers.items():
            OrderParser.merge_value(self.data[buyer], item_name, val)

        self.ent_item.delete(0, "end")
        self.txt_post.delete("1.0", "end")
        self._refresh_tree()

    def on_export(self):
        if self.df_display is None or self.df_display.empty:
            messagebox.showwarning("無資料", "目前沒有可匯出的資料")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel 活頁簿", "*.xlsx")]
        )

        if file_path:
            self._get_export_df().to_excel(file_path, index=False)
            messagebox.showinfo("已匯出", f"檔案已儲存至：\n{file_path}")

    def _refresh_tree(self):
        all_items = sorted(
            {item for orders in self.data.values() for item in orders}, key=get_sort_key
        )
        all_buyers = sorted(self.data.keys(), key=get_sort_key)

        rows = [self.data[buyer] for buyer in all_buyers]
        df = pd.DataFrame(rows, index=all_buyers).reindex(columns=all_items).fillna("")
        df.index.name = "姓名"
        df.reset_index(inplace=True)
        self.df_display = df

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=max(100, int(len(col) * 14)))

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def _get_export_df(self) -> pd.DataFrame:
        df = self.df_display.copy()
        df.insert(0, "", "")
        df.rename(columns={"姓名": ""}, inplace=True)
        return df


if __name__ == "__main__":
    App().mainloop()
