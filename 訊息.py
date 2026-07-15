import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import pyperclip


# ---------------------- 訂購訊息生成器 ----------------------
class OrderMessageGenerator(ctk.CTkFrame):
    def __init__(self, parent, status_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.status_callback = status_callback  # 用以更新狀態列
        self.data = None
        self.messages = {}  # 儲存每位訂購人的訂購訊息
        self.order_details = (
            {}
        )  # 儲存每位訂購人各產品訂購數量（字典：姓名 -> {產品名稱: 數量, ...}）
        self.copied_names = set()  # 紀錄已複製的訂購人（需要保持反白）
        self.filter_id = None  # 搜尋延遲 id
        self.editing = False  # 編輯模式旗標
        self.create_widgets()

    def set_status(self, message):
        if self.status_callback:
            self.status_callback(message)

    def create_widgets(self):
        # 上方工具列：上傳、搜尋、依產品排序、匯出
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=5)

        self.upload_btn = ctk.CTkButton(
            top_frame, text="上傳 Excel 檔案", command=self.upload_file
        )
        self.upload_btn.grid(row=0, column=0, padx=5, pady=5)

        search_label = ctk.CTkLabel(top_frame, text="搜尋訂購人：")
        search_label.grid(row=0, column=1, padx=5, pady=5)

        self.search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(top_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        search_entry.bind("<KeyRelease>", lambda event: self.delayed_filter())
        top_frame.grid_columnconfigure(2, weight=1)

        sort_label = ctk.CTkLabel(top_frame, text="依產品排序:")
        sort_label.grid(row=0, column=3, padx=5, pady=5)

        self.product_sort_combobox = ctk.CTkComboBox(
            top_frame, values=[], width=150, state="readonly"
        )
        self.product_sort_combobox.grid(row=0, column=4, padx=5, pady=5)

        sort_btn = ctk.CTkButton(top_frame, text="排序", command=self.sort_by_product)
        sort_btn.grid(row=0, column=5, padx=5, pady=5)

        export_btn = ctk.CTkButton(
            top_frame, text="匯出訊息", command=self.export_messages
        )
        export_btn.grid(row=0, column=6, padx=5, pady=5)

        # 主區域：左側訂購人列表 / 右側訂購訊息顯示與編輯
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 左側：訂購人列表（由於 customtkinter 尚無 Listbox 元件，保留 tk.Listbox）
        left_frame = ctk.CTkFrame(main_frame, border_width=2)
        left_frame.pack(side="left", fill="y", padx=(0, 5), pady=5)
        list_label = ctk.CTkLabel(left_frame, text="訂購人列表")
        list_label.pack(anchor="nw", padx=5, pady=5)
        self.order_listbox = tk.Listbox(
            left_frame, width=30, exportselection=False, font=("微軟正黑體", 14)
        )
        self.order_listbox.pack(side="left", fill="y", expand=True, padx=5, pady=5)
        list_scrollbar = tk.Scrollbar(
            left_frame, orient="vertical", command=self.order_listbox.yview
        )
        list_scrollbar.pack(side="right", fill="y")
        self.order_listbox.config(yscrollcommand=list_scrollbar.set)
        self.order_listbox.bind("<<ListboxSelect>>", self.show_order_detail)

        # 右側：訂購訊息詳情與編輯（使用 CTkTextbox 取代 ScrolledText）
        right_frame = ctk.CTkFrame(main_frame, border_width=2)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0), pady=5)
        detail_label = ctk.CTkLabel(right_frame, text="訂購訊息")
        detail_label.pack(anchor="nw", padx=5, pady=5)
        self.detail_text = ctk.CTkTextbox(right_frame, wrap="word")
        self.detail_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.detail_text.configure(state="disabled")

        btn_frame = ctk.CTkFrame(right_frame)
        btn_frame.pack(pady=5)
        copy_btn = ctk.CTkButton(btn_frame, text="複製訊息", command=self.copy_message)
        copy_btn.grid(row=0, column=0, padx=5)
        # 新增「編輯訂單」按鈕，點選後可切換編輯/儲存模式
        self.edit_btn = ctk.CTkButton(
            btn_frame, text="編輯訂單", command=self.toggle_edit_order
        )
        self.edit_btn.grid(row=0, column=1, padx=5)

    def delayed_filter(self):
        if self.filter_id:
            self.after_cancel(self.filter_id)
        self.filter_id = self.after(300, self.filter_orders)

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        try:
            self.data = pd.read_excel(file_path, engine="openpyxl")
            self.generate_order_messages()
            self.populate_order_list()
            self.update_product_sort_options()
            self.set_status("Excel檔案上傳成功")
        except FileNotFoundError:
            messagebox.showerror("錯誤", "找不到檔案，請確認路徑是否正確")
            self.set_status("檔案讀取失敗")
        except ValueError:
            messagebox.showerror("錯誤", "Excel 檔案格式有誤，請確認檔案內容")
            self.set_status("檔案讀取失敗")
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取檔案時出現錯誤：{str(e)}")
            self.set_status("檔案讀取失敗")

    def parse_order_quantity(self, raw_quantity):
        if pd.isna(raw_quantity) or str(raw_quantity).strip() == "":
            return 0, "你沒訂"
        try:
            numeric_quantity = float(raw_quantity)
            if numeric_quantity.is_integer():
                numeric_quantity = int(numeric_quantity)
            return numeric_quantity, numeric_quantity
        except Exception:
            return raw_quantity, str(raw_quantity).strip()

    def generate_order_messages(self):
        """
        根據 Excel 資料生成訂購訊息：
          - 第一行為各品項價格（空值顯示「無價格」）。
          - 第二行起為訂購資料，從第3欄開始為各產品訂購數量。
        並計算每筆訂購的總金額（僅針對價格為數值的產品）。
        當訂購數量儲存格有內容（例如 "M+1"），則顯示原字串；僅當儲存格空白或數值為 0 時，
        顯示「你沒訂」且不列入訂購項目。
        """
        if self.data is None or self.data.empty:
            return
        self.messages.clear()
        self.order_details.clear()
        today_date = datetime.now().strftime("%m/%d")
        header_row = self.data.iloc[0]
        orders_data = self.data.iloc[1:].fillna("")
        for _, row in orders_data.iterrows():
            name = row.iloc[1]  # 假設第2欄為訂購人姓名
            if not str(name).strip():
                continue
            message_lines = [
                f"*熊熊媽團團轉{today_date}訂購清單*",
                "",
                "#取貨時間三點到七點",
                "#本日到貨狀況請留意公告",
                "================",
                f"訂購人：{name}",
            ]
            has_order = False
            order_items = []  # 收集有訂購的品項（供排序）
            product_quantities = {}  # 記錄各產品訂購數量
            total_amount = 0.0  # 計算該訂單總金額
            for col_index in range(2, len(row)):
                item_name = self.data.columns[col_index]
                raw_price = self.data.iloc[0, col_index]
                numeric_price = None
                if pd.isna(raw_price) or str(raw_price).strip() == "":
                    item_price_str = "無價格"
                else:
                    try:
                        numeric_price = float(str(raw_price))
                        if numeric_price.is_integer():
                            numeric_price = int(numeric_price)
                            item_price_str = str(numeric_price)
                        else:
                            item_price_str = f"{numeric_price:.2f}"
                    except Exception:
                        item_price_str = str(raw_price)
                raw_quantity = row.iloc[col_index]
                quantity, item_quantity = self.parse_order_quantity(raw_quantity)
                product_quantities[item_name] = quantity
                if (isinstance(quantity, (int, float)) and quantity > 0) or (
                    not isinstance(quantity, (int, float))
                    and str(quantity).strip() not in ["", "0"]
                ):
                    has_order = True
                    order_items.append((item_name, item_price_str, item_quantity))
                    if isinstance(quantity, (int, float)) and (
                        numeric_price is not None
                    ):
                        total_amount += numeric_price * quantity
            self.order_details[name] = product_quantities
            if has_order:
                order_items.sort(key=lambda x: x[0])
                for item_name, item_price_str, item_quantity in order_items:
                    message_lines.extend(
                        [
                            f"訂購商品：{item_name}",
                            f"品項單價：{item_price_str}",
                            f"數量品項：{item_quantity}",
                            "----------------",
                        ]
                    )
                message_lines.append("已讀請回覆訊息喔~")
                final_message = "\n".join(message_lines)
                self.messages[name] = final_message

    def populate_order_list(self):
        self.order_listbox.delete(0, tk.END)
        sorted_names = sorted(self.messages.keys())
        for idx, name in enumerate(sorted_names):
            self.order_listbox.insert(tk.END, name)
            if name in self.copied_names:
                self.order_listbox.itemconfig(idx, bg="black", fg="white")

    def filter_orders(self):
        search_term = self.search_var.get().strip().lower()
        self.order_listbox.delete(0, tk.END)
        sorted_names = sorted(self.messages.keys())
        for idx, name in enumerate(sorted_names):
            if search_term in name.lower():
                self.order_listbox.insert(tk.END, name)
                if name in self.copied_names:
                    self.order_listbox.itemconfig(idx, bg="black", fg="white")

    def show_order_detail(self, event):
        selection = self.order_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        name = self.order_listbox.get(index)
        message = self.messages.get(name, "")
        if not self.editing:
            self.detail_text.configure(state="normal")
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.insert(tk.END, message)
            self.detail_text.configure(state="disabled")
        else:
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.insert(tk.END, message)

    def copy_message(self):
        selection = self.order_listbox.curselection()
        if not selection:
            self.set_status("請先選擇一個訂購人。")
            return
        index = selection[0]
        name = self.order_listbox.get(index)
        message = self.messages.get(name, "")
        if message:
            pyperclip.copy(message)
            self.set_status("訊息已複製到剪貼簿")
            self.order_listbox.itemconfig(index, bg="black", fg="white")
            self.copied_names.add(name)
        else:
            self.set_status("無法複製空訊息")

    def toggle_edit_order(self):
        selection = self.order_listbox.curselection()
        if not selection:
            self.set_status("請先選擇一個訂購人。")
            return
        index = selection[0]
        name = self.order_listbox.get(index)
        if not self.editing:
            # 進入編輯模式
            self.detail_text.configure(state="normal")
            self.edit_btn.configure(text="儲存訂單")
            self.editing = True
            self.set_status("編輯模式：請修改訂單內容後按儲存")
        else:
            # 儲存修改並離開編輯模式
            new_text = self.detail_text.get("1.0", tk.END).strip()
            if new_text:
                self.messages[name] = new_text
                self.detail_text.configure(state="disabled")
                self.edit_btn.configure(text="編輯訂單")
                self.editing = False
                self.set_status(f"訂單 {name} 已更新")
            else:
                self.set_status("訂單內容不可為空")

    def update_product_sort_options(self):
        if self.data is not None:
            product_names = list(self.data.columns[2:])
            product_names.insert(0, "依字母排序")
            self.product_sort_combobox.configure(values=product_names)
            if product_names:
                self.product_sort_combobox.set(product_names[0])

    def sort_by_product(self):
        selected_product = self.product_sort_combobox.get()
        if not selected_product:
            messagebox.showwarning("提示", "請先選擇一個產品。")
            return

        def has_order(name):
            """判斷此人是否有訂購 selected_product。"""
            q = self.order_details.get(name, {}).get(selected_product, 0)
            if isinstance(q, (int, float)):
                return q > 0
            return str(q).strip() not in ["", "0"]

        # 依 flag (0=有訂購, 1=沒訂購) + 姓名 Unicode 排序
        sorted_names = sorted(
            self.messages.keys(), key=lambda n: (0 if has_order(n) else 1, n)
        )

        # 更新 Listbox
        self.order_listbox.delete(0, tk.END)
        for idx, n in enumerate(sorted_names):
            self.order_listbox.insert(tk.END, n)
            if n in self.copied_names:
                self.order_listbox.itemconfig(idx, bg="black", fg="white")

        self.set_status("依產品排序完成")

    def export_messages(self):
        if not self.messages:
            messagebox.showinfo("提示", "目前無任何訊息可供匯出。")
            self.set_status("無訊息匯出")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                for name, message in self.messages.items():
                    f.write(f"訂購人：{name}\n")
                    f.write(message)
                    f.write("\n" + "=" * 40 + "\n")
            messagebox.showinfo("提示", "訊息匯出成功。")
            self.set_status("訊息匯出成功")
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出訊息時發生錯誤：{str(e)}")
            self.set_status("訊息匯出失敗")


# ---------------------- 主程式 ----------------------
def main():
    ctk.set_appearance_mode("System")  # 可設定 "Light", "Dark" 或 "System"
    root = ctk.CTk()
    root.title("訂購訊息生成器")
    root.geometry("900x700")

    status_var = tk.StringVar()
    status_var.set("準備就緒")
    status_bar = ctk.CTkLabel(root, textvariable=status_var, anchor="w")
    status_bar.pack(side="bottom", fill="x")

    def update_status(message, duration=5000):
        status_var.set(message)
        root.after(duration, lambda: status_var.set("準備就緒"))

    # 使用 CTkTabview 取代 Notebook
    tabview = ctk.CTkTabview(root)
    tabview.pack(fill="both", expand=True, padx=10, pady=10)
    tabview.add("訂購訊息生成器")

    order_generator = OrderMessageGenerator(
        tabview.tab("訂購訊息生成器"), status_callback=update_status
    )
    order_generator.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
