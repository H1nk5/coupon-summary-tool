import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading

from summary_core import aggregate, generate_report, read_data

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("代金券汇总工具")
        self.geometry("700x420")
        self.resizable(False, False)

        self.兰考路径 = ctk.StringVar()
        self.民权路径 = ctk.StringVar()
        self.输出路径 = ctk.StringVar()

        # === 标题栏 ===
        header = ctk.CTkFrame(self, fg_color="#1a73e8", corner_radius=0, height=50)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="代金券汇总工具",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white",
        ).pack(expand=True)

        # === 主体 ===
        body = ctk.CTkFrame(self, fg_color="#f5f5f5", corner_radius=0)
        body.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(body, fg_color="transparent")
        inner.pack(pady=25, padx=40, fill="x")

        # -- 兰考 --
        row1 = ctk.CTkFrame(inner, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(row1, text="兰考表格", font=ctk.CTkFont(size=13), width=80, anchor="e").pack(side="left")
        ctk.CTkEntry(
            row1, textvariable=self.兰考路径, height=38,
            font=ctk.CTkFont(size=12), placeholder_text="请选择兰考.xlsx",
        ).pack(side="left", fill="x", expand=True, padx=(10, 10))
        ctk.CTkButton(
            row1, text="浏览", width=90, height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self.选择文件(self.兰考路径),
        ).pack(side="left")

        # -- 民权 --
        row2 = ctk.CTkFrame(inner, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(row2, text="民权表格", font=ctk.CTkFont(size=13), width=80, anchor="e").pack(side="left")
        ctk.CTkEntry(
            row2, textvariable=self.民权路径, height=38,
            font=ctk.CTkFont(size=12), placeholder_text="请选择民权.xlsx",
        ).pack(side="left", fill="x", expand=True, padx=(10, 10))
        ctk.CTkButton(
            row2, text="浏览", width=90, height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self.选择文件(self.民权路径),
        ).pack(side="left")

        # -- 输出 --
        row3 = ctk.CTkFrame(inner, fg_color="transparent")
        row3.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(row3, text="输出位置", font=ctk.CTkFont(size=13), width=80, anchor="e").pack(side="left")
        ctk.CTkEntry(
            row3, textvariable=self.输出路径, height=38,
            font=ctk.CTkFont(size=12), placeholder_text="选择保存位置",
        ).pack(side="left", fill="x", expand=True, padx=(10, 10))
        ctk.CTkButton(
            row3, text="浏览", width=90, height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.选择输出,
        ).pack(side="left")

        # -- 进度条 --
        self.进度 = ctk.CTkProgressBar(inner, height=14, corner_radius=7)
        self.进度.pack(fill="x", pady=(0, 6))
        self.进度.set(0)

        # -- 状态 --
        self.状态 = ctk.CTkLabel(inner, text="就绪", font=ctk.CTkFont(size=12), text_color="gray")
        self.状态.pack(pady=(0, 15))

        # -- 开始按钮 --
        self.按钮 = ctk.CTkButton(
            inner, text="开始汇总", height=48, width=220,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#1a73e8", hover_color="#1558b0",
            command=self.开始,
        )
        self.按钮.pack()

    def 选择文件(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            var.set(path)

    def 选择输出(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile="代金券汇总表.xlsx",
        )
        if path:
            self.输出路径.set(path)

    def 更新状态(self, text, color="gray"):
        if threading.current_thread() is not threading.main_thread():
            self.after(0, lambda: self.更新状态(text, color))
            return
        self.状态.configure(text=text, text_color=color)
        self.update_idletasks()

    def 更新进度(self, value):
        self.after(0, lambda: self.进度.set(value))

    def 开始(self):
        paths = (self.兰考路径.get(), self.民权路径.get(), self.输出路径.get())
        if not all(paths):
            messagebox.showwarning("提示", "请先选择所有文件路径")
            return
        self.按钮.configure(state="disabled")
        threading.Thread(target=self.执行, args=paths, daemon=True).start()

    def 读取数据(self, filepath):
        return read_data(filepath)

    def 聚合(self, data):
        return aggregate(data)

    def 执行(self, 兰考路径, 民权路径, 输出路径):
        try:
            self.更新状态("正在读取、汇总并生成表格...", "#1a73e8")
            result = generate_report(兰考路径, 民权路径, 输出路径, self.更新进度)
            self.更新状态(
                f"完成! {result['days']} 天 | {result['first_date']} ~ {result['last_date']}",
                "#16a34a",
            )
            self.after(0, lambda: messagebox.showinfo(
                "完成",
                f"汇总完成!\n共 {result['days']} 天\n{result['first_date']} ~ {result['last_date']}\n\n已保存: {输出路径}",
            ))

        except Exception as e:
            self.更新状态(f"出错: {e}", "#dc2626")
            self.after(0, lambda error=str(e): messagebox.showerror("错误", error))
        finally:
            self.after(0, lambda: self.按钮.configure(state="normal"))


if __name__ == "__main__":
    app = App()
    app.mainloop()
