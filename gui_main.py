"""
机制电价市场力分析工具 - GUI入口
使用PyInstaller打包为Windows独立exe
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 将项目根目录加入路径
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from main import run_analysis


class AnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("机制电价市场力分析工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # 数据
        self.previous_dir = tk.StringVar()
        self.current_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.df_abnormal = None

        # 设置默认值（如果存在）
        default_prev = os.path.join(project_root, 'Data', 'Input', 'Previous')
        default_curr = os.path.join(project_root, 'Data', 'Input', 'Current')
        default_out = os.path.join(project_root, 'Data', 'Output')
        if os.path.exists(default_prev):
            self.previous_dir.set(default_prev)
        if os.path.exists(default_curr):
            self.current_dir.set(default_curr)
        self.output_dir.set(default_out)

        self._build_ui()

    def _build_ui(self):
        # 标题
        title = tk.Label(self.root, text="机制电价市场力分析工具", font=("微软雅黑", 18, "bold"))
        title.pack(pady=15)

        # 输入区域
        input_frame = tk.LabelFrame(self.root, text="数据输入", font=("微软雅黑", 11))
        input_frame.pack(fill=tk.X, padx=20, pady=5)

        # 历史数据文件夹
        row1 = tk.Frame(input_frame)
        row1.pack(fill=tk.X, padx=10, pady=8)
        tk.Label(row1, text="历史机制电价竞价数据:", width=22, anchor="e", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.previous_dir, font=("微软雅黑", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(row1, text="浏览...", command=self._browse_previous, font=("微软雅黑", 9)).pack(side=tk.LEFT)

        # 最新数据文件夹
        row2 = tk.Frame(input_frame)
        row2.pack(fill=tk.X, padx=10, pady=8)
        tk.Label(row2, text="最新机制电价竞价数据:", width=22, anchor="e", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        tk.Entry(row2, textvariable=self.current_dir, font=("微软雅黑", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(row2, text="浏览...", command=self._browse_current, font=("微软雅黑", 9)).pack(side=tk.LEFT)

        # 输出目录
        row3 = tk.Frame(input_frame)
        row3.pack(fill=tk.X, padx=10, pady=8)
        tk.Label(row3, text="输出目录:", width=22, anchor="e", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        tk.Entry(row3, textvariable=self.output_dir, font=("微软雅黑", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(row3, text="浏览...", command=self._browse_output, font=("微软雅黑", 9)).pack(side=tk.LEFT)

        # 运行按钮 + 进度条
        ctrl_frame = tk.Frame(self.root)
        ctrl_frame.pack(fill=tk.X, padx=20, pady=10)

        self.run_btn = tk.Button(ctrl_frame, text="开始分析", command=self._start_analysis,
                                 font=("微软雅黑", 12, "bold"), bg="#4CAF50", fg="white",
                                 width=15, height=1)
        self.run_btn.pack(side=tk.LEFT, padx=5)

        self.save_btn = tk.Button(ctrl_frame, text="保存异常主体名单", command=self._save_abnormal,
                                  font=("微软雅黑", 11), state=tk.DISABLED,
                                  width=18, height=1)
        self.save_btn.pack(side=tk.LEFT, padx=10)

        self.progress = ttk.Progressbar(ctrl_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        self.progress_label = tk.Label(ctrl_frame, text="0%", font=("微软雅黑", 10), width=5)
        self.progress_label.pack(side=tk.LEFT)

        # 日志区域
        log_frame = tk.LabelFrame(self.root, text="运行日志", font=("微软雅黑", 11))
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 10),
                                state=tk.DISABLED, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 结果显示区域
        result_frame = tk.LabelFrame(self.root, text="异常市场主体识别结果", font=("微软雅黑", 11))
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        columns = ('市场主体名称', '所属发电集团', '行为类型', '报价偏离度', '收益偏离度')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=8)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor='center')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tree_scroll = ttk.Scrollbar(self.tree, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _browse_previous(self):
        path = filedialog.askdirectory(title="选择历史机制电价竞价数据文件夹")
        if path:
            self.previous_dir.set(path)

    def _browse_current(self):
        path = filedialog.askdirectory(title="选择最新机制电价竞价数据文件夹")
        if path:
            self.current_dir.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def _log(self, msg):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _set_progress(self, val):
        self.progress['value'] = val
        self.progress_label['text'] = f"{val}%"
        self.root.update_idletasks()

    def _start_analysis(self):
        prev = self.previous_dir.get().strip()
        curr = self.current_dir.get().strip()
        out = self.output_dir.get().strip()

        if not prev or not os.path.exists(prev):
            messagebox.showerror("错误", "请选择有效的历史数据文件夹")
            return
        if not curr or not os.path.exists(curr):
            messagebox.showerror("错误", "请选择有效的最新数据文件夹")
            return
        if not out:
            messagebox.showerror("错误", "请选择输出目录")
            return

        self.run_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.progress['value'] = 0
        self.progress_label['text'] = "0%"
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

        # 在后台线程运行分析，避免阻塞GUI
        thread = threading.Thread(target=self._run_analysis_thread,
                                  args=(prev, curr, out), daemon=True)
        thread.start()

    def _run_analysis_thread(self, prev, curr, out):
        try:
            result = run_analysis(
                prev, curr, out,
                progress_callback=self._set_progress,
                log_callback=self._log
            )
            self.df_abnormal = result['df_abnormal']

            # 在GUI线程更新结果
            self.root.after(0, self._show_results)
        except Exception as e:
            self.root.after(0, lambda: self._on_error(str(e)))

    def _show_results(self):
        # 填充Treeview
        for _, row in self.df_abnormal.iterrows():
            self.tree.insert('', tk.END, values=(
                row['市场主体名称'],
                row['所属发电集团'],
                row['行为类型'],
                row['报价偏离度'],
                row['收益偏离度'],
            ))

        self.run_btn.config(state=tk.NORMAL)
        self.save_btn.config(state=tk.NORMAL)
        messagebox.showinfo("完成", f"分析完成！\n共识别异常主体 {len(self.df_abnormal)} 个\n结果已保存到: {self.output_dir.get()}")

    def _on_error(self, msg):
        self.run_btn.config(state=tk.NORMAL)
        messagebox.showerror("错误", f"分析过程中出错:\n{msg}")

    def _save_abnormal(self):
        if self.df_abnormal is None or self.df_abnormal.empty:
            messagebox.showwarning("提示", "暂无异常主体数据可保存")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存异常主体名单",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                self.df_abnormal.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"异常主体名单已保存到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败:\n{str(e)}")


def main():
    root = tk.Tk()
    app = AnalysisApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
