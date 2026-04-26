# -*- coding: utf-8 -*-
"""Markdown转Docx GUI界面"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path

# 导入转换器
from converter import convert_file


class Md2DocxGUI:
    """Markdown转Docx图形界面"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Markdown转Docx")
        self.root.geometry("500x350")
        self.root.resizable(False, False)

        # 居中窗口
        self._center_window()

        # 文件路径
        self.md_file_path = None

        # 设置样式
        self._setup_styles()

        # 创建界面
        self._create_widgets()

        # 绑定拖拽事件
        self._bind_drag_events()

    def _center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Microsoft YaHei', 16, 'bold'))
        style.configure('Info.TLabel', font=('Microsoft YaHei', 10))
        style.configure('Success.TLabel', font=('Microsoft YaHei', 10), foreground='green')
        style.configure('Error.TLabel', font=('Microsoft YaHei', 10), foreground='red')

    def _create_widgets(self):
        """创建界面组件"""
        # 标题
        title_label = ttk.Label(
            self.root,
            text="Markdown 转 Docx",
            style='Title.TLabel'
        )
        title_label.pack(pady=20)

        # 文件选择区域
        file_frame = ttk.Frame(self.root)
        file_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)

        # 文件路径显示
        self.path_var = tk.StringVar(value="点击下方按钮选择Markdown文件或拖拽文件到此处")
        path_label = ttk.Label(
            file_frame,
            textvariable=self.path_var,
            wraplength=400,
            style='Info.TLabel',
            anchor=tk.CENTER
        )
        path_label.pack(fill=tk.BOTH, expand=True, pady=10)

        # 按钮区域
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        # 选择文件按钮
        self.select_btn = ttk.Button(
            btn_frame,
            text="选择Markdown文件",
            command=self._select_file,
            width=20
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)

        # 转换按钮
        self.convert_btn = ttk.Button(
            btn_frame,
            text="开始转换",
            command=self._start_convert,
            state=tk.DISABLED,
            width=20
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)

        # 状态显示
        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(
            self.root,
            textvariable=self.status_var,
            style='Info.TLabel'
        )
        self.status_label.pack(pady=10)

        # 进度条
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate',
            length=400
        )

    def _bind_drag_events(self):
        """绑定拖拽事件"""
        # Windows 拖拽支持
        try:
            import ctypes
            from ctypes import wintypes

            # 注册拖拽接受
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', self._on_drop)
        except Exception:
            # 如果不支持拖拽，忽略
            pass

    def _on_drop(self, event):
        """处理拖拽事件"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.md'):
                self._set_file_path(file_path)
            else:
                messagebox.showerror("错误", "请选择Markdown文件(.md)")

    def _select_file(self):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Markdown文件",
            filetypes=[("Markdown文件", "*.md"), ("所有文件", "*.*")]
        )
        if file_path:
            self._set_file_path(file_path)

    def _set_file_path(self, path: str):
        """设置文件路径"""
        self.md_file_path = path
        self.path_var.set(path)
        self.convert_btn.config(state=tk.NORMAL)
        self.status_var.set("")

    def _start_convert(self):
        """开始转换"""
        if not self.md_file_path:
            return

        # 检查文件是否存在
        if not os.path.exists(self.md_file_path):
            messagebox.showerror("错误", "文件不存在")
            return

        # 禁用按钮
        self.select_btn.config(state=tk.DISABLED)
        self.convert_btn.config(state=tk.DISABLED)
        self.progress.pack(pady=5)
        self.progress.start()
        self.status_var.set("正在转换...")

        # 在新线程中执行转换
        thread = threading.Thread(target=self._convert, daemon=True)
        thread.start()

    def _convert(self):
        """执行转换"""
        try:
            output_path = convert_file(self.md_file_path)

            # 转换完成，在主线程中更新界面
            self.root.after(0, self._convert_success, output_path)
        except Exception as e:
            self.root.after(0, self._convert_error, str(e))

    def _convert_success(self, output_path: str):
        """转换成功"""
        self.progress.stop()
        self.progress.pack_forget()

        self.status_var.set("转换成功！")
        messagebox.showinfo("成功", f"文件已转换为:\n{output_path}")

        # 恢复按钮
        self.select_btn.config(state=tk.NORMAL)
        self.convert_btn.config(state=tk.NORMAL)
        self.status_var.set("")

    def _convert_error(self, error: str):
        """转换失败"""
        self.progress.stop()
        self.progress.pack_forget()

        self.status_var.set("转换失败")
        messagebox.showerror("错误", f"转换失败:\n{error}")

        # 恢复按钮
        self.select_btn.config(state=tk.NORMAL)
        self.convert_btn.config(state=tk.NORMAL)

    def run(self):
        """运行应用"""
        self.root.mainloop()


def main():
    """主函数"""
    app = Md2DocxGUI()
    app.run()


if __name__ == '__main__':
    main()
