# gui.py
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, Any

def get_default_config_path() -> str:
    """获取 config_pure.json 的默认路径"""
    if getattr(__import__('sys'), 'frozen', False):
        return os.path.join(os.path.dirname(__import__('sys').executable), "config_pure.json")
    else:
        return os.path.join(os.path.dirname(__file__), "config_pure.json")

def load_config() -> Dict[str, Any]:
    """加载配置，失败返回默认"""
    config_path = get_default_config_path()
    default = {
        "template_dir": "templates",
        "products": []
    }
    if not os.path.isfile(config_path):
        return default
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # 补全 template_dir
        data.setdefault("template_dir", default["template_dir"])
        data.setdefault("products", default["products"])
        return data
    except Exception:
        return default

def save_config(config: Dict[str, Any]) -> None:
    """保存配置"""
    config_path = get_default_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"配置已保存: {config_path}")
    except Exception as e:
        messagebox.showerror("保存失败", f"无法保存配置文件:\n{e}")

def show_settings_gui() -> bool:
    """显示设置界面，返回 True 表示用户点击了“开始打印”"""
    config = load_config()
    root = tk.Tk()
    root.title("OuterBoxPrinter 设置")
    root.geometry("560x320")
    root.resizable(False, False)

    # === 模板目录 ===
    tk.Label(root, text="模板目录 (.btw 文件存放位置):", font=("Segoe UI", 10)).pack(pady=(20, 5), anchor="w", padx=20)

    dir_frame = tk.Frame(root)
    dir_frame.pack(fill="x", padx=20, pady=5)

    dir_var = tk.StringVar(value=config["template_dir"])
    entry = tk.Entry(dir_frame, textvariable=dir_var, width=50, font=("Consolas", 10))
    entry.pack(side="left", expand=True, fill="x")

    def browse_dir():
        path = filedialog.askdirectory(
            title="选择模板目录",
            initialdir=os.path.dirname(get_default_config_path())
        )
        if path:
            dir_var.set(path)

    tk.Button(dir_frame, text="浏览...", command=browse_dir).pack(side="right", padx=(5, 0))

    # === 当前路径预览 ===
    def update_preview(*args):
        path = dir_var.get()
        if os.path.isabs(path):
            abs_path = path
        else:
            abs_path = os.path.abspath(os.path.join(os.path.dirname(get_default_config_path()), path))
        preview_label.config(text=f"完整路径: {abs_path}")

    dir_var.trace_add("write", update_preview)
    preview_label = tk.Label(root, text="", fg="gray", font=("Consolas", 9))
    preview_label.pack(pady=5)

    # === 产品列表预览 ===
    tk.Label(root, text="产品列表预览:", font=("Segoe UI", 10)).pack(pady=(15, 5), anchor="w", padx=20)
    listbox = tk.Listbox(root, height=6, font=("Consolas", 9))
    listbox.pack(fill="x", padx=20, pady=5)

    for prod in config.get("products", []):
        name = prod.get("name", "未知")
        tmpl = prod.get("template", "未设置")
        listbox.insert(tk.END, f"{name} → {tmpl}.btw")

    if not config.get("products"):
        listbox.insert(tk.END, "<无产品>")

    # === 按钮 ===
    button_frame = tk.Frame(root)
    button_frame.pack(pady=20)

    result = {"start": False}

    def on_save():
        new_config = {
            "template_dir": dir_var.get().strip(),
            "products": config.get("products", [])
        }
        save_config(new_config)
        messagebox.showinfo("成功", "设置已保存！")

    def on_start():
        if not dir_var.get().strip():
            messagebox.showwarning("警告", "请设置模板目录！")
            return
        on_save()
        result["start"] = True
        root.quit()

    tk.Button(button_frame, text="保存设置", width=12, command=on_save).pack(side="left", padx=10)
    tk.Button(button_frame, text="开始打印", width=12, bg="#4CAF50", fg="white", command=on_start).pack(side="left", padx=10)
    tk.Button(button_frame, text="退出", width=12, command=root.quit).pack(side="left", padx=10)

    # 初始化预览
    update_preview()

    # 居中窗口
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()
    root.destroy()
    return result["start"]