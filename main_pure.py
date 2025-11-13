# main_pure.py
import os
import sys
import json
from typing import List, Dict

# 导入 GUI
try:
    from gui import show_settings_gui, get_default_config_path
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

# ==================== 路径工具 ====================
def get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_template_dir() -> str:
    config_path = get_default_config_path()
    if not os.path.isfile(config_path):
        return os.path.join(get_app_dir(), "templates")
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        raw = config.get("template_dir", "templates").strip()
        return raw if os.path.isabs(raw) else os.path.join(get_app_dir(), raw)
    except:
        return os.path.join(get_app_dir(), "templates")

def get_template_path(name: str) -> str:
    if not name.endswith(".btw"):
        name += ".btw"
    return os.path.join(get_template_dir(), name)

# ==================== 核心逻辑 ====================
def select_template(name: str) -> str:
    path = get_template_path(name)
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"\n模板未找到: {os.path.basename(path)}\n"
            f"请放入: {get_template_dir()}\n"
        )
    print(f"使用模板: {path}")
    return path

def load_products() -> List[Dict]:
    config_path = get_default_config_path()
    if not os.path.isfile(config_path):
        raise FileNotFoundError(f"配置文件未找到: {config_path}")
    with open(config_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get("products", [])

# ==================== 主程序 ====================
def run_printing():
    print("正在启动打印流程...\n")
    template_dir = get_template_dir()
    print(f"模板目录: {template_dir}")

    if not os.path.exists(template_dir):
        os.makedirs(template_dir, exist_ok=True)
        print(f"已创建模板目录: {template_dir}")

    products = load_products()
    if not products:
        print("无产品可打印。")
        input("按 Enter 退出...")
        return

    success = 0
    for prod in products:
        name = prod.get("name", "未知")
        tmpl = prod.get("template")
        if not tmpl:
            print(f"{name} 缺少 template，跳过。")
            continue
        try:
            select_template(tmpl)
            print(f"{name} 打印完成")
            success += 1
        except Exception as e:
            print(f"{name} 失败: {e}")

    print(f"\n完成: 成功 {success}/{len(products)}")
    input("按 Enter 退出...")

def main():
    print("OuterBoxPrinter_PureData")

    # 首次运行或配置文件缺失 → 弹出 GUI
    config_path = get_default_config_path()
    should_show_gui = (
        GUI_AVAILABLE and
        (not os.path.exists(config_path) or
         not os.path.exists(get_template_dir()))
    )

    if should_show_gui:
        print("首次运行，打开设置界面...")
        if show_settings_gui():
            run_printing()
        else:
            print("用户取消，退出。")
    else:
        # 直接运行
        run_printing()

if __name__ == "__main__":
    main()