"""
Personal Assistant
项目复制 / 文件搬运 / 效率工具
Author: MeiKuaiKuai
"""

import os
import sys
import json
import re
import shutil
import subprocess
import threading
import time
import stat
import logging
import traceback
import platform
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# ============================================================
# A. 多语言文本
# ============================================================

TEXTS = {
  "zh": {
    "support_btn": "  支持与联系  ",
    "tab_copy": "  项目复制  ",
    "tab_transfer": "  文件搬运  ",
    "source_project": "源项目",
    "browse": "浏览...",
    "remember_dir": "记住此目录",
    "copy_name": "复制名称",
    "auto_hint": "自动生成名称，可手动修改",
    "start_copy": "开始复制",
    "copying": "复制中...",
    "open_dest": "打开目标文件夹",
    "action": "操作",
    "ready": "就绪",
    "favorites": "常用地址",
    "fill_src": "填入源",
    "fill_dst": "填入目标",
    "add_fav": "+添加",
    "del_fav": "-删除",
    "source_a": "源地址 (A)",
    "dest_b": "目标地址 (B)",
    "select_all": "全选",
    "deselect_all": "取消全选",
    "refresh": "刷新列表",
    "sort_label": "排序:",
    "sort_name_asc": "名称 A→Z",
    "sort_name_desc": "名称 Z→A",
    "sort_time_new": "时间 新→旧",
    "sort_time_old": "时间 旧→新",
    "sort_size_big": "大小 大→小",
    "sort_size_small": "大小 小→大",
    "op_label": "操作:",
    "op_copy": "复制",
    "op_move": "剪切(移动)",
    "rename_label": "重命名:",
    "rename_hint": "(仅单选有效，留空不改名)",
    "execute": "执行搬运",
    "transferring": "搬运中...",
    "delete_btn": "删除选中",
    "deleting_btn": "删除中...",
    "exec_section": "执行",
    "folder_tag": "[文件夹]",
    "file_tag": "[文件]",
    "footer": "Personal Assistant v1.0  |  by MeiKuaiKuai",
    # 弹窗消息
    "m_del_confirm": "确定要删除以下 {} 个项目吗？\n\n{}\n\n此操作不可撤销！",
    "m_del_done": "已删除 {} 个项目。",
    "m_del_err": "删除时发生错误:\n{}",
    "t_del_confirm": "确认删除",
    "t_del_fail": "删除失败",
    "m_no_source": "请先选择源项目文件夹。",
    "m_source_gone": "源项目文件夹不存在。",
    "m_no_name": "请输入复制名称。",
    "m_bad_chars": '名称不能包含以下字符: \\ / : * ? " < > |',
    "m_dest_exists": "目标文件夹已存在:\n{}\n请修改名称。",
    "m_disk": "检测到 {} 磁盘，使用 {} 线程复制...",
    "m_scanning": "正在扫描文件...",
    "m_progress": "进度: {}/{} 个文件 ({}/{})",
    "m_speed": "速度: {}/s  |  预计剩余: {}",
    "m_cur_file": "正在复制: {}",
    "m_done": "完成！共 {} 个文件 ({})，耗时 {} 秒",
    "m_avg": "平均速度: {}",
    "m_ok_copy": "项目复制成功！\n共 {} 个文件 ({})\n耗时 {} 秒，平均 {}",
    "m_err_copy": "发生错误:\n{}\n\n已复制的部分文件保留在目标文件夹中。",
    "m_err": "错误: {}",
    "m_no_src_valid": "请选择有效的源地址。",
    "m_no_dst": "请选择目标地址。",
    "m_no_items": "请在文件列表中选择要搬运的项目。",
    "m_mkdir": "目标目录不存在:\n{}\n是否创建？",
    "m_conflict": "目标已存在同名项:\n{}\n请重命名或手动处理。",
    "m_cur_proc": "正在处理: {}",
    "m_tr_done": "完成！{} {} 个文件 ({})，耗时 {} 秒",
    "m_tr_ok": "{}成功！\n共 {} 个文件 ({})\n耗时 {} 秒，平均 {}",
    "m_tr_err": "发生错误:\n{}",
    "m_deleting": "正在删除源文件...",
    "m_quit": "操作进行中，确定要取消并退出吗？",
    "m_op_copy": "复制",
    "m_op_cut": "剪切",
    "m_calc": "计算中...",
    "t_hint": "提示",
    "t_err": "错误",
    "t_done": "完成",
    "t_confirm": "确认",
    "t_fail_copy": "复制失败",
    "t_fail_tr": "搬运失败",
    "d_sel_proj": "选择项目文件夹",
    "d_sel_src": "选择源文件夹",
    "d_sel_dst": "选择目标文件夹",
    "d_sel_fav": "选择要添加的常用地址",
    # 支持弹窗
    "s_title": "支持与联系",
    "s_thanks": "感谢你使用",
    "s_body": (
        "这款工具由独立开发者 MeiKuaiKuai 精心打造，\n"
        "只为让每一位工程师的日常工作更加高效、从容。\n"
        "\n"
        "如果它曾为你节省了几分钟时间，\n"
        "或让繁琐的操作变得简单了一点点——\n"
        "那就是我最大的成就感。\n"
        "\n"
        "你的支持，是我持续创造的动力。\n"
        "一杯咖啡虽小，却意味着认可与鼓励。"
    ),
    "s_wechat": "添加微信",
    "s_pay": "请我喝杯咖啡",
    "s_wechat_hint": "扫码加好友",
    "s_pay_hint": "扫码支持一下",
    "s_footer": "MeiKuaiKuai 用心出品",
    "s_close": "  关闭  ",
    "s_no_pil": "请先安装 Pillow:\npip install Pillow",
    "log_btn": "  错误日志  ",
    "log_title": "错误日志",
    "log_empty": "暂无错误记录，运行正常！",
    "log_copy_ok": "日志内容已复制到剪贴板，可直接粘贴反馈给开发者。",
    "log_copy_btn": "复制日志内容",
    "log_open_btn": "打开日志文件",
    "log_clear_btn": "清空日志",
    "log_cleared": "日志已清空。",
    "m_move_partial": "剪切完成，但以下文件删除失败（已复制到目标）：\n{}",
    "m_move_skip": "跳过(权限不足): {}",
    "m_disk_transfer": "源:{} 目标:{} → 使用 {} 线程",
  },
  "en": {
    "support_btn": "  Support & Contact  ",
    "tab_copy": "  Project Copy  ",
    "tab_transfer": "  File Transfer  ",
    "source_project": "Source Project",
    "browse": "Browse...",
    "remember_dir": "Remember Directory",
    "copy_name": "Copy Name",
    "auto_hint": "Auto-generated, editable",
    "start_copy": "Start Copy",
    "copying": "Copying...",
    "open_dest": "Open Destination",
    "action": "Actions",
    "ready": "Ready",
    "favorites": "Favorites",
    "fill_src": "To Src",
    "fill_dst": "To Dest",
    "add_fav": "+Add",
    "del_fav": "-Del",
    "source_a": "Source (A)",
    "dest_b": "Destination (B)",
    "select_all": "Select All",
    "deselect_all": "Deselect",
    "refresh": "Refresh",
    "sort_label": "Sort:",
    "sort_name_asc": "Name A→Z",
    "sort_name_desc": "Name Z→A",
    "sort_time_new": "Time New→Old",
    "sort_time_old": "Time Old→New",
    "sort_size_big": "Size Big→Small",
    "sort_size_small": "Size Small→Big",
    "op_label": "Mode:",
    "op_copy": "Copy",
    "op_move": "Cut (Move)",
    "rename_label": "Rename:",
    "rename_hint": "(Single only, empty=keep)",
    "execute": "Execute",
    "transferring": "Transferring...",
    "delete_btn": "Delete Selected",
    "deleting_btn": "Deleting...",
    "exec_section": "Execute",
    "folder_tag": "[Folder]",
    "file_tag": "[File]",
    "footer": "Personal Assistant v1.0  |  by MeiKuaiKuai",
    "m_del_confirm": "Delete the following {} item(s)?\n\n{}\n\nThis cannot be undone!",
    "m_del_done": "Deleted {} item(s).",
    "m_del_err": "Error during deletion:\n{}",
    "t_del_confirm": "Confirm Delete",
    "t_del_fail": "Delete Failed",
    "m_no_source": "Please select a source folder.",
    "m_source_gone": "Source folder does not exist.",
    "m_no_name": "Please enter a copy name.",
    "m_bad_chars": 'Name cannot contain: \\ / : * ? " < > |',
    "m_dest_exists": "Destination exists:\n{}\nPlease rename.",
    "m_disk": "Detected {} disk, {} threads...",
    "m_scanning": "Scanning files...",
    "m_progress": "Progress: {}/{} files ({}/{})",
    "m_speed": "Speed: {}/s  |  ETA: {}",
    "m_cur_file": "Copying: {}",
    "m_done": "Done! {} files ({}), {:.1f}s",
    "m_avg": "Avg speed: {}",
    "m_ok_copy": "Copy successful!\n{} files ({})\n{:.1f}s, avg {}",
    "m_err_copy": "Error:\n{}\n\nPartial copy preserved.",
    "m_err": "Error: {}",
    "m_no_src_valid": "Please select a valid source.",
    "m_no_dst": "Please select a destination.",
    "m_no_items": "Please select items to transfer.",
    "m_mkdir": "Destination missing:\n{}\nCreate it?",
    "m_conflict": "Already exists:\n{}\nPlease rename.",
    "m_cur_proc": "Processing: {}",
    "m_tr_done": "Done! {} {} files ({}), {:.1f}s",
    "m_tr_ok": "{} successful!\n{} files ({})\n{:.1f}s, avg {}",
    "m_tr_err": "Error:\n{}",
    "m_deleting": "Deleting source files...",
    "m_quit": "Operation in progress. Cancel and exit?",
    "m_op_copy": "Copied",
    "m_op_cut": "Moved",
    "m_calc": "Calculating...",
    "t_hint": "Notice",
    "t_err": "Error",
    "t_done": "Done",
    "t_confirm": "Confirm",
    "t_fail_copy": "Copy Failed",
    "t_fail_tr": "Transfer Failed",
    "d_sel_proj": "Select Project Folder",
    "d_sel_src": "Select Source Folder",
    "d_sel_dst": "Select Destination Folder",
    "d_sel_fav": "Select Folder to Add",
    "s_title": "Support & Contact",
    "s_thanks": "Thank you for using",
    "s_body": (
        "This tool is independently crafted by MeiKuaiKuai,\n"
        "built with care for engineers who deserve\n"
        "a better, smoother workflow.\n"
        "\n"
        "If it has ever saved you a few minutes,\n"
        "or made a tedious task just a bit easier —\n"
        "that alone makes it all worthwhile.\n"
        "\n"
        "Your support fuels the next update.\n"
        "A cup of coffee is small,\n"
        "but it means recognition and encouragement."
    ),
    "s_wechat": "Add WeChat",
    "s_pay": "Buy Me a Coffee",
    "s_wechat_hint": "Scan to connect",
    "s_pay_hint": "Scan to support",
    "s_footer": "Made with dedication by MeiKuaiKuai",
    "s_close": "  Close  ",
    "s_no_pil": "Please install Pillow:\npip install Pillow",
    "log_btn": "  Error Log  ",
    "log_title": "Error Log",
    "log_empty": "No errors recorded. Everything is working fine!",
    "log_copy_ok": "Log copied to clipboard. You can paste it to report issues.",
    "log_copy_btn": "Copy Log",
    "log_open_btn": "Open Log File",
    "log_clear_btn": "Clear Log",
    "log_cleared": "Log cleared.",
    "m_move_partial": "Move completed, but failed to delete these files (already copied):\n{}",
    "m_move_skip": "Skipped (access denied): {}",
    "m_disk_transfer": "Src:{} Dst:{} → {} threads",
  }
}


# ============================================================
# B. 配置管理
# ============================================================

def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


CONFIG_FILE = os.path.join(get_app_dir(), "config.json")


def load_config():
    defaults = {"default_directory": "", "favorites": [], "language": "zh"}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        for k, v in defaults.items():
            config.setdefault(k, v)
        return config
    except (FileNotFoundError, json.JSONDecodeError):
        return defaults


def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# ============================================================
# C. 日志系统
# ============================================================

LOG_FILE = os.path.join(get_app_dir(), "error_log.txt")

def setup_logging():
    logger = logging.getLogger("PA")
    logger.setLevel(logging.DEBUG)
    # 文件输出
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger.addHandler(fh)
    return logger

logger = setup_logging()

def get_system_info():
    """收集系统信息，用于错误报告"""
    info = []
    info.append(f"OS: {platform.system()} {platform.version()}")
    info.append(f"Python: {platform.python_version()}")
    info.append(f"App Dir: {get_app_dir()}")
    try:
        import ctypes
        info.append(f"Admin: {ctypes.windll.shell32.IsUserAnAdmin() != 0}")
    except Exception:
        info.append("Admin: N/A")
    return " | ".join(info)

def log_error(context, error, extra=""):
    """记录错误到日志文件"""
    msg = f"[{context}] {type(error).__name__}: {error}"
    if extra:
        msg += f" | {extra}"
    msg += f"\n  System: {get_system_info()}"
    msg += f"\n  Traceback:\n{''.join(traceback.format_exception(type(error), error, error.__traceback__))}"
    logger.error(msg)

def _rmtree_onerror(func, path, exc_info):
    """shutil.rmtree 错误回调：尝试去掉只读属性后重试"""
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        func(path)
    except Exception as e:
        log_error("rmtree_retry", e, f"path={path}")


# ============================================================
# D. 通用工具函数
# ============================================================

INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')
SSD_WORKERS = 8
HDD_WORKERS = 2


def detect_disk_type(path):
    if sys.platform != "win32":
        return "HDD", HDD_WORKERS
    try:
        letter = os.path.splitdrive(os.path.abspath(path))[0][0]
        cmd = (f'powershell -NoProfile -Command "'
               f'(Get-Partition -DriveLetter {letter} '
               f'| Get-Disk | Get-PhysicalDisk).MediaType"')
        r = subprocess.run(cmd, capture_output=True, text=True,
                           timeout=10, shell=True)
        if "SSD" in r.stdout or "Solid State" in r.stdout:
            return "SSD", SSD_WORKERS
        return "HDD", HDD_WORKERS
    except Exception:
        return "HDD", HDD_WORKERS


def detect_workers_for_transfer(src, dst):
    _, w1 = detect_disk_type(src)
    _, w2 = detect_disk_type(dst)
    return min(w1, w2)


def format_size(b):
    if b < 1024:
        return f"{b} B"
    if b < 1024**2:
        return f"{b/1024:.1f} KB"
    if b < 1024**3:
        return f"{b/1024**2:.1f} MB"
    return f"{b/1024**3:.2f} GB"


def format_eta(s):
    if s < 60:
        return f"{int(s)} s"
    if s < 3600:
        m, sec = divmod(int(s), 60)
        return f"{m}m {sec}s"
    h, rem = divmod(int(s), 3600)
    m, _ = divmod(rem, 60)
    return f"{h}h {m}m"


def generate_copy_name(source_path):
    src = Path(source_path)
    name = src.name
    m = re.match(r'^(.+)_(\d+)$', name)
    root = m.group(1) if m else name
    mx = 0
    try:
        for item in src.parent.iterdir():
            if item.is_dir():
                m2 = re.match(rf'^{re.escape(root)}_(\d+)$', item.name)
                if m2 and int(m2.group(1)) > mx:
                    mx = int(m2.group(1))
    except OSError:
        pass
    return f"{root}_{mx + 1}"


def open_folder(path):
    if not path or not os.path.isdir(path):
        return
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ============================================================
# D. 复制引擎
# ============================================================

def _collect_copy_tasks(src, dst):
    tasks = []
    for dp, _, fns in os.walk(src):
        rd = os.path.relpath(dp, src)
        dd = os.path.join(dst, rd)
        os.makedirs(dd, exist_ok=True)
        for fn in fns:
            sf = os.path.join(dp, fn)
            df = os.path.join(dd, fn)
            try:
                sz = os.path.getsize(sf)
            except OSError:
                sz = 0
            tasks.append((sf, df, sz))
    return tasks


def _collect_transfer_tasks(items, dest_dir, rename_map=None):
    tasks = []
    for p in items:
        name = os.path.basename(p)
        dname = name
        if rename_map and name in rename_map:
            dname = rename_map[name]
        if os.path.isfile(p):
            try:
                sz = os.path.getsize(p)
            except OSError:
                sz = 0
            tasks.append((p, os.path.join(dest_dir, dname), sz))
        elif os.path.isdir(p):
            tasks.extend(_collect_copy_tasks(p, os.path.join(dest_dir, dname)))
    return tasks


def run_copy_engine(tasks, progress_cb, done_cb, error_cb, cancel_event,
                    workers=2, delete_sources=None):
    try:
        total_files = len(tasks)
        total_bytes = sum(s for _, _, s in tasks)
        if total_files == 0:
            done_cb(0, 0, 0.0)
            return

        progress_cb({
            "phase": "copy", "copied_files": 0, "total_files": total_files,
            "copied_bytes": 0, "total_bytes": total_bytes,
            "speed_bps": 0, "eta_str": "...", "current_file": "",
        })

        t0 = time.time()
        copied_files = 0
        copied_bytes = 0
        last_up = 0
        lock = threading.Lock()
        copy_errors = []
        base = os.path.dirname(tasks[0][0]) if tasks else ""

        def _do(task):
            if cancel_event.is_set():
                return None, 0, None
            s, d, sz = task
            try:
                shutil.copy2(s, d)
                return s, sz, None
            except Exception as e:
                log_error("copy_file", e, f"src={s} dst={d}")
                return s, sz, str(e)

        with ThreadPoolExecutor(max_workers=workers) as pool:
            futs = {pool.submit(_do, t): t for t in tasks}
            for f in as_completed(futs):
                if cancel_event.is_set():
                    pool.shutdown(wait=False, cancel_futures=True)
                    return
                s, sz, err = f.result()
                if s is None:
                    continue
                if err:
                    copy_errors.append(f"{os.path.basename(s)}: {err}")
                with lock:
                    copied_files += 1
                    copied_bytes += sz
                    now = time.time()
                    if now - last_up >= 0.1:
                        el = now - t0
                        spd = copied_bytes / el if el > 0 else 0
                        rem = total_bytes - copied_bytes
                        eta = rem / spd if spd > 0 else 0
                        try:
                            disp = os.path.relpath(s, base)
                        except ValueError:
                            disp = os.path.basename(s)
                        progress_cb({
                            "phase": "copy",
                            "copied_files": copied_files,
                            "total_files": total_files,
                            "copied_bytes": copied_bytes,
                            "total_bytes": total_bytes,
                            "speed_bps": spd,
                            "eta_str": format_eta(eta),
                            "current_file": disp,
                        })
                        last_up = now

        # 删除源文件（剪切模式）
        delete_errors = []
        if delete_sources and not cancel_event.is_set():
            progress_cb({"phase": "scan", "message": "deleting..."})
            for sp in delete_sources:
                try:
                    if os.path.isdir(sp):
                        shutil.rmtree(sp, onerror=_rmtree_onerror)
                    elif os.path.isfile(sp):
                        try:
                            os.remove(sp)
                        except PermissionError:
                            os.chmod(sp, stat.S_IWRITE | stat.S_IREAD)
                            os.remove(sp)
                except Exception as e:
                    log_error("delete_source", e, f"path={sp}")
                    delete_errors.append(os.path.basename(sp))

        all_errors = copy_errors + [f"[Delete] {x}" for x in delete_errors]
        if all_errors:
            logger.warning(f"Operation completed with {len(all_errors)} errors")

        done_cb(copied_files, total_bytes, time.time() - t0,
                delete_errors=delete_errors, copy_errors=copy_errors)
    except Exception as e:
        log_error("copy_engine", e)
        error_cb(e)


# ============================================================
# E. GUI 主类
# ============================================================

class PersonalAssistant(tk.Tk):

    def __init__(self):
        super().__init__()
        self.config_data = load_config()
        self.lang = self.config_data.get("language", "zh")
        self.cancel_event = threading.Event()
        self.copying = False
        self.last_dest = None
        self.last_transfer_dest = None
        # 存储需要切换语言的控件 [(widget, key, field), ...]
        self._i18n = []

        self.title("Personal Assistant — by MeiKuaiKuai")
        self.geometry("700x650")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")

        self._build_top_bar()
        self._build_notebook()
        self._build_footer()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def t(self, key):
        return TEXTS.get(self.lang, TEXTS["zh"]).get(key, key)

    def _reg(self, widget, key, field="text"):
        """注册控件到多语言系统"""
        self._i18n.append((widget, key, field))
        if field == "text":
            widget.config(text=self.t(key))
        return widget

    def _switch_language(self, lang):
        self.lang = lang
        self.config_data["language"] = lang
        save_config(self.config_data)
        # 更新按钮样式
        if lang == "zh":
            self._zh_btn.config(relief="sunken", bg="#1abc9c", fg="white")
            self._en_btn.config(relief="flat", bg="#465c6e", fg="#bdc3c7")
        else:
            self._en_btn.config(relief="sunken", bg="#1abc9c", fg="white")
            self._zh_btn.config(relief="flat", bg="#465c6e", fg="#bdc3c7")
        # 更新所有注册控件
        for widget, key, field in self._i18n:
            try:
                widget[field] = self.t(key)
            except Exception:
                pass
        # 更新 Notebook 标签名
        self.notebook.tab(0, text=self.t("tab_copy"))
        self.notebook.tab(1, text=self.t("tab_transfer"))
        # 更新排序下拉框选项文字
        self._update_sort_options()
        # 刷新文件列表（标签前缀变了）
        self._tr_refresh_list()

    # ---- 顶部栏 ----

    def _build_top_bar(self):
        top = tk.Frame(self, bg="#2c3e50", height=50)
        top.pack(fill="x")
        top.pack_propagate(False)

        tk.Label(top, text="Personal Assistant",
                 font=("Segoe UI", 16, "bold"),
                 fg="white", bg="#2c3e50").pack(side="left", padx=15)

        tk.Label(top, text="v1.0", font=("Segoe UI", 9),
                 fg="#95a5a6", bg="#2c3e50").pack(side="left", pady=(8, 0))

        # 语言切换
        lang_frame = tk.Frame(top, bg="#2c3e50")
        lang_frame.pack(side="right", padx=8)

        self._en_btn = tk.Button(
            lang_frame, text="EN", font=("Segoe UI", 9, "bold"), width=4,
            fg="#bdc3c7", bg="#465c6e", relief="flat", bd=0, cursor="hand2",
            command=lambda: self._switch_language("en"))
        self._en_btn.pack(side="right", padx=1)

        self._zh_btn = tk.Button(
            lang_frame, text="中文", font=("Segoe UI", 9, "bold"), width=4,
            fg="white", bg="#1abc9c", relief="sunken", bd=0, cursor="hand2",
            command=lambda: self._switch_language("zh"))
        self._zh_btn.pack(side="right", padx=1)

        # 如果当前语言是英文，初始化按钮样式
        if self.lang == "en":
            self._zh_btn.config(relief="flat", bg="#465c6e", fg="#bdc3c7")
            self._en_btn.config(relief="sunken", bg="#1abc9c", fg="white")

        # 支持按钮
        self._support_btn = tk.Button(
            top, font=("Segoe UI", 10), fg="white", bg="#e74c3c",
            activebackground="#c0392b", activeforeground="white",
            relief="flat", cursor="hand2", bd=0,
            command=self._show_support_dialog)
        self._reg(self._support_btn, "support_btn")
        self._support_btn.pack(side="right", padx=8, pady=10)

        # 日志按钮
        self._log_btn = tk.Button(
            top, font=("Segoe UI", 10), fg="white", bg="#f39c12",
            activebackground="#e67e22", activeforeground="white",
            relief="flat", cursor="hand2", bd=0,
            command=self._show_log_dialog)
        self._reg(self._log_btn, "log_btn")
        self._log_btn.pack(side="right", pady=10)

    # ---- 日志弹窗 ----

    def _show_log_dialog(self):
        d = tk.Toplevel(self)
        d.title(self.t("log_title"))
        d.geometry("650x500")
        d.resizable(True, True)
        d.configure(bg="white")
        d.transient(self)
        d.grab_set()

        tk.Frame(d, bg="#f39c12", height=6).pack(fill="x")

        tk.Label(d, text=self.t("log_title"),
                 font=("Segoe UI", 16, "bold"), fg="#2c3e50", bg="white"
                 ).pack(pady=(15, 5))

        # 系统信息
        sys_info = get_system_info()
        tk.Label(d, text=sys_info, font=("Consolas", 8),
                 fg="#999", bg="white").pack(pady=(0, 10))

        # 日志内容
        tf = tk.Frame(d, bg="white")
        tf.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        log_text = tk.Text(tf, wrap="word", font=("Consolas", 9),
                           bg="#f8f8f8", fg="#333", relief="solid", bd=1)
        sb = ttk.Scrollbar(tf, command=log_text.yview)
        log_text.config(yscrollcommand=sb.set)
        log_text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 读取日志
        try:
            if os.path.isfile(LOG_FILE):
                with open(LOG_FILE, "r", encoding="utf-8") as f:
                    content = f.read()
                if content.strip():
                    log_text.insert("1.0", content)
                else:
                    log_text.insert("1.0", self.t("log_empty"))
            else:
                log_text.insert("1.0", self.t("log_empty"))
        except Exception as e:
            log_text.insert("1.0", f"Error reading log: {e}")
        log_text.config(state="disabled")

        # 按钮栏
        btn_frame = tk.Frame(d, bg="white")
        btn_frame.pack(pady=(0, 15))

        def _copy_log():
            try:
                log_text.config(state="normal")
                content = log_text.get("1.0", "end-1c")
                log_text.config(state="disabled")
                full_report = f"=== System Info ===\n{sys_info}\n\n=== Log ===\n{content}"
                self.clipboard_clear()
                self.clipboard_append(full_report)
                messagebox.showinfo(self.t("t_done"), self.t("log_copy_ok"))
            except Exception as e:
                messagebox.showerror(self.t("t_err"), str(e))

        def _open_log():
            if os.path.isfile(LOG_FILE):
                if sys.platform == "win32":
                    os.startfile(LOG_FILE)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", LOG_FILE])
                else:
                    subprocess.Popen(["xdg-open", LOG_FILE])

        def _clear_log():
            try:
                with open(LOG_FILE, "w", encoding="utf-8") as f:
                    f.write("")
                log_text.config(state="normal")
                log_text.delete("1.0", tk.END)
                log_text.insert("1.0", self.t("log_empty"))
                log_text.config(state="disabled")
                messagebox.showinfo(self.t("t_done"), self.t("log_cleared"))
            except Exception as e:
                messagebox.showerror(self.t("t_err"), str(e))

        tk.Button(btn_frame, text=self.t("log_copy_btn"),
                  font=("Segoe UI", 10), fg="white", bg="#3498db",
                  activebackground="#2980b9", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=15, pady=5,
                  command=_copy_log).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t("log_open_btn"),
                  font=("Segoe UI", 10), fg="white", bg="#2ecc71",
                  activebackground="#27ae60", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=15, pady=5,
                  command=_open_log).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t("log_clear_btn"),
                  font=("Segoe UI", 10), fg="white", bg="#e74c3c",
                  activebackground="#c0392b", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=15, pady=5,
                  command=_clear_log).pack(side="left", padx=5)

    # ---- 支持弹窗 ----

    def _show_support_dialog(self):
        d = tk.Toplevel(self)
        d.title(self.t("s_title"))
        d.geometry("580x700")
        d.resizable(False, False)
        d.configure(bg="white")
        d.transient(self)
        d.grab_set()

        tk.Frame(d, bg="#2c3e50", height=8).pack(fill="x")

        tk.Label(d, text=self.t("s_thanks"),
                 font=("Segoe UI", 11), fg="#7f8c8d", bg="white"
                 ).pack(pady=(25, 0))
        tk.Label(d, text="Personal Assistant",
                 font=("Segoe UI", 22, "bold"), fg="#2c3e50", bg="white"
                 ).pack(pady=(2, 15))

        ttk.Separator(d, orient="horizontal").pack(fill="x", padx=40)

        tk.Label(d, text=self.t("s_body"),
                 font=("Segoe UI", 10), fg="#555", bg="white",
                 justify="center", wraplength=480).pack(pady=(20, 20))

        ttk.Separator(d, orient="horizontal").pack(fill="x", padx=40)

        qr = tk.Frame(d, bg="white")
        qr.pack(pady=(15, 10))

        # 加微信
        wf = tk.Frame(qr, bg="white")
        wf.pack(side="left", padx=25)
        tk.Label(wf, text=self.t("s_wechat"),
                 font=("Segoe UI", 11, "bold"), fg="#07C160", bg="white"
                 ).pack(pady=(0, 8))
        self._show_qr(wf, os.path.join("支付系统", "加微信.jpg"))
        tk.Label(wf, text=self.t("s_wechat_hint"),
                 font=("Segoe UI", 9), fg="#999", bg="white").pack(pady=(6, 0))

        # 微信支付
        pf = tk.Frame(qr, bg="white")
        pf.pack(side="left", padx=25)
        tk.Label(pf, text=self.t("s_pay"),
                 font=("Segoe UI", 11, "bold"), fg="#e74c3c", bg="white"
                 ).pack(pady=(0, 8))
        self._show_qr(pf, os.path.join("支付系统", "微信支付.jpg"))
        tk.Label(pf, text=self.t("s_pay_hint"),
                 font=("Segoe UI", 9), fg="#999", bg="white").pack(pady=(6, 0))

        tk.Label(d, text=self.t("s_footer"),
                 font=("Segoe UI", 9), fg="#bdc3c7", bg="white"
                 ).pack(side="bottom", pady=(0, 15))

        tk.Button(d, text=self.t("s_close"),
                  font=("Segoe UI", 10), fg="white", bg="#2c3e50",
                  activebackground="#34495e", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0,
                  command=d.destroy).pack(side="bottom", pady=(0, 10))

    def _show_qr(self, parent, rel_path, size=180):
        """加载并显示二维码图片"""
        img_path = get_resource_path(rel_path)

        if not HAS_PIL:
            tk.Label(parent, text=self.t("s_no_pil"),
                     font=("Segoe UI", 9), fg="#e74c3c", bg="#fef0f0",
                     width=24, height=8, relief="solid", bd=1).pack()
            return

        if not os.path.isfile(img_path):
            tk.Label(parent, text=f"Image not found:\n{rel_path}",
                     font=("Segoe UI", 9), fg="#999", bg="#f0f0f0",
                     width=24, height=8, relief="solid", bd=1).pack()
            return

        try:
            img = Image.open(img_path)
            img = img.resize((size, size), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(parent, image=photo, bg="white",
                           relief="solid", bd=1)
            lbl.image = photo
            lbl.pack()
        except Exception as e:
            tk.Label(parent, text=f"Load error:\n{e}",
                     font=("Segoe UI", 8), fg="#e74c3c", bg="#fef0f0",
                     width=24, height=8, relief="solid", bd=1,
                     wraplength=160).pack()

    # ---- 标签页 ----

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        tab1 = ttk.Frame(self.notebook)
        tab2 = ttk.Frame(self.notebook)
        self.notebook.add(tab1, text=self.t("tab_copy"))
        self.notebook.add(tab2, text=self.t("tab_transfer"))
        self._build_tab_copy(tab1)
        self._build_tab_transfer(tab2)

    def _build_footer(self):
        f = ttk.Label(self, foreground="gray")
        self._reg(f, "footer")
        f.pack(side="bottom", pady=(0, 5))

    # ============================================================
    # Tab 1: 项目复制
    # ============================================================

    def _build_tab_copy(self, p):
        pad = {"padx": 10, "pady": 5}

        self.cp_source_var = tk.StringVar()
        self.cp_name_var = tk.StringVar()
        self.cp_dest_var = tk.StringVar()
        self.cp_remember_var = tk.BooleanVar(
            value=bool(self.config_data.get("default_directory")))
        self.cp_status_var = tk.StringVar(value=self.t("ready"))
        self.cp_speed_var = tk.StringVar()
        self.cp_file_var = tk.StringVar()
        self.cp_progress_var = tk.DoubleVar()

        sf = ttk.LabelFrame(p, padding=10)
        self._reg(sf, "source_project", "text")
        sf.pack(fill="x", **pad)

        r = ttk.Frame(sf)
        r.pack(fill="x")
        ttk.Entry(r, textvariable=self.cp_source_var,
                  state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.cp_browse_btn = ttk.Button(r, command=self._cp_browse)
        self._reg(self.cp_browse_btn, "browse")
        self.cp_browse_btn.pack(side="right")

        self.cp_remember_cb = ttk.Checkbutton(
            sf, variable=self.cp_remember_var, command=self._cp_on_remember)
        self._reg(self.cp_remember_cb, "remember_dir")
        self.cp_remember_cb.pack(anchor="w", pady=(5, 0))

        nf = ttk.LabelFrame(p, padding=10)
        self._reg(nf, "copy_name", "text")
        nf.pack(fill="x", **pad)

        self.cp_name_entry = ttk.Entry(nf, textvariable=self.cp_name_var)
        self.cp_name_entry.pack(fill="x")
        h = ttk.Label(nf, foreground="gray")
        self._reg(h, "auto_hint")
        h.pack(anchor="w", pady=(2, 0))
        ttk.Label(nf, textvariable=self.cp_dest_var,
                  foreground="#0066cc").pack(anchor="w", pady=(2, 0))
        self.cp_name_var.trace_add("write", self._cp_update_preview)

        af = ttk.LabelFrame(p, padding=10)
        self._reg(af, "action", "text")
        af.pack(fill="x", **pad)

        br = ttk.Frame(af)
        br.pack(fill="x")
        self.cp_copy_btn = ttk.Button(br, command=self._cp_start)
        self._reg(self.cp_copy_btn, "start_copy")
        self.cp_copy_btn.pack(side="left", fill="x", expand=True, ipady=8)
        self.cp_open_btn = ttk.Button(
            br, command=lambda: open_folder(self.last_dest))
        self._reg(self.cp_open_btn, "open_dest")
        self.cp_open_btn.pack(side="right", padx=(8, 0), ipady=8)
        self.cp_open_btn.config(state="disabled")

        ttk.Progressbar(af, variable=self.cp_progress_var,
                        maximum=100).pack(fill="x", pady=(8, 0))
        ttk.Label(af, textvariable=self.cp_status_var,
                  foreground="gray").pack(anchor="w", pady=(4, 0))
        ttk.Label(af, textvariable=self.cp_speed_var,
                  foreground="#006600").pack(anchor="w", pady=(2, 0))
        ttk.Label(af, textvariable=self.cp_file_var,
                  foreground="gray").pack(anchor="w", pady=(2, 0))

    def _cp_browse(self):
        ini = self.config_data.get("default_directory", "")
        if not ini or not os.path.isdir(ini):
            ini = str(Path.home())
        f = filedialog.askdirectory(title=self.t("d_sel_proj"), initialdir=ini)
        if not f:
            return
        self.cp_source_var.set(f)
        self.cp_name_var.set(generate_copy_name(f))
        if self.cp_remember_var.get():
            self.config_data["default_directory"] = str(Path(f).parent)
            save_config(self.config_data)

    def _cp_on_remember(self):
        if self.cp_remember_var.get() and self.cp_source_var.get():
            self.config_data["default_directory"] = str(
                Path(self.cp_source_var.get()).parent)
        else:
            self.config_data["default_directory"] = ""
        save_config(self.config_data)

    def _cp_update_preview(self, *_):
        s = self.cp_source_var.get()
        n = self.cp_name_var.get().strip()
        if s and n:
            self.cp_dest_var.set(f"{self.t('copy_name')}: "
                                f"{os.path.join(str(Path(s).parent), n)}")
        else:
            self.cp_dest_var.set("")

    def _cp_start(self):
        src = self.cp_source_var.get().strip()
        name = self.cp_name_var.get().strip()
        if not src:
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_source"))
            return
        if not os.path.isdir(src):
            messagebox.showerror(self.t("t_err"), self.t("m_source_gone"))
            return
        if not name:
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_name"))
            return
        if INVALID_CHARS.search(name):
            messagebox.showwarning(self.t("t_hint"), self.t("m_bad_chars"))
            return
        dst = os.path.join(str(Path(src).parent), name)
        if os.path.exists(dst):
            messagebox.showwarning(self.t("t_hint"),
                                   self.t("m_dest_exists").format(dst))
            return

        self.copying = True
        self.last_dest = dst
        self.cancel_event.clear()
        self._cp_set_enabled(False)
        self.cp_progress_var.set(0)

        dt, w = detect_disk_type(src)
        self.cp_status_var.set(self.t("m_disk").format(dt, w))

        def pcb(info):
            self.after(0, self._cp_progress, info)

        def dcb(c, tb, el, **kw):
            self.after(0, self._cp_done, c, tb, el, kw)

        def ecb(e):
            self.after(0, self._cp_error, e)

        def run():
            try:
                pcb({"phase": "scan", "message": self.t("m_scanning")})
                tasks = _collect_copy_tasks(src, dst)
                run_copy_engine(tasks, pcb, dcb, ecb, self.cancel_event, w)
            except Exception as e:
                log_error("cp_start", e)
                ecb(e)

        threading.Thread(target=run, daemon=True).start()

    def _cp_progress(self, info):
        if info.get("phase") == "scan":
            self.cp_status_var.set(info["message"])
            self.cp_speed_var.set("")
            self.cp_file_var.set("")
            return
        c, t = info["copied_files"], info["total_files"]
        if t > 0:
            self.cp_progress_var.set(c / t * 100)
        self.cp_status_var.set(self.t("m_progress").format(
            c, t, format_size(info["copied_bytes"]),
            format_size(info["total_bytes"])))
        self.cp_speed_var.set(self.t("m_speed").format(
            format_size(int(info["speed_bps"])), info["eta_str"]))
        if info.get("current_file"):
            self.cp_file_var.set(self.t("m_cur_file").format(
                info["current_file"]))

    def _cp_done(self, count, tb, elapsed, extras=None):
        extras = extras or {}
        copy_errors = extras.get("copy_errors", [])
        self.cp_progress_var.set(100)
        ss = format_size(tb)
        avg = format_size(int(tb / elapsed)) + "/s" if elapsed > 0 else ""
        self.cp_status_var.set(self.t("m_done").format(count, ss, f"{elapsed:.1f}"))
        self.cp_speed_var.set(self.t("m_avg").format(avg) if avg else "")
        self.cp_file_var.set("")
        self.copying = False
        self._cp_set_enabled(True)
        self.cp_open_btn.config(state="normal")
        msg = self.t("m_ok_copy").format(count, ss, elapsed, avg)
        if copy_errors:
            msg += "\n\n" + self.t("m_move_partial").format(
                "\n".join(copy_errors[:10]))
        messagebox.showinfo(self.t("t_done"), msg)
        s = self.cp_source_var.get()
        if s:
            self.cp_name_var.set(generate_copy_name(s))

    def _cp_error(self, e):
        log_error("cp_error_ui", e)
        self.cp_status_var.set(self.t("m_err").format(e))
        self.cp_speed_var.set("")
        self.cp_file_var.set("")
        self.copying = False
        self._cp_set_enabled(True)
        messagebox.showerror(self.t("t_fail_copy"),
                             self.t("m_err_copy").format(e))

    def _cp_set_enabled(self, on):
        st = "normal" if on else "disabled"
        self.cp_browse_btn.config(state=st)
        self.cp_name_entry.config(state=st)
        self.cp_copy_btn.config(state=st)
        self.cp_remember_cb.config(state=st)
        self.cp_copy_btn.config(
            text=self.t("start_copy") if on else self.t("copying"))

    # ============================================================
    # Tab 2: 文件搬运
    # ============================================================

    def _build_tab_transfer(self, p):
        pad = {"padx": 10, "pady": 4}

        self.tr_src_var = tk.StringVar()
        self.tr_dst_var = tk.StringVar()
        self.tr_rename_var = tk.StringVar()
        self.tr_op_var = tk.StringVar(value="copy")
        self.tr_status_var = tk.StringVar(value=self.t("ready"))
        self.tr_speed_var = tk.StringVar()
        self.tr_file_var = tk.StringVar()
        self.tr_progress_var = tk.DoubleVar()

        # 常用地址
        ff = ttk.LabelFrame(p, padding=8)
        self._reg(ff, "favorites", "text")
        ff.pack(fill="x", **pad)

        fr = ttk.Frame(ff)
        fr.pack(fill="x")
        self.fav_combo = ttk.Combobox(fr, state="readonly", width=50)
        self.fav_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self._refresh_favs()

        self._fav_src_btn = ttk.Button(fr, width=7,
            command=lambda: self._fav_fill(self.tr_src_var))
        self._reg(self._fav_src_btn, "fill_src")
        self._fav_src_btn.pack(side="left", padx=2)

        self._fav_dst_btn = ttk.Button(fr, width=7,
            command=lambda: self._fav_fill(self.tr_dst_var))
        self._reg(self._fav_dst_btn, "fill_dst")
        self._fav_dst_btn.pack(side="left", padx=2)

        self._fav_add_btn = ttk.Button(fr, width=6, command=self._fav_add)
        self._reg(self._fav_add_btn, "add_fav")
        self._fav_add_btn.pack(side="left", padx=2)

        self._fav_del_btn = ttk.Button(fr, width=6, command=self._fav_remove)
        self._reg(self._fav_del_btn, "del_fav")
        self._fav_del_btn.pack(side="left", padx=2)

        # 源地址
        sf = ttk.LabelFrame(p, padding=8)
        self._reg(sf, "source_a", "text")
        sf.pack(fill="x", **pad)

        r = ttk.Frame(sf)
        r.pack(fill="x")
        ttk.Entry(r, textvariable=self.tr_src_var
                  ).pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.tr_src_browse = ttk.Button(r, command=self._tr_browse_src)
        self._reg(self.tr_src_browse, "browse")
        self.tr_src_browse.pack(side="right")

        lf = ttk.Frame(sf)
        lf.pack(fill="both", expand=True, pady=(5, 0))
        self.tr_listbox = tk.Listbox(lf, selectmode=tk.EXTENDED, height=6)
        sb = ttk.Scrollbar(lf, command=self.tr_listbox.yview)
        self.tr_listbox.config(yscrollcommand=sb.set)
        self.tr_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        sr = ttk.Frame(sf)
        sr.pack(fill="x", pady=(3, 0))
        self._sel_all_btn = ttk.Button(sr, width=6, command=self._tr_sel_all)
        self._reg(self._sel_all_btn, "select_all")
        self._sel_all_btn.pack(side="left", padx=(0, 5))

        self._desel_btn = ttk.Button(sr, width=8, command=self._tr_desel_all)
        self._reg(self._desel_btn, "deselect_all")
        self._desel_btn.pack(side="left", padx=(0, 5))

        self._refresh_btn = ttk.Button(sr, width=8,
                                       command=self._tr_refresh_list)
        self._reg(self._refresh_btn, "refresh")
        self._refresh_btn.pack(side="left")

        # 排序
        self._sort_lbl = ttk.Label(sr)
        self._reg(self._sort_lbl, "sort_label")
        self._sort_lbl.pack(side="left", padx=(15, 3))

        self._sort_keys = [
            "sort_name_asc", "sort_name_desc",
            "sort_time_new", "sort_time_old",
            "sort_size_big", "sort_size_small",
        ]
        self.tr_sort_var = tk.StringVar(value=self.t("sort_name_asc"))
        self.tr_sort_combo = ttk.Combobox(
            sr, textvariable=self.tr_sort_var, state="readonly", width=14)
        self._update_sort_options()
        self.tr_sort_combo.pack(side="left")
        self.tr_sort_combo.bind("<<ComboboxSelected>>",
                                lambda _: self._tr_refresh_list())

        self.tr_src_var.trace_add("write", lambda *_: self._tr_refresh_list())

        # 目标地址
        df = ttk.LabelFrame(p, padding=8)
        self._reg(df, "dest_b", "text")
        df.pack(fill="x", **pad)

        r2 = ttk.Frame(df)
        r2.pack(fill="x")
        ttk.Entry(r2, textvariable=self.tr_dst_var
                  ).pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.tr_dst_browse = ttk.Button(r2, command=self._tr_browse_dst)
        self._reg(self.tr_dst_browse, "browse")
        self.tr_dst_browse.pack(side="right")

        # 选项行
        of = ttk.Frame(p)
        of.pack(fill="x", **pad)

        self._op_lbl = ttk.Label(of)
        self._reg(self._op_lbl, "op_label")
        self._op_lbl.pack(side="left")

        self._op_copy_rb = ttk.Radiobutton(of, variable=self.tr_op_var,
                                           value="copy")
        self._reg(self._op_copy_rb, "op_copy")
        self._op_copy_rb.pack(side="left", padx=(5, 10))

        self._op_move_rb = ttk.Radiobutton(of, variable=self.tr_op_var,
                                           value="move")
        self._reg(self._op_move_rb, "op_move")
        self._op_move_rb.pack(side="left", padx=(0, 20))

        self._ren_lbl = ttk.Label(of)
        self._reg(self._ren_lbl, "rename_label")
        self._ren_lbl.pack(side="left")

        self.tr_rename_entry = ttk.Entry(of, textvariable=self.tr_rename_var,
                                         width=20)
        self.tr_rename_entry.pack(side="left", padx=(5, 0))

        self._ren_hint = ttk.Label(of, foreground="gray")
        self._reg(self._ren_hint, "rename_hint")
        self._ren_hint.pack(side="left", padx=(5, 0))

        # 执行
        ef = ttk.LabelFrame(p, padding=8)
        self._reg(ef, "exec_section", "text")
        ef.pack(fill="x", **pad)

        br = ttk.Frame(ef)
        br.pack(fill="x")
        self.tr_exec_btn = ttk.Button(br, command=self._tr_start)
        self._reg(self.tr_exec_btn, "execute")
        self.tr_exec_btn.pack(side="left", fill="x", expand=True, ipady=6)

        self.tr_del_btn = tk.Button(
            br, command=self._tr_delete,
            fg="white", bg="#e74c3c", activebackground="#c0392b",
            activeforeground="white", relief="flat", cursor="hand2", bd=0)
        self._reg(self.tr_del_btn, "delete_btn")
        self.tr_del_btn.pack(side="left", padx=(8, 0), ipady=6, ipadx=8)

        self.tr_open_btn = ttk.Button(
            br, command=lambda: open_folder(self.last_transfer_dest))
        self._reg(self.tr_open_btn, "open_dest")
        self.tr_open_btn.pack(side="right", padx=(8, 0), ipady=6)
        self.tr_open_btn.config(state="disabled")

        ttk.Progressbar(ef, variable=self.tr_progress_var,
                        maximum=100).pack(fill="x", pady=(6, 0))
        ttk.Label(ef, textvariable=self.tr_status_var,
                  foreground="gray").pack(anchor="w", pady=(3, 0))
        ttk.Label(ef, textvariable=self.tr_speed_var,
                  foreground="#006600").pack(anchor="w", pady=(1, 0))
        ttk.Label(ef, textvariable=self.tr_file_var,
                  foreground="gray").pack(anchor="w", pady=(1, 0))

    # -- 常用地址 --

    def _refresh_favs(self):
        favs = self.config_data.get("favorites", [])
        self.fav_combo["values"] = favs
        if favs:
            self.fav_combo.current(0)

    def _fav_fill(self, var):
        v = self.fav_combo.get()
        if v:
            var.set(v)

    def _fav_add(self):
        f = filedialog.askdirectory(title=self.t("d_sel_fav"))
        if not f:
            return
        favs = self.config_data.get("favorites", [])
        if f not in favs:
            favs.append(f)
            self.config_data["favorites"] = favs
            save_config(self.config_data)
            self._refresh_favs()
            self.fav_combo.set(f)

    def _fav_remove(self):
        v = self.fav_combo.get()
        if not v:
            return
        favs = self.config_data.get("favorites", [])
        if v in favs:
            favs.remove(v)
            self.config_data["favorites"] = favs
            save_config(self.config_data)
            self._refresh_favs()

    # -- Tab2 事件 --

    def _tr_browse_src(self):
        f = filedialog.askdirectory(title=self.t("d_sel_src"))
        if f:
            self.tr_src_var.set(f)

    def _tr_browse_dst(self):
        f = filedialog.askdirectory(title=self.t("d_sel_dst"))
        if f:
            self.tr_dst_var.set(f)

    def _update_sort_options(self):
        """更新排序下拉框的选项文字"""
        vals = [self.t(k) for k in self._sort_keys]
        cur = self.tr_sort_combo.current()
        self.tr_sort_combo["values"] = vals
        if cur >= 0:
            self.tr_sort_combo.current(cur)
        else:
            self.tr_sort_combo.current(0)

    def _get_sort_index(self):
        """获取当前排序选项的索引"""
        return self.tr_sort_combo.current() if self.tr_sort_combo.current() >= 0 else 0

    def _tr_refresh_list(self):
        self.tr_listbox.delete(0, tk.END)
        src = self.tr_src_var.get().strip()
        if not src or not os.path.isdir(src):
            return
        try:
            items = os.listdir(src)
        except OSError:
            return

        # 收集文件信息
        entries = []
        for name in items:
            full = os.path.join(src, name)
            is_dir = os.path.isdir(full)
            try:
                stat = os.stat(full)
                mtime = stat.st_mtime
                size = 0 if is_dir else stat.st_size
            except OSError:
                mtime, size = 0, 0
            entries.append((name, is_dir, mtime, size, full))

        # 分离文件夹和文件，分别排序（文件夹始终在前）
        dirs = [e for e in entries if e[1]]
        files = [e for e in entries if not e[1]]

        sort_idx = self._get_sort_index()
        if sort_idx == 0:    # 名称 A→Z
            key, rev = lambda e: e[0].lower(), False
        elif sort_idx == 1:  # 名称 Z→A
            key, rev = lambda e: e[0].lower(), True
        elif sort_idx == 2:  # 时间 新→旧
            key, rev = lambda e: e[2], True
        elif sort_idx == 3:  # 时间 旧→新
            key, rev = lambda e: e[2], False
        elif sort_idx == 4:  # 大小 大→小
            key, rev = lambda e: e[3], True
        elif sort_idx == 5:  # 大小 小→大
            key, rev = lambda e: e[3], False
        else:
            key, rev = lambda e: e[0].lower(), False

        dirs.sort(key=key, reverse=rev)
        files.sort(key=key, reverse=rev)
        entries = dirs + files

        for name, is_dir, mtime, size, full in entries:
            if is_dir:
                self.tr_listbox.insert(
                    tk.END, f"{self.t('folder_tag')} {name}")
            else:
                self.tr_listbox.insert(
                    tk.END, f"{self.t('file_tag')} {name}  ({format_size(size)})")

    def _tr_sel_all(self):
        self.tr_listbox.select_set(0, tk.END)

    def _tr_desel_all(self):
        self.tr_listbox.select_clear(0, tk.END)

    def _tr_delete(self):
        """删除选中的文件/文件夹"""
        selected = self._tr_get_paths()
        if not selected:
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_items"))
            return

        # 显示要删除的项目名称列表（最多显示10个）
        names = [os.path.basename(p) for p in selected]
        display = "\n".join(names[:10])
        if len(names) > 10:
            display += f"\n... (+{len(names) - 10})"

        if not messagebox.askyesno(
                self.t("t_del_confirm"),
                self.t("m_del_confirm").format(len(selected), display)):
            return

        # 执行删除
        self.tr_del_btn.config(state="disabled",
                               text=self.t("deleting_btn"))
        self.update_idletasks()

        errors = []
        deleted = 0
        for p in selected:
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p, onerror=_rmtree_onerror)
                elif os.path.isfile(p):
                    try:
                        os.remove(p)
                    except PermissionError:
                        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
                        os.remove(p)
                deleted += 1
            except Exception as e:
                log_error("delete_item", e, f"path={p}")
                errors.append(f"{os.path.basename(p)}: {e}")

        self.tr_del_btn.config(state="normal",
                               text=self.t("delete_btn"))

        if errors:
            messagebox.showerror(
                self.t("t_del_fail"),
                self.t("m_del_err").format("\n".join(errors)))
        else:
            messagebox.showinfo(
                self.t("t_done"),
                self.t("m_del_done").format(deleted))

        self._tr_refresh_list()

    def _tr_get_paths(self):
        src = self.tr_src_var.get().strip()
        if not src:
            return []
        ft = self.t("folder_tag")
        fit = self.t("file_tag")
        # 也兼容另一种语言的前缀（防止切换语言后旧列表项解析失败）
        all_folder_tags = [TEXTS["zh"]["folder_tag"], TEXTS["en"]["folder_tag"]]
        all_file_tags = [TEXTS["zh"]["file_tag"], TEXTS["en"]["file_tag"]]
        paths = []
        for idx in self.tr_listbox.curselection():
            text = self.tr_listbox.get(idx)
            name = text
            for tag in all_folder_tags:
                if text.startswith(f"{tag} "):
                    name = text[len(tag) + 1:]
                    break
            else:
                for tag in all_file_tags:
                    if text.startswith(f"{tag} "):
                        name = text[len(tag) + 1:]
                        paren = name.rfind("  (")
                        if paren > 0:
                            name = name[:paren]
                        break
            full = os.path.join(src, name)
            if os.path.exists(full):
                paths.append(full)
        return paths

    def _tr_start(self):
        src = self.tr_src_var.get().strip()
        dst = self.tr_dst_var.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_src_valid"))
            return
        if not dst:
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_dst"))
            return
        selected = self._tr_get_paths()
        if not selected:
            messagebox.showwarning(self.t("t_hint"), self.t("m_no_items"))
            return
        if not os.path.isdir(dst):
            if messagebox.askyesno(self.t("t_hint"),
                                   self.t("m_mkdir").format(dst)):
                os.makedirs(dst, exist_ok=True)
            else:
                return

        rename_map = None
        nn = self.tr_rename_var.get().strip()
        if nn and len(selected) == 1:
            if INVALID_CHARS.search(nn):
                messagebox.showwarning(self.t("t_hint"), self.t("m_bad_chars"))
                return
            rename_map = {os.path.basename(selected[0]): nn}

        for ip in selected:
            iname = os.path.basename(ip)
            dname = iname
            if rename_map and iname in rename_map:
                dname = rename_map[iname]
            if os.path.exists(os.path.join(dst, dname)):
                messagebox.showwarning(
                    self.t("t_hint"),
                    self.t("m_conflict").format(os.path.join(dst, dname)))
                return

        is_move = self.tr_op_var.get() == "move"
        op_t = self.t("m_op_cut") if is_move else self.t("m_op_copy")

        self.copying = True
        self.last_transfer_dest = dst
        self.cancel_event.clear()
        self._tr_set_enabled(False)
        self.tr_progress_var.set(0)

        # 检测磁盘类型并显示
        dt_src, w_src = detect_disk_type(src)
        dt_dst, w_dst = detect_disk_type(dst)
        w = min(w_src, w_dst)
        self.tr_status_var.set(
            self.t("m_disk_transfer").format(dt_src, dt_dst, w))
        self.tr_speed_var.set("")
        self.tr_file_var.set("")
        self.update_idletasks()

        def pcb(info):
            self.after(0, self._tr_progress, info)

        def dcb(c, tb, el, **kw):
            self.after(0, self._tr_done, c, tb, el, op_t, is_move, kw)

        def ecb(e):
            self.after(0, self._tr_error, e)

        def run():
            try:
                pcb({"phase": "scan", "message": self.t("m_scanning")})
                tasks = _collect_transfer_tasks(selected, dst, rename_map)
                ds = selected if is_move else None
                run_copy_engine(tasks, pcb, dcb, ecb,
                                self.cancel_event, w, ds)
            except Exception as e:
                log_error("tr_start", e)
                ecb(e)

        threading.Thread(target=run, daemon=True).start()

    def _tr_progress(self, info):
        if info.get("phase") == "scan":
            self.tr_status_var.set(info.get("message", ""))
            self.tr_speed_var.set("")
            self.tr_file_var.set("")
            return
        c, t = info["copied_files"], info["total_files"]
        if t > 0:
            self.tr_progress_var.set(c / t * 100)
        self.tr_status_var.set(self.t("m_progress").format(
            c, t, format_size(info["copied_bytes"]),
            format_size(info["total_bytes"])))
        self.tr_speed_var.set(self.t("m_speed").format(
            format_size(int(info["speed_bps"])), info["eta_str"]))
        if info.get("current_file"):
            self.tr_file_var.set(self.t("m_cur_proc").format(
                info["current_file"]))

    def _tr_done(self, count, tb, elapsed, op_t, is_move, extras=None):
        extras = extras or {}
        delete_errors = extras.get("delete_errors", [])
        copy_errors = extras.get("copy_errors", [])
        self.tr_progress_var.set(100)
        ss = format_size(tb)
        avg = format_size(int(tb / elapsed)) + "/s" if elapsed > 0 else ""
        self.tr_status_var.set(
            self.t("m_tr_done").format(op_t, count, ss, f"{elapsed:.1f}"))
        self.tr_speed_var.set(self.t("m_avg").format(avg) if avg else "")
        self.tr_file_var.set("")
        self.copying = False
        self._tr_set_enabled(True)
        self.tr_open_btn.config(state="normal")
        msg = self.t("m_tr_ok").format(op_t, count, ss, elapsed, avg)
        if delete_errors:
            msg += "\n\n" + self.t("m_move_partial").format(
                "\n".join(delete_errors[:10]))
            if len(delete_errors) > 10:
                msg += f"\n... (+{len(delete_errors) - 10})"
        if copy_errors:
            msg += "\n\n" + self.t("m_move_partial").format(
                "\n".join(copy_errors[:10]))
        messagebox.showinfo(self.t("t_done"), msg)
        if is_move:
            self._tr_refresh_list()

    def _tr_error(self, e):
        log_error("tr_error_ui", e)
        self.tr_status_var.set(self.t("m_err").format(e))
        self.tr_speed_var.set("")
        self.tr_file_var.set("")
        self.copying = False
        self._tr_set_enabled(True)
        messagebox.showerror(self.t("t_fail_tr"), self.t("m_tr_err").format(e))

    def _tr_set_enabled(self, on):
        st = "normal" if on else "disabled"
        self.tr_src_browse.config(state=st)
        self.tr_dst_browse.config(state=st)
        self.tr_exec_btn.config(state=st)
        self.tr_del_btn.config(state=st)
        self.tr_rename_entry.config(state=st)
        self.tr_exec_btn.config(
            text=self.t("execute") if on else self.t("transferring"))

    # ---- 公共 ----

    def _on_close(self):
        if self.copying:
            if messagebox.askyesno(self.t("t_confirm"), self.t("m_quit")):
                self.cancel_event.set()
                self.destroy()
        else:
            self.destroy()


# ============================================================
# F. 入口
# ============================================================

if __name__ == "__main__":
    app = PersonalAssistant()
    app.mainloop()
