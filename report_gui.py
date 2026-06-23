import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date, timezone, timedelta
import re
import textwrap
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("天气通报生成器")
        # ---------- 设置图标 ----------
        # 获取图标路径（兼容打包后和源码运行）
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, "icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        self.root.geometry("1280x750")
        self.root.resizable(True, True)

        # ---------- 样式 ----------
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=("微软雅黑", 12))
        style.configure("TLabel", background="#f0f4f8", foreground="#2c3e50")
        style.configure("TFrame", background="#f0f4f8")
        style.configure("TLabelframe", background="#f0f4f8", foreground="#2c3e50", borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", font=("微软雅黑", 13, "bold"), background="#f0f4f8", foreground="#1a5276")
        style.configure("TEntry", fieldbackground="white", borderwidth=1, relief="solid")
        style.configure("TCombobox", fieldbackground="white", borderwidth=1, relief="solid")
        style.configure("TButton", background="#2980b9", foreground="white", borderwidth=0, focuscolor="none")
        style.map("TButton", background=[("active", "#3498db")])
        style.configure("Accent.TButton", background="#27ae60", foreground="white")
        style.map("Accent.TButton", background=[("active", "#2ecc71")])

        # ---------- 路径配置 ----------
        if getattr(sys, 'frozen', False):
            # 打包后的临时目录
            self.app_dir = os.path.dirname(sys.executable)
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.stations_file = os.path.join(self.app_dir, "stations.json")
        self.config_file = os.path.join(self.app_dir, "config.json")

        # ---------- 站点数据 ----------
        self.default_stations = [
            "龙港芦浦华中", "龙港白沙河东", "龙港舥艚中段",
            "云岩瑞联", "芦浦东门垟", "云岩鲸头"
        ]
        self.stations = self.load_stations()

        # ---------- 保存目录（记忆上次，两个标签页独立） ----------
        saved = self.load_config()
        pre_dir = saved.get("pre_save_dir", "")
        live_dir = saved.get("live_save_dir", "")
        weather_dir = saved.get("weather_save_dir", "")
        cwd = os.getcwd()
        self.pre_save_dir = tk.StringVar(value=pre_dir if pre_dir and os.path.isdir(pre_dir) else cwd)
        self.live_save_dir = tk.StringVar(value=live_dir if live_dir and os.path.isdir(live_dir) else cwd)
        self.weather_save_dir = tk.StringVar(value=weather_dir if weather_dir and os.path.isdir(weather_dir) else cwd)
        risk_dir = saved.get("risk_save_dir", "")
        self.risk_save_dir = tk.StringVar(value=risk_dir if risk_dir and os.path.isdir(risk_dir) else cwd)
        self.risk_issue_number = saved.get("risk_issue_number", 1)
        lzy_dir = saved.get("lzy_save_dir", "")
        self.lzy_save_dir = tk.StringVar(value=lzy_dir if lzy_dir else cwd)

        # ---------- 主布局 ----------
        main_panel = ttk.Frame(root)
        main_panel.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------- 右侧站点管理面板 ----------
        self.right_frame = ttk.Frame(main_panel, padding=10, relief="solid", borderwidth=1)
        self.right_frame.pack(side="right", fill="y", padx=(10, 0))
        self.build_station_panel(self.right_frame)

        # ---------- 所有保存栏（先创建，避免标签页切换事件中引用未创建对象） ----------
        self.save_bar_pre = ttk.Frame(root)
        ttk.Label(self.save_bar_pre, text="保存目录：").pack(side="left")
        self.entry_save_pre = ttk.Entry(self.save_bar_pre, textvariable=self.pre_save_dir, width=32)
        self.entry_save_pre.pack(side="left", padx=5)
        ttk.Button(self.save_bar_pre, text="浏览...", command=self.browse_folder_pre).pack(side="left", padx=3)
        ttk.Button(self.save_bar_pre, text="✨ 生成 TXT 文件", command=self.generate_pre_report, style="Accent.TButton").pack(side="right", padx=5)

        self.save_bar_weather = ttk.Frame(root)
        ttk.Label(self.save_bar_weather, text="保存目录：").pack(side="left")
        self.entry_save_weather = ttk.Entry(self.save_bar_weather, textvariable=self.weather_save_dir, width=32)
        self.entry_save_weather.pack(side="left", padx=5)
        ttk.Button(self.save_bar_weather, text="浏览...", command=self.browse_folder_weather).pack(side="left", padx=3)
        ttk.Button(self.save_bar_weather, text="✨ 生成 TXT 文件", command=self.generate_weather_alert, style="Accent.TButton").pack(side="right", padx=5)

        self.save_bar_risk = ttk.Frame(root)
        ttk.Label(self.save_bar_risk, text="保存目录：").pack(side="left")
        self.entry_save_risk = ttk.Entry(self.save_bar_risk, textvariable=self.risk_save_dir, width=32)
        self.entry_save_risk.pack(side="left", padx=5)
        ttk.Button(self.save_bar_risk, text="浏览...", command=self.browse_folder_risk).pack(side="left", padx=3)
        ttk.Button(self.save_bar_risk, text="📄 生成 Word 文档", command=self.generate_risk_alert, style="Accent.TButton").pack(side="right", padx=5)

        self.save_bar_live = ttk.Frame(root)
        ttk.Label(self.save_bar_live, text="保存目录：").pack(side="left")
        self.entry_save_live = ttk.Entry(self.save_bar_live, textvariable=self.live_save_dir, width=32)
        self.entry_save_live.pack(side="left", padx=5)
        ttk.Button(self.save_bar_live, text="浏览...", command=self.browse_folder_live).pack(side="left", padx=3)
        ttk.Button(self.save_bar_live, text="✨ 生成 TXT 文件", command=self.generate_live_report, style="Accent.TButton").pack(side="right", padx=5)

        self.save_bar_lzy = ttk.Frame(root)
        ttk.Label(self.save_bar_lzy, text="保存目录：").pack(side="left")
        self.entry_save_lzy = ttk.Entry(self.save_bar_lzy, textvariable=self.lzy_save_dir, width=32)
        self.entry_save_lzy.pack(side="left", padx=5)
        ttk.Button(self.save_bar_lzy, text="浏览...", command=self._browse_lzy_dir).pack(side="left", padx=3)
        ttk.Button(self.save_bar_lzy, text="📊 写入 Excel", command=self._lzy_write_excel,
                   style="Accent.TButton").pack(side="right", padx=5)

        # ---------- 标签页（Notebook） ----------
        self.notebook = ttk.Notebook(main_panel)
        self.notebook.pack(side="left", fill="both", expand=True)

        # 首页标签页（纯展示文档）
        tab_home = ttk.Frame(self.notebook)
        self.notebook.add(tab_home, text="📖 首页")
        self.build_homepage(tab_home)

        # 叫应名单标签页（表格展示）
        tab_call = ttk.Frame(self.notebook)
        self.notebook.add(tab_call, text="📞 叫应名单")
        self.build_call_list(tab_call)

        # 天气提醒标签页（第 3 项）
        tab_weather = ttk.Frame(self.notebook)
        self.notebook.add(tab_weather, text="🌤天气提醒")
        self.weather_canvas, self.weather_scroll_frame = self._make_scrollable(tab_weather)
        self.build_weather_alert_form()

        # 气象灾害风险提示单标签页（第 4 项）
        tab_risk = ttk.Frame(self.notebook)
        self.notebook.add(tab_risk, text="⚠气象灾害风险提示单")
        self.risk_canvas, self.risk_scroll_frame = self._make_scrollable(tab_risk)
        self.build_risk_alert_form()

        # 预通报标签页（第 5 项）
        tab_pre = ttk.Frame(self.notebook)
        self.notebook.add(tab_pre, text="📋 预通报")
        self.pre_canvas, self.pre_scroll_frame = self._make_scrollable(tab_pre)
        self.build_pre_report_form()

        # 实况通报标签页（第 6 项）
        tab_live = ttk.Frame(self.notebook)
        self.notebook.add(tab_live, text="📊 实况通报")
        self.live_canvas, self.live_scroll_frame = self._make_scrollable(tab_live)
        self.build_live_report_form()

        # 两直一白标签页（第 7 项）
        tab_lzy = ttk.Frame(self.notebook)
        self.notebook.add(tab_lzy, text="📋 两直一白")
        self.lzy_canvas, self.lzy_scroll_frame = self._make_scrollable(tab_lzy)
        self.build_liangzhiyibai()

        # 服务记录标签页（可编辑表格，第 8 项）
        tab_svc = ttk.Frame(self.notebook)
        self.notebook.add(tab_svc, text="📝 服务记录")
        self.build_service_record(tab_svc)

        # 标签页切换绑定（放在所有标签页和保存栏创建之后）
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # ---------- 全局鼠标滚轮处理 ----------
        self._setup_global_wheel()

    def _setup_global_wheel(self):
        """统一全局鼠标滚轮：非编辑状态时页面自由滚动"""
        self._wheel_canvases = {}  # tab_index -> canvas
        # 各标签页对应的 canvas（由 build 方法填充）
        # 在首次 tab 切换时延迟构建映射
        self._wheel_map_built = False

        def _build_wheel_map():
            """构建 tab_index -> canvas 映射"""
            self._wheel_canvases.clear()
            for i in range(self.notebook.index("end")):
                tab_text = self.notebook.tab(i, "text")
                if "首页" in tab_text and hasattr(self, 'home_canvas'):
                    self._wheel_canvases[i] = self.home_canvas
                elif "叫应名单" in tab_text and hasattr(self, 'call_canvas'):
                    self._wheel_canvases[i] = self.call_canvas
                elif "天气提醒" in tab_text and hasattr(self, 'weather_canvas'):
                    self._wheel_canvases[i] = self.weather_canvas
                elif "气象灾害" in tab_text and hasattr(self, 'risk_canvas'):
                    self._wheel_canvases[i] = self.risk_canvas
                elif "预通报" in tab_text and hasattr(self, 'pre_canvas'):
                    self._wheel_canvases[i] = self.pre_canvas
                elif "实况通报" in tab_text and hasattr(self, 'live_canvas'):
                    self._wheel_canvases[i] = self.live_canvas
                elif "两直一白" in tab_text and hasattr(self, 'lzy_canvas'):
                    self._wheel_canvases[i] = self.lzy_canvas
            self._wheel_map_built = True

        def _global_wheel(event):
            if not self._wheel_map_built:
                _build_wheel_map()
            tab_idx = self.notebook.index("current")
            canvas = self._wheel_canvases.get(tab_idx)
            if canvas is None:
                return
            if not canvas.winfo_exists():
                return
            # 检查鼠标是否在该 canvas 区域内
            cx0 = canvas.winfo_rootx()
            cy0 = canvas.winfo_rooty()
            cx1 = cx0 + canvas.winfo_width()
            cy1 = cy0 + canvas.winfo_height()
            if not (cx0 <= event.x_root <= cx1 and cy0 <= event.y_root <= cy1):
                return
            # 非编辑状态下页面自由滚动（文本控件已自适应高度，无需拦截）
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.root.bind_all("<MouseWheel>", _global_wheel)

    # ========== 配置持久化 ==========
    def load_config(self):
        """读取保存的目录配置"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                legacy = data.get("save_dir", "")
                return {
                    "pre_save_dir": data.get("pre_save_dir", legacy),
                    "live_save_dir": data.get("live_save_dir", legacy),
                    "svc_save_dir": data.get("svc_save_dir", legacy),
                    "weather_save_dir": data.get("weather_save_dir", legacy),
                    "risk_save_dir": data.get("risk_save_dir", legacy),
                    "risk_issue_number": data.get("risk_issue_number", 1),
                    "lzy_save_dir": data.get("lzy_save_dir", ""),
                }
            except:
                pass
        return {"pre_save_dir": "", "live_save_dir": "", "svc_save_dir": "", "weather_save_dir": "", "risk_save_dir": "", "risk_issue_number": 1, "lzy_save_dir": ""}

    def save_config(self):
        """保存目录配置"""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump({
                    "pre_save_dir": self.pre_save_dir.get(),
                    "live_save_dir": self.live_save_dir.get(),
                    "svc_save_dir": self.svc_save_dir.get(),
                    "weather_save_dir": self.weather_save_dir.get(),
                    "risk_save_dir": self.risk_save_dir.get(),
                    "risk_issue_number": self.risk_issue_number,
                    "lzy_save_dir": self.lzy_save_dir.get(),
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("配置保存失败", f"无法保存目录配置:\n{e}")

    # ========== 站点持久化 ==========
    def load_stations(self):
        if os.path.exists(self.stations_file):
            try:
                with open(self.stations_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
            except:
                pass
        self.save_stations(self.default_stations)
        return self.default_stations.copy()

    def save_stations(self, stations_list=None):
        if stations_list is None:
            stations_list = self.stations
        try:
            with open(self.stations_file, "w", encoding="utf-8") as f:
                json.dump(stations_list, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("保存站点失败", f"无法写入站点文件:\n{e}")

    # ========== 站点管理面板 ==========
    def build_station_panel(self, parent):
        ttk.Label(parent, text="📌 站点管理", font=("微软雅黑", 14, "bold")).pack(anchor="w", pady=(0, 10))
        self.station_listbox = tk.Listbox(parent, height=12, width=22, font=("微软雅黑", 12),
                                          selectmode=tk.SINGLE, exportselection=False)
        self.station_listbox.pack(fill="both", expand=True, pady=(0, 10))
        self.refresh_station_listbox()

        add_frame = ttk.Frame(parent)
        add_frame.pack(fill="x", pady=(0, 5))
        self.new_station_entry = ttk.Entry(add_frame, font=("微软雅黑", 12))
        self.new_station_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(add_frame, text="添加", command=self.add_station).pack(side="right")
        ttk.Button(parent, text="删除选中站点", command=self.delete_station).pack(fill="x", pady=(0, 10))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=5)
        ttk.Label(parent, text="💡 输入关键字即可筛选站点\n（自动保存）", font=("微软雅黑", 10), foreground="gray").pack()

    def refresh_station_listbox(self):
        self.station_listbox.delete(0, tk.END)
        for s in self.stations:
            self.station_listbox.insert(tk.END, s)

    def add_station(self):
        name = self.new_station_entry.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入站点名称")
            return
        if name in self.stations:
            messagebox.showwarning("提示", "站点已存在")
            return
        self.stations.append(name)
        self.refresh_station_listbox()
        self.new_station_entry.delete(0, tk.END)
        self.update_all_station_combos()
        self.save_stations()

    def delete_station(self):
        sel = self.station_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的站点")
            return
        idx = sel[0]
        del self.stations[idx]
        self.refresh_station_listbox()
        self.update_all_station_combos()
        self.save_stations()

    def update_all_station_combos(self):
        for cb in self.station_combos:
            cb["values"] = self.stations

    # ========== 智能搜索（自定义弹窗版 — 不抢焦点，弹出后仍可编辑） ==========
    def setup_autocomplete(self, combobox, all_values=None):
        """用自定义 Toplevel+Listbox 替代原生下拉框。
           all_values: 完整选项列表，默认 self.stations。"""

        if all_values is None:
            all_values = self.stations
        combobox._all_values = all_values     # 记在控件上，过滤用

        # ── 创建自定义弹出窗 ──
        popup = tk.Toplevel(combobox)
        popup.withdraw()
        popup.overrideredirect(True)        # 无标题栏
        popup.attributes('-topmost', True)  # 置顶
        popup.configure(bg="#3498db")       # 外框颜色 = 蓝色边框

        inner = tk.Frame(popup, bg="white")
        inner.pack(padx=1, pady=1)          # 1px 留白即蓝色边框

        listbox = tk.Listbox(inner, font=("微软雅黑", 12),
                             selectmode=tk.SINGLE, exportselection=False,
                             bg="white", fg="#2c3e50",
                             selectbackground="#3498db",
                             selectforeground="white",
                             relief="flat", borderwidth=0,
                             highlightthickness=0)
        listbox.pack(fill="both", expand=True)

        # ── 定时器 ──
        combobox._autocomplete_timer = None

        # ── 弹窗显示/隐藏 ──
        def _show_popup():
            values = combobox['values']
            if not values:
                _hide_popup()
                return
            listbox.delete(0, tk.END)
            for v in values:
                listbox.insert(tk.END, v)
            n = min(len(values), 8)
            listbox.config(height=n)
            popup.update_idletasks()
            x = combobox.winfo_rootx()
            y = combobox.winfo_rooty() + combobox.winfo_height()
            w = combobox.winfo_width()
            popup.geometry(f"{w}x{listbox.winfo_reqheight() + 2}+{x}+{y}")
            popup.deiconify()

        def _hide_popup():
            popup.withdraw()

        # ── 点击弹窗以外任意位置 → 自动消失 ──
        def _on_global_click(event):
            if not popup.winfo_viewable():
                return
            # 沿控件树向上查找，判断点击是否发生在此 combobox 上
            w = event.widget
            while w is not None and w != self.root:
                if w == combobox:
                    return   # 点击 combobox 本身 → 不隐藏
                try:
                    w = w.master
                except Exception:
                    break
            _hide_popup()

        self.root.bind('<Button-1>', _on_global_click, add='+')

        def _check_and_hide():
            """失焦后延迟检查：鼠标在弹窗上则不隐藏"""
            if not popup.winfo_viewable():
                return
            px, py = popup.winfo_pointerxy()
            x0, y0 = popup.winfo_rootx(), popup.winfo_rooty()
            x1, y1 = x0 + popup.winfo_width(), y0 + popup.winfo_height()
            if not (x0 <= px <= x1 and y0 <= py <= y1):
                _hide_popup()

        # ── 列表项点击 / 回车 → 填入并关闭 ──
        def _on_list_select(event):
            if listbox.curselection():
                text = listbox.get(listbox.curselection()[0])
                combobox.set(text)
                _hide_popup()
                combobox.focus_set()
                combobox.icursor(len(text))

        listbox.bind('<ButtonRelease-1>', _on_list_select)
        listbox.bind('<Return>', _on_list_select)

        # ── 输入框失焦 → 延迟隐藏弹窗 ──
        combobox.bind('<FocusOut>', lambda e: combobox.after(150, _check_and_hide))

        # ── 鼠标点击处理 ──
        def on_click(event):
            w = event.widget.winfo_width()
            if event.x > w - 20:
                # 点击箭头区域 → 显示完整列表（阻止原生下拉框）
                combobox['values'] = list(combobox._all_values)
                _show_popup()
                return 'break'
            # 点击输入区域 → 正常编辑

        combobox.bind('<Button-1>', on_click)

        # ── 键盘处理 ──
        def on_keyrelease(event):
            if event.keysym in ("Up", "Down", "Left", "Right",
                                "Return", "Tab", "Escape",
                                "Control_L", "Control_R",
                                "Shift_L", "Shift_R",
                                "Alt_L", "Alt_R",
                                "Home", "End", "Prior", "Next",
                                "Caps_Lock", "Num_Lock"):
                if event.keysym == "Escape":
                    _hide_popup()
                elif event.keysym == "Down" and popup.winfo_viewable():
                    listbox.focus_set()
                    listbox.selection_set(0)
                elif event.keysym == "Return" and popup.winfo_viewable():
                    _on_list_select(None)
                return

            # 立刻过滤选项
            value = event.widget.get()
            if value == '':
                combobox['values'] = list(combobox._all_values)
            else:
                data = [s for s in combobox._all_values if value.lower() in s.lower()]
                combobox['values'] = data

            # 防抖 250ms 后弹出自定义弹窗（不打断 IME 组字）
            if combobox._autocomplete_timer is not None:
                combobox.after_cancel(combobox._autocomplete_timer)
            combobox._autocomplete_timer = combobox.after(250, _show_popup)

        combobox.bind('<KeyRelease>', on_keyrelease)

    # ========== 滚动容器 ==========
    def _make_scrollable(self, parent):
        """创建带滚动条的容器，返回 (canvas, inner_frame) 供填充"""
        canvas = tk.Canvas(parent, highlightthickness=0, bg="#f0f4f8")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, padding=10)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 滚轮由全局处理器 _global_wheel_handler 统一管理
        return canvas, inner

    # ========== 首页（展示文档） ==========
    def _get_docx_path(self):
        """获取 docx 文件路径（兼容打包后）"""
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = self.app_dir
        for f in os.listdir(base):
            if f.endswith('.docx'):
                return os.path.join(base, f)
        return None

    def _parse_docx(self):
        """解析 docx 返回段落列表：[{'text': str, 'font': str, 'size': int, 'bold': bool, 'align': str}]"""
        docx_path = self._get_docx_path()
        if not docx_path:
            return [{'text': '未找到文档文件，请将 .docx 文件放在程序同目录下。', 'font': '微软雅黑', 'size': 28, 'bold': False, 'align': 'left'}]

        NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        with zipfile.ZipFile(docx_path, 'r') as z:
            doc_xml = z.read('word/document.xml')

        root = ET.fromstring(doc_xml)
        body = root.find(f'.//{NS}body')
        if body is None:
            return []

        paragraphs = []
        for p in body.findall(f'.//{NS}p'):
            # 段落属性
            pPr = p.find(f'{NS}pPr')
            align = 'left'
            if pPr is not None:
                jc = pPr.find(f'{NS}jc')
                if jc is not None:
                    align = jc.get(f'{NS}val', 'left')

            # 收集 run
            runs = []
            for r in p.findall(f'.//{NS}r'):
                rPr = r.find(f'{NS}rPr')
                bold = False; font = '仿宋'; size = 28  # 默认 14pt
                if rPr is not None:
                    b = rPr.find(f'{NS}b')
                    bold = b is not None
                    rf = rPr.find(f'{NS}rFonts')
                    if rf is not None:
                        font = rf.get(f'{NS}eastAsia', '') or rf.get(f'{NS}ascii', '') or '仿宋'
                    sz = rPr.find(f'{NS}sz')
                    if sz is not None:
                        size = int(sz.get(f'{NS}val', '28'))
                t = r.find(f'{NS}t')
                text = t.text if t is not None and t.text else ''
                # 保留空白 runs（如空格）
                if t is not None and t.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve':
                    text = t.text or ''
                runs.append({'text': text, 'bold': bold, 'font': font, 'size': size})

            full_text = ''.join(r['text'] for r in runs)
            if full_text.strip() or full_text:
                # 取第一个非空 run 的格式作为段落主格式
                main_run = next((r for r in runs if r['text'].strip()), runs[0]) if runs else {'font': '仿宋', 'size': 28, 'bold': False}
                paragraphs.append({
                    'text': full_text,
                    'font': main_run['font'],
                    'size': main_run['size'],
                    'bold': main_run['bold'],
                    'align': align,
                    'runs': runs if any(r['bold'] != main_run['bold'] or r['font'] != main_run['font'] for r in runs) else None,
                })

        return paragraphs

    # ========== 首页：可编辑 block 系统 ==========
    def build_homepage(self, parent):
        """构建首页：可编辑 block 列表"""
        self.homepage_blocks_file = os.path.join(self.app_dir, "homepage_blocks.json")
        self.homepage_blocks = []
        self._homepage_load()

        # 滚动容器
        self.home_canvas = tk.Canvas(parent, highlightthickness=0, bg="#f0f4f8")
        home_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.home_canvas.yview)
        self.home_inner = ttk.Frame(self.home_canvas)
        self.home_inner.bind("<Configure>", lambda e: self.home_canvas.configure(
            scrollregion=self.home_canvas.bbox("all")))
        cw = self.home_canvas.create_window((0, 0), window=self.home_inner, anchor="nw")
        def _resize(event):
            self.home_canvas.itemconfig(cw, width=event.width)
        self.home_canvas.bind("<Configure>", _resize)
        self.home_canvas.configure(yscrollcommand=home_scrollbar.set)
        self.home_canvas.pack(side="left", fill="both", expand=True)
        home_scrollbar.pack(side="right", fill="y")

        # 滚轮由全局处理器 _global_wheel_handler 统一管理
        # 渲染所有 blocks + 底部按钮
        self._homepage_render_all()

    def _homepage_load(self):
        """加载首页数据（优先 JSON，否则从 docx 解析为单个文本 block）"""
        if os.path.exists(self.homepage_blocks_file):
            try:
                with open(self.homepage_blocks_file, "r", encoding="utf-8") as f:
                    self.homepage_blocks = json.load(f)
                return
            except Exception:
                pass
        # 从 docx 解析为单个文本 block
        paragraphs = self._parse_docx()
        header_map = {
            "影响提示": "一、影响提示", "警戒提醒": "二、警戒提醒",
            "精细预警": "三、精细预警", "分级叫应": "四、分级叫应",
            "实况通报": "五、实况通报",
        }
        lines = []
        for p in paragraphs:
            s = p['text'].strip()
            lines.append(header_map.get(s, p['text']))
        self.homepage_blocks = [{"type": "text", "content": "\n".join(lines)}]
        self._homepage_save()

    def _homepage_save(self):
        try:
            with open(self.homepage_blocks_file, "w", encoding="utf-8") as f:
                json.dump(self.homepage_blocks, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ==================== 全量渲染 ====================
    def _homepage_render_all(self):
        """清空容器，重新渲染所有 blocks + 底部添加按钮"""
        for w in self.home_inner.winfo_children():
            w.destroy()

        for bi, block in enumerate(self.homepage_blocks):
            self._homepage_render_block(bi, block, edit_mode=False)

        # 底部添加按钮（始终在最下方）
        add_bar = ttk.Frame(self.home_inner, padding=10)
        add_bar.pack(fill="x")
        ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
        btn_frame = ttk.Frame(add_bar)
        btn_frame.pack()
        ttk.Button(btn_frame, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)

    # ==================== 渲染单个 block ====================
    def _homepage_render_block(self, bi, block, edit_mode=False):
        """渲染单个 block（view 或 edit 模式）"""
        if block["type"] == "text":
            self._homepage_render_text_block(bi, block, edit_mode)
        elif block["type"] == "table":
            self._homepage_render_table_block(bi, block, edit_mode)

    # ==================== 文本 block ====================
    def _homepage_render_text_block(self, bi, block, edit_mode=False):
        frm = ttk.LabelFrame(self.home_inner, padding=8)
        frm.pack(fill="x", padx=15, pady=(8, 4))

        if edit_mode:
            # ---- 格式工具栏 ----
            tb = ttk.Frame(frm)
            tb.pack(fill="x", pady=(0, 5))

            fonts = ["宋体", "黑体", "仿宋", "微软雅黑", "楷体", "Arial"]
            font_cb = ttk.Combobox(tb, values=fonts, width=10, state="readonly")
            font_cb.set("微软雅黑")
            font_cb.pack(side="left", padx=2)

            sizes = [str(s) for s in [10, 11, 12, 14, 16, 18, 20, 24, 28]]
            size_cb = ttk.Combobox(tb, values=sizes, width=4, state="readonly")
            size_cb.set("12")
            size_cb.pack(side="left", padx=2)

            # 辅助：获取操作范围（有选区用选区，否则全文）
            def _get_range(tw):
                try:
                    r = tw.tag_ranges(tk.SEL)
                    if r:
                        return (r[0], r[1])
                except Exception:
                    pass
                return ("1.0", "end-1c")

            # 格式切换函数
            def make_toggle(tag, text_w):
                def _toggle():
                    try:
                        s, e = _get_range(text_w)
                        current = text_w.tag_names(s)
                        if tag in current:
                            text_w.tag_remove(tag, s, e)
                        else:
                            text_w.tag_add(tag, s, e)
                    except Exception:
                        pass
                return _toggle

            def make_apply_font(cb, text_w):
                def _apply():
                    try:
                        s, e = _get_range(text_w)
                        fname = cb.get()
                        fsize = int(size_cb.get())
                        # 移除该范围内已有的字体 tag
                        for t in text_w.tag_names(s):
                            if t.startswith("font_"):
                                text_w.tag_remove(t, s, e)
                        tname = f"font_{fname}_{fsize}"
                        text_w.tag_configure(tname, font=(fname, fsize))
                        text_w.tag_add(tname, s, e)
                    except Exception:
                        pass
                return _apply

            def make_apply_color(text_w):
                def _apply():
                    try:
                        from tkinter import colorchooser
                        c = colorchooser.askcolor(title="选择文字颜色")
                        if c and c[1]:
                            s, e = _get_range(text_w)
                            # 移除该范围内已有的颜色 tag（名称含 color_ 的）
                            for t in text_w.tag_names(s):
                                if t.startswith("color_"):
                                    text_w.tag_remove(t, s, e)
                            # 新颜色 tag（去掉 # 避免潜在问题）
                            color_hex = c[1].replace("#", "")
                            tname = f"color_{color_hex}"
                            text_w.tag_configure(tname, foreground=c[1])
                            text_w.tag_add(tname, s, e)
                    except Exception:
                        pass
                return _apply

            def make_align(align, text_w):
                def _apply():
                    try:
                        s, e = _get_range(text_w)
                        tname = f"align_{align}"
                        text_w.tag_configure(tname, justify=align)
                        text_w.tag_add(tname, s, e)
                    except Exception:
                        pass
                return _apply

            # 编辑区（先创建，供按钮引用）
            text_w = tk.Text(frm, wrap=tk.WORD, font=("微软雅黑", 12),
                              relief="solid", borderwidth=1, padx=10, pady=10,
                              height=4, undo=True)
            # 还原已保存的格式
            segments = block.get("segments", None)
            tag_configs = block.get("tag_configs", {})
            if segments:
                for tname, cfg in tag_configs.items():
                    try:
                        tk_cfg = {}
                        if "font" in cfg:
                            font_str = cfg["font"]
                            tk_cfg["font"] = eval(font_str) if font_str.startswith("(") else font_str
                        if "foreground" in cfg:
                            tk_cfg["foreground"] = cfg["foreground"]
                        if "justify" in cfg:
                            tk_cfg["justify"] = cfg["justify"]
                        if "underline" in cfg:
                            tk_cfg["underline"] = True
                        text_w.tag_configure(tname, **tk_cfg)
                    except Exception:
                        pass
                for seg in segments:
                    txt = seg["text"]
                    tags = seg.get("tags", [])
                    if tags:
                        text_w.insert(tk.END, txt, tuple(tags))
                    else:
                        text_w.insert(tk.END, txt)
            else:
                text_w.insert("1.0", block["content"])

            # 配置基础 tag
            text_w.tag_configure("bold", font=("微软雅黑", 12, "bold"))
            text_w.tag_configure("italic", font=("微软雅黑", 12, "italic"))
            text_w.tag_configure("underline", underline=True)

            # 格式按钮（放在编辑区上方）
            ttk.Button(tb, text="B", width=2, command=make_toggle("bold", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="I", width=2, command=make_toggle("italic", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="U", width=2, command=make_toggle("underline", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="A", width=2, command=make_apply_color(text_w)).pack(side="left", padx=3)
            ttk.Button(tb, text="≡", width=2, command=make_align("left", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="⊜", width=2, command=make_align("center", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="⊐", width=2, command=make_align("right", text_w)).pack(side="left", padx=1)
            ttk.Button(tb, text="应用字体", width=7, command=make_apply_font(font_cb, text_w)).pack(side="left", padx=5)

            # 自适应高度：根据内容计算所需行数
            def _auto_height(tw=None):
                tw = tw or text_w
                try:
                    tw.update_idletasks()
                    result = tw.tk.call(tw._w, 'count', '-displaylines', '1.0', 'end-1c')
                    actual_lines = int(result) if result else int(tw.index("end-1c").split(".")[0])
                except Exception:
                    actual_lines = int(tw.index("end-1c").split(".")[0])
                tw.configure(height=max(actual_lines + 2, 6))
            text_w.pack(side="left", fill="both", expand=True)
            # 初始化自适应高度
            _auto_height(text_w)
            # 输入时动态调整高度（防抖 100ms）
            _height_timer = [None]
            def _on_text_change(event=None):
                if _height_timer[0] is not None:
                    text_w.after_cancel(_height_timer[0])
                _height_timer[0] = text_w.after(100, lambda: _auto_height(text_w))
            text_w.bind("<KeyRelease>", _on_text_change)
            # 窗口大小变化时也调整高度
            text_w.bind("<Configure>", lambda e: _auto_height(text_w) if e.widget == text_w and e.width > 10 else None)

            # 确认 / 取消按钮
            btn_row = ttk.Frame(frm)
            btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
            ttk.Button(btn_row, text="✅ 确认",
                       command=lambda w=text_w, b=bi: self._homepage_text_confirm(w, b)).pack(side="right", padx=5)
            ttk.Button(btn_row, text="❌ 取消",
                       command=lambda b=bi, orig=dict(block): self._homepage_text_cancel(b, orig)).pack(side="right", padx=5)
        else:
            # ---- 只读显示（还原格式） ----
            text_w = tk.Text(frm, wrap=tk.WORD, font=("微软雅黑", 12),
                              bg="#f0f4f8", fg="#2c3e50",
                              relief="flat", borderwidth=0,
                              height=1,  # 初始最小高度，后续自动调整
                              state="disabled")
            text_w.pack(fill="both", expand=True)
            text_w.configure(state="normal")

            segments = block.get("segments", None)
            tag_configs = block.get("tag_configs", {})
            if segments:
                # 预注册所有 tag 配置
                for tname, cfg in tag_configs.items():
                    try:
                        tk_cfg = {}
                        if "font" in cfg:
                            font_str = cfg["font"]
                            tk_cfg["font"] = eval(font_str) if font_str.startswith("(") else font_str
                        if "foreground" in cfg:
                            tk_cfg["foreground"] = cfg["foreground"]
                        if "justify" in cfg:
                            tk_cfg["justify"] = cfg["justify"]
                        if "underline" in cfg:
                            tk_cfg["underline"] = True
                        text_w.tag_configure(tname, **tk_cfg)
                    except Exception:
                        pass
                for seg in segments:
                    txt = seg["text"]
                    tags = seg.get("tags", [])
                    if tags:
                        text_w.insert(tk.END, txt, tuple(tags))
                    else:
                        text_w.insert(tk.END, txt)
            else:
                text_w.insert("1.0", block["content"])
            text_w.configure(state="disabled")
            # 自动调整高度：用 Tcl count -displaylines 计算含自动换行的真实行数
            _view_height_busy = [False]
            def _view_auto_height(tw=text_w):
                if _view_height_busy[0]:
                    return
                _view_height_busy[0] = True
                try:
                    tw.update_idletasks()
                    try:
                        result = tw.tk.call(tw._w, 'count', '-displaylines', '1.0', 'end-1c')
                        actual_lines = int(result) if result else int(tw.index("end-1c").split(".")[0])
                    except Exception:
                        actual_lines = int(tw.index("end-1c").split(".")[0])
                    new_h = max(actual_lines + 1, 4)
                    if int(tw.cget("height")) != new_h:
                        tw.configure(height=new_h)
                finally:
                    _view_height_busy[0] = False
            _view_auto_height()
            text_w.bind("<Configure>", lambda e, tw=text_w: (None if e.widget != tw or e.width <= 10
                else tw.after(80, _view_auto_height)))

            # ---- 拖拽调整高度把手 ----
            grip = tk.Frame(frm, height=6, cursor="sb_v_double_arrow", bg="#c0c8d0")
            grip.pack(fill="x", pady=(2, 0))
            _drag_grip = [False, 0, 0]  # [dragging, start_y, start_height]
            def _grip_press(event):
                _drag_grip[0] = True
                _drag_grip[1] = event.y_root
                _drag_grip[2] = int(text_w.cget("height"))
            def _grip_move(event):
                if not _drag_grip[0]:
                    return
                dy = event.y_root - _drag_grip[1]
                new_h = max(3, _drag_grip[2] + dy // 16)
                if int(text_w.cget("height")) != new_h:
                    text_w.configure(height=new_h)
                    # 取消自动高度（手动设置优先）
                    text_w.unbind("<Configure>")
            def _grip_stop(event):
                _drag_grip[0] = False
                # 保存手动高度到 block
                block["_manual_height"] = int(text_w.cget("height"))
            grip.bind("<ButtonPress-1>", _grip_press)
            grip.bind("<B1-Motion>", _grip_move)
            grip.bind("<ButtonRelease-1>", _grip_stop)
            # 恢复手动高度
            if block.get("_manual_height"):
                try:
                    text_w.configure(height=max(int(block["_manual_height"]), 3))
                except Exception:
                    pass

            # 编辑 / 删除按钮
            btn_row = ttk.Frame(frm)
            btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
            ttk.Button(btn_row, text="✏️ 编辑",
                       command=lambda b=bi: self._homepage_edit_block(b)).pack(side="right", padx=5)
            ttk.Button(btn_row, text="🗑 删除",
                       command=lambda b=bi: self._homepage_delete_block(b)).pack(side="right", padx=5)

    def _homepage_text_confirm(self, text_w, bi):
        """确认文本编辑：保存文本 + 格式信息"""
        content = text_w.get("1.0", "end-1c")
        if not content.strip():
            del self.homepage_blocks[bi]
            self._homepage_save()
            self._homepage_render_all()
            return

        # 使用 dump 提取完整格式信息
        # dump 返回: [('text', 'H', []), ('tagon', 'bold', {}), ('text', 'e', ['bold']), ...]
        dumped = text_w.dump("1.0", "end-1c", tag=True, text=True)
        # 简化存储：逐段记录 [text, active_tags, tag_configs]
        segments = []
        current_text = ""
        current_tags = set()
        all_tag_configs = {}

        for item in dumped:
            if item[0] == "text":
                current_text += item[1]
            elif item[0] == "tagon":
                tname = item[1]
                if current_text:
                    segments.append({"text": current_text, "tags": list(current_tags)})
                    current_text = ""
                current_tags.add(tname)
                # 记录 tag 配置
                if tname not in all_tag_configs:
                    cfg = {}
                    try:
                        fi = text_w.tag_cget(tname, "font")
                        if fi:
                            cfg["font"] = str(fi)
                    except Exception:
                        pass
                    try:
                        fg = text_w.tag_cget(tname, "foreground")
                        if fg:
                            cfg["foreground"] = str(fg)
                    except Exception:
                        pass
                    try:
                        jf = text_w.tag_cget(tname, "justify")
                        if jf and jf != "left":
                            cfg["justify"] = str(jf)
                    except Exception:
                        pass
                    try:
                        ul = text_w.tag_cget(tname, "underline")
                        if ul and ul != "0":
                            cfg["underline"] = "1"
                    except Exception:
                        pass
                    if cfg:
                        all_tag_configs[tname] = cfg
            elif item[0] == "tagoff":
                tname = item[1]
                if current_text:
                    segments.append({"text": current_text, "tags": list(current_tags)})
                    current_text = ""
                current_tags.discard(tname)

        if current_text:
            segments.append({"text": current_text, "tags": list(current_tags) if current_tags else []})

        self.homepage_blocks[bi]["content"] = content
        self.homepage_blocks[bi]["segments"] = segments
        self.homepage_blocks[bi]["tag_configs"] = all_tag_configs
        self._homepage_save()
        self._homepage_render_all()

    def _homepage_text_cancel(self, bi, original):
        """取消文本编辑：恢复原始完整状态"""
        self.homepage_blocks[bi] = original
        self._homepage_save()
        self._homepage_render_all()

    # ==================== 表格 block ====================
    # ==================== 网格表格（Entry 网格） ====================
    class _GridTable:
        """基于 Entry 网格的表格，支持单元格独立着色"""
        def __init__(self, parent, headers, rows, cell_styles=None, edit_mode=False):
            self.parent = parent
            self.headers = list(headers)
            self.rows = [list(r) for r in rows]
            self.edit_mode = edit_mode
            self.cell_styles = cell_styles or {}
            self.sel_cells = set()
            self.anchor = None
            self.ncols = len(headers)
            self.nrows = len(rows)
            self._drag_moved = False

            self.frame = tk.Frame(parent, bg="#d0d0d0")
            self._build()

        def _build(self):
            for w in self.frame.winfo_children():
                w.destroy()
            ncols = max(self.ncols, 1)
            nrows = self.nrows
            col_w = max(80, min(180, 900 // ncols))

            # 表头
            for ci in range(ncols):
                h = self.headers[ci] if ci < len(self.headers) else f"列{ci + 1}"
                if self.edit_mode:
                    w = tk.Entry(self.frame, font=("微软雅黑", 11, "bold"),
                                 justify="center", relief="flat", bg="#d9e1e8")
                    w.insert(0, h)
                else:
                    w = tk.Label(self.frame, text=h, font=("微软雅黑", 11, "bold"),
                                 bg="#d9e1e8", fg="#2c3e50", anchor="center")
                w.grid(row=0, column=ci, sticky="nsew", padx=1, pady=1)
                self.frame.columnconfigure(ci, weight=1, minsize=col_w)

            # 数据行
            for ri in range(nrows):
                for ci in range(ncols):
                    val = self.rows[ri][ci] if ci < len(self.rows[ri]) else ""
                    style = self.cell_styles.get((ri, ci), {})
                    if self.edit_mode:
                        w = tk.Entry(self.frame, font=("微软雅黑", 11),
                                     relief="flat", justify="left",
                                     bg=style.get("bg", "white"),
                                     fg=style.get("fg", "black"))
                        w.insert(0, str(val))
                        w.bind("<ButtonPress-1>", lambda e, r=ri, c=ci: self._on_press(r, c, e))
                        w.bind("<B1-Motion>", lambda e, r=ri, c=ci: self._on_move(r, c))
                        w.bind("<Button-3>", lambda e, r=ri, c=ci: self._on_right_click(r, c, e))
                        w.bind("<Tab>", lambda e, r=ri, c=ci: self._on_tab(r, c))
                        w.bind("<Shift-Tab>", lambda e, r=ri, c=ci: self._on_shift_tab(r, c))
                    else:
                        w = tk.Label(self.frame, text=str(val), font=("微软雅黑", 11),
                                     bg=style.get("bg", "white"),
                                     fg=style.get("fg", "#2c3e50"),
                                     anchor="w", padx=4)
                    w.grid(row=ri + 1, column=ci, sticky="nsew", padx=1, pady=1)

        def _apply_selection(self):
            for ri in range(self.nrows):
                for ci in range(self.ncols):
                    style = self.cell_styles.get((ri, ci), {})
                    bg = "#cce5ff" if (ri, ci) in self.sel_cells else style.get("bg", "white")
                    fg = style.get("fg", "black")
                    for w in self.frame.grid_slaves(row=ri + 1, column=ci):
                        try: w.configure(bg=bg, fg=fg)
                        except Exception: pass

        def _on_press(self, ri, ci, event):
            self._drag_moved = False
            if event.state & 0x0004:
                if (ri, ci) in self.sel_cells:
                    self.sel_cells.discard((ri, ci))
                else:
                    self.sel_cells.add((ri, ci))
            elif event.state & 0x20000 and self.anchor:
                ar, ac = self.anchor
                lo_r, hi_r = min(ar, ri), max(ar, ri)
                lo_c, hi_c = min(ac, ci), max(ac, ci)
                self.sel_cells.clear()
                for r in range(lo_r, hi_r + 1):
                    for c in range(lo_c, hi_c + 1):
                        self.sel_cells.add((r, c))
            else:
                self.sel_cells.clear()
                self.sel_cells.add((ri, ci))
            self.anchor = (ri, ci)
            self._apply_selection()

        def _on_move(self, ri, ci):
            if self.anchor and (ri, ci) != self.anchor:
                self._drag_moved = True
                ar, ac = self.anchor
                lo_r, hi_r = min(ar, ri), max(ar, ri)
                lo_c, hi_c = min(ac, ci), max(ac, ci)
                self.sel_cells.clear()
                for r in range(lo_r, hi_r + 1):
                    for c in range(lo_c, hi_c + 1):
                        self.sel_cells.add((r, c))
                self._apply_selection()

        def _on_tab(self, ri, ci):
            nc = ci + 1; nr = ri
            if nc >= self.ncols: nc = 0; nr = ri + 1
            if nr >= self.nrows: nr = 0
            self.sel_cells.clear(); self.sel_cells.add((nr, nc))
            self.anchor = (nr, nc); self._apply_selection()
            for w in self.frame.grid_slaves(row=nr + 1, column=nc):
                w.focus_set(); return "break"
            return "break"

        def _on_shift_tab(self, ri, ci):
            nc = ci - 1; nr = ri
            if nc < 0: nc = self.ncols - 1; nr = ri - 1
            if nr < 0: nr = self.nrows - 1
            self.sel_cells.clear(); self.sel_cells.add((nr, nc))
            self.anchor = (nr, nc); self._apply_selection()
            for w in self.frame.grid_slaves(row=nr + 1, column=nc):
                w.focus_set(); return "break"
            return "break"

        def _on_right_click(self, ri, ci, event):
            if (ri, ci) not in self.sel_cells:
                self.sel_cells.clear(); self.sel_cells.add((ri, ci))
                self.anchor = (ri, ci); self._apply_selection()
            menu = tk.Menu(self.frame, tearoff=0)
            count = len(self.sel_cells)
            label = f"（{count}格）" if count > 1 else ""
            fmt_menu = tk.Menu(menu, tearoff=0)
            def _bg():
                from tkinter import colorchooser
                c = colorchooser.askcolor(title=f"选择{count}个格子的背景色")
                if c and c[1]:
                    for (r, c2) in list(self.sel_cells):
                        if (r, c2) not in self.cell_styles: self.cell_styles[(r, c2)] = {}
                        self.cell_styles[(r, c2)]["bg"] = c[1]
                    self._apply_selection()
            fmt_menu.add_command(label=f"填充背景色...{label}", command=_bg)
            def _fg():
                from tkinter import colorchooser
                c = colorchooser.askcolor(title=f"选择{count}个格子的文字色")
                if c and c[1]:
                    for (r, c2) in list(self.sel_cells):
                        if (r, c2) not in self.cell_styles: self.cell_styles[(r, c2)] = {}
                        self.cell_styles[(r, c2)]["fg"] = c[1]
                    self._apply_selection()
            fmt_menu.add_command(label=f"文字颜色...{label}", command=_fg)
            fmt_menu.add_separator()
            if len(self.sel_cells) >= 2:
                fmt_menu.add_command(label="🔗 合并选中格子", command=self._merge_cells)
            menu.add_cascade(label="单元格格式", menu=fmt_menu)
            menu.post(event.x_root, event.y_root)

        def _merge_cells(self):
            if len(self.sel_cells) < 2: return
            sc = sorted(self.sel_cells)
            fr, fc = sc[0]
            combined = self.rows[fr][fc] if fc < len(self.rows[fr]) else ""
            for r, c in sc[1:]:
                if c < len(self.rows[r]) and self.rows[r][c].strip():
                    combined += str(self.rows[r][c])
                while len(self.rows[r]) <= c: self.rows[r].append("")
                self.rows[r][c] = ""
            while len(self.rows[fr]) <= fc: self.rows[fr].append("")
            self.rows[fr][fc] = combined
            self._rebuild()

        def get_data(self):
            hdrs = []
            for ci in range(self.ncols):
                txt = self.headers[ci] if ci < len(self.headers) else f"列{ci + 1}"
                if self.edit_mode:
                    for w in self.frame.grid_slaves(row=0, column=ci):
                        if isinstance(w, tk.Entry): txt = w.get().strip() or txt
                hdrs.append(txt)
            data_rows = []
            for ri in range(self.nrows):
                row_vals = list(self.rows[ri])
                while len(row_vals) < self.ncols: row_vals.append("")
                data_rows.append(row_vals[:self.ncols])
            styles_export = {}
            for (r, c), s in self.cell_styles.items():
                styles_export[f"{r},{c}"] = s
            return hdrs, data_rows, styles_export

        def _rebuild(self):
            self.ncols = len(self.headers)
            self.nrows = len(self.rows)
            for w in self.frame.winfo_children(): w.destroy()
            self._build()
            self._apply_selection()

        def add_row_at_end(self):
            self.rows.append([""] * self.ncols); self._rebuild()

        def insert_row_above(self):
            if self.sel_cells:
                r = min(rc[0] for rc in self.sel_cells)
                self.rows.insert(r, [""] * self.ncols)
            else:
                self.rows.append([""] * self.ncols)
            self._rebuild()

        def delete_selected_rows(self):
            rows_to_del = set(r for (r, c) in self.sel_cells)
            if not rows_to_del: return
            for r in sorted(rows_to_del, reverse=True):
                if 0 <= r < len(self.rows): self.rows.pop(r)
            self.sel_cells.clear(); self._rebuild()

        def add_column_at_end(self):
            self.headers.append(f"列{self.ncols + 1}")
            for row in self.rows:
                while len(row) < len(self.headers): row.append("")
            self._rebuild()

        def delete_last_column(self):
            if self.ncols <= 1: return
            self.headers.pop()
            for row in self.rows:
                if len(row) > len(self.headers): row.pop()
            self._rebuild()

    # ---- _homepage_render_table_block（重写为网格表格） ----
    def _homepage_render_table_block(self, bi, block, edit_mode=False):
        frm = ttk.LabelFrame(self.home_inner, padding=8)
        frm.pack(fill="x", padx=15, pady=(8, 4))

        headers = list(block.get("headers", ["列1", "列2", "列3"]))
        rows = block.get("rows", [["", "", ""]])
        for r in rows:
            while len(r) < len(headers):
                r.append("")
        saved_styles = block.get("cell_styles", {})
        cell_styles = {}
        for key, style in saved_styles.items():
            parts = key.split(",")
            if len(parts) == 2:
                try: cell_styles[(int(parts[0]), int(parts[1]))] = style
                except Exception: pass

        gt = self._GridTable(frm, headers, rows, cell_styles, edit_mode)
        gt.frame.pack(fill="both", expand=True)

        if edit_mode:
            op_row = ttk.Frame(frm)
            op_row.pack(fill="x", pady=(5, 0))
            ttk.Label(op_row, text="行：", font=("微软雅黑", 10)).pack(side="left", padx=(0, 2))
            ttk.Button(op_row, text="➕ 末尾", command=gt.add_row_at_end).pack(side="left", padx=1)
            ttk.Button(op_row, text="📌 插入", command=gt.insert_row_above).pack(side="left", padx=1)
            ttk.Button(op_row, text="🗑 删除", command=gt.delete_selected_rows).pack(side="left", padx=1)
            ttk.Separator(op_row, orient="vertical").pack(side="left", fill="y", padx=5, pady=2)
            ttk.Label(op_row, text="列：", font=("微软雅黑", 10)).pack(side="left", padx=(0, 2))
            ttk.Button(op_row, text="➕ 末尾", command=gt.add_column_at_end).pack(side="left", padx=1)
            ttk.Button(op_row, text="🗑 末尾", command=gt.delete_last_column).pack(side="left", padx=1)

            btn_row = ttk.Frame(frm)
            btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
            def _confirm():
                hdrs, data_rows, styles = gt.get_data()
                block["headers"] = hdrs
                block["rows"] = data_rows
                block["cell_styles"] = styles
                self._homepage_save()
                self._homepage_render_all()
            ttk.Button(btn_row, text="✅ 确认", command=_confirm).pack(side="right", padx=5)
            ttk.Button(btn_row, text="❌ 取消",
                       command=lambda: self._homepage_render_all()).pack(side="right", padx=5)
        else:
            btn_row = ttk.Frame(frm)
            btn_row.pack(fill="x", pady=(5, 0))
            ttk.Button(btn_row, text="✏️ 编辑此表格",
                       command=lambda b=bi: self._homepage_edit_block(b)).pack(side="left", padx=2)
            ttk.Button(btn_row, text="🗑 删除此表格",
                       command=lambda b=bi: self._homepage_delete_block(b)).pack(side="right", padx=2)

    def _homepage_edit_block(self, bi):
        """进入编辑模式：只渲染该 block 为编辑态"""
        # 清除容器
        for w in self.home_inner.winfo_children():
            w.destroy()
        for i, blk in enumerate(self.homepage_blocks):
            self._homepage_render_block(i, blk, edit_mode=(i == bi))
        # 底部按钮
        add_bar = ttk.Frame(self.home_inner, padding=10)
        add_bar.pack(fill="x")
        ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
        ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
        ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)

    def _homepage_delete_block(self, bi):
        """删除 block（带确认弹窗）"""
        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg="#e74c3c")
        inner = tk.Frame(popup, bg="white", padx=20, pady=18)
        inner.pack(padx=2, pady=2)
        ttk.Label(inner, text="确定要删除吗？", font=("微软雅黑", 12, "bold")).pack(pady=(0, 12))
        bf = ttk.Frame(inner)
        bf.pack()
        def _do():
            popup.destroy()
            del self.homepage_blocks[bi]
            self._homepage_save()
            self._homepage_render_all()
            try:
                self.root.unbind('<Button-1>', _bid)
            except Exception:
                pass
        def _cancel():
            popup.destroy()
            try:
                self.root.unbind('<Button-1>', _bid)
            except Exception:
                pass
        ttk.Button(bf, text="确定", command=_do).pack(side="left", padx=8)
        ttk.Button(bf, text="取消", command=_cancel).pack(side="left", padx=8)
        # 收集弹窗控件
        pw = {popup, inner, bf}
        def _coll(w):
            pw.add(w)
            for c in w.winfo_children():
                _coll(c)
        _coll(inner)
        def _click(event):
            if not popup.winfo_exists():
                return
            w = event.widget
            while w:
                if w in pw:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            _cancel()
        _bid = self.root.bind('<Button-1>', _click, add='+')
        popup.update_idletasks()
        px = self.root.winfo_rootx() + self.root.winfo_width() // 2 - 80
        py = self.root.winfo_rooty() + self.root.winfo_height() // 2 - 50
        popup.geometry(f"+{px}+{py}")
        popup.focus_set()

    def _homepage_add_text(self):
        """添加文本 block 并进入编辑"""
        self.homepage_blocks.append({"type": "text", "content": ""})
        self._homepage_save()
        bi = len(self.homepage_blocks) - 1
        # 渲染：其他 view，新 block edit
        for w in self.home_inner.winfo_children():
            w.destroy()
        for i, blk in enumerate(self.homepage_blocks):
            self._homepage_render_block(i, blk, edit_mode=(i == bi))
        add_bar = ttk.Frame(self.home_inner, padding=10)
        add_bar.pack(fill="x")
        ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
        ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
        ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)

    def _homepage_add_table(self):
        """添加表格 block 并进入编辑"""
        self.homepage_blocks.append({
            "type": "table",
            "headers": ["列1", "列2", "列3"],
            "rows": [["", "", ""], ["", "", ""], ["", "", ""]],
        })
        self._homepage_save()
        bi = len(self.homepage_blocks) - 1
        for w in self.home_inner.winfo_children():
            w.destroy()
        for i, blk in enumerate(self.homepage_blocks):
            self._homepage_render_block(i, blk, edit_mode=(i == bi))
        add_bar = ttk.Frame(self.home_inner, padding=10)
        add_bar.pack(fill="x")
        ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
        ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
        ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)

    # ========== 叫应名单（解析 docx 含表格） ==========
    def _get_call_docx_path(self):
        """获取叫应名单 docx 路径"""
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = self.app_dir
        for f in os.listdir(base):
            if "叫应" in f and f.endswith(".docx"):
                return os.path.join(base, f)
        return None

    def _parse_call_docx(self):
        """解析叫应名单 docx，返回 sections 列表"""
        docx_path = self._get_call_docx_path()
        if not docx_path:
            return [{'type': 'para', 'text': '未找到叫应名单文档。', 'bold': False, 'font': '微软雅黑', 'size': '28'}]

        NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        with zipfile.ZipFile(docx_path, 'r') as z:
            doc_xml = z.read('word/document.xml')

        root = ET.fromstring(doc_xml)
        body = root.find(f'.//{NS}body')
        if body is None:
            return []

        def _get_run_info(r):
            rPr = r.find(f'{NS}rPr')
            bold = False; font = ''; size = ''
            if rPr is not None:
                b = rPr.find(f'{NS}b')
                bold = b is not None
                rf = rPr.find(f'{NS}rFonts')
                if rf is not None:
                    font = rf.get(f'{NS}eastAsia', '') or rf.get(f'{NS}ascii', '')
                sz = rPr.find(f'{NS}sz')
                if sz is not None:
                    size = sz.get(f'{NS}val', '')
            t = r.find(f'{NS}t')
            text = t.text if t is not None and t.text else ''
            return {'text': text, 'bold': bold, 'font': font, 'size': size}

        sections = []
        for elem in body:
            tag = elem.tag.replace(NS, '')
            if tag == 'p':
                runs = [_get_run_info(r) for r in elem.findall(f'.//{NS}r')]
                full_text = ''.join(r['text'] for r in runs)
                if full_text.strip():
                    main = runs[0] if runs else {'bold': False, 'font': '', 'size': ''}
                    sections.append({
                        'type': 'para',
                        'text': full_text.strip(),
                        'bold': main['bold'],
                        'font': main['font'],
                        'size': main['size'],
                    })
            elif tag == 'tbl':
                rows = []
                for tr in elem.findall(f'.//{NS}tr'):
                    cells = []
                    for tc in tr.findall(f'.//{NS}tc'):
                        cell_runs = [_get_run_info(r) for r in tc.findall(f'.//{NS}r')]
                        cell_text = ''.join(r['text'] for r in cell_runs)
                        is_bold = any(r['bold'] for r in cell_runs)
                        cells.append({'text': cell_text.strip(), 'bold': is_bold})
                    rows.append(cells)
                if rows:
                    sections.append({'type': 'table', 'rows': rows})

        return sections

    # ========== 叫应名单持久化 ==========
    def _call_list_load(self):
        """加载叫应名单数据（优先 JSON，否则从 docx 解析）"""
        call_json = os.path.join(self.app_dir, "call_list.json")
        if os.path.exists(call_json):
            try:
                with open(call_json, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # 验证结构
                if isinstance(data, list) and all(
                    isinstance(t, dict) and "headers" in t and "rows" in t
                    for t in data
                ):
                    return data, True  # (data, from_json)
            except Exception:
                pass
        # 从 docx 解析
        sections = self._parse_call_docx()
        tables = []
        for s in sections:
            if s['type'] == 'table' and s['rows']:
                rows = s['rows']
                headers = [c['text'] for c in rows[0]]
                # 数据行（去掉表头行）
                data_rows = [[c['text'] for c in row] for row in rows[1:]]
                tables.append({"headers": headers, "rows": data_rows})
        return tables, False

    def _call_list_save(self):
        """保存叫应名单到 JSON"""
        if not hasattr(self, '_call_list_tables_data'):
            return
        call_json = os.path.join(self.app_dir, "call_list.json")
        try:
            with open(call_json, "w", encoding="utf-8") as f:
                json.dump(self._call_list_tables_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _call_list_recalc_widths(self, tree, headers, rows):
        """重新计算列宽"""
        col_count = len(headers)
        max_lens = [len(str(h)) for h in headers]
        for row in rows:
            for ci, val in enumerate(row):
                if ci < col_count:
                    txt = str(val)
                    if len(txt) > 15 and txt.isdigit():
                        txt = txt[:11] + " / " + txt[11:]
                    max_lens[ci] = max(max_lens[ci], len(txt))
        for ci in range(col_count):
            width = min(max(max_lens[ci] * 18, 60), 300)
            tree.column(f"#{ci + 1}", width=width)

    def _call_list_refresh_tree(self, tree, headers, rows):
        """刷新 treeview 内容（保留结构，更新数据）"""
        for iid in tree.get_children():
            tree.delete(iid)
        for row in rows:
            # 确保行长度匹配
            vals = list(row)
            while len(vals) < len(headers):
                vals.append("")
            tree.insert("", tk.END, values=vals[:len(headers)])
        tree.configure(height=max(len(rows), 4))

    def _call_list_edit_cell(self, tree, ti, event):
        """编辑叫应名单单元格"""
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y)
        if not item_id:
            return
        ci = int(col_id.replace("#", "")) - 1
        if ci < 0:
            return
        cur_vals = list(tree.item(item_id, "values"))
        cur_text = cur_vals[ci] if ci < len(cur_vals) else ""
        x, y, w, h = tree.bbox(item_id, column=f"#{ci + 1}")
        if not w or not h:
            return
        e = ttk.Entry(tree, width=max(w // 10, 10), font=("微软雅黑", 11))
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, cur_text)
        e.lift()
        e.focus_set()

        destroyed = [False]
        def _save():
            if destroyed[0]:
                return
            destroyed[0] = True
            new_val = e.get().strip()
            cur_vals[ci] = new_val
            tree.item(item_id, values=cur_vals)
            # 同步到持久化数据
            headers = self._call_list_tables_data[ti]["headers"]
            row_idx = tree.index(item_id)
            if row_idx < len(self._call_list_tables_data[ti]["rows"]):
                self._call_list_tables_data[ti]["rows"][row_idx] = list(cur_vals[:len(headers)])
            if e.winfo_exists():
                e.destroy()
            self._call_list_save()
            self._call_list_recalc_widths(tree, headers, self._call_list_tables_data[ti]["rows"])
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        def _on_root_click(event):
            if destroyed[0]:
                return
            w = event.widget
            while w is not None and w != self.root:
                if w == e:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            tree.after(50, _save)

        _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')
        e.bind("<Return>", lambda ev: _save())
        e.bind("<FocusOut>", lambda ev: tree.after(100, _save))

    def _call_list_delete_rows(self, tree, ti):
        """删除叫应名单表格选中行"""
        sel = tree.selection()
        if not sel:
            return
        data_rows = self._call_list_tables_data[ti]["rows"]
        headers = self._call_list_tables_data[ti]["headers"]
        # 从后往前删（避免索引变化）
        indices = sorted([tree.index(iid) for iid in sel], reverse=True)
        for idx in indices:
            if 0 <= idx < len(data_rows):
                data_rows.pop(idx)
        self._call_list_refresh_tree(tree, headers, data_rows)
        self._call_list_save()

    def _call_list_insert_row(self, tree, ti):
        """在选中行下方插入空行"""
        sel = tree.selection()
        headers = self._call_list_tables_data[ti]["headers"]
        data_rows = self._call_list_tables_data[ti]["rows"]
        if sel:
            after_idx = tree.index(sel[-1])
        else:
            after_idx = len(data_rows) - 1
        new_row = [""] * len(headers)
        data_rows.insert(after_idx + 1, new_row)
        self._call_list_refresh_tree(tree, headers, data_rows)
        children = tree.get_children()
        if after_idx + 1 < len(children):
            tree.selection_set(children[after_idx + 1])
            tree.see(children[after_idx + 1])
        self._call_list_save()

    def _call_list_insert_above(self, tree, ti):
        """在选中行上方插入空行"""
        sel = tree.selection()
        headers = self._call_list_tables_data[ti]["headers"]
        data_rows = self._call_list_tables_data[ti]["rows"]
        if sel:
            insert_idx = tree.index(sel[0])
        else:
            insert_idx = 0
        new_row = [""] * len(headers)
        data_rows.insert(insert_idx, new_row)
        self._call_list_refresh_tree(tree, headers, data_rows)
        children = tree.get_children()
        if insert_idx < len(children):
            tree.selection_set(children[insert_idx])
            tree.see(children[insert_idx])
        self._call_list_save()

    def _call_list_move_row(self, tree, ti, direction):
        """上下移动选中行（支持多选）"""
        sel = tree.selection()
        if not sel:
            return
        data_rows = self._call_list_tables_data[ti]["rows"]
        headers = self._call_list_tables_data[ti]["headers"]
        # 获取选中行索引并排序
        indices = sorted([tree.index(iid) for iid in sel])
        if direction == "up" and indices[0] > 0:
            for idx in indices:
                data_rows[idx], data_rows[idx - 1] = data_rows[idx - 1], data_rows[idx]
            new_indices = [i - 1 for i in indices]
        elif direction == "down" and indices[-1] < len(data_rows) - 1:
            for idx in reversed(indices):
                data_rows[idx], data_rows[idx + 1] = data_rows[idx + 1], data_rows[idx]
            new_indices = [i + 1 for i in indices]
        else:
            return
        self._call_list_refresh_tree(tree, headers, data_rows)
        children = tree.get_children()
        new_sel = [children[i] for i in new_indices if i < len(children)]
        if new_sel:
            tree.selection_set(new_sel)
            tree.see(new_sel[0])
        self._call_list_save()

    def build_call_list(self, parent):
        """构建叫应名单：标题 + 可编辑表格"""
        # 滚动容器
        canvas = tk.Canvas(parent, highlightthickness=0, bg="#f0f4f8")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cw = canvas.create_window((0, 0), window=inner, anchor="nw")
        def _resize(event):
            canvas.itemconfig(cw, width=event.width)
        canvas.bind("<Configure>", _resize)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 存储 canvas 引用供全局滚轮使用
        self.call_canvas = canvas

        # 加载数据
        tables_data, from_json = self._call_list_load()
        self._call_list_tables_data = tables_data

        sections = self._parse_call_docx()
        paras = [s for s in sections if s['type'] == 'para']

        # 表格样式
        style = ttk.Style()
        style.configure("Call.Treeview", font=("微软雅黑", 11), rowheight=32)
        style.configure("Call.Treeview.Heading", font=("微软雅黑", 11, "bold"))
        style.configure("Bold.Treeview", font=("微软雅黑", 11, "bold"))
        title_font = ("微软雅黑", 14, "bold")

        # 渲染段落标题
        for s in paras:
            lbl = ttk.Label(inner, text=s['text'], font=title_font,
                            foreground="#2c3e50", background="#f0f4f8",
                            anchor="center")
            lbl.pack(fill="x", pady=(15, 10), padx=20)

        # 存储每个表格的 treeview 引用
        self._call_trees = []

        # 渲染可编辑表格
        for ti, tbl in enumerate(tables_data):
            headers = list(tbl["headers"])
            rows = tbl["rows"]
            col_count = len(headers)

            # 表格容器
            tbl_frame = ttk.Frame(inner, padding=5)
            tbl_frame.pack(fill="x", padx=20, pady=(0, 5))

            # 计算列宽
            max_lens = [len(str(h)) for h in headers]
            for row in rows:
                for ci, val in enumerate(row):
                    if ci < col_count:
                        txt = str(val)
                        if len(txt) > 15 and txt.isdigit():
                            txt = txt[:11] + " / " + txt[11:]
                        max_lens[ci] = max(max_lens[ci], len(txt))

            col_ids = [f"col{i}" for i in range(col_count)]
            tree = ttk.Treeview(tbl_frame, columns=col_ids, show="headings",
                                selectmode="extended",
                                style="Call.Treeview", height=max(len(rows), 4))
            for ci, hid in enumerate(headers):
                tree.heading(col_ids[ci], text=hid)
                width = min(max(max_lens[ci] * 18, 60), 300)
                tree.column(col_ids[ci], width=width, anchor="center", minwidth=50)

            # 填充数据
            for row in rows:
                vals = list(row)
                while len(vals) < col_count:
                    vals.append("")
                tree.insert("", tk.END, values=vals[:col_count])

            tree.pack(fill="x")

            # ---- 鼠标拖动多选 ----
            _call_drag = {"start": None, "dragging": False}
            def _on_drag_start(event):
                item_id = tree.identify_row(event.y)
                if item_id:
                    _call_drag["start"] = item_id
                    _call_drag["dragging"] = True
            def _on_drag_move(event):
                if not _call_drag.get("dragging"):
                    return
                item_id = tree.identify_row(event.y)
                if item_id and item_id != _call_drag.get("start"):
                    sel = set(tree.selection())
                    sel.add(item_id)
                    all_items = tree.get_children()
                    try:
                        lo = all_items.index(_call_drag["start"])
                        hi = all_items.index(item_id)
                    except ValueError:
                        return
                    if lo > hi:
                        lo, hi = hi, lo
                    for i in range(lo, hi + 1):
                        sel.add(all_items[i])
                    tree.selection_set(list(sel))
                    _call_drag["start"] = item_id
            def _on_drag_stop(event):
                _call_drag["dragging"] = False
                _call_drag["start"] = None
            tree.bind("<ButtonPress-1>", _on_drag_start, add="+")
            tree.bind("<B1-Motion>", _on_drag_move, add="+")
            tree.bind("<ButtonRelease-1>", _on_drag_stop, add="+")

            # ---- 操作栏 ----
            toolbar = ttk.Frame(tbl_frame)
            toolbar.pack(fill="x", pady=(5, 0))

            ttk.Button(toolbar, text="⬆ 上移",
                       command=lambda t=tree, i=ti: self._call_list_move_row(t, i, "up")
                       ).pack(side="left", padx=2)
            ttk.Button(toolbar, text="⬇ 下移",
                       command=lambda t=tree, i=ti: self._call_list_move_row(t, i, "down")
                       ).pack(side="left", padx=2)
            ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=5, pady=2)
            ttk.Button(toolbar, text="➕ 插入行",
                       command=lambda t=tree, i=ti: self._call_list_insert_row(t, i)
                       ).pack(side="left", padx=2)
            ttk.Button(toolbar, text="🗑 删除行",
                       command=lambda t=tree, i=ti: self._call_list_delete_rows(t, i)
                       ).pack(side="left", padx=2)

            # ---- 双击：表头编辑 / 单元格编辑 ----
            def _make_call_double_click(t, tbl_idx, col_ids_list, headers_list):
                def _on_double(event):
                    region = t.identify_region(event.x, event.y)
                    if region == "heading":
                        # 编辑表头
                        cid = t.identify_column(event.x)
                        if not cid:
                            return
                        ci = int(cid.replace("#", "")) - 1
                        if ci < 0 or ci >= len(col_ids_list):
                            return
                        x_offset = 0
                        for c in range(ci):
                            x_offset += t.column(f"#{c + 1}", "width")
                        w = t.column(f"#{ci + 1}", "width")
                        e = ttk.Entry(tbl_frame, width=max(w // 10, 8),
                                     font=("微软雅黑", 11, "bold"))
                        e.place(x=x_offset + 2, y=0, width=w - 4, height=28)
                        e.insert(0, headers_list[ci])
                        e.lift()
                        e.focus_set()
                        def _save():
                            headers_list[ci] = e.get().strip() or headers_list[ci]
                            t.heading(col_ids_list[ci], text=headers_list[ci])
                            self._call_list_tables_data[tbl_idx]["headers"] = list(headers_list)
                            self._call_list_save()
                            e.destroy()
                        e.bind("<Return>", lambda ev: _save())
                        e.bind("<FocusOut>", lambda ev: t.after(100, _save))
                    elif region == "cell":
                        # 编辑单元格
                        self._call_list_edit_cell(t, tbl_idx, event)
                return _on_double
            tree.bind("<Double-1>", _make_call_double_click(tree, ti, col_ids, headers))

            # ---- 右键菜单 ----
            def _make_right_click(t, tbl_idx):
                def _right_click(event):
                    menu = tk.Menu(self.root, tearoff=0)
                    sel = t.selection()
                    menu.add_command(label="⬆ 上移",
                        command=lambda: self._call_list_move_row(t, tbl_idx, "up"))
                    menu.add_command(label="⬇ 下移",
                        command=lambda: self._call_list_move_row(t, tbl_idx, "down"))
                    menu.add_separator()
                    menu.add_command(label="➕ 在上方插入",
                        command=lambda: self._call_list_insert_above(t, tbl_idx))
                    menu.add_command(label="➕ 在下方插入",
                        command=lambda: self._call_list_insert_row(t, tbl_idx))
                    if sel:
                        menu.add_separator()
                        menu.add_command(label="🗑 删除选中行",
                            command=lambda: self._call_list_delete_rows(t, tbl_idx))
                    menu.post(event.x_root, event.y_root)
                return _right_click
            tree.bind("<Button-3>", _make_right_click(tree, ti))

            # 存储 treeview 引用
            self._call_trees.append(tree)

        # 首次从 docx 解析则自动保存
        if not from_json and tables_data:
            self._call_list_save()

    # ========== 两直一白 ==========
    def build_liangzhiyibai(self):
        """构建两直一白标签页：数据表单 + 表格 + Excel 写入"""
        inner = self.lzy_scroll_frame
        self.lzy_data_file = os.path.join(self.app_dir, "liangzhiyibai_data.json")

        # ---- 列定义（与 Excel 表头一致） ----
        self.lzy_columns = [
            "序号", "时间", "灾害性天气", "类型",
            "叫应相关党政领导、部门负责人及基层责任人（人次）",
            "人工电话（人次）", "12379语音电话（人次）",
            "基层行政责任人和转移责任人（人次）",
            "家庭和个人（人次）", "应急广播", "闪信提醒（人次）", "汇总"
        ]
        # 表单字段（排除自增的序号和时间）
        self.lzy_fields = self.lzy_columns[2:]  # 从"灾害性天气"开始

        # ---- 加载持久化数据 ----
        self._lzy_load_data()

        # ===== 头部：日期 + 天气类型 =====
        f_head = ttk.LabelFrame(inner, text="数据录入", padding=10)
        f_head.pack(fill="x", pady=(5, 8))

        today = date.today()
        hdr = ttk.Frame(f_head)
        hdr.pack(fill="x")
        ttk.Label(hdr, text="日期：", font=("微软雅黑", 12)).pack(side="left")
        self.lzy_month = ttk.Combobox(hdr, values=[f"{m}月" for m in range(1, 13)],
                                       width=5, state="normal")
        self.lzy_month.set(f"{today.month}月")
        self.lzy_month.pack(side="left")
        self.lzy_day = ttk.Combobox(hdr, values=[f"{d}日" for d in range(1, 32)],
                                     width=5, state="normal")
        self.lzy_day.set(f"{today.day}日")
        self.lzy_day.pack(side="left")

        ttk.Label(hdr, text="  灾害性天气：", font=("微软雅黑", 12)).pack(side="left")
        self.lzy_disaster = ttk.Combobox(hdr, values=["强对流", "台风", "暴雨", "大风", "雷电", "冰雹"],
                                          width=8, state="normal")
        self.lzy_disaster.pack(side="left", padx=(0, 15))

        ttk.Label(hdr, text="类型：", font=("微软雅黑", 12)).pack(side="left")
        self.lzy_type = ttk.Combobox(hdr, values=["黄色预警", "橙色预警", "红色预警", "警报", "紧急警报"],
                                      width=10, state="normal")
        self.lzy_type.pack(side="left")

        # ===== 表单输入区（网格布局） =====
        f_form = ttk.LabelFrame(inner, text="填写字段", padding=10)
        f_form.pack(fill="x", pady=(0, 8))

        self.lzy_entries = {}
        # 前 9 个字段用 2 列网格，最后一个"汇总"独占一行
        fields_grid = self.lzy_fields[:-1]  # 除去"汇总"
        for i, field in enumerate(fields_grid):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(f_form, text=field + "：", font=("微软雅黑", 11)).grid(
                row=row, column=col, sticky="e", padx=(10, 2), pady=4)
            entry = ttk.Entry(f_form, width=22, font=("微软雅黑", 11))
            entry.grid(row=row, column=col + 1, sticky="w", padx=(2, 20), pady=4)
            self.lzy_entries[field] = entry

        # 汇总：独占一行，使用 Text
        summary_row = (len(fields_grid) + 1) // 2
        ttk.Label(f_form, text="汇总：", font=("微软雅黑", 11)).grid(
            row=summary_row, column=0, sticky="ne", padx=(10, 2), pady=4)
        self.lzy_summary_text = tk.Text(f_form, wrap=tk.WORD, font=("微软雅黑", 11),
                                         relief="solid", borderwidth=1,
                                         width=60, height=4, padx=6, pady=4)
        self.lzy_summary_text.grid(row=summary_row, column=1, columnspan=3,
                                    sticky="w", padx=(2, 20), pady=4)

        # ---- 操作按钮 ----
        btn_row = ttk.Frame(inner)
        btn_row.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_row, text="➕ 添加到表格", command=self._lzy_add_row,
                   style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(btn_row, text="🔄 清空表单", command=self._lzy_clear_form).pack(side="left", padx=5)

        # ===== 数据表格 =====
        f_table = ttk.LabelFrame(inner, text="已录入数据", padding=5)
        f_table.pack(fill="both", expand=True)

        # 表格容器（水平滚动）
        tbl_frame = ttk.Frame(f_table)
        tbl_frame.pack(fill="both", expand=True)

        lzy_style = ttk.Style()
        lzy_style.configure("Lzy.Treeview", font=("微软雅黑", 10), rowheight=28)

        col_ids = [f"c{i}" for i in range(len(self.lzy_columns))]
        self.lzy_tree = ttk.Treeview(tbl_frame, columns=col_ids, show="headings",
                                      selectmode="extended", height=8, style="Lzy.Treeview")
        col_widths = [50, 80, 80, 80, 180, 90, 110, 140, 90, 70, 90, 200]
        for ci, (col_name, w) in enumerate(zip(self.lzy_columns, col_widths)):
            self.lzy_tree.heading(col_ids[ci], text=col_name)
            self.lzy_tree.column(col_ids[ci], width=w, anchor="center" if ci < 4 else "w", minwidth=40)

        # 滚动条
        lzy_vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.lzy_tree.yview)
        lzy_hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.lzy_tree.xview)
        self.lzy_tree.configure(yscrollcommand=lzy_vsb.set, xscrollcommand=lzy_hsb.set)
        self.lzy_tree.grid(row=0, column=0, sticky="nsew")
        lzy_vsb.grid(row=0, column=1, sticky="ns")
        lzy_hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        # 表格操作按钮
        tbl_btn = ttk.Frame(f_table)
        tbl_btn.pack(fill="x", pady=(5, 0))
        ttk.Button(tbl_btn, text="🗑 删除选中行",
                   command=self._lzy_delete_rows).pack(side="left", padx=2)
        ttk.Button(tbl_btn, text="✏️ 编辑选中行（回填表单）",
                   command=self._lzy_edit_selected).pack(side="left", padx=2)

        # 刷新表格
        self._lzy_refresh_tree()

        # ---- 保存栏（嵌入标签页底部） ----
        sep = ttk.Separator(inner, orient="horizontal")
        sep.pack(fill="x", pady=(15, 5))
        bar = ttk.Frame(inner, padding=5)
        bar.pack(fill="x")
        ttk.Label(bar, text="保存目录：").pack(side="left")
        self.entry_save_lzy2 = ttk.Entry(bar, textvariable=self.lzy_save_dir, width=32)
        self.entry_save_lzy2.pack(side="left", padx=5)
        ttk.Button(bar, text="浏览...", command=self._browse_lzy_dir).pack(side="left", padx=3)
        ttk.Button(bar, text="📊 写入 Excel", command=self._lzy_write_excel,
                   style="Accent.TButton").pack(side="right", padx=5)

    # ==================== 两直一白 辅助方法 ====================
    def _lzy_load_data(self):
        """加载持久化数据"""
        if os.path.exists(self.lzy_data_file):
            try:
                with open(self.lzy_data_file, "r", encoding="utf-8") as f:
                    self.lzy_rows = json.load(f)
                return
            except Exception:
                pass
        self.lzy_rows = []

    def _lzy_save_data(self):
        """保存到 JSON"""
        try:
            with open(self.lzy_data_file, "w", encoding="utf-8") as f:
                json.dump(self.lzy_rows, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _lzy_add_row(self):
        """从表单收集数据，添加到表格（或更新正在编辑的行）"""
        month = self.lzy_month.get().strip()
        day = self.lzy_day.get().strip()
        if not month or not day:
            messagebox.showwarning("提示", "请选择日期")
            return
        time_str = f"{month}{day}"

        disaster = self.lzy_disaster.get().strip()
        if not disaster:
            messagebox.showwarning("提示", "请选择灾害性天气")
            return

        row_data = {
            "时间": time_str,
            "灾害性天气": disaster,
            "类型": self.lzy_type.get().strip(),
        }
        for field in self.lzy_fields:
            if field in ("灾害性天气", "类型"):
                continue
            if field == "汇总":
                row_data[field] = self.lzy_summary_text.get("1.0", "end-1c").strip()
            else:
                entry = self.lzy_entries.get(field)
                row_data[field] = entry.get().strip() if entry else ""

        edit_idx = getattr(self, '_lzy_edit_idx', None)
        if edit_idx is not None and 0 <= edit_idx < len(self.lzy_rows):
            # 更新正在编辑的行（保留原序号）
            old_seq = self.lzy_rows[edit_idx].get("序号", "")
            row_data["序号"] = old_seq
            self.lzy_rows[edit_idx] = row_data
            self._lzy_edit_idx = None
            self._lzy_edit_btn.config(text="✏️ 编辑选中行（回填表单）")
        else:
            # 新增行：序号自增
            max_seq = 0
            for r in self.lzy_rows:
                try:
                    max_seq = max(max_seq, int(r.get("序号", 0)))
                except Exception:
                    pass
            row_data["序号"] = str(max_seq + 1)
            self.lzy_rows.append(row_data)

        self._lzy_save_data()
        self._lzy_refresh_tree()
        self._lzy_clear_form()

    def _lzy_clear_form(self):
        """清空表单"""
        self.lzy_disaster.set("")
        self.lzy_type.set("")
        for entry in self.lzy_entries.values():
            entry.delete(0, tk.END)
        self.lzy_summary_text.delete("1.0", tk.END)
        today = date.today()
        self.lzy_month.set(f"{today.month}月")
        self.lzy_day.set(f"{today.day}日")
        # 退出编辑模式
        self._lzy_edit_idx = None
        if hasattr(self, '_lzy_edit_btn'):
            self._lzy_edit_btn.config(text="✏️ 编辑选中行（回填表单）")

    def _lzy_refresh_tree(self):
        """刷新 Treeview"""
        for iid in self.lzy_tree.get_children():
            self.lzy_tree.delete(iid)
        for row in self.lzy_rows:
            vals = [row.get(col, "") for col in self.lzy_columns]
            self.lzy_tree.insert("", tk.END, values=vals)
        if self.lzy_rows:
            self.lzy_tree.configure(height=min(len(self.lzy_rows) + 2, 16))

    def _lzy_delete_rows(self):
        """删除选中行"""
        sel = self.lzy_tree.selection()
        if not sel:
            return
        indices = sorted([self.lzy_tree.index(iid) for iid in sel], reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.lzy_rows):
                self.lzy_rows.pop(idx)
        self._lzy_save_data()
        self._lzy_refresh_tree()

    def _lzy_edit_selected(self):
        """将选中行数据回填到表单（编辑模式）"""
        sel = self.lzy_tree.selection()
        if not sel:
            return
        idx = self.lzy_tree.index(sel[0])
        if idx >= len(self.lzy_rows):
            return
        row = self.lzy_rows[idx]

        # 记录正在编辑的行
        self._lzy_edit_idx = idx
        self._lzy_edit_btn.config(text="✅ 确认修改（当前编辑第{}行）".format(idx + 1))
        # 解析时间
        time_str = row.get("时间", "")
        m = re.match(r'(\d{1,2})月(\d{1,2})日', time_str)
        if m:
            self.lzy_month.set(f"{int(m.group(1))}月")
            self.lzy_day.set(f"{int(m.group(2))}日")
        self.lzy_disaster.set(row.get("灾害性天气", ""))
        self.lzy_type.set(row.get("类型", ""))
        for field in self.lzy_fields:
            if field in ("灾害性天气", "类型", "汇总"):
                continue
            entry = self.lzy_entries.get(field)
            if entry:
                entry.delete(0, tk.END)
                entry.insert(0, row.get(field, ""))
        self.lzy_summary_text.delete("1.0", tk.END)
        self.lzy_summary_text.insert("1.0", row.get("汇总", ""))

    def _browse_lzy_dir(self):
        """浏览：优先选择已存在的 Excel 文件，也可选目录"""
        # 先尝试选择文件
        file = filedialog.askopenfilename(
            title="选择要写入的 Excel 文件（或取消后选目录）",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if file:
            self.lzy_save_dir.set(os.path.dirname(file))
            self._lzy_target_file = file
            self.save_config()
            return
        # 取消选文件 → 选目录
        folder = filedialog.askdirectory(title="选择保存目录（将在该目录创建两直一白.xlsx）")
        if folder:
            self.lzy_save_dir.set(folder)
            self._lzy_target_file = None
            self.save_config()

    def _lzy_write_excel(self):
        """写入 Excel：首次创建带表头，后续追加行"""
        try:
            self._lzy_write_excel_impl()
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("写入 Excel 失败", str(e))

    def _lzy_write_excel_impl(self):
        if not self.lzy_rows:
            messagebox.showwarning("提示", "没有数据可写入")
            return

        save_dir = self.lzy_save_dir.get()
        if not save_dir or not os.path.isdir(save_dir):
            messagebox.showwarning("提示", "请先选择保存目录")
            return

        # 确定目标文件
        target = getattr(self, '_lzy_target_file', None)
        if not target or not os.path.exists(os.path.dirname(target)):
            target = os.path.join(save_dir, "两直一白.xlsx")
        self._lzy_target_file = target

        import shutil

        NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        template = os.path.join(self.app_dir, "两直一白.xlsx")

        if not os.path.exists(target):
            # 首次：复制模板
            if os.path.exists(template):
                shutil.copy2(template, target)
            else:
                messagebox.showerror("错误", f"模板文件不存在：\n{template}")
                return

        # 读取现有 Excel
        with zipfile.ZipFile(target, 'r') as z:
            sst_xml = z.read('xl/sharedStrings.xml')
            sheet_xml = z.read('xl/worksheets/sheet1.xml')
            styles_xml = z.read('xl/styles.xml')
            other_files = {}
            for name in z.namelist():
                if name not in ('xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml',
                                'xl/styles.xml'):
                    other_files[name] = z.read(name)

        sst_root = ET.fromstring(sst_xml)
        sheet_root = ET.fromstring(sheet_xml)
        styles_root = ET.fromstring(styles_xml)

        # --- 查找或创建数据单元格样式（微软雅黑 Light, 14pt, 居中, 细边框） ---
        fonts_elem = styles_root.find(f'{{{NS}}}fonts')
        if fonts_elem is None:
            fonts_elem = ET.SubElement(styles_root, f'{{{NS}}}fonts')

        # 查找已有匹配字体（精确匹配 微软雅黑 Light 14pt）
        data_font_id = None
        for fi, f in enumerate(fonts_elem.findall(f'{{{NS}}}font')):
            sz = f.find(f'{{{NS}}}sz')
            name = f.find(f'{{{NS}}}name')
            if (sz is not None and sz.get('val') == '14' and
                name is not None and name.get('val') == '微软雅黑 Light'):
                data_font_id = fi
                break

        # 没找到则创建
        if data_font_id is None:
            font_count = len(fonts_elem.findall(f'{{{NS}}}font'))
            data_font = ET.SubElement(fonts_elem, f'{{{NS}}}font')
            ET.SubElement(data_font, f'{{{NS}}}sz').set('val', '14')
            ET.SubElement(data_font, f'{{{NS}}}name').set('val', '微软雅黑 Light')
            data_font_id = font_count
            fonts_elem.set('count', str(font_count + 1))

        # 查找已有匹配 xf
        cell_xfs_elem = styles_root.find(f'{{{NS}}}cellXfs')
        if cell_xfs_elem is None:
            cell_xfs_elem = ET.SubElement(styles_root, f'{{{NS}}}cellXfs')
        data_xf_id = None
        for xi, xf in enumerate(cell_xfs_elem.findall(f'{{{NS}}}xf')):
            al = xf.find(f'{{{NS}}}alignment')
            if (xf.get('fontId') == str(data_font_id) and
                xf.get('borderId') == '1' and
                al is not None and
                al.get('horizontal') == 'center' and
                al.get('vertical') == 'center'):
                data_xf_id = xi
                break

        # 没找到则创建
        if data_xf_id is None:
            xf_count = len(cell_xfs_elem.findall(f'{{{NS}}}xf'))
            data_xf = ET.SubElement(cell_xfs_elem, f'{{{NS}}}xf', {
                'fontId': str(data_font_id),
                'fillId': '0',
                'borderId': '1',
                'applyFont': '1',
                'applyAlignment': '1',
                'applyBorder': '1',
            })
            ET.SubElement(data_xf, f'{{{NS}}}alignment', {
                'horizontal': 'center',
                'vertical': 'center',
            })
            data_xf_id = xf_count
            cell_xfs_elem.set('count', str(xf_count + 1))

        # --- 现有的 shared strings ---
        strings = []
        for si in sst_root:
            t = si.find(f'{{{NS}}}t')
            if t is not None and t.text is not None:
                strings.append(t.text)
            else:
                strings.append('')
        str_index = {s: i for i, s in enumerate(strings)}

        def _add_str(s):
            s = str(s) if s is not None else ''
            if s not in str_index:
                str_index[s] = len(strings)
                strings.append(s)
            return str_index[s]

        # --- 找到最后一行的序号 ---
        sheet_data_elem = sheet_root.find(f'{{{NS}}}sheetData')
        if sheet_data_elem is None:
            messagebox.showerror("错误", "Excel 文件中未找到 sheetData")
            return

        max_row = 0
        last_seq = 0
        for row_elem in sheet_data_elem.findall(f'{{{NS}}}row'):
            r = int(row_elem.get('r', 0))
            max_row = max(max_row, r)
            if r == max_row:
                for c in row_elem.findall(f'{{{NS}}}c'):
                    if c.get('r', '').startswith('A'):
                        v = c.find(f'{{{NS}}}v')
                        if v is not None and v.text:
                            if c.get('t') == 's':
                                try:
                                    si = int(v.text)
                                    if 0 <= si < len(strings):
                                        last_seq = int(strings[si])
                                except (ValueError, IndexError):
                                    pass
                            elif c.get('t') is None or c.get('t') == 'n':
                                try:
                                    last_seq = int(v.text)
                                except ValueError:
                                    pass
                        break

        # 序号的起始值至少从表格现有行数+1开始（表头占第1行）
        next_seq = max(last_seq + 1, max_row) if max_row > 1 else last_seq + 1
        for ri, row_data in enumerate(self.lzy_rows):
            row_data["序号"] = str(next_seq + ri)

        # --- 追加新行 ---
        for ri, row_data in enumerate(self.lzy_rows):
            row_num = max_row + 1 + ri
            row_elem = ET.SubElement(sheet_data_elem, f'{{{NS}}}row', {'r': str(row_num)})
            for ci, col_name in enumerate(self.lzy_columns):
                col_letter = chr(ord('A') + ci) if ci < 26 else (
                    'A' + chr(ord('A') + ci - 26))
                cell_ref = f'{col_letter}{row_num}'
                val = str(row_data.get(col_name, ''))
                if val:
                    si = _add_str(val)
                    c = ET.SubElement(row_elem, f'{{{NS}}}c',
                                      {'r': cell_ref, 't': 's',
                                       's': str(data_xf_id)})
                    ET.SubElement(c, f'{{{NS}}}v').text = str(si)
                else:
                    c = ET.SubElement(row_elem, f'{{{NS}}}c',
                                      {'r': cell_ref, 's': str(data_xf_id)})
                    ET.SubElement(c, f'{{{NS}}}v').text = ''

        # --- 重建 SST ---
        sst_ns = f'{{{NS}}}'
        new_sst = ET.Element(f'{sst_ns}sst', {
            'xmlns': NS,
            'count': str(len(strings)),
            'uniqueCount': str(len(strings)),
        })
        for s in strings:
            si_elem = ET.SubElement(new_sst, f'{sst_ns}si')
            t = ET.SubElement(si_elem, f'{sst_ns}t')
            t.text = s if s else ''
            if s and (' ' in s or '\n' in s):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        # --- 写出新的 xlsx ---
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as zout:
            zout.writestr('xl/sharedStrings.xml',
                          ET.tostring(new_sst, encoding='unicode',
                                      xml_declaration=True))
            zout.writestr('xl/worksheets/sheet1.xml',
                          ET.tostring(sheet_root, encoding='unicode',
                                      xml_declaration=True))
            zout.writestr('xl/styles.xml',
                          ET.tostring(styles_root, encoding='unicode',
                                      xml_declaration=True))
            for name, data in other_files.items():
                zout.writestr(name, data)

        # 清空已写入的数据
        self.lzy_rows.clear()
        self._lzy_save_data()
        self._lzy_refresh_tree()

        try:
            os.startfile(os.path.normpath(os.path.dirname(target)))
        except Exception:
            pass
        messagebox.showinfo("成功", f"数据已写入 Excel：\n{target}")

    # ========== 服务记录（可编辑表格 + Excel 导出） ==========
    def build_service_record(self, parent):
        """构建服务记录标签页"""
        # ---- 数据文件路径 ----
        self.service_data_file = os.path.join(self.app_dir, "service_records.json")

        # ---- 事项下拉选项 ----
        self.service_items = [
            "发布雷电黄色预警信号", "发布雷电橙色预警信号", "发布雷电红色预警信号",
            "发布暴雨黄色预警信号", "发布暴雨橙色预警信号", "发布暴雨红色预警信号",
            "发布大风黄色预警信号", "发布大风橙色预警信号", "发布大风红色预警信号",
            "发布雷暴大风黄色预警信号", "发布雷暴大风橙色预警信号", "发布雷暴大风红色预警信号",
            "入驻防指",
        ]

        # ---- 服务方式选项 ----
        self.service_modes = ["浙政钉", "微信群", "电话", "短信", "闪信", "语音外呼"]

        # ---- 头部：日期 + 天气类型 ----
        header_frame = ttk.Frame(parent, padding=10)
        header_frame.pack(fill="x")

        today = date.today()
        ttk.Label(header_frame, text="", font=("微软雅黑", 13, "bold")).pack(side="left")  # spacer
        self.svc_year = ttk.Combobox(header_frame, values=[str(y) for y in range(today.year - 2, today.year + 3)],
                                     width=5, state="normal")
        self.svc_year.set(str(today.year))
        self.svc_year.pack(side="left")
        ttk.Label(header_frame, text="年", font=("微软雅黑", 12)).pack(side="left")
        self.svc_month = ttk.Combobox(header_frame, values=[f"{m:02d}" for m in range(1, 13)],
                                      width=4, state="normal")
        self.svc_month.set(f"{today.month:02d}")
        self.svc_month.pack(side="left")
        ttk.Label(header_frame, text="月", font=("微软雅黑", 12)).pack(side="left")
        self.svc_day = ttk.Combobox(header_frame, values=[f"{d:02d}" for d in range(1, 32)],
                                    width=4, state="normal")
        self.svc_day.set(f"{today.day:02d}")
        self.svc_day.pack(side="left")
        ttk.Label(header_frame, text="日", font=("微软雅黑", 12)).pack(side="left")

        self.svc_weather = ttk.Combobox(header_frame, values=["强对流", "台风", "暴雨", "大风"],
                                        width=6, state="normal")
        self.svc_weather.set("")
        self.svc_weather.pack(side="left", padx=5)
        ttk.Label(header_frame, text="天气服务记录", font=("微软雅黑", 13, "bold")).pack(side="left", padx=5)

        # ---- 表格 ----
        table_frame = ttk.Frame(parent, padding=5)
        table_frame.pack(fill="both", expand=True)

        columns = ("序号", "时间", "事项", "内容", "服务方式", "备注")
        # 增大行高便于阅读
        svc_style = ttk.Style()
        svc_style.configure("Svc.Treeview", font=("微软雅黑", 11), rowheight=36)
        self.svc_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                     selectmode="extended", height=12, style="Svc.Treeview")
        col_widths = [50, 120, 180, 300, 160, 120]
        for col, w in zip(columns, col_widths):
            self.svc_tree.heading(col, text=col)
            self.svc_tree.column(col, width=w, anchor="center" if col in ("序号", "时间") else "w", minwidth=40)

        # 滚动条
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.svc_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.svc_tree.xview)
        self.svc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.svc_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # 单击编辑（改为 ButtonRelease 以兼容拖动选择）
        self.svc_tree.bind("<ButtonRelease-1>", self._on_svc_click_or_edit)

        # 双击表头编辑
        self.svc_tree.bind("<Double-1>", self._on_svc_double_click)

        # 右键菜单：单元格格式化
        self.svc_tree.bind("<Button-3>", self._on_svc_right_click)

        # 鼠标拖动多选
        self._svc_drag_data = {"start": None, "dragging": False, "moved": False}
        self.svc_tree.bind("<ButtonPress-1>", self._on_svc_drag_start, add="+")
        self.svc_tree.bind("<B1-Motion>", self._on_svc_drag_move, add="+")
        self.svc_tree.bind("<ButtonRelease-1>", self._on_svc_drag_stop, add="+")

        # ---- 底栏：保存目录 + 清空 + 生成 Excel ----
        sep2 = ttk.Separator(parent, orient="horizontal")
        sep2.pack(fill="x", pady=(0, 0))
        bottom_bar = ttk.Frame(parent, padding=5)
        bottom_bar.pack(fill="x", pady=(0, 5))

        ttk.Label(bottom_bar, text="保存目录：", font=("微软雅黑", 11)).pack(side="left")
        saved = self.load_config()
        svc_dir = saved.get("svc_save_dir", "")
        self.svc_save_dir = tk.StringVar(value=svc_dir if svc_dir and os.path.isdir(svc_dir) else os.getcwd())
        self.entry_svc_save = ttk.Entry(bottom_bar, textvariable=self.svc_save_dir, width=28)
        self.entry_svc_save.pack(side="left", padx=5)
        ttk.Button(bottom_bar, text="浏览...", command=self._browse_svc_dir).pack(side="left", padx=3)
        ttk.Separator(bottom_bar, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        ttk.Button(bottom_bar, text="➕ 添加行", command=self._add_empty_row).pack(side="left", padx=2)
        ttk.Button(bottom_bar, text="📌 插入行", command=self._insert_row_above).pack(side="left", padx=2)
        ttk.Button(bottom_bar, text="🗑 删除选中行", command=self._delete_selected_row).pack(side="left", padx=2)
        ttk.Button(bottom_bar, text="🔄 清空", command=self._clear_service_data).pack(side="left", padx=2)
        ttk.Separator(bottom_bar, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        ttk.Button(bottom_bar, text="📊 生成 Excel 文档", command=self._export_service_excel,
                   style="Accent.TButton").pack(side="right", padx=5)

        # 加载持久化数据（无数据时预置 3 行空行）
        self._load_service_data()
        if not self.svc_tree.get_children():
            for _ in range(3):
                self._add_empty_row(save=False)

    def _browse_svc_dir(self):
        folder = filedialog.askdirectory(title="选择服务记录保存目录")
        if folder:
            self.svc_save_dir.set(folder)
            self.save_config()

    def add_service_row(self, time_val="", item_val="", content_val="", mode_val="", remark_val=""):
        """外部调用：向服务记录表格添加一行（预通报/实况通报自动填入时使用）"""
        items = self.svc_tree.get_children()
        next_num = len(items) + 1
        row_vals = (str(next_num), time_val, item_val, content_val, mode_val, remark_val)
        self.svc_tree.insert("", tk.END, values=row_vals)
        self._save_service_data()

    def _renumber_rows(self):
        """重新编号所有行"""
        for i, iid in enumerate(self.svc_tree.get_children(), 1):
            vals = list(self.svc_tree.item(iid, "values"))
            vals[0] = str(i)
            self.svc_tree.item(iid, values=vals)

    def _add_empty_row(self, save=True):
        """在末尾添加一个空白行"""
        items = self.svc_tree.get_children()
        next_num = len(items) + 1
        self.svc_tree.insert("", tk.END, values=(str(next_num), "", "", "", "", ""))
        if save:
            self._save_service_data()

    def _insert_row_above(self):
        """在选中行上方插入空白行"""
        sel = self.svc_tree.selection()
        if sel:
            # 在选中行之前插入
            before_iid = sel[0]
            idx = self.svc_tree.index(before_iid)
            self.svc_tree.insert("", idx, values=("", "", "", "", "", ""))
        else:
            # 没有选中行则添加到末尾
            self.svc_tree.insert("", tk.END, values=("", "", "", "", "", ""))
        self._renumber_rows()
        self._save_service_data()

    def _delete_selected_row(self):
        """删除选中行"""
        sel = self.svc_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先点击选中要删除的行")
            return
        # 先关闭正在编辑的控件
        self._svc_close_editor()
        for iid in sel:
            self.svc_tree.delete(iid)
        self._renumber_rows()
        self._save_service_data()

    def _svc_close_editor(self):
        """关闭当前活跃的编辑控件"""
        w = getattr(self, '_svc_active_widget', None)
        if w is not None and w.winfo_exists():
            w.destroy()
        self._svc_active_widget = None
        self._svc_editing = False

    def _edit_service_cell(self, event):
        """单击单元格编辑"""
        # 先关闭旧编辑器
        self._svc_close_editor()

        region = self.svc_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self.svc_tree.identify_column(event.x)
        item_id = self.svc_tree.identify_row(event.y)
        if not item_id:
            return
        col_idx = int(col_id.replace("#", "")) - 1
        col_name = self.svc_tree["columns"][col_idx]

        if col_name == "序号":
            return  # 序号不可编辑

        self._svc_editing = True  # 编辑中，防止重复触发

        cur_values = list(self.svc_tree.item(item_id, "values"))
        cur_text = cur_values[col_idx] if col_idx < len(cur_values) else ""

        # 先选中该行（视觉反馈）
        self.svc_tree.selection_set(item_id)

        if col_name == "事项":
            self._edit_with_combobox(item_id, col_idx, cur_text, self.service_items)
        elif col_name == "服务方式":
            self._edit_with_checkboxes(item_id, col_idx, cur_text)
        else:
            self._edit_with_entry(item_id, col_idx, cur_text, col_name)

    def _on_svc_double_click(self, event):
        """双击：表头编辑 或 单元格编辑"""
        # 如果正在编辑中，忽略（防止重复编辑）
        if getattr(self, '_svc_editing', False):
            return
        region = self.svc_tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.svc_tree.identify_column(event.x)
            if not col_id:
                return
            ci = int(col_id.replace("#", "")) - 1
            col_name = self.svc_tree["columns"][ci]
            # 定位表头位置
            x = 0
            for c in range(ci):
                x += self.svc_tree.column(f"#{c + 1}", "width")
            w = self.svc_tree.column(f"#{ci + 1}", "width")
            e = ttk.Entry(self.svc_tree, width=max(w // 10, 10),
                         font=("微软雅黑", 11, "bold"))
            e.place(x=x + 2, y=0, width=w - 4, height=30)
            e.insert(0, self.svc_tree.heading(col_name, "text"))
            e.lift()
            e.focus_set()
            def _save_header():
                new_text = e.get().strip() or col_name
                self.svc_tree.heading(col_name, text=new_text)
                e.destroy()
            e.bind("<Return>", lambda ev: _save_header())
            e.bind("<FocusOut>", lambda ev: self.svc_tree.after(100, _save_header))
        elif region == "cell":
            # 回退到单击编辑逻辑
            self._edit_service_cell(event)

    def _on_svc_right_click(self, event):
        """右键菜单：单元格格式化 + 行操作"""
        item_id = self.svc_tree.identify_row(event.y)
        col_id = self.svc_tree.identify_column(event.x)

        menu = tk.Menu(self.root, tearoff=0)

        # 单元格操作（仅当点击在有效单元格上时显示）
        has_cell = False
        if col_id and item_id:
            ci = int(col_id.replace("#", "")) - 1
            if ci >= 0:
                has_cell = True
                menu.add_command(label="🎨 填充背景色...",
                    command=lambda iid=item_id, c=ci: self._svc_cell_bgcolor(iid, c))
                menu.add_command(label="✏️ 文字颜色...",
                    command=lambda iid=item_id, c=ci: self._svc_cell_fgcolor(iid, c))
                menu.add_command(label="𝐁 加粗",
                    command=lambda iid=item_id: self._svc_cell_bold(iid))
                menu.add_separator()
                menu.add_command(label="📝 编辑单元格",
                    command=lambda: self._edit_service_cell(event))

        # 行操作（始终可用）
        if has_cell:
            menu.add_separator()
        menu.add_command(label="📌 在上方插入行",
            command=self._insert_row_above)
        sel = self.svc_tree.selection()
        if len(sel) >= 2:
            menu.add_separator()
            menu.add_command(label="🔗 合并选中行",
                command=lambda s=sel: self._svc_merge_rows(s))
        menu.post(event.x_root, event.y_root)

    def _svc_cell_bgcolor(self, iid, ci):
        """设置单元格背景色"""
        try:
            from tkinter import colorchooser
            c = colorchooser.askcolor(title="选择背景色")
            if c and c[1]:
                tag_name = f"svc_bg_{c[1].replace('#', '')}_{iid}"
                self.svc_tree.tag_configure(tag_name, background=c[1])
                cur_tags = list(self.svc_tree.item(iid, "tags"))
                # 移除旧背景 tag
                cur_tags = [t for t in cur_tags if not t.startswith("svc_bg_")]
                cur_tags.append(tag_name)
                self.svc_tree.item(iid, tags=tuple(cur_tags))
        except Exception:
            pass

    def _svc_cell_fgcolor(self, iid, ci):
        """设置单元格文字颜色"""
        try:
            from tkinter import colorchooser
            c = colorchooser.askcolor(title="选择文字颜色")
            if c and c[1]:
                tag_name = f"svc_fg_{c[1].replace('#', '')}_{iid}"
                self.svc_tree.tag_configure(tag_name, foreground=c[1])
                cur_tags = list(self.svc_tree.item(iid, "tags"))
                cur_tags = [t for t in cur_tags if not t.startswith("svc_fg_")]
                cur_tags.append(tag_name)
                self.svc_tree.item(iid, tags=tuple(cur_tags))
        except Exception:
            pass

    def _svc_cell_bold(self, iid):
        """切换行加粗"""
        try:
            cur_tags = list(self.svc_tree.item(iid, "tags"))
            if "svc_bold" in cur_tags:
                cur_tags.remove("svc_bold")
            else:
                self.svc_tree.tag_configure("svc_bold", font=("微软雅黑", 11, "bold"))
                cur_tags.append("svc_bold")
            self.svc_tree.item(iid, tags=tuple(cur_tags))
        except Exception:
            pass

    def _svc_merge_rows(self, selection):
        """垂直合并选中行：将下面行的内容追加到第一行"""
        if len(selection) < 2:
            return
        iids = list(selection)
        first_iid = iids[0]
        first_vals = list(self.svc_tree.item(first_iid, "values"))
        for iid in iids[1:]:
            vals = list(self.svc_tree.item(iid, "values"))
            for ci in range(min(len(first_vals), len(vals))):
                if vals[ci].strip():
                    sep = "\n" if first_vals[ci].strip() else ""
                    first_vals[ci] = first_vals[ci] + sep + vals[ci]
            self.svc_tree.delete(iid)
        self.svc_tree.item(first_iid, values=first_vals)
        self._renumber_rows()
        self._save_service_data()

    def _on_svc_drag_start(self, event):
        """鼠标拖动开始（仅做记录，不改变选择）"""
        item_id = self.svc_tree.identify_row(event.y)
        if item_id:
            self._svc_drag_data["start"] = item_id
            self._svc_drag_data["dragging"] = True
            self._svc_drag_data["moved"] = False

    def _on_svc_drag_move(self, event):
        """鼠标拖动中"""
        if not self._svc_drag_data.get("dragging"):
            return
        item_id = self.svc_tree.identify_row(event.y)
        if item_id and item_id != self._svc_drag_data.get("start"):
            self._svc_drag_data["moved"] = True
            sel = set(self.svc_tree.selection())
            sel.add(item_id)
            all_items = self.svc_tree.get_children()
            start_iid = self._svc_drag_data["start"]
            try:
                start_idx = all_items.index(start_iid)
                cur_idx = all_items.index(item_id)
            except ValueError:
                return
            lo, hi = min(start_idx, cur_idx), max(start_idx, cur_idx)
            for i in range(lo, hi + 1):
                sel.add(all_items[i])
            self.svc_tree.selection_set(list(sel))
            self._svc_drag_data["start"] = item_id

    def _on_svc_drag_stop(self, event):
        """鼠标拖动结束"""
        self._svc_drag_data["dragging"] = False
        self._svc_drag_data["start"] = None

    def _on_svc_click_or_edit(self, event):
        """ButtonRelease 时：如果没有拖动，执行单元格编辑"""
        if self._svc_drag_data.get("moved"):
            self._svc_drag_data["moved"] = False
            return  # 拖动选择，不编辑
        # 防止重复编辑（编辑进行中时不再触发）
        if getattr(self, '_svc_editing', False):
            return
        self._edit_service_cell(event)

    def _edit_with_entry(self, item_id, col_idx, cur_text, col_name):
        """Entry / Text 编辑（内容列用 Text 支持换行）"""
        x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")

        if col_name == "内容":
            # 使用多行 Text 控件，支持自动换行
            text_w = tk.Text(self.svc_tree, width=max(w // 8, 20), height=5,
                            font=("微软雅黑", 11), wrap=tk.WORD,
                            padx=4, pady=4, relief="solid", borderwidth=1)
            text_w.place(x=x, y=y, width=w, height=max(h * 3, 80))
            text_w.insert("1.0", cur_text)
            text_w.lift()  # 确保在最上层
            text_w.focus_set()
            widget = text_w
            self._svc_active_widget = widget

            def _get_val():
                return text_w.get("1.0", "end-1c").strip()

            def _delayed_destroy():
                if text_w.winfo_exists():
                    text_w.destroy()
        else:
            entry = ttk.Entry(self.svc_tree, width=max(w // 10, 10))
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, cur_text)
            entry.lift()
            entry.focus_set()
            widget = entry
            self._svc_active_widget = widget

            def _get_val():
                return entry.get().strip()

            def _delayed_destroy():
                if entry.winfo_exists():
                    entry.destroy()

        destroyed = [False]

        def _save():
            if destroyed[0]:
                return
            destroyed[0] = True
            self._svc_editing = False
            new_val = _get_val()
            values = list(self.svc_tree.item(item_id, "values"))
            values[col_idx] = new_val
            self.svc_tree.item(item_id, values=values)
            _delayed_destroy()
            self._save_service_data()
            # 清理全局点击绑定
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        def _on_root_click(event):
            if destroyed[0]:
                return
            # 判断点击是否在编辑控件内部
            w = event.widget
            while w is not None and w != self.root:
                if w == widget:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            self.svc_tree.after(50, _save)

        _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')

        widget.bind("<Return>", lambda e: _save())
        widget.bind("<FocusOut>", lambda e: self.svc_tree.after(150, _save))

    def _edit_with_combobox(self, item_id, col_idx, cur_text, options):
        """事项列 Combobox 编辑（简洁自包含的智能下拉）"""
        x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")
        if not h:  # bbox 失败则放弃
            return

        # ── 编辑框（只读，点击/输入触发下拉）──
        cb = ttk.Combobox(self.svc_tree, values=options, width=max(w // 10, 15))
        cb.configure(state="normal")
        cb.place(x=x, y=y, width=max(w, 180), height=h + 20)
        cb.set(cur_text)
        cb.lift()
        cb.focus_set()
        self._svc_active_widget = cb

        destroyed = [False]
        popup = [None]
        listbox = [None]
        timer = [None]

        def _cleanup():
            if popup[0] and popup[0].winfo_exists():
                popup[0].destroy()
            popup[0] = None
            listbox[0] = None

        def _save_and_close():
            if destroyed[0]:
                return
            destroyed[0] = True
            self._svc_editing = False
            _cleanup()
            new_val = cb.get().strip()
            values = list(self.svc_tree.item(item_id, "values"))
            values[col_idx] = new_val
            self.svc_tree.item(item_id, values=values)
            if cb.winfo_exists():
                cb.destroy()
            self._save_service_data()
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        # ── 自定义弹出列表 ──
        def _show_popup():
            _cleanup()
            vals = cb['values']
            if not vals:
                return
            popup[0] = tk.Toplevel(self.root)
            popup[0].overrideredirect(True)
            popup[0].attributes('-topmost', True)
            popup[0].configure(bg="#3498db")
            inner = tk.Frame(popup[0], bg="white")
            inner.pack(padx=1, pady=1)
            listbox[0] = tk.Listbox(inner, font=("微软雅黑", 11),
                                    selectmode=tk.SINGLE, exportselection=False,
                                    bg="white", fg="#2c3e50",
                                    selectbackground="#3498db", selectforeground="white",
                                    relief="flat", borderwidth=0, highlightthickness=0)
            for v in vals:
                listbox[0].insert(tk.END, v)
            n = min(len(vals), 8)
            listbox[0].config(height=n)
            listbox[0].pack(fill="both", expand=True)
            popup[0].update_idletasks()
            px = cb.winfo_rootx()
            py = cb.winfo_rooty() + cb.winfo_height()
            pw = cb.winfo_width()
            popup[0].geometry(f"{pw}x{listbox[0].winfo_reqheight() + 2}+{px}+{py}")
            popup[0].deiconify()

            def _on_list_click(event):
                if listbox[0] and listbox[0].curselection():
                    text = listbox[0].get(listbox[0].curselection()[0])
                    cb.set(text)
                    cb.icursor(len(text))
                    _save_and_close()

            def _on_list_key(event):
                if event.keysym == "Return":
                    _on_list_click(event)
                elif event.keysym == "Escape":
                    _cleanup()
                    cb.focus_set()
            listbox[0].bind('<ButtonRelease-1>', _on_list_click)
            listbox[0].bind('<KeyRelease>', _on_list_key)

        # ── 键盘输入 → 过滤 + 弹窗 ──
        def _on_keyrelease(event):
            if destroyed[0]:
                return
            if event.keysym in ("Escape",):
                _cleanup()
                return
            if event.keysym == "Return":
                # 如果弹窗可见且有选中项，选它
                if popup[0] and popup[0].winfo_viewable() and listbox[0]:
                    sel = listbox[0].curselection()
                    if sel:
                        cb.set(listbox[0].get(sel[0]))
                _save_and_close()
                return
            if event.keysym in ("Up", "Down", "Prior", "Next"):
                # 确保弹窗显示
                if not popup[0] or not popup[0].winfo_viewable():
                    cb['values'] = list(options)
                    _show_popup()
                if listbox[0] and listbox[0].size():
                    listbox[0].focus_set()
                    # 移动选择
                    sel = listbox[0].curselection()
                    cur_idx = sel[0] if sel else -1
                    if event.keysym == "Down":
                        new_idx = min(cur_idx + 1, listbox[0].size() - 1)
                    elif event.keysym == "Up":
                        new_idx = max(cur_idx - 1, 0)
                    elif event.keysym == "Prior":
                        new_idx = max(cur_idx - 5, 0)
                    else:  # Next
                        new_idx = min(cur_idx + 5, listbox[0].size() - 1)
                    listbox[0].selection_clear(0, tk.END)
                    listbox[0].selection_set(new_idx)
                    listbox[0].see(new_idx)
                    # 同步填入输入框
                    cb.set(listbox[0].get(new_idx))
                    cb.icursor(len(cb.get()))
                return
            if event.keysym in ("Left", "Right", "Tab", "Control_L", "Control_R",
                                "Shift_L", "Shift_R", "Alt_L", "Alt_R",
                                "Home", "End", "Caps_Lock"):
                return

            # 过滤
            value = event.widget.get()
            if value:
                filtered = [s for s in options if value.lower() in s.lower()]
            else:
                filtered = list(options)
            cb['values'] = filtered
            # 防抖 200ms 后弹出
            if timer[0] is not None:
                cb.after_cancel(timer[0])
            timer[0] = cb.after(200, _show_popup)

        cb.bind('<KeyRelease>', _on_keyrelease)

        # ── 点击箭头区域 → 显示完整列表 ──
        def _on_click(event):
            if event.x > event.widget.winfo_width() - 25:
                cb['values'] = list(options)
                _show_popup()
                return 'break'

        cb.bind('<Button-1>', _on_click)

        # ── 失焦延迟保存 ──
        def _on_focusout(event):
            def _check():
                if destroyed[0]:
                    return
                if popup[0] and popup[0].winfo_viewable():
                    # 检查：焦点是否转移到了弹窗内的 listbox？
                    fw = popup[0].winfo_containing(
                        popup[0].winfo_pointerx(),
                        popup[0].winfo_pointery())
                    if fw and (fw == listbox[0] or fw == popup[0] or
                               (hasattr(fw, 'master') and
                                (fw.master == popup[0] or
                                 (hasattr(fw.master, 'master') and fw.master.master == popup[0])))):
                        return  # 焦点/鼠标在弹窗上，不关
                    # 额外检查：焦点控件是否是 listbox
                    focus_w = popup[0].focus_get()
                    if focus_w and (focus_w == listbox[0]):
                        return
                self.svc_tree.after(100, _save_and_close)
            self.svc_tree.after(150, _check)

        cb.bind('<FocusOut>', _on_focusout)

        # ── 全局点击外部 → 保存 ──
        def _on_root_click(event):
            if destroyed[0]:
                return
            w = event.widget
            while w is not None and w != self.root:
                if w == cb:
                    return
                if w == popup[0]:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            self.svc_tree.after(50, _save_and_close)

        _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')

    def _edit_with_checkboxes(self, item_id, col_idx, cur_text):
        """服务方式列多选框编辑（点击外部自动关闭，无需 grab）"""
        x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")
        selected = set(s.strip() for s in cur_text.split("、") if s.strip())

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)  # 始终置顶，不被其他控件遮挡
        popup.configure(bg="#3498db")
        inner = tk.Frame(popup, bg="white", padx=10, pady=10)
        inner.pack(padx=1, pady=1)
        ttk.Label(inner, text="选择服务方式", font=("微软雅黑", 12, "bold")).pack(anchor="w", pady=(0, 8))

        vars_ = {}
        for mode in self.service_modes:
            v = tk.BooleanVar(value=mode in selected)
            cb = tk.Checkbutton(inner, text=mode, variable=v,
                               font=("微软雅黑", 11),
                               bg="white", anchor="w",
                               selectcolor="white",
                               activebackground="white")
            cb.pack(anchor="w", fill="x")
            vars_[mode] = v

        def _confirm():
            self._svc_editing = False
            result = "、".join(m for m in self.service_modes if vars_[m].get())
            values = list(self.svc_tree.item(item_id, "values"))
            values[col_idx] = result
            self.svc_tree.item(item_id, values=values)
            popup.destroy()
            self._save_service_data()
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        btn_frame = ttk.Frame(inner)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_frame, text="确定", command=_confirm).pack(side="right", padx=3)
        def _cancel():
            self._svc_editing = False
            popup.destroy()
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass
        ttk.Button(btn_frame, text="取消", command=_cancel).pack(side="right", padx=3)

        # 收集弹窗内所有子 widget 用于点击检测
        popup_widgets = {popup, inner, btn_frame}
        def _collect_children(w):
            popup_widgets.add(w)
            for child in w.winfo_children():
                _collect_children(child)
        _collect_children(inner)

        def _on_root_click(event):
            if not popup.winfo_exists() or not popup.winfo_viewable():
                return
            w = event.widget
            while w is not None:
                if w in popup_widgets:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            # 点击弹窗外 → 等同点击"确定"
            _confirm()

        _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')

        def _on_popup_destroy(event):
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass
        popup.bind('<Destroy>', _on_popup_destroy)

        popup.update_idletasks()
        px = self.svc_tree.winfo_rootx() + x
        py = self.svc_tree.winfo_rooty() + y + h
        popup.geometry(f"+{px}+{py}")
        popup.focus_set()
        # 不用 grab_set()，让外部点击可以触发 root 的 Button-1 绑定
        self.root.wait_window(popup)

    def _save_service_data(self):
        """持久化表格数据到 JSON"""
        items = self.svc_tree.get_children()
        data = []
        for iid in items:
            vals = self.svc_tree.item(iid, "values")
            data.append(list(vals))
        try:
            with open(self.service_data_file, "w", encoding="utf-8") as f:
                json.dump({"date": date.today().isoformat(), "weather": self.svc_weather.get(), "rows": data},
                          f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_service_data(self):
        """从 JSON 加载表格数据"""
        if not os.path.exists(self.service_data_file):
            return
        try:
            with open(self.service_data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.svc_tree.delete(*self.svc_tree.get_children())
            # 恢复头部
            saved_date = data.get("date", "")
            if saved_date:
                try:
                    d = date.fromisoformat(saved_date)
                    self.svc_year.set(str(d.year))
                    self.svc_month.set(f"{d.month:02d}")
                    self.svc_day.set(f"{d.day:02d}")
                except Exception:
                    pass
            saved_weather = data.get("weather", "")
            if saved_weather:
                self.svc_weather.set(saved_weather)
            # 恢复行
            for row in data.get("rows", []):
                self.svc_tree.insert("", tk.END, values=row)
        except Exception:
            pass

    def _clear_service_data(self):
        """清空所有数据"""
        if messagebox.askyesno("确认清空", "确定要清空表格中的所有数据吗？此操作不可恢复。"):
            self.svc_tree.delete(*self.svc_tree.get_children())
            self._save_service_data()

    def _get_svc_header_title(self):
        """获取表头标题"""
        y = self.svc_year.get()
        m = self.svc_month.get()
        d = self.svc_day.get()
        w = self.svc_weather.get()
        return f"{y}年{m}月{d}日{w}天气服务记录"

    def _export_service_excel(self):
        """导出为 Excel (.xlsx) 文件"""
        try:
            title = self._get_svc_header_title()
            # 文件名以表头文字命名：2026年06月22日强对流天气服务记录.xlsx
            filename = f"{title}.xlsx"
            save_path = os.path.join(self.svc_save_dir.get(), filename)

            if os.path.exists(save_path):
                base, ext = os.path.splitext(filename)
                count = 2
                while True:
                    new_name = f"{base}_{count}{ext}"
                    save_path = os.path.join(self.svc_save_dir.get(), new_name)
                    if not os.path.exists(save_path):
                        break
                    count += 1

            # 手动构建 xlsx（zip 包 xml）
            self._write_xlsx(save_path, title)

            try:
                os.startfile(os.path.normpath(os.path.dirname(save_path)))
            except Exception:
                pass
            messagebox.showinfo("成功", f"Excel 文件已生成：\n{save_path}")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{e}")

    def _write_xlsx(self, path, title):
        """用 XML 构建 xlsx 文件（按模板格式）"""
        columns = ["序号", "时间", "事项", "内容", "服务方式", "备注"]
        rows_data = []
        for iid in self.svc_tree.get_children():
            rows_data.append(list(self.svc_tree.item(iid, "values")))

        NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

        # --- 分析时间列：尝试转为 Excel 时间序列号 ---
        def _parse_time(val):
            """尝试将 'HHMM' 或 'HH:MM' 或 'HHMM:HHMM' 格式转为 Excel 时间序列号"""
            if not val:
                return None, val
            val = val.strip()
            for sep in [';', '；']:
                val = val.split(sep)[0]
            val = val.strip()
            if ':' in val:
                parts = val.split(':')
                if len(parts) >= 2:
                    try:
                        h, m = int(parts[0]), int(parts[1])
                        return (h * 60 + m) / 1440.0, f"{h:02d}:{m:02d}"
                    except ValueError:
                        pass
            # Try 4-digit format like "1430"
            if len(val) >= 4 and val[:4].isdigit():
                try:
                    h, m = int(val[:2]), int(val[2:4])
                    if 0 <= h <= 23 and 0 <= m <= 59:
                        return (h * 60 + m) / 1440.0, f"{h:02d}:{m:02d}"
                except ValueError:
                    pass
            return None, val

        # --- 构建 sharedStrings ---
        all_strings = []
        str_index = {}
        def _add_str(s):
            s = str(s) if s is not None else ''
            if s not in str_index:
                str_index[s] = len(all_strings)
                all_strings.append(s)
            return str_index[s]

        title_idx = _add_str(title)
        for col in columns:
            _add_str(col)

        # 收集数据并分析时间
        time_serials = []  # per row: Excel serial or None
        for row in rows_data:
            time_val = row[1] if len(row) > 1 else ''
            serial, _ = _parse_time(time_val)
            time_serials.append(serial)
            for cell in row:
                _add_str(str(cell))

        # --- Shared Strings XML ---
        sst_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst {NS} count="{len(all_strings)}" uniqueCount="{len(all_strings)}">'
        for s in all_strings:
            # 彻底转义所有 XML 敏感字符和控制字符
            escaped = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                       .replace('"', "&quot;").replace("'", "&apos;"))
            # 移除非法 XML 控制字符（除 \t \n \r）
            escaped = ''.join(c for c in escaped if ord(c) >= 0x20 or c in '\t\n\r')
            sst_xml += f'<si><t>{escaped}</t></si>'
        sst_xml += '</sst>'

        # --- Styles ---
        # Fonts: 0=default(宋体11), 1=title(宋体26bold), 2=header(宋体20bold), 3=time(宋体11)
        # Borders: 0=none, 1=thin all sides
        # numFmt: 0=general, 20=h:mm, 166=yyyy/m/d h:mm
        styles_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet {NS}>'
        styles_xml += '<fonts count="4">'
        styles_xml += '<font><sz val="11"/><name val="宋体"/></font>'            # 0
        styles_xml += '<font><sz val="26"/><name val="宋体"/><b/></font>'       # 1 - title
        styles_xml += '<font><sz val="20"/><name val="宋体"/><b/></font>'       # 2 - header
        styles_xml += '<font><sz val="11"/><name val="宋体"/></font>'           # 3 - time
        styles_xml += '</fonts>'
        styles_xml += '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        styles_xml += '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>'
        styles_xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        styles_xml += '<cellXfs count="6">'
        styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'  # xf0 default
        # xf1: title (center, vertical center)
        styles_xml += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        # xf2: header (center, thin border)
        styles_xml += '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        # xf3: data center with border (序号, 时间, 服务方式)
        styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        # xf4: data center with border + time format
        styles_xml += '<xf numFmtId="20" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        # xf5: content left-align with border + wrap
        styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
        styles_xml += '</cellXfs>'
        styles_xml += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        styles_xml += '</styleSheet>'

        # --- Sheet data ---
        sheet_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet {NS} {NS_R}>'
        # Page setup: landscape A4
        sheet_xml += '<sheetPr/>'
        sheet_xml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane yOffset="2" xOffset="0" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        sheet_xml += '<sheetFormatPr defaultRowHeight="15"/>'
        # Column widths (matching template)
        col_widths = [12.67, 16, 34.13, 56.63, 27.44, 28.33]
        sheet_xml += '<cols>'
        for i, w in enumerate(col_widths):
            sheet_xml += f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
        sheet_xml += '</cols>'
        sheet_xml += '<sheetData>'

        # Row 1: Title (merged A1:F1, s=1)
        sheet_xml += '<row r="1" ht="32.4">'
        sheet_xml += f'<c r="A1" t="s" s="1"><v>{title_idx}</v></c>'
        for ci in range(1, 6):
            sheet_xml += f'<c r="{chr(65 + ci)}1" s="1"/>'
        sheet_xml += '</row>'

        # Row 2: Headers (s=2)
        sheet_xml += '<row r="2" ht="25.8">'
        for ci, col in enumerate(columns):
            si = str_index[col]
            sheet_xml += f'<c r="{chr(65 + ci)}2" t="s" s="2"><v>{si}</v></c>'
        sheet_xml += '</row>'

        # Data rows
        for ri, row in enumerate(rows_data, 3):
            sheet_xml += f'<row r="{ri}">'
            for ci, val in enumerate(row):
                cell_ref = f"{chr(65 + ci)}{ri}"
                if ci == 0:  # 序号 - center
                    si = str_index[str(val)]
                    sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                elif ci == 1:  # 时间 - try Excel time or string
                    serial = time_serials[ri - 3]
                    if serial is not None:
                        sheet_xml += f'<c r="{cell_ref}" s="4"><v>{serial}</v></c>'
                    else:
                        si = str_index[str(val)]
                        sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                elif ci == 2 or ci == 4:  # 事项, 服务方式 - center
                    si = str_index[str(val)]
                    sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                elif ci == 3 or ci == 5:  # 内容, 备注 - left wrap
                    si = str_index[str(val)]
                    sheet_xml += f'<c r="{cell_ref}" t="s" s="5"><v>{si}</v></c>'
            sheet_xml += '</row>'

        sheet_xml += '</sheetData>'
        sheet_xml += f'<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>'
        sheet_xml += f'<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.3" footer="0.3"/>'
        sheet_xml += f'<pageSetup orientation="landscape" paperSize="9" fitToHeight="1" fitToWidth="1"/>'
        sheet_xml += '</worksheet>'

        # --- Workbook ---
        workbook_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {NS} {NS_R}><sheets><sheet name="服务记录" sheetId="1" r:id="rId1"/></sheets></workbook>'

        # --- Package relationships (_rels/.rels) → 指向 xl/workbook.xml ---
        rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

        # --- Workbook relationships (xl/_rels/workbook.xml.rels) → 指向内部部件 ---
        wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

        # --- [Content_Types].xml ---
        content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

        # --- Write ZIP ---
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', content_types)
            zf.writestr('_rels/.rels', rels_xml)
            zf.writestr('xl/workbook.xml', workbook_xml)
            zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
            zf.writestr('xl/sharedStrings.xml', sst_xml)
            zf.writestr('xl/styles.xml', styles_xml)
            zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)

    # ========== 实况通报表单 ==========
    def build_live_report_form(self):
        self.station_combos = []

        f_title = ttk.LabelFrame(self.live_scroll_frame, text="通报标题", padding=8)
        f_title.pack(fill="x", pady=(0, 8))
        row = ttk.Frame(f_title)
        row.pack(anchor="w")
        ttk.Label(row, text="【").pack(side="left")
        self.entry_title1 = ttk.Entry(row, width=9)
        self.entry_title1.pack(side="left", padx=2)
        ttk.Label(row, text=":").pack(side="left")
        self.entry_title2 = ttk.Entry(row, width=9)
        self.entry_title2.pack(side="left", padx=2)
        ttk.Label(row, text="实况通报】").pack(side="left")
        ttk.Label(row, text="（例如：20260604 或 14:30）", font=("微软雅黑", 9), foreground="gray").pack(side="left", padx=5)

        f1 = ttk.LabelFrame(self.live_scroll_frame, text="近1小时实况", padding=8)
        f1.pack(fill="x", pady=(0, 8))
        self._build_1h(f1)

        f3 = ttk.LabelFrame(self.live_scroll_frame, text="近3小时实况", padding=8)
        f3.pack(fill="x", pady=(0, 8))
        self._build_3h(f3)

        ff = ttk.LabelFrame(self.live_scroll_frame, text="临近预报", padding=8)
        ff.pack(fill="x", pady=(0, 8))
        self._build_forecast(ff)

        # ---- 保存栏（嵌入标签页底部） ----
        sep = ttk.Separator(self.live_scroll_frame, orient="horizontal")
        sep.pack(fill="x", pady=(15, 5))
        bar = ttk.Frame(self.live_scroll_frame, padding=5)
        bar.pack(fill="x")
        ttk.Label(bar, text="保存目录：").pack(side="left")
        self.entry_save_live2 = ttk.Entry(bar, textvariable=self.live_save_dir, width=32)
        self.entry_save_live2.pack(side="left", padx=5)
        ttk.Button(bar, text="浏览...", command=self.browse_folder_live).pack(side="left", padx=3)
        ttk.Button(bar, text="✨ 生成 TXT 文件", command=self.generate_live_report,
                   style="Accent.TButton").pack(side="right", padx=5)

    # ========== 预通报表单 ==========
    def build_pre_report_form(self):
        # --- 通报标题 ---
        f_title = ttk.LabelFrame(self.pre_scroll_frame, text="通报标题", padding=8)
        f_title.pack(fill="x", pady=(0, 8))
        row = ttk.Frame(f_title)
        row.pack(anchor="w")
        ttk.Label(row, text="【").pack(side="left")
        self.pre_title1 = ttk.Entry(row, width=9)
        self.pre_title1.pack(side="left", padx=2)
        ttk.Label(row, text=":").pack(side="left")
        self.pre_title2 = ttk.Entry(row, width=9)
        self.pre_title2.pack(side="left", padx=2)
        ttk.Label(row, text="强天气预通报】").pack(side="left")
        ttk.Label(row, text="（例如：20260604 或 14:30）", font=("微软雅黑", 9), foreground="gray").pack(side="left", padx=5)

        # --- 影响系统与实况 ---
        f_sys = ttk.LabelFrame(self.pre_scroll_frame, text="影响系统与实况", padding=8)
        f_sys.pack(fill="x", pady=(0, 8))

        r1 = ttk.Frame(f_sys)
        r1.pack(anchor="w", pady=3)
        ttk.Label(r1, text="受").pack(side="left")
        self.pre_system = ttk.Combobox(r1, values=["强对流云团", "冷空气", "高空槽东移", "台风"], width=14, state="normal")
        self.pre_system.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_system, ["强对流云团", "冷空气", "高空槽东移", "台风"])
        ttk.Label(r1, text="影响，目前").pack(side="left")

        r2 = ttk.Frame(f_sys)
        r2.pack(anchor="w", pady=3)
        self.pre_area = ttk.Entry(r2, width=20)
        self.pre_area.pack(side="left", padx=2)
        ttk.Label(r2, text="等周边地区出现").pack(side="left")
        self.pre_weather = ttk.Combobox(r2, values=["强对流", "暴雨", "大风"], width=10, state="normal")
        self.pre_weather.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_weather, ["强对流", "暴雨", "大风"])
        ttk.Label(r2, text="天气，").pack(side="left")

        # --- 预警信息（三行独立） ---
        warn_values = ["暴雨黄色", "暴雨橙色", "暴雨红色"]
        org_values = ["平阳", "苍南", "泰顺", "文成"]

        r3a = ttk.Frame(f_sys)
        r3a.pack(anchor="w", pady=3)
        self.pre_org1 = ttk.Combobox(r3a, values=org_values, width=10, state="normal")
        self.pre_org1.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_org1, org_values)
        ttk.Label(r3a, text="已发布").pack(side="left")
        self.pre_warn1 = ttk.Combobox(r3a, values=warn_values, width=10, state="normal")
        self.pre_warn1.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_warn1, warn_values)
        ttk.Label(r3a, text="预警，").pack(side="left")

        r3b = ttk.Frame(f_sys)
        r3b.pack(anchor="w", pady=3)
        self.pre_org2 = ttk.Combobox(r3b, values=org_values, width=10, state="normal")
        self.pre_org2.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_org2, org_values)
        ttk.Label(r3b, text="发布").pack(side="left")
        self.pre_warn2 = ttk.Combobox(r3b, values=warn_values, width=10, state="normal")
        self.pre_warn2.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_warn2, warn_values)
        ttk.Label(r3b, text="预警，").pack(side="left")

        r3c = ttk.Frame(f_sys)
        r3c.pack(anchor="w", pady=3)
        self.pre_org3 = ttk.Combobox(r3c, values=org_values, width=10, state="normal")
        self.pre_org3.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_org3, org_values)
        ttk.Label(r3c, text="发布").pack(side="left")
        self.pre_warn3 = ttk.Combobox(r3c, values=warn_values, width=10, state="normal")
        self.pre_warn3.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_warn3, warn_values)
        ttk.Label(r3c, text="预警。").pack(side="left")

        # --- 预报与提醒 ---
        f_fc = ttk.LabelFrame(self.pre_scroll_frame, text="预报与提醒", padding=8)
        f_fc.pack(fill="x", pady=(0, 8))

        r4 = ttk.Frame(f_fc)
        r4.pack(anchor="w", pady=3)
        ttk.Label(r4, text="预计").pack(side="left")
        self.pre_forecast = ttk.Entry(r4, width=30)
        self.pre_forecast.pack(side="left", padx=2)

        r5 = ttk.Frame(f_fc)
        r5.pack(anchor="w", pady=3)
        ttk.Label(r5, text="有雷雨时可伴有短时强降水、雷电、").pack(side="left")
        self.pre_wind = ttk.Combobox(r5, values=["7-9级雷暴大风", "8-10级雷暴大风"], width=16, state="normal")
        self.pre_wind.pack(side="left", padx=2)
        self.setup_autocomplete(self.pre_wind, ["7-9级雷暴大风", "8-10级雷暴大风"])
        ttk.Label(r5, text="、冰雹等强对流天气，请各有关方面注意加强防范。").pack(side="left")

        # ---- 保存栏（嵌入标签页底部） ----
        sep = ttk.Separator(self.pre_scroll_frame, orient="horizontal")
        sep.pack(fill="x", pady=(15, 5))
        bar = ttk.Frame(self.pre_scroll_frame, padding=5)
        bar.pack(fill="x")
        ttk.Label(bar, text="保存目录：").pack(side="left")
        self.entry_save_pre2 = ttk.Entry(bar, textvariable=self.pre_save_dir, width=32)
        self.entry_save_pre2.pack(side="left", padx=5)
        ttk.Button(bar, text="浏览...", command=self.browse_folder_pre).pack(side="left", padx=3)
        ttk.Button(bar, text="✨ 生成 TXT 文件", command=self.generate_pre_report,
                   style="Accent.TButton").pack(side="right", padx=5)

    # ========== 天气提醒表单 ==========
    def build_weather_alert_form(self):
        """构建天气提醒表单"""
        weather_types = [
            "有阵雨或雷雨天气",
            "多阵雨天气",
            "有台风影响",
            "有冷空气影响",
        ]

        # --- 通报标题 ---
        f_title = ttk.LabelFrame(self.weather_scroll_frame, text="提醒标题", padding=8)
        f_title.pack(fill="x", pady=(0, 8))
        row_title = ttk.Frame(f_title)
        row_title.pack(anchor="w")
        ttk.Label(row_title, text="【").pack(side="left")
        self.weather_day1 = ttk.Entry(row_title, width=6)
        self.weather_day1.pack(side="left", padx=2)
        ttk.Label(row_title, text="日-").pack(side="left")
        self.weather_day2 = ttk.Entry(row_title, width=6)
        self.weather_day2.pack(side="left", padx=2)
        ttk.Label(row_title, text="日龙港天气提醒】").pack(side="left")
        ttk.Label(row_title, text="（例如：12-13）", font=("微软雅黑", 9), foreground="gray").pack(side="left", padx=5)

        # --- 预报内容 ---
        f_content = ttk.LabelFrame(self.weather_scroll_frame, text="预报内容", padding=8)
        f_content.pack(fill="x", pady=(0, 8))

        r1 = ttk.Frame(f_content)
        r1.pack(anchor="w", pady=4)
        ttk.Label(r1, text="预计未来").pack(side="left")
        self.weather_fc_day1 = ttk.Entry(r1, width=6)
        self.weather_fc_day1.pack(side="left", padx=2)
        ttk.Label(r1, text="-").pack(side="left")
        self.weather_fc_day2 = ttk.Entry(r1, width=6)
        self.weather_fc_day2.pack(side="left", padx=2)
        ttk.Label(r1, text="天").pack(side="left")

        r2 = ttk.Frame(f_content)
        r2.pack(anchor="w", pady=4)
        ttk.Label(r2, text="龙港市").pack(side="left")
        self.weather_type = ttk.Combobox(r2, values=weather_types, width=20, state="normal")
        self.weather_type.pack(side="left", padx=2)
        self.setup_autocomplete(self.weather_type, weather_types)
        ttk.Label(r2, text="。").pack(side="left")

        # --- 降水 ---
        f_rain = ttk.LabelFrame(self.weather_scroll_frame, text="降水", padding=8)
        f_rain.pack(fill="x", pady=(0, 8))
        rain_frame = ttk.Frame(f_rain)
        rain_frame.pack(fill="x", pady=4)
        ttk.Label(rain_frame, text="【降水】").pack(anchor="nw")
        rain_text_frame = ttk.Frame(f_rain)
        rain_text_frame.pack(fill="both", expand=True)
        self.weather_rain = tk.Text(rain_text_frame, font=("微软雅黑", 12), wrap=tk.WORD,
                                     width=60, height=4, relief="solid", borderwidth=1,
                                     padx=5, pady=5)
        rain_scroll = ttk.Scrollbar(rain_text_frame, orient="vertical", command=self.weather_rain.yview)
        self.weather_rain.configure(yscrollcommand=rain_scroll.set)
        self.weather_rain.pack(side="left", fill="both", expand=True)
        rain_scroll.pack(side="right", fill="y")

        # --- 重点关注 ---
        f_focus = ttk.LabelFrame(self.weather_scroll_frame, text="重点关注", padding=8)
        f_focus.pack(fill="x", pady=(0, 8))
        focus_frame = ttk.Frame(f_focus)
        focus_frame.pack(fill="x", pady=4)
        ttk.Label(focus_frame, text="【重点关注】").pack(anchor="nw")
        focus_text_frame = ttk.Frame(f_focus)
        focus_text_frame.pack(fill="both", expand=True)
        self.weather_focus = tk.Text(focus_text_frame, font=("微软雅黑", 12), wrap=tk.WORD,
                                      width=60, height=6, relief="solid", borderwidth=1,
                                      padx=5, pady=5)
        focus_scroll = ttk.Scrollbar(focus_text_frame, orient="vertical", command=self.weather_focus.yview)
        self.weather_focus.configure(yscrollcommand=focus_scroll.set)
        self.weather_focus.pack(side="left", fill="both", expand=True)
        focus_scroll.pack(side="right", fill="y")

        # ---- 保存栏（嵌入标签页底部） ----
        sep = ttk.Separator(self.weather_scroll_frame, orient="horizontal")
        sep.pack(fill="x", pady=(15, 5))
        bar = ttk.Frame(self.weather_scroll_frame, padding=5)
        bar.pack(fill="x")
        ttk.Label(bar, text="保存目录：").pack(side="left")
        self.entry_save_weather2 = ttk.Entry(bar, textvariable=self.weather_save_dir, width=32)
        self.entry_save_weather2.pack(side="left", padx=5)
        ttk.Button(bar, text="浏览...", command=self.browse_folder_weather).pack(side="left", padx=3)
        ttk.Button(bar, text="✨ 生成 TXT 文件", command=self.generate_weather_alert,
                   style="Accent.TButton").pack(side="right", padx=5)

    # ========== 气象灾害风险提示单表单 ==========
    def build_risk_alert_form(self):
        """构建气象灾害风险提示单表单"""
        now = datetime.now()

        # --- 标题 ---
        f_title = ttk.LabelFrame(self.risk_scroll_frame, text="提示单标题", padding=8)
        f_title.pack(fill="x", pady=(0, 8))

        title_label = ttk.Label(f_title, text="气象灾害风险提示单",
                                font=("微软雅黑", 18, "bold"), anchor="center")
        title_label.pack(fill="x")

        # 期数行
        num_row = ttk.Frame(f_title)
        num_row.pack(anchor="center", pady=5)
        ttk.Label(num_row, text=str(now.year), font=("微软雅黑", 13, "bold")).pack(side="left")
        ttk.Label(num_row, text="年第", font=("微软雅黑", 12)).pack(side="left")
        self.risk_issue_num = ttk.Entry(num_row, width=6, font=("微软雅黑", 12))
        self.risk_issue_num.pack(side="left", padx=2)
        self.risk_issue_num.insert(0, str(self.risk_issue_number))
        ttk.Label(num_row, text="期", font=("微软雅黑", 12)).pack(side="left")

        # 发布单位 + 发布时间行
        pub_row = ttk.Frame(f_title)
        pub_row.pack(anchor="w", pady=3)
        ttk.Label(pub_row, text="发布单位：", font=("微软雅黑", 12)).pack(side="left")
        self.risk_pub_unit = ttk.Entry(pub_row, width=22, font=("微软雅黑", 12))
        self.risk_pub_unit.pack(side="left", padx=2)
        self.risk_pub_unit.insert(0, "龙港市气象局")

        time_str = now.strftime("%Y年%m月%d日%H时")
        ttk.Label(pub_row, text="    发布时间：", font=("微软雅黑", 12)).pack(side="left")
        self.risk_pub_time_label = ttk.Label(pub_row, text=time_str, font=("微软雅黑", 12, "bold"))
        self.risk_pub_time_label.pack(side="left", padx=2)

        # --- 正文 ---
        f_body = ttk.LabelFrame(self.risk_scroll_frame, text="预报正文", padding=8)
        f_body.pack(fill="x", pady=(0, 8))
        body_label_row = ttk.Frame(f_body)
        body_label_row.pack(anchor="w", pady=(0, 3))
        ttk.Label(body_label_row, text="预计", font=("微软雅黑", 12)).pack(side="left")
        body_text_frame = ttk.Frame(f_body)
        body_text_frame.pack(fill="both", expand=True)
        self.risk_body_text = tk.Text(body_text_frame, font=("微软雅黑", 12), wrap=tk.WORD,
                                       width=70, height=5, relief="solid", borderwidth=1,
                                       padx=8, pady=5)
        body_scroll = ttk.Scrollbar(body_text_frame, orient="vertical", command=self.risk_body_text.yview)
        self.risk_body_text.configure(yscrollcommand=body_scroll.set)
        self.risk_body_text.pack(side="left", fill="both", expand=True)
        body_scroll.pack(side="right", fill="y")

        # --- 表1：气象灾害风险预警信息 ---
        f_t1 = ttk.LabelFrame(self.risk_scroll_frame, text="表1：气象灾害风险预警信息", padding=8)
        f_t1.pack(fill="x", pady=(0, 8))

        # 日期
        date_row = ttk.Frame(f_t1)
        date_row.pack(anchor="w", pady=3)
        ttk.Label(date_row, text="日期：", font=("微软雅黑", 11, "bold")).pack(side="left")
        self.risk_t1_month = ttk.Combobox(date_row, values=[f"{m}" for m in range(1, 13)],
                                           width=4, state="normal", font=("微软雅黑", 11))
        self.risk_t1_month.set(str(now.month))
        self.risk_t1_month.pack(side="left")
        ttk.Label(date_row, text="月", font=("微软雅黑", 11)).pack(side="left")
        self.risk_t1_day = ttk.Combobox(date_row, values=[f"{d}" for d in range(1, 32)],
                                         width=4, state="normal", font=("微软雅黑", 11))
        self.risk_t1_day.set(str(now.day))
        self.risk_t1_day.pack(side="left")
        ttk.Label(date_row, text="日", font=("微软雅黑", 11)).pack(side="left")

        # 预警风险可能性
        warn_label_row = ttk.Frame(f_t1)
        warn_label_row.pack(anchor="w", pady=(6, 2))
        ttk.Label(warn_label_row, text="预警风险可能性：", font=("微软雅黑", 11, "bold")).pack(side="left")

        self._warn_options = [
            "雷电黄色", "雷电橙色", "雷电红色",
            "暴雨黄色", "暴雨橙色", "暴雨红色",
            "大风黄色", "大风橙色", "大风红色",
            "雷暴大风黄色", "雷暴大风橙色", "雷暴大风红色",
            "冰雹黄色", "冰雹橙色", "冰雹红色",
        ]

        high_row = ttk.Frame(f_t1)
        high_row.pack(anchor="w", pady=2, padx=(20, 0))
        ttk.Label(high_row, text="高：", width=4, font=("微软雅黑", 11)).pack(side="left")
        self.risk_t1_high = ttk.Entry(high_row, width=45, font=("微软雅黑", 11))
        self.risk_t1_high.pack(side="left")
        ttk.Button(high_row, text="选择...",
                   command=lambda: self._show_multi_select_popup(self.risk_t1_high, self._warn_options, "选择预警类型")).pack(side="left", padx=5)

        mid_row = ttk.Frame(f_t1)
        mid_row.pack(anchor="w", pady=2, padx=(20, 0))
        ttk.Label(mid_row, text="中：", width=4, font=("微软雅黑", 11)).pack(side="left")
        self.risk_t1_mid = ttk.Entry(mid_row, width=45, font=("微软雅黑", 11))
        self.risk_t1_mid.pack(side="left")
        ttk.Button(mid_row, text="选择...",
                   command=lambda: self._show_multi_select_popup(self.risk_t1_mid, self._warn_options, "选择预警类型")).pack(side="left", padx=5)

        low_row = ttk.Frame(f_t1)
        low_row.pack(anchor="w", pady=2, padx=(20, 0))
        ttk.Label(low_row, text="低：", width=4, font=("微软雅黑", 11)).pack(side="left")
        self.risk_t1_low = ttk.Entry(low_row, width=45, font=("微软雅黑", 11))
        self.risk_t1_low.pack(side="left")
        ttk.Button(low_row, text="选择...",
                   command=lambda: self._show_multi_select_popup(self.risk_t1_low, self._warn_options, "选择预警类型")).pack(side="left", padx=5)

        # 影响时段
        period_row = ttk.Frame(f_t1)
        period_row.pack(anchor="w", pady=3)
        ttk.Label(period_row, text="影响时段：", font=("微软雅黑", 11, "bold")).pack(side="left")
        self.risk_t1_period = ttk.Entry(period_row, width=30, font=("微软雅黑", 11))
        self.risk_t1_period.pack(side="left", padx=2)

        # 关注区域
        area_row = ttk.Frame(f_t1)
        area_row.pack(anchor="w", pady=3)
        ttk.Label(area_row, text="关注区域：", font=("微软雅黑", 11, "bold")).pack(side="left")
        self.risk_t1_area = ttk.Entry(area_row, width=30, font=("微软雅黑", 11))
        self.risk_t1_area.pack(side="left", padx=2)
        self.risk_t1_area.insert(0, "全市")

        # 可能出现灾害种类
        self._disaster_options = ["雷电灾害", "地质灾害", "城市积涝", "泥石流", "滑坡"]
        disaster_row = ttk.Frame(f_t1)
        disaster_row.pack(anchor="w", pady=3)
        ttk.Label(disaster_row, text="可能出现灾害种类：", font=("微软雅黑", 11, "bold")).pack(side="left")
        self.risk_t1_disaster = ttk.Entry(disaster_row, width=40, font=("微软雅黑", 11))
        self.risk_t1_disaster.pack(side="left", padx=2)
        ttk.Button(disaster_row, text="选择...",
                   command=lambda: self._show_multi_select_popup(self.risk_t1_disaster, self._disaster_options, "选择灾害种类")).pack(side="left", padx=5)

        # 天气影响
        impact_row = ttk.Frame(f_t1)
        impact_row.pack(anchor="w", pady=3)
        ttk.Label(impact_row, text="天气影响：", font=("微软雅黑", 11, "bold")).pack(side="left")
        self.risk_t1_weather_impact = ttk.Entry(impact_row, width=40, font=("微软雅黑", 11))
        self.risk_t1_weather_impact.pack(side="left", padx=2)
        self.risk_t1_weather_impact.insert(0, "强对流天气")

        # --- 表2：更长时间强天气风险提示 ---
        f_t2 = ttk.LabelFrame(self.risk_scroll_frame, text="表2：更长时间强天气风险提示", padding=8)
        f_t2.pack(fill="x", pady=(0, 8))

        self.risk_t2_container = ttk.Frame(f_t2)
        self.risk_t2_container.pack(fill="x")

        btn_row2 = ttk.Frame(f_t2)
        btn_row2.pack(anchor="w", pady=(5, 0))
        ttk.Button(btn_row2, text="➕ 添加行", command=self._add_risk_t2_row).pack(side="left", padx=3)
        ttk.Button(btn_row2, text="🗑 删除行", command=self._delete_risk_t2_row).pack(side="left", padx=3)

        self.risk_t2_rows = []  # list of dicts: {month, day, disaster, area, disaster_type, cause, frame}
        self._add_risk_t2_row()

        # --- 落款 ---
        f_sig = ttk.LabelFrame(self.risk_scroll_frame, text="落款", padding=8)
        f_sig.pack(fill="x", pady=(0, 8))
        sig_row = ttk.Frame(f_sig)
        sig_row.pack(anchor="w", pady=5)
        ttk.Label(sig_row, text="制作：", font=("微软雅黑", 12)).pack(side="left")
        self.risk_maker = ttk.Entry(sig_row, width=12, font=("微软雅黑", 12))
        self.risk_maker.pack(side="left", padx=2)
        ttk.Label(sig_row, text="    审核：", font=("微软雅黑", 12)).pack(side="left")
        self.risk_reviewer = ttk.Entry(sig_row, width=12, font=("微软雅黑", 12))
        self.risk_reviewer.pack(side="left", padx=2)
        ttk.Label(sig_row, text="    签发：", font=("微软雅黑", 12)).pack(side="left")
        self.risk_approver = ttk.Entry(sig_row, width=12, font=("微软雅黑", 12))
        self.risk_approver.pack(side="left", padx=2)

        # ---- 保存栏（嵌入标签页底部） ----
        sep = ttk.Separator(self.risk_scroll_frame, orient="horizontal")
        sep.pack(fill="x", pady=(15, 5))
        bar = ttk.Frame(self.risk_scroll_frame, padding=5)
        bar.pack(fill="x")
        ttk.Label(bar, text="保存目录：").pack(side="left")
        self.entry_save_risk2 = ttk.Entry(bar, textvariable=self.risk_save_dir, width=32)
        self.entry_save_risk2.pack(side="left", padx=5)
        ttk.Button(bar, text="浏览...", command=self.browse_folder_risk).pack(side="left", padx=3)
        ttk.Button(bar, text="📄 生成 Word 文档", command=self.generate_risk_alert,
                   style="Accent.TButton").pack(side="right", padx=5)

    # ========== 表2 行管理 ==========
    def _add_risk_t2_row(self):
        """添加一行到表2"""
        now = datetime.now()
        row_frame = ttk.Frame(self.risk_t2_container)
        row_frame.pack(anchor="w", pady=3, fill="x")

        ttk.Label(row_frame, text="日期：", font=("微软雅黑", 11)).pack(side="left")
        month_cb = ttk.Combobox(row_frame, values=[f"{m}" for m in range(1, 13)],
                                 width=4, state="normal", font=("微软雅黑", 11))
        month_cb.set(str(now.month))
        month_cb.pack(side="left")
        ttk.Label(row_frame, text="月", font=("微软雅黑", 11)).pack(side="left")
        day_cb = ttk.Combobox(row_frame, values=[f"{d}" for d in range(1, 32)],
                               width=4, state="normal", font=("微软雅黑", 11))
        day_cb.set(str(now.day))
        day_cb.pack(side="left")
        ttk.Label(row_frame, text="日", font=("微软雅黑", 11)).pack(side="left")

        ttk.Label(row_frame, text="  强天气灾害：", font=("微软雅黑", 11)).pack(side="left")
        disaster_entry = ttk.Entry(row_frame, width=12, font=("微软雅黑", 11))
        disaster_entry.pack(side="left", padx=2)

        ttk.Label(row_frame, text="  关注区域：", font=("微软雅黑", 11)).pack(side="left")
        area_entry = ttk.Entry(row_frame, width=10, font=("微软雅黑", 11))
        area_entry.pack(side="left", padx=2)

        ttk.Label(row_frame, text="  可能出现灾害种类：", font=("微软雅黑", 11)).pack(side="left")
        disaster_type_entry = ttk.Entry(row_frame, width=16, font=("微软雅黑", 11))
        disaster_type_entry.pack(side="left", padx=2)
        ttk.Button(row_frame, text="选择",
                   command=lambda e=disaster_type_entry: self._show_multi_select_popup(e, self._disaster_options, "选择灾害种类")).pack(side="left", padx=3)

        ttk.Label(row_frame, text="  致灾原因：", font=("微软雅黑", 11)).pack(side="left")
        cause_entry = ttk.Entry(row_frame, width=10, font=("微软雅黑", 11))
        cause_entry.pack(side="left", padx=2)

        row_data = {
            "frame": row_frame,
            "month": month_cb,
            "day": day_cb,
            "disaster": disaster_entry,
            "area": area_entry,
            "disaster_type": disaster_type_entry,
            "cause": cause_entry,
        }
        self.risk_t2_rows.append(row_data)

    def _delete_risk_t2_row(self):
        """删除表2最后一行"""
        if not self.risk_t2_rows:
            messagebox.showwarning("提示", "没有可删除的行")
            return
        row = self.risk_t2_rows.pop()
        row["frame"].destroy()

    def _show_multi_select_popup(self, target_entry, options, title="选择"):
        """通用多选弹窗：将选中项以'、'分隔写回 target_entry"""
        # 解析当前值
        current = set(s.strip() for s in target_entry.get().split("、") if s.strip())

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg="#3498db")
        inner = tk.Frame(popup, bg="white", padx=10, pady=10)
        inner.pack(padx=1, pady=1)

        ttk.Label(inner, text=title, font=("微软雅黑", 12, "bold")).pack(anchor="w", pady=(0, 8))

        vars_ = {}
        for opt in options:
            v = tk.BooleanVar(value=opt in current)
            cb = tk.Checkbutton(inner, text=opt, variable=v,
                               font=("微软雅黑", 11),
                               bg="white", anchor="w",
                               selectcolor="white",
                               activebackground="white")
            cb.pack(anchor="w", fill="x")
            vars_[opt] = v

        def _confirm():
            result = "、".join(opt for opt in options if vars_[opt].get())
            target_entry.delete(0, tk.END)
            target_entry.insert(0, result)
            popup.destroy()
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        btn_frame = ttk.Frame(inner)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_frame, text="确定", command=_confirm).pack(side="right", padx=3)

        def _cancel():
            popup.destroy()
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass

        ttk.Button(btn_frame, text="取消", command=_cancel).pack(side="right", padx=3)

        # 收集弹窗内所有子 widget
        popup_widgets = {popup, inner, btn_frame}
        def _collect(w):
            popup_widgets.add(w)
            for child in w.winfo_children():
                _collect(child)
        _collect(inner)

        def _on_root_click(event):
            if not popup.winfo_exists() or not popup.winfo_viewable():
                return
            w = event.widget
            while w is not None:
                if w in popup_widgets:
                    return
                try:
                    w = w.master
                except Exception:
                    break
            _confirm()

        _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')

        def _on_popup_destroy(event):
            try:
                self.root.unbind('<Button-1>', _bind_id)
            except Exception:
                pass
        popup.bind('<Destroy>', _on_popup_destroy)

        popup.update_idletasks()
        # 定位到触发按钮附近
        px = target_entry.winfo_rootx()
        py = target_entry.winfo_rooty() + target_entry.winfo_height()
        popup.geometry(f"+{px}+{py}")
        popup.focus_set()

    def _build_1h(self, parent):
        r1 = ttk.Frame(parent)
        r1.pack(anchor="w", pady=4)
        ttk.Label(r1, text="受").pack(side="left")
        self.cb_1h_system = ttk.Combobox(r1, values=["强对流云团", "冷空气", "高空槽东移", "台风"], width=15, state="normal")
        self.cb_1h_system.pack(side="left", padx=2)
        self.setup_autocomplete(self.cb_1h_system,
                                ["强对流云团", "冷空气", "高空槽东移", "台风"])
        self.station_combos.append(self.cb_1h_system)
        ttk.Label(r1, text="影响，近1小时，龙港市出现").pack(side="left")
        self.cb_1h_weather = ttk.Combobox(r1, values=["强降水", "雷电", "雷暴大风", "8级以上大风"], width=15, state="normal")
        self.cb_1h_weather.pack(side="left", padx=2)
        self.setup_autocomplete(self.cb_1h_weather,
                                ["强降水", "雷电", "雷暴大风", "8级以上大风"])
        self.station_combos.append(self.cb_1h_weather)
        ttk.Label(r1, text="，全市面雨量").pack(side="left")
        self.entry_1h_avg = ttk.Entry(r1, width=7)
        self.entry_1h_avg.pack(side="left", padx=2)
        ttk.Label(r1, text="毫米；").pack(side="left")

        r2 = ttk.Frame(parent)
        r2.pack(anchor="w", pady=4)
        ttk.Label(r2, text="单站总雨量前三分别为").pack(side="left")

        self.entries_1h_stations = []
        for i in range(3):
            row_sta = ttk.Frame(parent)
            row_sta.pack(anchor="w", pady=4)
            cb_sta = ttk.Combobox(row_sta, values=self.stations, width=14, state="normal")
            cb_sta.pack(side="left", padx=2)
            self.setup_autocomplete(cb_sta)
            self.station_combos.append(cb_sta)
            val_entry = ttk.Entry(row_sta, width=7)
            val_entry.pack(side="left", padx=2)
            ttk.Label(row_sta, text="毫米").pack(side="left")
            if i < 2:
                ttk.Label(row_sta, text="、").pack(side="left", padx=2)
            else:
                ttk.Label(row_sta, text="。").pack(side="left")
            self.entries_1h_stations.append((cb_sta, val_entry))

    def _build_3h(self, parent):
        r1 = ttk.Frame(parent)
        r1.pack(anchor="w", pady=4)
        ttk.Label(r1, text="近3小时，全市面雨量").pack(side="left")
        self.entry_3h_avg = ttk.Entry(r1, width=7)
        self.entry_3h_avg.pack(side="left", padx=2)
        ttk.Label(r1, text="毫米；").pack(side="left")

        r2 = ttk.Frame(parent)
        r2.pack(anchor="w", pady=4)
        ttk.Label(r2, text="单站总雨量前三分别为").pack(side="left")

        self.entries_3h_stations = []
        for i in range(3):
            row_sta = ttk.Frame(parent)
            row_sta.pack(anchor="w", pady=4)
            cb_sta = ttk.Combobox(row_sta, values=self.stations, width=14, state="normal")
            cb_sta.pack(side="left", padx=2)
            self.setup_autocomplete(cb_sta)
            self.station_combos.append(cb_sta)
            val_entry = ttk.Entry(row_sta, width=7)
            val_entry.pack(side="left", padx=2)
            ttk.Label(row_sta, text="毫米").pack(side="left")
            if i < 2:
                ttk.Label(row_sta, text="、").pack(side="left", padx=2)
            else:
                ttk.Label(row_sta, text="。").pack(side="left")
            self.entries_3h_stations.append((cb_sta, val_entry))

        r3 = ttk.Frame(parent)
        r3.pack(anchor="w", pady=4)
        ttk.Label(r3, text="单站1小时雨量最大出现在").pack(side="left")
        self.cb_3h_max_sta = ttk.Combobox(r3, values=self.stations, width=14, state="normal")
        self.cb_3h_max_sta.pack(side="left", padx=2)
        self.setup_autocomplete(self.cb_3h_max_sta)
        self.station_combos.append(self.cb_3h_max_sta)
        self.entry_3h_max_val = ttk.Entry(r3, width=7)
        self.entry_3h_max_val.pack(side="left", padx=2)
        ttk.Label(r3, text="毫米。").pack(side="left")

        r4 = ttk.Frame(parent)
        r4.pack(anchor="w", pady=4)
        ttk.Label(r4, text="风力最大出现在").pack(side="left")
        self.cb_3h_wind_sta = ttk.Combobox(r4, values=self.stations, width=14, state="normal")
        self.cb_3h_wind_sta.pack(side="left", padx=2)
        self.setup_autocomplete(self.cb_3h_wind_sta)
        self.station_combos.append(self.cb_3h_wind_sta)
        self.entry_3h_wind_speed = ttk.Entry(r4, width=7)
        self.entry_3h_wind_speed.pack(side="left", padx=2)
        ttk.Label(r4, text="米/秒（").pack(side="left")
        self.entry_3h_wind_scale = ttk.Entry(r4, width=5)
        self.entry_3h_wind_scale.pack(side="left", padx=2)
        ttk.Label(r4, text="级）。").pack(side="left")

    def _build_forecast(self, parent):
        r1 = ttk.Frame(parent)
        r1.pack(anchor="w", pady=4)
        ttk.Label(r1, text="【临近预报】预计未来").pack(side="left")
        self.entry_fc_time = ttk.Entry(r1, width=13)
        self.entry_fc_time.pack(side="left", padx=2)
        ttk.Label(r1, text="小时，龙港市仍有").pack(side="left")
        self.entry_fc_weather = ttk.Entry(r1, width=13)
        self.entry_fc_weather.pack(side="left", padx=2)
        ttk.Label(r1, text="，雨量").pack(side="left")
        self.entry_fc_rain1 = ttk.Entry(r1, width=7)
        self.entry_fc_rain1.pack(side="left", padx=2)
        ttk.Label(r1, text="-").pack(side="left")
        self.entry_fc_rain2 = ttk.Entry(r1, width=7)
        self.entry_fc_rain2.pack(side="left", padx=2)
        ttk.Label(r1, text="毫米。").pack(side="left")

    # ========== 标签页切换 ==========
    def on_tab_changed(self, event):
        try:
            self._do_tab_changed(event)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("标签页切换错误", f"{e}")

    def _do_tab_changed(self, event):
        tab_id = self.notebook.select()
        tab_text = self.notebook.tab(tab_id, "text")
        # 所有保存栏已嵌入各标签页内部，root 级栏不再使用
        self.right_frame.pack_forget()
        if "实况通报" in tab_text:
            self.right_frame.pack(side="right", fill="y", padx=(10, 0))

    # ========== 预通报 TXT 生成 ==========
    def generate_pre_report(self):
        try:
            title1 = self.pre_title1.get().strip()
            title2 = self.pre_title2.get().strip()
            system = self.pre_system.get().strip()
            area = self.pre_area.get().strip()
            weather = self.pre_weather.get().strip()
            org1 = self.pre_org1.get().strip()
            org2 = self.pre_org2.get().strip()
            org3 = self.pre_org3.get().strip()
            warn1 = self.pre_warn1.get().strip()
            warn2 = self.pre_warn2.get().strip()
            warn3 = self.pre_warn3.get().strip()
            forecast = self.pre_forecast.get().strip()
            wind = self.pre_wind.get().strip()

            head = f"【{title1}:{title2}强天气预通报】"

            # 构建预警部分（为空自动跳过）
            warn_parts = []
            if org1 and warn1:
                warn_parts.append(f"{org1}已发布{warn1}预警")
            if org2 and warn2:
                warn_parts.append(f"{org2}发布{warn2}预警")
            if org3 and warn3:
                warn_parts.append(f"{org3}发布{warn3}预警")

            if warn_parts:
                p1 = (f"受{system}影响，目前{area}等周边地区出现{weather}天气，"
                      f"{'，'.join(warn_parts)}。")
            else:
                p1 = f"受{system}影响，目前{area}等周边地区出现{weather}天气。"

            p2 = (f"预计{forecast}，有雷雨时可伴有短时强降水、雷电、{wind}、"
                  f"冰雹等强对流天气，请各有关方面注意加强防范。")

            full_text = head + p1 + p2
            wrapped = textwrap.fill(full_text, width=70, break_long_words=False, replace_whitespace=False)

            now_str = datetime.now().strftime("%Y%m%d%H%M")
            filename = f"{now_str}预通报.txt"
            save_dir = self.pre_save_dir.get()
            save_path = os.path.join(save_dir, filename)

            if os.path.exists(save_path):
                base, ext = os.path.splitext(filename)
                count = 2
                while True:
                    new_name = f"{base}_{count}{ext}"
                    save_path = os.path.join(save_dir, new_name)
                    if not os.path.exists(save_path):
                        break
                    count += 1

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(wrapped)

            # 尝试打开文件所在目录（网络路径可能失败，不影响主流程）
            try:
                os.startfile(os.path.normpath(os.path.dirname(save_path)))
            except Exception:
                pass

            # 自动填入服务记录表格
            try:
                self.add_service_row(
                    time_val=f"{title1}:{title2}",
                    item_val="强天气预通报",
                    content_val=full_text,
                )
            except Exception:
                pass

            messagebox.showinfo("成功", f"文件已生成：\n{save_path}")

        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    def browse_folder_pre(self):
        folder = filedialog.askdirectory(title="选择预通报保存目录")
        if folder:
            self.pre_save_dir.set(folder)
            self.save_config()

    def browse_folder_live(self):
        folder = filedialog.askdirectory(title="选择实况通报保存目录")
        if folder:
            self.live_save_dir.set(folder)
            self.save_config()

    def browse_folder_risk(self):
        folder = filedialog.askdirectory(title="选择风险提示单保存目录")
        if folder:
            self.risk_save_dir.set(folder)
            self.save_config()

    def browse_folder_weather(self):
        folder = filedialog.askdirectory(title="选择天气提醒保存目录")
        if folder:
            self.weather_save_dir.set(folder)
            self.save_config()

    # ========== 天气提醒 TXT 生成 ==========
    def generate_weather_alert(self):
        try:
            day1 = self.weather_day1.get().strip()
            day2 = self.weather_day2.get().strip()
            fc_day1 = self.weather_fc_day1.get().strip()
            fc_day2 = self.weather_fc_day2.get().strip()
            weather_type = self.weather_type.get().strip()
            rain = self.weather_rain.get("1.0", "end-1c").strip()
            focus = self.weather_focus.get("1.0", "end-1c").strip()

            head = f"【{day1}日-{day2}日龙港天气提醒】"
            p1 = f"预计未来{fc_day1}-{fc_day2}天龙港市{weather_type}。"
            p2 = f"【降水】{rain}"
            p3 = f"【重点关注】{focus}"

            full_text = head + p1 + "\n" + p2 + "\n" + p3
            wrapped = textwrap.fill(full_text, width=70, break_long_words=False, replace_whitespace=False)

            now_str = datetime.now().strftime("%Y%m%d%H%M")
            filename = f"{now_str}天气提醒.txt"
            save_dir = self.weather_save_dir.get()
            save_path = os.path.join(save_dir, filename)

            if os.path.exists(save_path):
                base, ext = os.path.splitext(filename)
                count = 2
                while True:
                    new_name = f"{base}_{count}{ext}"
                    save_path = os.path.join(save_dir, new_name)
                    if not os.path.exists(save_path):
                        break
                    count += 1

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(wrapped)

            # 尝试打开文件所在目录（网络路径可能失败，不影响主流程）
            try:
                os.startfile(os.path.normpath(os.path.dirname(save_path)))
            except Exception:
                pass

            # 自动填入服务记录表格
            try:
                self.add_service_row(
                    time_val=f"{day1}日-{day2}日",
                    item_val="天气提醒",
                    content_val=full_text,
                )
            except Exception:
                pass

            messagebox.showinfo("成功", f"文件已生成：\n{save_path}")

        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    # ========== 气象灾害风险提示单 Word 生成 ==========
    def _set_cell_text(self, cell, text, font_name='宋体', size=12, bold=False, align='center'):
        """设置表格单元格文本（含东亚字体支持）"""
        # 清除现有段落
        for p in cell.paragraphs:
            for r in p.runs:
                r._element.getparent().remove(r._element)
        para = cell.paragraphs[0]
        if align == 'center':
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = para.add_run(str(text))
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(size)
        run.bold = bold
        # 垂直居中
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        tcPr.append(vAlign)

    def _set_warning_cell(self, cell, text, font_name='宋体_GB2312', size=10.5):
        """设置预警风险可能性单元格——按颜色渲染，每种预警换一行"""
        # 清除现有段落
        for p in cell.paragraphs:
            for r in p.runs:
                r._element.getparent().remove(r._element)

        items = [s.strip() for s in text.split("、") if s.strip()]
        if not items:
            return

        # 颜色映射
        def _get_color(item):
            if "红色" in item:
                return "FF6666"
            elif "橙色" in item:
                return "FFA500"
            elif "黄色" in item:
                return "FFFF00"
            return None

        for i, item in enumerate(items):
            if i == 0:
                para = cell.paragraphs[0]
            else:
                para = cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 段落间距收紧
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(1)

            run = para.add_run(item)
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            run.font.size = Pt(size)
            run.bold = True

            color = _get_color(item)
            if color:
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), color)
                run._element.rPr.append(shd)

        # 垂直居中
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        tcPr.append(vAlign)

    def generate_risk_alert(self):
        try:
            bj_tz = timezone(timedelta(hours=8))
            bj_now = datetime.now(bj_tz)

            # ---- 收集数据 ----
            pub_unit = self.risk_pub_unit.get().strip()
            pub_time = bj_now.strftime("%Y年%m月%d日%H时")
            body_text = self.risk_body_text.get("1.0", "end-1c").strip()
            current_year = str(bj_now.year)

            # 期数自增
            try:
                user_num = int(self.risk_issue_num.get().strip())
            except ValueError:
                user_num = self.risk_issue_number
            issue_num_to_use = user_num
            self.risk_issue_number = user_num + 1
            self.risk_issue_num.delete(0, tk.END)
            self.risk_issue_num.insert(0, str(self.risk_issue_number))
            self.save_config()

            # 表1 数据
            t1_month = self.risk_t1_month.get().strip()
            t1_day = self.risk_t1_day.get().strip()
            t1_date = f"{t1_month}月{t1_day}日" if t1_month and t1_day else ""
            t1_high = self.risk_t1_high.get().strip()
            t1_mid = self.risk_t1_mid.get().strip()
            t1_low = self.risk_t1_low.get().strip()
            t1_period = self.risk_t1_period.get().strip()
            t1_area = self.risk_t1_area.get().strip()
            t1_disaster = self.risk_t1_disaster.get().strip()
            t1_impact = self.risk_t1_weather_impact.get().strip()

            # 表2 数据（过滤全空行）
            t2_data = []
            for row in self.risk_t2_rows:
                r_month = row["month"].get().strip()
                r_day = row["day"].get().strip()
                r_date = f"{r_month}月{r_day}日" if r_month and r_day else ""
                r_disaster = row["disaster"].get().strip()
                r_area = row["area"].get().strip()
                r_type = row["disaster_type"].get().strip()
                r_cause = row["cause"].get().strip()
                if any([r_disaster, r_area, r_type, r_cause]):
                    t2_data.append({
                        "date": r_date, "disaster": r_disaster,
                        "area": r_area, "disaster_type": r_type, "cause": r_cause,
                    })

            # 落款
            maker = self.risk_maker.get().strip()
            reviewer = self.risk_reviewer.get().strip()
            approver = self.risk_approver.get().strip()

            # ---- 构建 DOCX ----
            doc = Document()

            # 页面设置
            section = doc.sections[0]
            section.page_width = Cm(21)
            section.page_height = Cm(29.7)
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(3.18)
            section.right_margin = Cm(3.18)

            STYLE_HEI = '黑体'
            STYLE_SONG = '宋体_GB2312'

            # 标题
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run("气象灾害风险提示单")
            r_title.font.name = STYLE_HEI
            r_title._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE_HEI)
            r_title.font.size = Pt(24)
            p_title.paragraph_format.space_after = Pt(6)

            # 期数
            p_num = doc.add_paragraph()
            p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_num = p_num.add_run(f"{current_year}年第{issue_num_to_use}期")
            r_num.font.name = STYLE_HEI
            r_num._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE_HEI)
            r_num.font.size = Pt(12)
            p_num.paragraph_format.space_after = Pt(12)

            # 发布信息行（带下边框）
            p_info = doc.add_paragraph()
            r_info = p_info.add_run(f"发布单位：{pub_unit}              发布时间：{pub_time}")
            r_info.font.name = STYLE_HEI
            r_info._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE_HEI)
            r_info.font.size = Pt(12)
            pPr_info = p_info._element.get_or_add_pPr()
            pBdr_info = OxmlElement('w:pBdr')
            bottom_info = OxmlElement('w:bottom')
            bottom_info.set(qn('w:val'), 'single')
            bottom_info.set(qn('w:sz'), '12')
            bottom_info.set(qn('w:space'), '4')
            bottom_info.set(qn('w:color'), '000000')
            pBdr_info.append(bottom_info)
            pPr_info.append(pBdr_info)
            p_info.paragraph_format.space_after = Pt(12)

            # 正文
            p_body = doc.add_paragraph()
            full_body = f"预计{body_text}" if body_text else "预计"
            r_body = p_body.add_run(full_body)
            r_body.font.name = STYLE_SONG
            r_body._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE_SONG)
            r_body.font.size = Pt(12)
            r_body.bold = True
            p_body.paragraph_format.first_line_indent = Pt(24)
            p_body.paragraph_format.space_after = Pt(12)

            # ---- 表1 ----
            table1 = doc.add_table(rows=3, cols=8)
            table1.style = 'Table Grid'
            # 设置列宽 (8 列，参考模板比例)
            col_widths_t1 = [Cm(2.0), Cm(1.2), Cm(1.3), Cm(1.4), Cm(1.5), Cm(1.4), Cm(2.4), Cm(2.9)]
            for ci, w in enumerate(col_widths_t1):
                for row in table1.rows:
                    row.cells[ci].width = w

            # 合并单元格
            table1.cell(0, 1).merge(table1.cell(0, 3))  # 预警风险可能性 跨 3 列
            table1.cell(0, 0).merge(table1.cell(1, 0))  # 日期 垂直合并
            table1.cell(0, 4).merge(table1.cell(1, 4))  # 影响时段 垂直合并
            table1.cell(0, 5).merge(table1.cell(1, 5))  # 关注区域 垂直合并
            table1.cell(0, 6).merge(table1.cell(1, 6))  # 可能出现灾害种类 垂直合并
            table1.cell(0, 7).merge(table1.cell(1, 7))  # 天气影响 垂直合并

            # 表头文字
            self._set_cell_text(table1.cell(0, 0), "日期", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(0, 1), "预警风险可能性", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(0, 4), "影响时段", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(0, 5), "关注区域", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(0, 6), "可能出现灾害种类", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(0, 7), "天气影响", STYLE_HEI, 12)

            # 子表头 (高/中/低)
            self._set_cell_text(table1.cell(1, 1), "高", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(1, 2), "中", STYLE_HEI, 12)
            self._set_cell_text(table1.cell(1, 3), "低", STYLE_HEI, 12)

            # 数据行
            self._set_cell_text(table1.cell(2, 0), t1_date, STYLE_SONG, 10.5, bold=True)
            self._set_warning_cell(table1.cell(2, 1), t1_high)
            self._set_warning_cell(table1.cell(2, 2), t1_mid)
            self._set_warning_cell(table1.cell(2, 3), t1_low)
            self._set_cell_text(table1.cell(2, 4), t1_period, STYLE_SONG, 10.5, bold=True)
            self._set_cell_text(table1.cell(2, 5), t1_area, STYLE_SONG, 10.5, bold=True)
            self._set_cell_text(table1.cell(2, 6), t1_disaster, STYLE_SONG, 10.5, bold=True)
            self._set_cell_text(table1.cell(2, 7), t1_impact, STYLE_SONG, 10.5, bold=True)

            # 两张表格之间的空行
            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_before = Pt(6)
            spacer.paragraph_format.space_after = Pt(6)

            # ---- 表2（仅非空时生成） ----
            if t2_data:
                n_data = len(t2_data)
                table2 = doc.add_table(rows=2 + n_data, cols=5)
                table2.style = 'Table Grid'
                col_widths_t2 = [Cm(2.3), Cm(2.7), Cm(2.3), Cm(3.9), Cm(2.7)]
                for ci, w in enumerate(col_widths_t2):
                    for row in table2.rows:
                        row.cells[ci].width = w

                headers_t2 = ["日期", "强天气灾害", "关注区域", "可能出现灾害种类", "致灾原因"]
                for ci in range(5):
                    table2.cell(0, ci).merge(table2.cell(1, ci))
                    self._set_cell_text(table2.cell(0, ci), headers_t2[ci], STYLE_HEI, 12)

                for ri, rd in enumerate(t2_data):
                    vals = [rd["date"], rd["disaster"], rd["area"], rd["disaster_type"], rd["cause"]]
                    for ci, val in enumerate(vals):
                        self._set_cell_text(table2.cell(2 + ri, ci), val, STYLE_SONG, 10.5, bold=True)

            # ---- 落款（上黑线 + 签名 + 下黑线） ----
            # 落款前空行（紧缩间距）
            p_spacer = doc.add_paragraph()
            p_spacer.paragraph_format.space_before = Pt(4)
            p_spacer.paragraph_format.space_after = Pt(0)

            # 上黑线段落（无额外间距）
            p_line1 = doc.add_paragraph()
            p_line1.paragraph_format.space_before = Pt(0)
            p_line1.paragraph_format.space_after = Pt(2)
            pPr_l1 = p_line1._element.get_or_add_pPr()
            pBdr_l1 = OxmlElement('w:pBdr')
            top_l1 = OxmlElement('w:top')
            top_l1.set(qn('w:val'), 'single')
            top_l1.set(qn('w:sz'), '12')
            top_l1.set(qn('w:space'), '1')
            top_l1.set(qn('w:color'), '000000')
            pBdr_l1.append(top_l1)
            pPr_l1.append(pBdr_l1)

            # 签名行（紧贴黑线）
            p_sig = doc.add_paragraph()
            p_sig.paragraph_format.space_before = Pt(2)
            p_sig.paragraph_format.space_after = Pt(2)
            sig_parts = []
            if maker:
                sig_parts.append(f"制作：{maker}")
            if reviewer:
                sig_parts.append(f"审核：{reviewer}")
            if approver:
                sig_parts.append(f"签发：{approver}")
            sig_text = "              ".join(sig_parts) if sig_parts else "制作：              审核：              签发："
            r_sig = p_sig.add_run(sig_text)
            r_sig.font.name = STYLE_HEI
            r_sig._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE_HEI)
            r_sig.font.size = Pt(12)

            # 下黑线段落（无额外间距）
            p_line2 = doc.add_paragraph()
            p_line2.paragraph_format.space_before = Pt(0)
            p_line2.paragraph_format.space_after = Pt(0)
            pPr_l2 = p_line2._element.get_or_add_pPr()
            pBdr_l2 = OxmlElement('w:pBdr')
            bottom_l2 = OxmlElement('w:bottom')
            bottom_l2.set(qn('w:val'), 'single')
            bottom_l2.set(qn('w:sz'), '12')
            bottom_l2.set(qn('w:space'), '1')
            bottom_l2.set(qn('w:color'), '000000')
            pBdr_l2.append(bottom_l2)
            pPr_l2.append(pBdr_l2)

            # ---- 保存文件 ----
            now_str = bj_now.strftime("%Y%m%d%H%M")
            filename = f"气象灾害风险提示单（{current_year}年第{issue_num_to_use}期）.docx"
            save_dir = self.risk_save_dir.get()
            save_path = os.path.join(save_dir, filename)

            if os.path.exists(save_path):
                base, ext = os.path.splitext(filename)
                count = 2
                while True:
                    new_name = f"{base}_{count}{ext}"
                    save_path = os.path.join(save_dir, new_name)
                    if not os.path.exists(save_path):
                        break
                    count += 1

            doc.save(save_path)

            # 打开目录
            try:
                os.startfile(os.path.normpath(os.path.dirname(save_path)))
            except Exception:
                pass

            # 自动填入服务记录
            try:
                self.add_service_row(
                    time_val=f"{current_year}年第{issue_num_to_use}期",
                    item_val="气象灾害风险提示单",
                    content_val=f"已生成：{save_path}",
                )
            except Exception:
                pass

            messagebox.showinfo("成功", f"Word 文档已生成：\n{save_path}")

        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    def generate_live_report(self):
        try:
            title1 = self.entry_title1.get().strip()
            title2 = self.entry_title2.get().strip()
            sys_1h = self.cb_1h_system.get().strip()
            weather_1h = self.cb_1h_weather.get().strip()
            avg_1h = self.entry_1h_avg.get().strip()
            stations_1h = [(e[0].get().strip(), e[1].get().strip()) for e in self.entries_1h_stations]

            avg_3h = self.entry_3h_avg.get().strip()
            stations_3h = [(e[0].get().strip(), e[1].get().strip()) for e in self.entries_3h_stations]
            max_sta_3h = self.cb_3h_max_sta.get().strip()
            max_val_3h = self.entry_3h_max_val.get().strip()
            wind_sta_3h = self.cb_3h_wind_sta.get().strip()
            wind_speed_3h = self.entry_3h_wind_speed.get().strip()
            wind_scale_3h = self.entry_3h_wind_scale.get().strip()

            fc_time = self.entry_fc_time.get().strip()
            fc_weather = self.entry_fc_weather.get().strip()
            fc_rain1 = self.entry_fc_rain1.get().strip()
            fc_rain2 = self.entry_fc_rain2.get().strip()

            head = f"【{title1}:{title2}实况通报】"
            p1 = (f"受{sys_1h}影响，近1小时，龙港市出现{weather_1h}，"
                  f"全市面雨量{avg_1h}毫米；"
                  f"单站总雨量前三分别为{stations_1h[0][0]}{stations_1h[0][1]}毫米、"
                  f"{stations_1h[1][0]}{stations_1h[1][1]}毫米、"
                  f"{stations_1h[2][0]}{stations_1h[2][1]}毫米。")
            part1 = head + p1
            p3 = (f"近3小时，全市面雨量{avg_3h}毫米；"
                  f"单站总雨量前三分别为{stations_3h[0][0]}{stations_3h[0][1]}毫米、"
                  f"{stations_3h[1][0]}{stations_3h[1][1]}毫米、"
                  f"{stations_3h[2][0]}{stations_3h[2][1]}毫米。"
                  f"单站1小时雨量最大出现在{max_sta_3h}{max_val_3h}毫米。"
                  f"风力最大出现在{wind_sta_3h}{wind_speed_3h}米/秒（{wind_scale_3h}级）。")
            p4 = (f"【临近预报】预计未来{fc_time}小时，龙港市仍有{fc_weather}，"
                  f"雨量{fc_rain1}-{fc_rain2}毫米。")

            full_text = part1 + "\n" + p3 + "\n" + p4
            wrapped = textwrap.fill(full_text, width=70, break_long_words=False, replace_whitespace=False)

            now_str = datetime.now().strftime("%Y%m%d%H%M")
            filename = f"{now_str}实况通报.txt"
            save_dir = self.live_save_dir.get()
            save_path = os.path.join(save_dir, filename)

            if os.path.exists(save_path):
                base, ext = os.path.splitext(filename)
                count = 2
                while True:
                    new_name = f"{base}_{count}{ext}"
                    save_path = os.path.join(save_dir, new_name)
                    if not os.path.exists(save_path):
                        break
                    count += 1

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(wrapped)

            # 尝试打开文件所在目录（网络路径可能失败，不影响主流程）
            try:
                os.startfile(os.path.normpath(os.path.dirname(save_path)))
            except Exception:
                pass

            # 自动填入服务记录表格
            try:
                self.add_service_row(
                    time_val=f"{title1}:{title2}",
                    item_val="实况通报",
                    content_val=full_text,
                )
            except Exception:
                pass

            messagebox.showinfo("成功", f"文件已生成：\n{save_path}")

        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
