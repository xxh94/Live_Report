"""实况通报 + 站点面板"""
import os
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class LiveReportMixin:
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
