"""预通报：表单 + TXT 生成"""
import os
import textwrap
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class PreReportMixin:
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
