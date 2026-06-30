"""天气提醒：表单 + TXT 生成"""
import os
import textwrap
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class WeatherAlertMixin:
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
