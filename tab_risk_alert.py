"""气象灾害风险提示单：表单 + DOCX 生成"""
import os
from datetime import datetime, timezone, timedelta
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class RiskAlertMixin:
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
