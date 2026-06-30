"""两直一白：表单 + Excel 追加"""
import os
import shutil
import json
import zipfile
import re
import xml.etree.ElementTree as ET
from datetime import date
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class LiangzhiyibaiMixin:
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
