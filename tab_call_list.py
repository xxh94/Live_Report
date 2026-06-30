"""叫应名单：可编辑表格"""
import os, sys, json, zipfile, re
import xml.etree.ElementTree as ET
from tkinter import ttk, filedialog, messagebox
import tkinter as tk


class CallListMixin:
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
