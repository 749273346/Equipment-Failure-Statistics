import os
import sys
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import tempfile
import shutil
import json
import subprocess
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
import win32com.client
import pythoncom
import openpyxl
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import datetime
from copy import copy

# Configuration
DEFAULT_SOURCE_DIR = r"e:\QC-攻关小组\正在进行项目\设备故障统计\3-设备缺陷问题库及设备缺陷处理记录"
TARGET_EXCEL_PATH = r"e:\QC-攻关小组\正在进行项目\设备故障统计\设备缺陷问题库（日常巡视、故障处理问题库，广供记-002汇总表，202601起）.xlsx"

# Enable High DPI support
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

def _app_state_path():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base_dir = os.getcwd()
    return os.path.join(base_dir, ".app_state.json")

def _load_app_state():
    path = _app_state_path()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}

def _save_app_state(data):
    path = _app_state_path()
    tmp = f"{path}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, path)
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False

class DefectProcessor:
    def __init__(self, log_callback=print, progress_callback=None):
        self.log = log_callback
        self.progress = progress_callback
        self.stop_requested = False

    def _safe_temp_name(self, name):
        s = str(name or "")
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '：']:
            s = s.replace(ch, "_")
        s = s.strip()
        return s or "word.doc"

    def _open_word_doc(self, word, file_path):
        open_kwargs = dict(
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
            Visible=False,
            OpenAndRepair=True,
        )
        return word.Documents.Open(file_path, **open_kwargs)

    def _load_processed_paths_from_excel(self, target_excel):
        paths = set()
        try:
            wb = openpyxl.load_workbook(target_excel, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=4, min_col=14, max_col=14, values_only=True):
                v = row[0] if row else None
                if not isinstance(v, str):
                    continue
                s = v.strip()
                if not s:
                    continue
                paths.add(os.path.normcase(os.path.normpath(s)))
            try:
                wb.close()
            except Exception:
                pass
        except Exception:
            return set()
        return paths

    def _row_has_content(self, row_data):
        if not row_data or len(row_data) <= 1:
            return False
        for cell in row_data[1:]:
            if str(cell).strip():
                return True
        return False

    def _coerce_int(self, value, default=0):
        try:
            if value is None:
                return default
            s = str(value).strip()
            if s == "":
                return default
            return int(float(s))
        except Exception:
            return default

    def _find_last_valid_row(self, ws, min_row=3, serial_col=1, max_cols=13):
        for row in range(ws.max_row, min_row - 1, -1):
            serial_val = ws.cell(row=row, column=serial_col).value
            if serial_val is None or str(serial_val).strip() == "":
                continue
            has_any = False
            for c in range(2, max_cols + 1):
                v = ws.cell(row=row, column=c).value
                if v is not None and str(v).strip() != "":
                    has_any = True
                    break
            if has_any:
                return row
        return 0

    def _estimate_row_height(self, row_data, base_height=45, max_height=150):
        data = list(row_data or [])
        if data:
            last = data[-1]
            if isinstance(last, str):
                s = last.strip()
                if s and (":\\" in s or "\\" in s or "/" in s) and s.lower().endswith((".doc", ".docx")):
                    data = data[:-1]
        max_len = 0
        for cell_text in data:
            if cell_text is None:
                continue
            max_len = max(max_len, len(str(cell_text)))
        est_lines = (max_len / 25) + 1
        height = max(base_height, est_lines * 15)
        return min(height, max_height)

    def _apply_template_style(self, dst_cell, src_cell):
        try:
            dst_cell._style = copy(src_cell._style)
        except Exception:
            pass
        try:
            dst_cell.font = copy(src_cell.font)
        except Exception:
            pass
        try:
            dst_cell.border = copy(src_cell.border)
        except Exception:
            pass
        try:
            dst_cell.fill = copy(src_cell.fill)
        except Exception:
            pass
        try:
            dst_cell.number_format = src_cell.number_format
        except Exception:
            pass
        try:
            dst_cell.protection = copy(src_cell.protection)
        except Exception:
            pass
        try:
            base_alignment = copy(src_cell.alignment)
            try:
                dst_cell.alignment = base_alignment.copy(wrapText=True)
            except Exception:
                dst_cell.alignment = Alignment(horizontal=base_alignment.horizontal, vertical=base_alignment.vertical, wrap_text=True)
        except Exception:
            dst_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def _write_rows_to_excel(self, target_excel, extracted_rows, overwrite=False):
        wb = openpyxl.load_workbook(target_excel)
        ws = wb.active

        if overwrite:
            template_row = 3 if ws.max_row >= 3 else 1
            if ws.max_row >= 4:
                try:
                    ws.delete_rows(4, ws.max_row - 3)
                except Exception:
                    pass
            last_serial = 0
            current_row = 4
        else:
            last_valid_row = self._find_last_valid_row(ws, min_row=3, serial_col=1, max_cols=14)
            template_row = last_valid_row if last_valid_row >= 3 else 3
            last_serial = self._coerce_int(ws.cell(row=last_valid_row, column=1).value) if last_valid_row >= 3 else 0
            current_row = (last_valid_row if last_valid_row >= 3 else 2) + 1

        template_cells = [ws.cell(row=template_row, column=c) for c in range(1, 14)]
        template_height = ws.row_dimensions[template_row].height
        if template_height is None:
            template_height = 45
        serial = last_serial

        try:
            ws.column_dimensions[get_column_letter(14)].hidden = True
        except Exception:
            pass

        wrote = 0
        for row_data in extracted_rows:
            if not self._row_has_content(row_data):
                continue

            serial += 1
            if len(row_data) < 14:
                row_data = list(row_data) + [""] * (14 - len(row_data))
            row_data = row_data[:14]
            row_data[0] = str(serial)

            height = self._estimate_row_height(row_data, base_height=template_height)
            ws.row_dimensions[current_row].height = height

            for col_idx, value in enumerate(row_data, start=1):
                dst_cell = ws.cell(row=current_row, column=col_idx, value=value)
                if col_idx <= 13:
                    self._apply_template_style(dst_cell, template_cells[col_idx - 1])

            wrote += 1
            current_row += 1

        wb.save(target_excel)
        return wrote

    def _remove_rows_by_paths(self, target_excel, paths_to_remove):
        if not paths_to_remove:
            return 0
        
        try:
            wb = openpyxl.load_workbook(target_excel)
            ws = wb.active
            
            rows_to_delete = []
            
            # Scan all rows to find matches
            # Data starts at row 4
            for row in range(ws.max_row, 3, -1):
                cell_val = ws.cell(row=row, column=14).value
                if cell_val:
                    s_val = str(cell_val).strip()
                    if s_val:
                        norm_val = os.path.normcase(os.path.normpath(s_val))
                        if norm_val in paths_to_remove:
                            rows_to_delete.append(row)
            
            if not rows_to_delete:
                return 0
            
            self.log(f"发现 {len(rows_to_delete)} 条记录对应已删除的文件，正在清理...")
            
            for r in rows_to_delete:
                ws.delete_rows(r, 1)
                
            # Re-serialize
            serial = 0
            for row in range(4, ws.max_row + 1):
                has_any = False
                for c in range(2, 14):
                    v = ws.cell(row=row, column=c).value
                    if v is not None and str(v).strip() != "":
                        has_any = True
                        break
                
                if has_any:
                    serial += 1
                    ws.cell(row=row, column=1).value = serial

            wb.save(target_excel)
            return len(rows_to_delete)
            
        except Exception as e:
            self.log(f"清理删除文件数据时出错: {e}")
            return 0

    def process_source(self, source_path, target_excel, overwrite=False, incremental=False):
        self.log(f"开始处理: {source_path}")
        
        if not os.path.exists(target_excel):
            self.log(f"错误: 找不到目标Excel文件: {target_excel}")
            return False

        # 1. Collect all Word files
        word_files = []
        if os.path.isfile(source_path):
             if source_path.lower().endswith(('.doc', '.docx')) and not os.path.basename(source_path).startswith('~$'):
                word_files.append(source_path)
        elif os.path.isdir(source_path):
            self.log(f"正在扫描文件夹: {source_path}")
            for root, dirs, files in os.walk(source_path):
                for file in files:
                    if file.lower().endswith(('.doc', '.docx')) and not file.startswith('~$'):
                        word_files.append(os.path.join(root, file))
        else:
            self.log(f"错误: 找不到源文件或文件夹: {source_path}")
            return False

        try:
            word_files.sort()
        except Exception:
            pass

        if incremental:
            processed = self._load_processed_paths_from_excel(target_excel)
            if processed:
                # 1. Handle deleted files
                current_files_set = {os.path.normcase(os.path.normpath(p)) for p in word_files}
                deleted_files = processed - current_files_set
                
                if deleted_files:
                    self.log(f"发现 {len(deleted_files)} 个历史文件已被删除，正在同步清理Excel记录...")
                    removed_count = self._remove_rows_by_paths(target_excel, deleted_files)
                    self.log(f"已清理 {removed_count} 条无效记录。")

                # 2. Handle new files
                before = len(word_files)
                word_files = [p for p in word_files if os.path.normcase(os.path.normpath(p)) not in processed]
                
                if not word_files:
                    if deleted_files:
                        self.log("未发现新Word文档，同步完成。")
                    else:
                        self.log("未发现新Word文档，无需同步。")
                    
                    if self.progress:
                        self.progress(before, before, "完成")
                    return True
            else:
                self.log("提示: 未能从Excel读取历史路径，将执行全量同步。")

        total_files = len(word_files)
        self.log(f"共发现 {total_files} 个Word文件。")
        
        if self.progress:
            self.progress(0, total_files, "准备开始...")

        if not word_files:
            return True

        # 2. Extract data
        extracted_rows = []

        word = None
        temp_dir_obj = None
        try:
            pythoncom.CoInitialize()
            temp_dir_obj = tempfile.TemporaryDirectory()
            temp_dir = temp_dir_obj.name

            def close_word(app):
                if not app:
                    return
                try:
                    app.Quit()
                except Exception:
                    pass

            def kill_all_winword():
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                except Exception:
                    pass

            def create_word():
                kill_all_winword()
                app = win32com.client.DispatchEx("Word.Application")
                try:
                    app.Visible = False
                except Exception:
                    pass
                try:
                    app.DisplayAlerts = 0
                except Exception:
                    pass
                try:
                    app.AutomationSecurity = 3
                except Exception:
                    pass
                return app

            def is_rpc_unavailable(err):
                try:
                    args = getattr(err, "args", None)
                    if args and len(args) >= 1 and int(args[0]) == -2147023174:
                        return True
                except Exception:
                    pass
                return False

            def ensure_word_alive(app):
                if not app:
                    return False
                try:
                    _ = app.Version
                    return True
                except Exception:
                    return False

            word = create_word()
            time.sleep(0.8)
            consecutive_rpc_failures = 0

            for i, file_path in enumerate(word_files):
                if self.stop_requested:
                    self.log("用户停止了操作。")
                    break

                file_name = os.path.basename(file_path)
                self.log(f"正在读取 ({i+1}/{total_files}): {file_name}")
                if self.progress:
                    self.progress(i + 1, total_files, f"读取: {file_name}")

                success = False
                last_error = None

                for attempt in range(3):
                    doc = None
                    try:
                        if not ensure_word_alive(word):
                            close_word(word)
                            word = None
                            word = create_word()
                            time.sleep(0.8)

                        try:
                            doc = self._open_word_doc(word, file_path)
                        except Exception as e:
                            tmp_name = f"{i+1:04d}_{self._safe_temp_name(file_name)}"
                            tmp_path = os.path.join(temp_dir, tmp_name)
                            shutil.copy2(file_path, tmp_path)
                            doc = self._open_word_doc(word, tmp_path)

                        if doc.Tables.Count > 0:
                            table = doc.Tables(1)
                            row_count = table.Rows.Count
                            if row_count > 1:
                                for r in range(2, row_count + 1):
                                    row_data = []
                                    try:
                                        for c in range(1, 14):
                                            try:
                                                cell_text = table.Cell(r, c).Range.Text
                                                cell_text = cell_text.replace('\r', '').replace('\x07', '').strip()
                                                row_data.append(cell_text)
                                            except Exception:
                                                row_data.append("")

                                        has_content = False
                                        if len(row_data) > 1:
                                            for cell in row_data[1:]:
                                                if cell.strip():
                                                    has_content = True
                                                    break

                                        if has_content:
                                            row_data.append(file_path)
                                            extracted_rows.append(row_data)
                                    except Exception:
                                        pass
                            else:
                                self.log(f"  警告: {file_name} 表格行数不足")
                        else:
                            self.log(f"  警告: {file_name} 中没有表格")

                        try:
                            doc.Close(False)
                        except Exception:
                            pass
                        doc = None
                        success = True
                        break
                    except Exception as e:
                        last_error = e
                        try:
                            if doc:
                                doc.Close(False)
                        except Exception:
                            pass
                        if is_rpc_unavailable(e):
                            close_word(word)
                            word = None
                            kill_all_winword()
                            time.sleep(1.2)
                            try:
                                word = create_word()
                                time.sleep(0.8)
                            except Exception:
                                word = None
                            consecutive_rpc_failures += 1
                        else:
                            consecutive_rpc_failures = 0
                        time.sleep(0.8)

                if not success:
                    if last_error is None:
                        self.log(f"  错误: 无法读取文件 {file_name}")
                    else:
                        self.log(f"  错误: 无法读取文件 {file_name}（{type(last_error).__name__}: {last_error}）")

                    if is_rpc_unavailable(last_error) or consecutive_rpc_failures >= 2:
                        isolated_word = None
                        isolated_doc = None
                        isolated_error = None
                        try:
                            kill_all_winword()
                            isolated_word = win32com.client.DispatchEx("Word.Application")
                            try:
                                isolated_word.Visible = False
                            except Exception:
                                pass
                            try:
                                isolated_word.DisplayAlerts = 0
                            except Exception:
                                pass
                            try:
                                isolated_word.AutomationSecurity = 3
                            except Exception:
                                pass
                            time.sleep(0.8)

                            try:
                                isolated_doc = self._open_word_doc(isolated_word, file_path)
                            except Exception:
                                tmp_name = f"isolated_{i+1:04d}_{self._safe_temp_name(file_name)}"
                                tmp_path = os.path.join(temp_dir, tmp_name)
                                shutil.copy2(file_path, tmp_path)
                                isolated_doc = self._open_word_doc(isolated_word, tmp_path)

                            if isolated_doc.Tables.Count > 0:
                                table = isolated_doc.Tables(1)
                                row_count = table.Rows.Count
                                if row_count > 1:
                                    for r in range(2, row_count + 1):
                                        row_data = []
                                        try:
                                            for c in range(1, 14):
                                                try:
                                                    cell_text = table.Cell(r, c).Range.Text
                                                    cell_text = cell_text.replace('\r', '').replace('\x07', '').strip()
                                                    row_data.append(cell_text)
                                                except Exception:
                                                    row_data.append("")

                                            has_content = False
                                            if len(row_data) > 1:
                                                for cell in row_data[1:]:
                                                    if cell.strip():
                                                        has_content = True
                                                        break
                                            if has_content:
                                                row_data.append(file_path)
                                                extracted_rows.append(row_data)
                                        except Exception:
                                            pass
                                    self.log(f"  修复: 已通过隔离模式读取 {file_name}")
                                else:
                                    self.log(f"  警告: {file_name} 表格行数不足")
                            else:
                                self.log(f"  警告: {file_name} 中没有表格")
                        except Exception as e2:
                            isolated_error = e2
                        finally:
                            try:
                                if isolated_doc:
                                    isolated_doc.Close(False)
                            except Exception:
                                pass
                            try:
                                if isolated_word:
                                    isolated_word.Quit()
                            except Exception:
                                pass

                        if isolated_error is not None:
                            self.log(f"  错误: 隔离模式仍失败 {file_name}（{type(isolated_error).__name__}: {isolated_error}）")
        except Exception as e:
            self.log(f"错误: 读取Word时发生异常（{type(e).__name__}: {e}）")
            return False
        finally:
            try:
                close_word(word)
            except Exception:
                pass
            try:
                kill_all_winword()
            except Exception:
                pass
            try:
                if temp_dir_obj:
                    temp_dir_obj.cleanup()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        if self.stop_requested:
            return False

        if not extracted_rows:
            self.log("未提取到任何数据。")
            if self.progress: self.progress(total_files, total_files, "完成")
            return True

        self.log(f"提取完成，共 {len(extracted_rows)} 条记录。正在写入Excel...")
        if self.progress:
            self.progress(total_files, total_files, "正在写入Excel...")

        # 3. Write to Excel
        try:
            wrote = self._write_rows_to_excel(target_excel, extracted_rows, overwrite=overwrite)
            if overwrite:
                self.log(f"写入成功！已刷新 {wrote} 条记录，已保存到: {target_excel}")
            else:
                self.log(f"写入成功！新增 {wrote} 条记录，已保存到: {target_excel}")
            return True

        except PermissionError:
            self.log("错误: 目标Excel文件被占用 (Permission denied)。")
            messagebox.showwarning("文件被占用", "无法写入目标Excel文件。\n\n请检查该文件是否在Excel中打开。\n请关闭文件后再次点击“开始处理”。")
            return False
        except Exception as e:
            self.log(f"写入Excel失败: {e}")
            return False

class StatisticsPanel(ttk.Frame):
    def __init__(self, parent, excel_path, app_instance=None):
        super().__init__(parent)
        self.excel_path = excel_path
        self.app = app_instance
        self.df = None
        self._loaded_path = None
        self._loaded_mtime = None
        self._resize_job = None
        self._last_canvas_size = None
        self._redraw_job = None
        self._redraw_attempts = 0
        self._redraw_stable = 0
        self._redraw_last = None
        self._layout_mode = None
        self.file_path_map = {}
        self.setup_ui()

    def setup_ui(self):
        # Top Control Bar (Filter & Actions)
        control_frame = ttk.Frame(self, padding=5)
        control_frame.pack(fill=X)
        
        # Load Button
        self.btn_load = ttk.Button(control_frame, text="🔄 加载数据", command=self.load_data, bootstyle=PRIMARY)
        self.btn_load.pack(side=LEFT)
        
        # Sync Button
        if self.app:
            self.btn_sync = ttk.Button(control_frame, text="🔁 同步并刷新", command=self.on_sync, bootstyle=SUCCESS)
            self.btn_sync.pack(side=LEFT, padx=5)
        
        ttk.Separator(control_frame, orient=VERTICAL).pack(side=LEFT, padx=10, fill=Y)

        # Date Filter
        ttk.Label(control_frame, text="年份:").pack(side=LEFT, padx=5)
        self.year_var = tk.StringVar(value="全部")
        self.year_cb = ttk.Combobox(control_frame, textvariable=self.year_var, values=["全部"], width=8, state="readonly")
        self.year_cb.pack(side=LEFT)
        self.year_cb.bind("<<ComboboxSelected>>", self.apply_filter)
        
        ttk.Label(control_frame, text="月份:").pack(side=LEFT, padx=5)
        self.month_var = tk.StringVar(value="全部")
        months = ["全部"] + [f"{i}月" for i in range(1, 13)]
        self.month_cb = ttk.Combobox(control_frame, textvariable=self.month_var, values=months, width=8, state="readonly")
        self.month_cb.pack(side=LEFT)
        self.month_cb.bind("<<ComboboxSelected>>", self.apply_filter)
        
        # Export Button
        self.btn_export = ttk.Button(control_frame, text="📤 导出图表", command=self.export_chart, bootstyle=OUTLINE, state="disabled")
        self.btn_export.pack(side=RIGHT)

        # Status Label
        self.lbl_status = ttk.Label(control_frame, text="请先加载数据", bootstyle=SECONDARY)
        self.lbl_status.pack(side=LEFT, padx=20)

        # Content Area
        self.content_area = ttk.Frame(self)
        self.content_area.pack(fill=BOTH, expand=YES, pady=0)
        
        # View 1: Dashboard
        self.view_dashboard = ttk.Frame(self.content_area, padding=5)
        self.setup_dashboard_tab(self.view_dashboard)
        
        # View 2: Detail List
        self.view_details = ttk.Frame(self.content_area, padding=0)
        self.setup_details_tab(self.view_details)
        
        # Default View
        self.current_view = None
        self.switch_view("chart")

    def switch_view(self, view_name):
        if self.current_view:
            self.current_view.pack_forget()
            
        if view_name == "chart":
            self.view_dashboard.pack(fill=BOTH, expand=YES)
            self.current_view = self.view_dashboard
            self.request_redraw()
        elif view_name == "list":
            self.view_details.pack(fill=BOTH, expand=YES)
            self.current_view = self.view_details

    def setup_dashboard_tab(self, parent):
        # Summary Cards (Top)
        self.cards_frame = ttk.Frame(parent)
        self.cards_frame.pack(fill=X, pady=5)
        
        self.card_total = self.create_card(self.cards_frame, "缺陷总数", "0", "info")
        self.card_open = self.create_card(self.cards_frame, "未销号", "0", "danger")
        self.card_closed = self.create_card(self.cards_frame, "已销号", "0", "success")
        
        # Charts Area (Middle)
        self.charts_frame = ttk.Frame(parent)
        self.charts_frame.pack(fill=BOTH, expand=YES, pady=5)
        
        self.fig = Figure(figsize=(10, 5), dpi=100, constrained_layout=True)
        self.fig.patch.set_facecolor('#F8F9FA') # Match light theme bg roughly
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.charts_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=BOTH, expand=YES)
        self.canvas_widget.bind("<Configure>", self.on_resize, add="+")
        self.after(0, self._sync_figure_dpi_to_tk)

    def setup_details_tab(self, parent):
        # Treeview
        columns = ("serial", "location", "type", "status", "date", "action")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", bootstyle="primary")
        
        self.tree.heading("serial", text="序号")
        self.tree.heading("location", text="设备缺陷地点")
        self.tree.heading("type", text="设备缺陷类型")
        self.tree.heading("status", text="状态")
        self.tree.heading("date", text="销号时间")
        self.tree.heading("action", text="操作")
        
        self.tree.column("serial", width=60, anchor="center")
        self.tree.column("location", width=250)
        self.tree.column("type", width=150)
        self.tree.column("status", width=100, anchor="center")
        self.tree.column("date", width=150, anchor="center")
        self.tree.column("action", width=100, anchor="center")
        
        # Scrollbar
        vsb = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)
        vsb.pack(side=RIGHT, fill=Y)
        
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        # Tooltip or instructions removed as per user request


    def on_sync(self):
        if self.app:
            self.app.run_sync_process_from_stats()

    def on_resize(self, event):
        self._last_canvas_size = (event.width, event.height)
        if self._resize_job is not None:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(120, self._on_resize_debounced)

    def _get_tk_dpi(self):
        try:
            dpi = float(self.canvas_widget.winfo_fpixels('1i'))
            if dpi > 0:
                return dpi
        except Exception:
            pass
        return float(self.fig.dpi)

    def _sync_figure_dpi_to_tk(self):
        try:
            dpi = float(self._get_tk_dpi())
        except Exception:
            return
        if dpi <= 0:
            return
        try:
            current = float(self.fig.dpi)
        except Exception:
            current = dpi
        if abs(current - dpi) < 0.5:
            return
        try:
            self.fig.set_dpi(dpi)
        except Exception:
            return

    def _layout_mode_for_width(self, width):
        try:
            w = int(width)
        except Exception:
            w = 0
        if w > 1 and w < 800:
            return "vertical"
        return "horizontal"

    def _on_resize_debounced(self):
        self._resize_job = None
        self._sync_figure_dpi_to_tk()
        try:
            w = int(self.canvas_widget.winfo_width())
        except Exception:
            w = 0
        if w <= 1 and self._last_canvas_size:
            w = int(self._last_canvas_size[0])
        new_mode = self._layout_mode_for_width(w)
        if self.df is not None and new_mode != self._layout_mode:
            self.render_charts()
            return
        try:
            self.canvas.draw_idle()
        except Exception:
            pass

    def request_redraw(self):
        if self._redraw_job is not None:
            try:
                self.after_cancel(self._redraw_job)
            except Exception:
                pass
        self._redraw_attempts = 0
        self._redraw_stable = 0
        self._redraw_last = None
        self._redraw_job = self.after(0, self._redraw_tick)

    def _redraw_tick(self):
        self._redraw_job = None
        self._redraw_attempts += 1
        try:
            self.update_idletasks()
            w = int(self.canvas_widget.winfo_width())
            h = int(self.canvas_widget.winfo_height())
        except Exception:
            return

        if w < 60 or h < 60:
            if self._redraw_attempts < 15:
                self._redraw_job = self.after(60, self._redraw_tick)
            return

        if self._redraw_last == (w, h):
            self._redraw_stable += 1
        else:
            self._redraw_last = (w, h)
            self._redraw_stable = 0

        if self._redraw_stable >= 2 or self._redraw_attempts >= 15:
            try:
                self._sync_figure_dpi_to_tk()
                new_mode = self._layout_mode_for_width(w)
                if self.df is not None and new_mode != self._layout_mode:
                    self.render_charts()
                    return
                self.canvas.draw_idle()
            except Exception:
                return
            return

        self._redraw_job = self.after(80, self._redraw_tick)

    def create_card(self, parent, title, value, bootstyle="primary"):
        frame = ttk.Frame(parent, bootstyle="light", padding=10)
        # Use expand=YES to distribute width evenly
        frame.pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        
        # Use a localized style for the card content
        ttk.Label(frame, text=title, font=("Microsoft YaHei UI", 10), bootstyle="secondary").pack(anchor=W)
        lbl = ttk.Label(frame, text=value, font=("Microsoft YaHei UI", 24, "bold"), bootstyle=bootstyle)
        lbl.pack(pady=5)
        return lbl

    def load_data(self, force=False, silent=False):
        path = self.excel_path.get()
        if not os.path.exists(path):
            if not silent:
                messagebox.showerror("错误", "找不到Excel文件")
            return

        try:
            try:
                mtime = os.path.getmtime(path)
            except Exception:
                mtime = None
            if not force and self.df is not None and self._loaded_path == path and self._loaded_mtime == mtime:
                self.lbl_status.config(text=f"数据已就绪: {time.strftime('%H:%M:%S')}")
                self.request_redraw()
                return

            df = pd.read_excel(path, header=2)
            required_cols = ['设备缺陷类型', '销号时间', '设备缺陷地点']
            if not all(col in df.columns for col in required_cols):
                df = pd.read_excel(path) 
            
            self.df = df
            self._loaded_path = path
            self._loaded_mtime = mtime
            
            # Extract Years for Filter
            self.df['销号时间'] = pd.to_datetime(self.df['销号时间'], errors='coerce')
            years = sorted(self.df['销号时间'].dt.year.dropna().unique().astype(int).astype(str).tolist())
            self.year_cb['values'] = ["全部"] + years
            
            self.update_dashboard(self.df)
            self.lbl_status.config(text=f"数据已更新: {time.strftime('%H:%M:%S')}")
            self.btn_export.config(state="normal")
            self.request_redraw()
            
        except PermissionError:
            if not silent:
                messagebox.showwarning("提示", "请关闭Excel文件后再读取")
        except Exception as e:
            if not silent:
                messagebox.showerror("错误", str(e))

    def apply_filter(self, event=None):
        if self.df is None:
            return
            
        filtered_df = self.df.copy()
        
        year = self.year_var.get()
        month_str = self.month_var.get()
        
        if year != "全部":
            filtered_df = filtered_df[filtered_df['销号时间'].dt.year == int(year)]
            
        if month_str != "全部":
            month = int(month_str.replace("月", ""))
            filtered_df = filtered_df[filtered_df['销号时间'].dt.month == month]
            
        self.update_dashboard(filtered_df)
        self.request_redraw()

    def update_dashboard(self, df):
        # 1. Update Cards
        total = len(df)
        closed = df['销号时间'].notna().sum()
        pending = total - closed
        
        self.card_total.config(text=str(total))
        self.card_open.config(text=str(pending))
        self.card_closed.config(text=str(closed))
        
        # 2. Update Charts
        self.render_charts(df)
        
        # 3. Update Detail List
        self.update_detail_list(df)

    def update_detail_list(self, df):
        # Clear existing
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.file_path_map.clear()
        
        if df is None or df.empty:
            return

        for index, row in df.iterrows():
            serial = ""
            try:
                if '序号' in df.columns:
                    serial = row.get('序号', "")
                elif len(row) > 0:
                    serial = row.iloc[0]
            except Exception:
                serial = ""
            if serial is None or str(serial).strip() == "" or str(serial).strip().lower() == "nan":
                serial = index + 1
            loc = row.get('设备缺陷地点', '')
            dtype = row.get('设备缺陷类型', '')
            date_val = row.get('销号时间')
            date_ts = pd.to_datetime(date_val, errors='coerce')
            is_closed = pd.notna(date_ts)
            status_text = "✅ 已销号" if is_closed else "🔴 未销号"
            date_str = date_ts.strftime('%Y-%m-%d') if is_closed else "-"
            
            source_path = ""
            try:
                for v in reversed(list(row.values)):
                    if not isinstance(v, str):
                        continue
                    s = v.strip()
                    if not s:
                        continue
                    if (":\\" in s or "\\" in s or "/" in s) and s.lower().endswith((".doc", ".docx")):
                        source_path = s
                        break
            except Exception:
                source_path = ""
            
            # Insert into Treeview
            item_id = self.tree.insert("", "end", values=(serial, loc, dtype, status_text, date_str, "📂 打开"))
            
            # Store file path
            if source_path and os.path.exists(source_path):
                self.file_path_map[item_id] = source_path
            
            if not is_closed:
                self.tree.item(item_id, tags=("open",))
            else:
                self.tree.item(item_id, tags=("closed",))

        self.tree.tag_configure("open", foreground="red")
        self.tree.tag_configure("closed", foreground="green")

    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
            
        path = self.file_path_map.get(item_id)
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
        else:
            messagebox.showwarning("提示", "该条记录未关联到Word文档路径（可能是历史数据）。\n建议点击“同步并刷新”后再试。")

    def render_charts(self, df=None):
        if df is None:
            df = self.df
        if df is None:
            return

        try:
            self.update_idletasks()
        except Exception:
            pass

        self.fig.clear()
        self.fig.set_constrained_layout(True)
        
        # Check Theme for Colors
        current_theme = ttk.Style().theme_use()
        is_dark = "dark" in current_theme
        text_color = '#FFFFFF' if is_dark else '#333333'
        bg_color = '#222222' if is_dark else '#F8F9FA'
        grid_color = '#555555' if is_dark else '#DDDDDD'
        
        self.fig.patch.set_facecolor(bg_color)
        
        # Set Global Font & Colors
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['text.color'] = text_color
        plt.rcParams['axes.labelcolor'] = text_color
        plt.rcParams['xtick.color'] = text_color
        plt.rcParams['ytick.color'] = text_color
        plt.rcParams['axes.edgecolor'] = grid_color
        
        # Layout: Responsive GridSpec
        from matplotlib.gridspec import GridSpec
        
        # Check current width to decide layout
        try:
            current_width = int(self.canvas_widget.winfo_width())
        except Exception:
            current_width = 0
        if current_width <= 1 and self._last_canvas_size:
            current_width = int(self._last_canvas_size[0])
        self._layout_mode = self._layout_mode_for_width(current_width)
        if self._layout_mode == "vertical":
            # Vertical Stack (Small Screen)
            gs = GridSpec(2, 1, figure=self.fig, hspace=0.4)
            ax1 = self.fig.add_subplot(gs[0, 0])
            ax2 = self.fig.add_subplot(gs[1, 0])
        else:
            # Horizontal (Large Screen)
            gs = GridSpec(1, 3, figure=self.fig, wspace=0.3)
            ax1 = self.fig.add_subplot(gs[0, :2])
            ax2 = self.fig.add_subplot(gs[0, 2])

        ax1.set_facecolor(bg_color)
        
        type_counts = df['设备缺陷类型'].value_counts().head(8) # Show more items since we have space
        if not type_counts.empty:
            # Apple Style Colors: System Blue
            bars = ax1.bar(type_counts.index, type_counts.values, color='#007AFF', width=0.6, alpha=0.9)
            ax1.set_title("缺陷类型分布 (Top 8)", fontsize=12, pad=15, color=text_color, fontweight='bold')
            ax1.tick_params(axis='x', rotation=30, labelsize=9)
            ax1.grid(axis='y', linestyle='--', alpha=0.5, color=grid_color)
            
            # Remove top and right spines for cleaner look
            ax1.spines['top'].set_visible(False)
            ax1.spines['right'].set_visible(False)
            
            # Add value labels
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}',
                        ha='center', va='bottom', color=text_color, fontsize=9)
        else:
            ax1.text(0.5, 0.5, "暂无分类数据", ha='center', va='center', color=text_color, fontsize=12)
            ax1.axis('off')
            
        total = len(df)
        closed = df['销号时间'].notna().sum()
        pending = total - closed
        
        if total > 0:
            # Apple Style Colors: Green and Red/Orange
            colors = ['#34C759', '#FF3B30'] # iOS Green, iOS Red
            wedges, texts, autotexts = ax2.pie([closed, pending], labels=['已销号', '未销号'], 
                                             autopct='%1.1f%%', colors=colors,
                                             startangle=90, pctdistance=0.85,
                                             textprops={'color': text_color, 'fontsize': 10},
                                             wedgeprops={'width': 0.4, 'edgecolor': bg_color}) # Donut style
            
            # Center text
            ax2.text(0, 0, f"{int((closed/total)*100)}%", ha='center', va='center', fontsize=14, fontweight='bold', color=text_color)
            ax2.set_title("销号完成率", fontsize=12, pad=15, color=text_color, fontweight='bold')
        else:
            ax2.text(0.5, 0.5, "暂无数据", ha='center', va='center', color=text_color, fontsize=12)
            ax2.axis('off')

        self._sync_figure_dpi_to_tk()
        try:
            self.canvas.draw_idle()
        except Exception:
            self.canvas.draw_idle()

    def export_chart(self):
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"设备缺陷统计图表_{timestamp}.png"
        
        filename = filedialog.asksaveasfilename(
            initialfile=default_filename,
            defaultextension=".png", 
            filetypes=[("PNG Image", "*.png"), ("PDF Document", "*.pdf")]
        )
        
        if filename:
            try:
                self.fig.savefig(filename)
                messagebox.showinfo("成功", f"图表已保存至: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")

class App:
    def __init__(self, root):
        self.root = root
        # Initialize Theme
        self.style = ttk.Style(theme='cosmo')
        
        self.root.title("设备缺陷统计管理系统 V2.0")
        self.root.geometry("1100x750")
        
        # Default Maximized Window
        try:
            self.root.state('zoomed')
        except:
            self.root.attributes('-zoomed', True)

        # Logic Components
        self._app_state = _load_app_state()
        self.processor = DefectProcessor(self.log_message, self.update_progress)
        self.excel_path_var = tk.StringVar(value=self._app_state.get("excel_path") or TARGET_EXCEL_PATH)
        self._saved_source_path = self._app_state.get("source_path") or DEFAULT_SOURCE_DIR
        self._processing_lock = threading.Lock()
        self._is_processing = False

        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        except Exception:
            pass
        
        # --- Layout: Sidebar (Left) + Content (Right) ---
        self.setup_ui()

    def setup_ui(self):
        # 1. Sidebar
        self.sidebar = ttk.Frame(self.root, bootstyle="secondary", padding=10)
        self.sidebar.pack(side=LEFT, fill=Y)
        
        # App Title/Logo Area
        ttk.Label(self.sidebar, text="📊 统计助手", font=("Microsoft YaHei UI", 20, "bold"), bootstyle="inverse-secondary").pack(pady=(20, 40))
        
        # Navigation Buttons
        self.nav_btns = {}
        self.create_nav_btn("数据采集", "collect", "📚")
        self.create_nav_btn("统计分析", "stats", "📈")
        self.create_nav_btn("关于", "about", "ℹ️")
        
        self.nav_separator = ttk.Separator(self.sidebar)
        self.nav_separator.pack(fill=X, pady=20)
        
        # Sub-menu for Stats (Hidden by default)
        self.stats_sub_menu = ttk.Frame(self.sidebar, bootstyle="secondary")
        
        # Sub-menu item 1: Charts
        self.btn_view_chart = ttk.Button(self.stats_sub_menu, text="      📊    图表概览", 
                                       command=lambda: self.switch_stats_view("chart"),
                                       bootstyle="info-link", cursor="hand2")
        self.btn_view_chart.pack(pady=2, fill=X, padx=(10, 10))
        try: self.btn_view_chart.configure(anchor="w")
        except: pass
        
        # Sub-menu item 2: List
        self.btn_view_list = ttk.Button(self.stats_sub_menu, text="      📋    明细列表", 
                                      command=lambda: self.switch_stats_view("list"),
                                      bootstyle="secondary-link", cursor="hand2")
        self.btn_view_list.pack(pady=2, fill=X, padx=(10, 10))
        try: self.btn_view_list.configure(anchor="w")
        except: pass
        
        # Theme Toggle
        self.theme_var = tk.BooleanVar(value=False) # False=Light
        self.chk_theme = ttk.Checkbutton(self.sidebar, text="深色模式", variable=self.theme_var, 
                                       command=self.toggle_theme, bootstyle="round-toggle-inverse")
        self.chk_theme.pack(side=BOTTOM, pady=20)
        
        # 2. Content Area
        self.content_container = ttk.Frame(self.root, padding=5)
        self.content_container.pack(side=RIGHT, fill=BOTH, expand=YES)
        
        # Views Container
        self.views = {}
        
        # Initialize Views
        self.create_collect_view()
        self.create_stats_view()
        self.create_about_view()

        try:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, self._saved_source_path)
        except Exception:
            pass
        
        # Show default
        self.show_view("collect")

    def create_nav_btn(self, text, view_name, icon=""):
        # Use more spaces for better separation and anchor w for alignment
        btn = ttk.Button(self.sidebar, text=f"  {icon}    {text}", command=lambda: self.show_view(view_name),
                       bootstyle="secondary-link", cursor="hand2")
        btn.pack(pady=2, fill=X, padx=10) # fill=X ensures full width click area
        # Configure internal alignment of text to the left
        # Note: ttkbootstrap/ttk buttons alignment is handled by style or compound, 
        # but simple text alignment in button is often centered. 
        # We can try to force it via style or width.
        # Actually, for link style, anchor in pack might not be enough for text inside.
        # Let's use a standard button property if available or rely on pack fill.
        # However, ttk.Button 'anchor' option controls text position inside the widget.
        try:
            btn.configure(anchor="w") 
        except:
            pass
            
        self.nav_btns[view_name] = btn

    def switch_stats_view(self, view_type):
        if hasattr(self, 'stats_panel'):
            self.stats_panel.switch_view(view_type)
            # Update button styles
            if view_type == "chart":
                self.btn_view_chart.configure(bootstyle="info-link")
                self.btn_view_list.configure(bootstyle="secondary-link")
            else:
                self.btn_view_chart.configure(bootstyle="secondary-link")
                self.btn_view_list.configure(bootstyle="info-link")

    def show_view(self, view_name):
        # Hide all
        for v in self.views.values():
            v.pack_forget()
        
        # Show selected
        if view_name in self.views:
            self.views[view_name].pack(fill=BOTH, expand=YES)
            
        # Toggle Sub-menu
        if view_name == "stats":
            self.stats_sub_menu.pack(after=self.nav_separator, fill=X, pady=(0, 10))
            if hasattr(self, "stats_panel"):
                def refresh_stats():
                    try:
                        self.stats_panel.load_data(silent=True)
                        self.stats_panel.request_redraw()
                    except Exception:
                        return
                self.root.after(0, refresh_stats)
        else:
            self.stats_sub_menu.pack_forget()
            
        # Update Nav State (Visual feedback)
        for name, btn in self.nav_btns.items():
            if name == view_name:
                btn.configure(bootstyle="light") # Active look
            else:
                btn.configure(bootstyle="secondary-link")

    def toggle_theme(self):
        if self.theme_var.get():
            self.style.theme_use("darkly")
        else:
            self.style.theme_use("cosmo")
            
        # Update charts if they exist
        if hasattr(self, 'stats_panel'):
            self.stats_panel.render_charts()

    def on_close(self):
        try:
            state = {
                "excel_path": self.entry_dst.get() if hasattr(self, "entry_dst") else self.excel_path_var.get(),
                "source_path": self.entry_src.get() if hasattr(self, "entry_src") else DEFAULT_SOURCE_DIR,
                "saved_at": time.time(),
            }
            _save_app_state(state)
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def create_collect_view(self):
        view = ttk.Frame(self.content_container)
        self.views["collect"] = view
        
        # Header
        ttk.Label(view, text="数据采集与处理", font=("Microsoft YaHei UI", 18)).pack(anchor=W, pady=(0, 20))
        
        # Config Card
        config_frame = ttk.Labelframe(view, text="配置参数", padding=15, bootstyle="info")
        config_frame.pack(fill=X, pady=10)
        
        grid = ttk.Frame(config_frame)
        grid.pack(fill=X)
        
        # Row 0: Source
        ttk.Label(grid, text="数据源:").grid(row=0, column=0, sticky=W, pady=15)
        self.entry_src = ttk.Entry(grid)
        self.entry_src.grid(row=0, column=1, sticky=EW, padx=10)
        self.entry_src.insert(0, DEFAULT_SOURCE_DIR)
        
        # Buttons Frame for Source (Align Right)
        btn_frame_src = ttk.Frame(grid)
        btn_frame_src.grid(row=0, column=2, sticky=E)
        ttk.Button(btn_frame_src, text="📁 选文件夹", command=self.browse_folder, bootstyle="primary", width=12).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame_src, text="📄 选文件", command=self.browse_file, bootstyle="info", width=12).pack(side=LEFT, padx=5)
        
        # Row 1: Target
        ttk.Label(grid, text="目标Excel:").grid(row=1, column=0, sticky=W, pady=15)
        self.entry_dst = ttk.Entry(grid, textvariable=self.excel_path_var)
        self.entry_dst.grid(row=1, column=1, sticky=EW, padx=10)
        
        # Button for Target (Align Right, Wider to match above)
        btn_frame_dst = ttk.Frame(grid)
        btn_frame_dst.grid(row=1, column=2, sticky=E)
        ttk.Button(btn_frame_dst, text="📂 选择Excel", command=self.browse_dst, bootstyle="warning", width=27).pack(side=LEFT, padx=5)
        
        grid.columnconfigure(1, weight=1)
        
        # Action Card
        action_frame = ttk.Frame(view, padding=10)
        action_frame.pack(fill=X, pady=10)
        
        self.btn_run = ttk.Button(action_frame, text="▶ 开始处理", command=self.run_process_thread, bootstyle="success", width=20)
        self.btn_run.pack(side=LEFT)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_frame, variable=self.progress_var, maximum=100, bootstyle="success-striped")
        self.progress_bar.pack(side=LEFT, fill=X, expand=YES, padx=20)
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(action_frame, textvariable=self.status_var).pack(side=RIGHT)
        
        # Log Card
        ttk.Label(view, text="运行日志", font=("Microsoft YaHei UI", 12)).pack(anchor=W, pady=(20, 5))
        self.log_area = ScrolledText(view, height=10, state='disabled', font=("Consolas", 9))
        self.log_area.pack(fill=BOTH, expand=YES)

    def create_stats_view(self):
        self.stats_panel = StatisticsPanel(self.content_container, self.excel_path_var, app_instance=self)
        self.views["stats"] = self.stats_panel

    def create_about_view(self):
        view = ttk.Frame(self.content_container)
        self.views["about"] = view
        
        # Center container
        center_frame = ttk.Frame(view)
        center_frame.pack(expand=YES, fill=BOTH, padx=20, pady=20)
        
        # Card - Make it wider and more spacious
        card = ttk.Frame(center_frame, bootstyle="secondary", padding=(40, 40))
        card.pack(anchor=CENTER, fill=X, padx=50)
        
        # --- Top Section: Unit & Guidance ---
        top_section = ttk.Frame(card, bootstyle="secondary")
        top_section.pack(fill=X, pady=(0, 30))
        
        # Unit (Left)
        unit_frame = ttk.Frame(top_section, bootstyle="secondary")
        unit_frame.pack(side=LEFT, fill=X, expand=YES, anchor=N)
        
        ttk.Label(unit_frame, text="单位", font=("Microsoft YaHei UI", 16, "bold"), bootstyle="inverse-secondary").pack(anchor=W, pady=(0, 10))
        ttk.Label(unit_frame, text="惠州电务段汕头水电车间", font=("Microsoft YaHei UI", 14), bootstyle="inverse-secondary").pack(anchor=W)

        # Technical Guidance (Right)
        guide_frame = ttk.Frame(top_section, bootstyle="secondary")
        guide_frame.pack(side=LEFT, fill=X, expand=YES, anchor=N, padx=(20, 0))
        
        ttk.Label(guide_frame, text="技术指导", font=("Microsoft YaHei UI", 16, "bold"), bootstyle="inverse-secondary").pack(anchor=W, pady=(0, 10))
        ttk.Label(guide_frame, text="李海东、梁成欧、庄金旺、郭新城、洪映森", font=("Microsoft YaHei UI", 14), bootstyle="inverse-secondary").pack(anchor=W)
        
        # Separator
        ttk.Separator(card, bootstyle="secondary").pack(fill=X, pady=10)
        
        # --- Middle Section: Author & Contact ---
        mid_section = ttk.Frame(card, bootstyle="secondary")
        mid_section.pack(fill=X, pady=(30, 0))

        # Author (Left)
        author_frame = ttk.Frame(mid_section, bootstyle="secondary")
        author_frame.pack(side=LEFT, fill=X, expand=YES, anchor=N)
        
        ttk.Label(author_frame, text="作者", font=("Microsoft YaHei UI", 16, "bold"), bootstyle="inverse-secondary").pack(anchor=W, pady=(0, 10))
        ttk.Label(author_frame, text="智轨先锋组", font=("Microsoft YaHei UI", 14), bootstyle="inverse-secondary").pack(anchor=W)
        
        # Contact (Right)
        contact_frame = ttk.Frame(mid_section, bootstyle="secondary")
        contact_frame.pack(side=LEFT, fill=X, expand=YES, anchor=N, padx=(20, 0))
        
        ttk.Label(contact_frame, text="联系方式", font=("Microsoft YaHei UI", 16, "bold"), bootstyle="inverse-secondary").pack(anchor=W, pady=(0, 10))
        
        contact_grid = ttk.Frame(contact_frame, bootstyle="secondary")
        contact_grid.pack(anchor=W)
        
        ttk.Label(contact_grid, text="电话: 19119383440", font=("Microsoft YaHei UI", 14), bootstyle="inverse-secondary").pack(side=LEFT, padx=(0, 20))
        ttk.Label(contact_grid, text="微信: yh19119383440", font=("Microsoft YaHei UI", 14), bootstyle="inverse-secondary").pack(side=LEFT)

        # --- Footer Section: Badges ---
        footer = ttk.Frame(card, bootstyle="secondary")
        footer.pack(fill=X, pady=(40, 0))
        
        badges = [("🚄 广铁定制版", "info"), ("🎨 个性化引擎", "warning"), ("🛡️ 本地存储", "primary")]
        
        # Center the badges
        badge_container = ttk.Frame(footer, bootstyle="secondary")
        badge_container.pack(anchor=CENTER)
        
        for text, style in badges:
             lbl = ttk.Label(badge_container, text=f" {text} ", bootstyle=f"{style}-inverse", font=("Microsoft YaHei UI", 12))
             lbl.pack(side=LEFT, padx=15)

    def run_sync_process_from_stats(self):
        if messagebox.askyesno("确认", "是否检查并导入新增的Word文档？\n程序会根据Excel历史记录仅处理新增文件。"):
            self.run_process_thread(is_sync=True)

    # --- Actions ---
    def browse_folder(self):
        initial = self.entry_src.get()
        start_dir = initial if os.path.isdir(initial) else os.path.dirname(initial) if os.path.exists(initial) else None

        d = filedialog.askdirectory(initialdir=start_dir)
        if d:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, d)

    def browse_file(self):
        initial = self.entry_src.get()
        start_dir = os.path.dirname(initial) if os.path.exists(initial) else None

        f = filedialog.askopenfilename(
            initialdir=start_dir,
            filetypes=[("Word Documents", "*.doc;*.docx")]
        )
        if f:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, f)

    def browse_dst(self):
        f = filedialog.askopenfilename(initialdir=os.path.dirname(self.entry_dst.get()), filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.entry_dst.delete(0, tk.END)
            self.entry_dst.insert(0, f)

    def log_message(self, msg):
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg):
        try:
            # Try to use the inner text widget if available (ttkbootstrap ScrolledText)
            text_widget = getattr(self.log_area, 'text', self.log_area)
            text_widget.configure(state='normal')
            text_widget.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
            text_widget.see(tk.END)
            text_widget.configure(state='disabled')
        except Exception as e:
            print(f"Log error: {e}")

    def update_progress(self, current, total, status_msg):
        self.root.after(0, self._update_ui_progress, current, total, status_msg)

    def _update_ui_progress(self, current, total, status_msg):
        if total > 0:
            pct = (current / total) * 100
            self.progress_var.set(pct)
        self.status_var.set(f"{status_msg} ({current}/{total})")

    def run_process_thread(self, is_sync=False):
        src = self.entry_src.get()
        dst = self.entry_dst.get()

        with self._processing_lock:
            if self._is_processing:
                messagebox.showinfo("提示", "当前正在处理，请稍候完成后再操作。")
                return
            self._is_processing = True
        
        try:
            self.btn_run.config(state='disabled')
        except Exception:
            pass
        try:
            if hasattr(self, "stats_panel") and hasattr(self.stats_panel, "btn_sync"):
                self.stats_panel.btn_sync.config(state='disabled')
        except Exception:
            pass
        
        # Clear log
        try:
            text_widget = getattr(self.log_area, 'text', self.log_area)
            text_widget.configure(state='normal')
            text_widget.delete(1.0, tk.END)
            text_widget.configure(state='disabled')
        except Exception as e:
            print(f"Log clear error: {e}")
            
        self.progress_var.set(0)
        
        def task():
            try:
                success = self.processor.process_source(src, dst, overwrite=False, incremental=is_sync)
                if success:
                    if is_sync:
                        self.root.after(0, lambda: self.stats_panel.load_data(force=True, silent=True))
                        self.root.after(0, lambda: messagebox.showinfo("完成", "同步完成！如无新增Word文档则不会追加数据。"))
                    else:
                        self.root.after(0, lambda: messagebox.showinfo("完成", "数据采集处理完成！\n请切换到“统计分析”查看结果。"))
                else:
                    self.root.after(0, lambda: messagebox.showerror("失败", "处理过程中遇到错误，请检查日志。"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("异常", str(e)))
            finally:
                def finish():
                    with self._processing_lock:
                        self._is_processing = False
                    try:
                        self.btn_run.config(state='normal')
                    except Exception:
                        pass
                    try:
                        if hasattr(self, "stats_panel") and hasattr(self.stats_panel, "btn_sync"):
                            self.stats_panel.btn_sync.config(state='normal')
                    except Exception:
                        pass
                    try:
                        self.status_var.set("就绪")
                    except Exception:
                        pass

                self.root.after(0, finish)

        threading.Thread(target=task, daemon=True).start()

if __name__ == "__main__":
    import ttkbootstrap as ttk
    root = ttk.Window(themename="cosmo")
    app = App(root)
    root.mainloop()
