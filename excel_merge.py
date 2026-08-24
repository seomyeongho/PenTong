# -*- coding: utf-8 -*-
import os
import time
import webview
import pandas as pd
import numpy as np
import openpyxl

# 엑셀 제어를 위한 라이브러리
try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

class ExcelMergeService:
    def __init__(self):
        self.cache_file = None
        self.cache_sheet = None
        self.cache_df = None
        self.cache_mtime = None
        self.cache_merge_map = None
        self.cache_hidden_master = None

    def _get_sheet_info(self, file_path, sheet_name):
        """엑셀 파일에서 병합 셀 정보와 데이터프레임을 함께 추출하고 캐싱합니다."""
        mtime = os.path.getmtime(file_path)
        if self.cache_file == file_path and self.cache_sheet == sheet_name and self.cache_mtime == mtime:
            return self.cache_df, self.cache_merge_map, self.cache_hidden_master
        
        # 1. openpyxl로 병합 셀(MergeCells) 정보 완벽 추출
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb[sheet_name]
        
        merge_map = {}
        hidden_master_map = {}
        for m in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = m.bounds
            rowspan = max_row - min_row + 1
            colspan = max_col - min_col + 1
            merge_map[(min_row, min_col)] = (rowspan, colspan)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if r == min_row and c == min_col:
                        continue
                    hidden_master_map[(r, c)] = (min_row, min_col)
        wb.close()
        
        # 2. pandas로 데이터프레임 추출 (파일 잠금 방지를 위해 바이너리로 읽기)
        with open(file_path, "rb") as f:
            df = pd.read_excel(f, sheet_name=sheet_name, header=None, engine='openpyxl')
            
        self.cache_df = df
        self.cache_merge_map = merge_map
        self.cache_hidden_master = hidden_master_map
        self.cache_file = file_path
        self.cache_sheet = sheet_name
        self.cache_mtime = mtime
        return df, merge_map, hidden_master_map

    def select_files(self, window):
        file_types = ('Excel Files (*.xlsx;*.xls)', 'All files (*.*)')
        result = window.create_file_dialog(
            webview.OPEN_DIALOG, 
            allow_multiple=True, 
            file_types=file_types
        )
        return result

    def get_sheet_names(self, file_path):
        try:
            with open(file_path, "rb") as f:
                xls = pd.ExcelFile(f, engine='openpyxl')
                sheet_names = xls.sheet_names
            return {"success": True, "sheets": sheet_names, "path": file_path}
        except Exception as e:
            return {"success": False, "message": f"읽기 실패: {str(e)}", "path": file_path}

    def open_file(self, file_path):
        try:
            if os.path.exists(file_path):
                os.startfile(file_path)
                return {"success": True}
            else:
                return {"success": False, "message": "파일이 존재하지 않습니다."}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def merge_excel_sheets(self, merge_data):
        if not merge_data:
            return {"success": False, "message": "선택된 데이터가 없습니다."}

        if not HAS_WIN32:
            return {"success": False, "message": "필수 라이브러리(pywin32)가 설치되지 않았습니다."}

        all_data_frames = []
        try:
            for item in merge_data:
                file_path = item['path']
                selected_sheets = item['sheets']
                if not selected_sheets:
                    continue
                for sheet_name in selected_sheets:
                    try:
                        with open(file_path, "rb") as f:
                            df = pd.read_excel(f, sheet_name=sheet_name, engine='openpyxl')
                        if not df.empty:
                            all_data_frames.append(df)
                    except Exception:
                        continue

            if not all_data_frames:
                return {"success": False, "message": "취합할 데이터가 비어있습니다."}

            merged_df = pd.concat(all_data_frames, ignore_index=True)
            merged_df = merged_df.fillna("")

            pythoncom.CoInitialize()
            try:
                excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            except:
                excel = win32com.client.Dispatch("Excel.Application")
            
            excel.Visible = True 
            wb = excel.Workbooks.Add() 
            ws = wb.Worksheets(1)
            ws.Name = "취합결과"

            header = merged_df.columns.tolist()
            data = merged_df.values.tolist()
            final_data = [header] + data
            
            rows = len(final_data)
            cols = len(final_data[0])

            start_cell = ws.Cells(1, 1)
            end_cell = ws.Cells(rows, cols)
            ws.Range(start_cell, end_cell).Value = final_data

            ws.Rows(1).Font.Bold = True 
            ws.Columns.AutoFit() 

            return {"success": True, "message": "새 엑셀 문서에 데이터가 취합되었습니다."}
        except Exception as e:
            return {"success": False, "message": f"작업 중 오류 발생: {str(e)}"}

    def open_in_excel(self, data):
        if not HAS_WIN32:
            return {"success": False, "message": "필수 라이브러리(pywin32)가 설치되지 않았습니다."}
        try:
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            except:
                excel = win32com.client.Dispatch("Excel.Application")
            
            excel.Visible = True 
            wb = excel.Workbooks.Add() 
            ws = wb.Worksheets(1)
            ws.Name = "좌석배치결과"

            rows = len(data)
            cols = len(data[0]) if rows > 0 else 0

            if rows > 0 and cols > 0:
                start_cell = ws.Cells(1, 1)
                end_cell = ws.Cells(rows, cols)
                ws.Range(start_cell, end_cell).Value = data
                ws.Rows(1).Font.Bold = True 
                ws.Columns.AutoFit() 
                
            return {"success": True}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def _find_table_boundaries(self, df):
        df_clean = df.replace(r'^\s*$', np.nan, regex=True)
        row_densities = df_clean.notna().sum(axis=1)
        
        if row_densities.empty or row_densities.max() == 0:
            return [0], 1, 1

        max_density = row_densities.max()
        
        if max_density <= 1:
            first_row = int(row_densities[row_densities > 0].index[0])
            header_indices = [first_row]
            data_start_idx = first_row + 1
        else:
            threshold_data = max(2, int(max_density * 0.8))
            dense_start = 0
            for i in range(len(row_densities)):
                if row_densities.get(i, 0) >= threshold_data:
                    dense_start = i
                    break
            
            sparse_block = []
            for i in range(dense_start - 1, -1, -1):
                if row_densities.get(i, 0) > 0:
                    sparse_block.insert(0, i)
                else:
                    break
                    
            while len(sparse_block) > 0 and row_densities.get(sparse_block[0], 0) <= 1:
                sparse_block.pop(0)
                
            if not sparse_block:
                header_indices = [dense_start]
                data_start_idx = dense_start + 1
            else:
                header_indices = sparse_block
                data_start_idx = dense_start

        if max_density <= 2:
            footer_threshold = 0
        elif max_density <= 5:
            footer_threshold = 1
        else:
            footer_threshold = 2
            
        data_end_idx = df.index.max()
        for i in range(data_start_idx, len(df)):
            density = row_densities.get(i, 0)
            if density <= footer_threshold:
                data_end_idx = i - 1
                break
                
        if data_end_idx < data_start_idx:
            data_end_idx = data_start_idx
            
        return header_indices, data_start_idx, data_end_idx

    def _get_excel_column_name(self, n):
        name = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            name = chr(65 + remainder) + name
        return name

    def _build_preview_data(self, df_str, start_row, end_row, merge_map, hidden_master_map):
        """병합된 셀 정보를 포함하여 HTML에 렌더링할 완벽한 구조의 배열을 만듭니다."""
        preview_data = []
        max_cols = len(df_str.columns)
        window_start_excel_r = start_row + 1
        
        for i in range(start_row, end_row):
            excel_r = i + 1
            row_data = []
            for c_idx in range(max_cols):
                excel_c = c_idx + 1
                
                is_hidden = (excel_r, excel_c) in hidden_master_map
                rowspan, colspan = merge_map.get((excel_r, excel_c), (1, 1))
                
                if not is_hidden:
                    val = str(df_str.iloc[i, c_idx])
                else:
                    val = ""
                
                # 현재 표시할 윈도우(예: 20줄) 가장 첫 줄인데, 병합 셀의 주인이 창 밖에 있다면 강제로 끌어와서 보여줌
                if is_hidden:
                    master_r, master_c = hidden_master_map[(excel_r, excel_c)]
                    if excel_r == window_start_excel_r and master_r < window_start_excel_r:
                        is_hidden = False
                        master_rowspan, master_colspan = merge_map[(master_r, master_c)]
                        rowspan = master_rowspan - (window_start_excel_r - master_r)
                        colspan = master_colspan
                        val = str(df_str.iloc[master_r - 1, master_c - 1])
                
                # 병합된 셀이 윈도우 끝을 벗어나면 렌더링 깨짐을 막기 위해 높이를 자름
                if not is_hidden and excel_r + rowspan - 1 > end_row:
                    rowspan = end_row - excel_r + 1
                
                row_data.append({
                    "hidden": is_hidden,
                    "val": val,
                    "rowspan": rowspan,
                    "colspan": colspan
                })
            preview_data.append({"row": excel_r, "cells": row_data})
        return preview_data

    def analyze_excel_sheet(self, params):
        file_path = params.get('file_path')
        sheet_name = params.get('sheet_name')
        action = params.get('action') 
        
        try:
            df, merge_map, hidden_master = self._get_sheet_info(file_path, sheet_name)
            df_dropped = df.dropna(how='all')
            
            if df_dropped.empty:
                return {"success": False, "message": "시트에 데이터가 없습니다."}
                
            # [기능 1] 데이터 시작행 지정 시 거꾸로 추적하여 3행 제한으로 제목 감지
            if action == 'infer_header':
                data_start_str = str(params.get('data_start', '2'))
                df_clean = df.replace(r'^\s*$', np.nan, regex=True)
                row_densities = df_clean.notna().sum(axis=1)

                start_idx = int(data_start_str) - 1
                if start_idx < 1:
                    header_indices = [0]
                else:
                    header_indices = []
                    data_density = row_densities.get(start_idx, 1)

                    # 최대 3행까지만 위로 거슬러 올라감
                    for i in range(start_idx - 1, max(-1, start_idx - 4), -1):
                        density = row_densities.get(i, 0)
                        if density == 0:
                            break # 완전 빈칸 만나면 스톱
                        if data_density > 1 and density <= 1:
                            break # 위가 1칸뿐이면 문서 제목으로 간주하고 스톱
                            
                        header_indices.insert(0, i)

                    if not header_indices:
                        header_indices = [max(0, start_idx - 1)]

                # 역추적된 헤더 병합 텍스트 생성
                header_df = df.iloc[header_indices].copy()
                header_df = header_df.ffill(axis=0).ffill(axis=1)
                
                columns = []
                for col_idx in header_df.columns:
                    vals = []
                    for val in header_df[col_idx]:
                        if pd.notna(val) and str(val).strip():
                            val_str = str(val).strip()
                            if not vals or vals[-1] != val_str:
                                vals.append(val_str)
                    col_name = "_".join(vals)
                    
                    if len(col_name) > 10:
                        col_name = col_name[:10] + "..."
                        
                    if col_name:
                        columns.append({"index": col_idx + 1, "name": col_name})

                header_str = f"{header_indices[0]+1}-{header_indices[-1]+1}" if len(header_indices) > 1 else str(header_indices[0]+1)
                return {"success": True, "header_row": header_str, "columns": columns}

            # [기능 2] 최초 시트 분석 및 병합 정보가 적용된 20행 미리보기 데이터 생성
            header_indices, data_start_idx, data_end_idx = self._find_table_boundaries(df)
            
            header_df = df.iloc[header_indices].copy()
            header_df = header_df.ffill(axis=0).ffill(axis=1)
            
            columns = []
            for col_idx in header_df.columns:
                vals = []
                for val in header_df[col_idx]:
                    if pd.notna(val) and str(val).strip():
                        val_str = str(val).strip()
                        if not vals or vals[-1] != val_str:
                            vals.append(val_str)
                col_name = "_".join(vals)
                if len(col_name) > 10:
                    col_name = col_name[:10] + "..."
                if col_name:
                    columns.append({"index": col_idx + 1, "name": col_name})
            
            if len(header_indices) > 1:
                header_row_excel = f"{header_indices[0]+1}-{header_indices[-1]+1}"
            else:
                header_row_excel = str(header_indices[0]+1)

            df_str = df.fillna("").astype(str).replace(r'\.0$', '', regex=True)
            total_rows = len(df_str)
            max_cols = len(df_str.columns)
            
            # 3번 표: 헤더부터 약 20줄만 (병합 유지)
            top_start = max(0, header_indices[0] if header_indices else data_start_idx - 3)
            top_end = min(total_rows, top_start + 20)
            preview_top = self._build_preview_data(df_str, top_start, top_end, merge_map, hidden_master)

            # 4번 표: 맨 아래에서 20줄만 (병합 유지)
            bottom_start = max(0, total_rows - 20)
            preview_bottom = self._build_preview_data(df_str, bottom_start, total_rows, merge_map, hidden_master)
                    
            excel_cols = [self._get_excel_column_name(i+1) for i in range(max_cols)]
            
            return {
                "success": True,
                "header_row": header_row_excel,
                "data_start": str(data_start_idx + 1),
                "data_end": str(data_end_idx + 1),
                "columns": columns,
                "mtime": self.cache_mtime,
                "preview_top": preview_top,
                "preview_bottom": preview_bottom,
                "excel_cols": excel_cols
            }
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "message": f"분석 중 오류: {str(e)}"}

    def split_excel_by_column(self, params):
        if not HAS_WIN32:
            return {"success": False, "message": "필수 라이브러리(pywin32)가 설치되지 않았습니다."}
            
        file_path = params.get('file_path')
        sheet_name = params.get('sheet_name')
        data_start_str = str(params.get('data_start', '2'))
        data_end_str = str(params.get('data_end', '0'))
        target_col_idx = int(params.get('target_col_idx', 1))
        save_dir = params.get('save_dir')
        naming_pattern = params.get('naming_pattern', "{순번}_{원본파일명}_{기준값}")
        columns_info = params.get('columns', []) 
        last_mtime = params.get('mtime')

        # [수정] 메인 창(index.html)뿐만 아니라 내부 액자(iframe)까지 샅샅이 뒤져서 진행률을 전달하도록 개선
        def notify_progress(msg, current=0, total=0):
            try:
                if webview.windows:
                    win = webview.windows[0]
                    safe_msg = str(msg).replace('"', '\\"').replace('\n', '\\n')
                    js_code = f"""
                    if (window.updateSplitProgress) {{ window.updateSplitProgress("{safe_msg}", {current}, {total}); }}
                    var frames = document.getElementsByTagName('iframe');
                    for(var i=0; i<frames.length; i++) {{
                        if(frames[i].contentWindow && frames[i].contentWindow.updateSplitProgress) {{
                            frames[i].contentWindow.updateSplitProgress("{safe_msg}", {current}, {total});
                        }}
                    }}
                    """
                    win.evaluate_js(js_code)
            except Exception:
                pass

        if not all([file_path, sheet_name, save_dir, target_col_idx]):
            return {"success": False, "message": "필요한 정보가 모두 입력되지 않았습니다."}

        data_start = int(data_start_str.split('-')[0]) if '-' in data_start_str else int(data_start_str)
        data_end = int(data_end_str.split('-')[-1]) if '-' in data_end_str else int(data_end_str)

        current_mtime = os.path.getmtime(file_path)
        if last_mtime and float(last_mtime) != float(current_mtime):
            return {
                "success": False,
                "action": "re_analyze",
                "message": f"파일이 외부에서 수정되었습니다.\n구조가 변경되었을 수 있어 데이터를 다시 불러옵니다."
            }

        col_name_to_idx = {col['name']: col['index'] for col in columns_info}
        orig_filename = os.path.splitext(os.path.basename(file_path))[0]

        excel = None
        wb = None
        try:
            # Pandas의 병합 셀 인식 오류 원천 차단
            df, merge_map, hidden_master = self._get_sheet_info(file_path, sheet_name)

            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(file_path, ReadOnly=True)
            ws = wb.Sheets(sheet_name)

            used_range = ws.UsedRange
            last_row_actual = used_range.Row + used_range.Rows.Count - 1
            
            if data_end == 0 or data_end > last_row_actual:
                data_end = last_row_actual

            if data_start > data_end:
                 return {"success": False, "message": "데이터 시작 행이 데이터 종료 행보다 큽니다."}

            row_to_val = {}
            unique_splits = {}
            
            raw_values = ws.Range(ws.Cells(data_start, target_col_idx), ws.Cells(data_end, target_col_idx)).Value
            if not isinstance(raw_values, (list, tuple)):
                raw_values = [[raw_values]]
                
            for i, row in enumerate(raw_values):
                r_excel = data_start + i
                raw_val = row[0]
                
                # 병합 셀(MergeArea)을 통한 100% 확실한 값 추출
                if raw_val is None:
                    cell = ws.Cells(r_excel, target_col_idx)
                    if cell.MergeCells:
                        val = cell.MergeArea.Cells(1, 1).Value
                    else:
                        val = ""
                else:
                    val = raw_val
                    
                str_val = str(val).strip() if val is not None else ""
                
                if not str_val:
                    return {
                        "success": False, 
                        "message": f"⚠️ 분할 기준열({r_excel}행)에 빈 값이 있습니다.\n\n해당 열은 분할 기준이 되므로 비어있으면 안 됩니다.\n미리보기를 확인하여 데이터 종료행을 알맞게 위로 올려주세요."
                    }

                row_to_val[r_excel] = str_val
                
                if str_val != "" and str_val not in unique_splits:
                    unique_splits[str_val] = r_excel 

            if not unique_splits:
                return {"success": False, "message": "선택한 데이터 범위 내 분할 기준 열에 유효한 데이터가 없습니다."}

            df_str = df.fillna("").astype(str).replace(r'\.0$', '', regex=True)

            generated_filenames = {} 
            seen_filenames = set()
            
            seq = 1
            for val, r_excel in unique_splits.items():
                if r_excel - 1 < len(df_str):
                    rep_row_data = df_str.iloc[r_excel - 1].tolist()
                else:
                    rep_row_data = []
                    
                new_filename = naming_pattern
                new_filename = new_filename.replace("{원본파일명}", orig_filename)
                new_filename = new_filename.replace("{기준값}", val)
                new_filename = new_filename.replace("{순번}", str(seq))
                
                for col_name, c_idx in col_name_to_idx.items():
                    placeholder = f"{{{col_name}}}"
                    if placeholder in new_filename:
                        if c_idx - 1 < len(rep_row_data):
                            new_filename = new_filename.replace(placeholder, rep_row_data[c_idx - 1])
                        else:
                            new_filename = new_filename.replace(placeholder, "")
                
                safe_filename = "".join([c for c in new_filename if c not in '<>:"/\\|?*']).strip()
                if not safe_filename.endswith('.xlsx'):
                    safe_filename += '.xlsx'
                
                if safe_filename in seen_filenames:
                    return {
                        "success": False, 
                        "message": f"파일명 중복 오류가 예상됩니다.\n중복된 파일명: {safe_filename}\n순번({{순번}}) 등 변수를 추가하여 고유하게 만들어주세요."
                    }
                
                seen_filenames.add(safe_filename)
                generated_filenames[val] = safe_filename
                seq += 1

            created_files = []
            total_files = len(generated_filenames)
            
            # [추가] 검증을 모두 통과하면 본격적인 분할 시작을 알림!
            notify_progress(f"검증 완료! 총 {total_files}개의 파일로 분할 작업을 시작합니다...", 0, total_files)
            
            for val, filename in generated_filenames.items():
                ws.Copy()
                new_wb = excel.ActiveWorkbook
                new_ws = new_wb.Sheets(1)
                
                ranges_to_delete = []
                current_union = None
                count = 0
                
                for r in range(data_end, data_start - 1, -1):
                    str_val = row_to_val.get(r, "")
                    
                    if str_val != val:
                        if current_union is None:
                            current_union = new_ws.Rows(r)
                        else:
                            current_union = excel.Union(current_union, new_ws.Rows(r))
                        count += 1
                        
                        if count >= 100:
                            ranges_to_delete.append(current_union)
                            current_union = None
                            count = 0
                
                if current_union is not None:
                    ranges_to_delete.append(current_union)
                    
                for rng in ranges_to_delete:
                    rng.Delete()
                
                save_path = os.path.join(save_dir, filename)
                new_wb.SaveAs(save_path)
                new_wb.Close(False)
                new_wb = None 
                created_files.append(filename)

                # [추가] 파일이 하나씩 생성될 때마다 웹 화면으로 게이지 상태 업데이트 콜백
                notify_progress(f"[{filename}] 생성 중...", len(created_files), total_files)

            return {"success": True, "message": f"분할 작업이 완료되었습니다.\n총 {len(created_files)}개의 파일이 생성되었습니다."}

        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            print(err_msg)
            return {"success": False, "message": f"엑셀 처리 중 오류가 발생했습니다: {str(e)}"}
            
        finally:
            if excel:
                try:
                    excel.ScreenUpdating = True
                    excel.DisplayAlerts = True
                    if wb: 
                        wb.Close(False)
                    excel.Quit()
                except:
                    pass