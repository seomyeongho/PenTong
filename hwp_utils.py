# -*- coding: utf-8 -*-
import win32com.client as win32
import pythoncom
import os
import tempfile
import webview  # 파일 다이얼로그 호출을 위해 추가
import zipfile
import io

class HwpController:
    def __init__(self):
        self.hwp = None

    def open_hwp_and_insert(self, html_content):
        """
        한글 프로그램을 실행하고 빈 문서에 HTML 표를 삽입합니다.
        글자 깨짐 방지를 위해 임시 파일 방식으로 처리합니다.
        """
        try:
            # COM 객체 초기화 (스레드 문제 방지)
            pythoncom.CoInitialize()

            # 1. 한글 프로그램 실행 (또는 연결)
            try:
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                # 캐시 문제 발생 시 fallback
                self.hwp = win32.client.Dispatch("HWPFrame.HwpObject")

            # 2. 한글 창 보여주기 및 보안모듈 승인 시도
            self.hwp.XHwpWindows.Item(0).Visible = True
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

            # 3. 빈 문서 2개 뜨는 문제 해결
            try:
                if getattr(self.hwp, "IsModified", 0) != 0 or getattr(self.hwp, "Path", "") != "":
                    self.hwp.Run("FileNew")
            except Exception:
                self.hwp.Run("FileNew")

            # 4. 글자 깨짐 해결: HTML을 임시 파일로 저장 후 불러오기
            temp_dir = tempfile.gettempdir()
            temp_html_path = os.path.join(temp_dir, "temp_pentong_diff.html")
            
            # HTML 뼈대 구성 (UTF-8 인코딩 명시)
            full_html = f"""
            <!DOCTYPE html>
            <html>
            <head><meta charset="utf-8"></head>
            <body>
            {html_content}
            </body>
            </html>
            """

            with open(temp_html_path, "w", encoding="utf-8") as f:
                f.write(full_html)

            # 파일 끼워넣기 액션
            act = self.hwp.CreateAction("InsertFile")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("FileName", temp_html_path)
            pset.SetItem("Format", "HTML")
            act.Execute(pset)
            
            # 5. 표 여백 맞춤, 가운데 정렬, 제목 줄 반복 처리
            try:
                self.hwp.Run("MoveDocBegin") # 커서를 문서 제일 처음으로 이동
                ctrl = self.hwp.HeadCtrl     
                
                while ctrl is not None:
                    if ctrl.CtrlID == "tbl":  
                        anchor = ctrl.GetAnchorPos(0)
                        self.hwp.SetPos(anchor.Item("List"), anchor.Item("Para"), anchor.Item("Pos"))
                        self.hwp.FindCtrl() 
                        
                        # (1) 개체 속성 수정: 여백에 딱 맞추기
                        shape_act = self.hwp.CreateAction("ShapeObjDialog")
                        if shape_act:
                            shape_pset = shape_act.CreateSet()
                            shape_act.GetDefault(shape_pset)
                            shape_pset.SetItem("TreatAsChar", 0)  # 글자처럼 취급 해제 (페이지 분할 허용)
                            shape_pset.SetItem("TextWrap", 1)     # 1: 자리차지 (본문과 겹치지 않게, 오류 수정)
                            shape_pset.SetItem("HorzRelTo", 0)    # 0: 단 기준
                            shape_pset.SetItem("HorzAlign", 2)    # 2: 가운데 정렬
                            shape_pset.SetItem("WidthType", 1)    # 1: 단에 맞춤 (★ 종이 맞춤 오류 수정, 여백에 100% 딱 맞춤)
                            shape_act.Execute(shape_pset)
                        
                        # (2) 표 속성: 제목 줄 반복 및 셀 안에서 나눔
                        table_act = self.hwp.CreateAction("TablePropertyDialog")
                        if table_act:
                            table_pset = table_act.CreateSet()
                            table_act.GetDefault(table_pset)
                            table_pset.SetItem("PageBreak", 2)    # 2: 셀 안에서 나눔
                            table_pset.SetItem("HeaderRepeat", 1) # 1: 제목 줄 반복 켜기
                            table_act.Execute(table_pset)
                        
                        # (3) 첫 줄을 제목 줄로 지정
                        self.hwp.Run("ShapeEnter")  
                        cell_act = self.hwp.CreateAction("TablePropertyDialog")
                        if cell_act:
                            cell_pset = cell_act.CreateSet()
                            cell_act.GetDefault(cell_pset)
                            cell_pset.SetItem("Header", 1) # 현재 줄(첫 줄)을 제목 줄로 지정
                            cell_act.Execute(cell_pset)
                        
                        self.hwp.Run("Cancel")
                        self.hwp.Run("Cancel")
                    
                    ctrl = ctrl.Next
            except Exception as e:
                print("표 자동 조절 중 오류 발생:", e)
            
            # 임시 파일 삭제
            try:
                if os.path.exists(temp_html_path):
                    os.remove(temp_html_path)
            except Exception:
                pass
            
            return True, ""  

        except Exception as e:
            return False, f"한글 실행 중 오류가 발생했습니다: {str(e)}"
        finally:
            pythoncom.CoUninitialize()

    def select_hwp_files(self, window):
        """hwpx 파일만 선택하도록 변경 및 펜통 압축 파일 감지 로직 추가"""
        file_types = ('한글 HWPX 파일 (*.hwpx)', 'All files (*.*)')
        file_paths = window.create_file_dialog(
            webview.OPEN_DIALOG, 
            allow_multiple=True, 
            file_types=file_types
        )
        
        result = []
        if file_paths:
            for path in file_paths:
                size = os.path.getsize(path)
                name = os.path.basename(path)
                is_already_compressed = False
                
                # [정공법] HWPX(ZIP) 내부를 열어서 펜통이 남긴 이름표가 있는지 확인
                if path.lower().endswith('.hwpx'):
                    try:
                        with zipfile.ZipFile(path, 'r') as zf:
                            if 'pentong_compressed.meta' in zf.namelist():
                                is_already_compressed = True
                    except Exception:
                        pass
                        
                result.append({
                    "path": path, 
                    "name": name, 
                    "size": size,
                    "isAlreadyCompressed": is_already_compressed # 프론트로 전달
                })
        return result

    def compress_hwp_files(self, payload, window=None):
        """HWPX 내부의 이미지를 직접 찾아 리사이징 및 화질 압축(JPEG 변환)을 수행합니다."""
        paths = payload.get("paths", [])
        name_format = payload.get("nameFormat", "{name}_압축.hwpx")
        quality_level = payload.get("quality", 70)
        target_dir = payload.get("targetDir") 
        results = []
        
        try:
            from PIL import Image
            import io, zipfile
        except ImportError:
            return {"error": "이미지 압축을 위해 Pillow 라이브러리가 필요합니다. 터미널에 'pip install Pillow'를 입력해주세요."}

        if not target_dir or not os.path.isdir(target_dir):
            return {"error": "유효한 저장 폴더가 지정되지 않았습니다."}

        for path in paths:
            try:
                base_name, ext = os.path.splitext(os.path.basename(path))
                if ext.lower() != '.hwpx':
                    results.append({"path": path, "newSize": None, "savedName": None, "success": False, "error": "HWPX 파일만 지원합니다."})
                    continue

                new_file_name = name_format.replace("{name}", base_name)
                new_path = os.path.join(target_dir, new_file_name)
                
                # 중복 파일명 처리 (2), (3) ...
                if os.path.exists(new_path):
                    name_part, ext_part = os.path.splitext(new_file_name)
                    counter = 2
                    while os.path.exists(new_path):
                        new_path = os.path.join(target_dir, f"{name_part}({counter}){ext_part}")
                        counter += 1

                final_saved_name = os.path.basename(new_path)

                with zipfile.ZipFile(path, 'r') as zin:
                    with zipfile.ZipFile(new_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            
                            if 'BinData/' in item.filename and item.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                                try:
                                    img = Image.open(io.BytesIO(data))
                                    max_width = 1024
                                    if img.width > max_width:
                                        ratio = max_width / img.width
                                        new_size = (max_width, int(img.height * ratio))
                                        img = img.resize(new_size, Image.Resampling.LANCZOS)
                                    
                                    out_io = io.BytesIO()
                                    if img.mode in ("RGBA", "P"):
                                        img = img.convert("RGB")
                                        
                                    img.save(out_io, format='JPEG', quality=quality_level, optimize=True)
                                    data = out_io.getvalue()
                                except Exception as img_e:
                                    print(f"[{item.filename}] 이미지 압축 스킵 (원본 유지): {img_e}")
                            
                            zout.writestr(item, data)
                        
                        # ========================================================
                        # [핵심 추가] 압축 포장을 닫기 직전, 펜통 압축 완료 이름표 몰래 넣기
                        # ========================================================
                        zout.writestr('pentong_compressed.meta', b'compressed_by_pentong_v1')

                new_size = os.path.getsize(new_path)
                results.append({"path": path, "newSize": new_size, "savedName": final_saved_name, "success": True, "error": ""})
            except Exception as e:
                results.append({"path": path, "newSize": None, "savedName": None, "success": False, "error": str(e)})
                
        return results

 
    def get_hwp_info(self, paths):
        """[추가] 드래그 앤 드롭된 파일 경로들의 용량과 압축 여부 메타데이터를 추출합니다."""
        result = []
        if paths:
            for path in paths:
                if not os.path.exists(path): continue
                size = os.path.getsize(path)
                name = os.path.basename(path)
                is_already_compressed = False
                
                # HWPX(ZIP) 내부를 열어서 펜통이 남긴 이름표가 있는지 확인
                if path.lower().endswith('.hwpx'):
                    try:
                        with zipfile.ZipFile(path, 'r') as zf:
                            if 'pentong_compressed.meta' in zf.namelist():
                                is_already_compressed = True
                    except Exception:
                        pass
                        
                result.append({
                    "path": path, 
                    "name": name, 
                    "size": size,
                    "isAlreadyCompressed": is_already_compressed
                })
        return result