# 파일명: multi_pdf_merger.py

import os
import shutil
import tempfile
import webview
import traceback
import time
import platform
import subprocess

class MultiPdfMergerService:
    def select_mixed_files(self, window):
        try:
            file_types = (
                '지원하는 문서 및 이미지 (*.hwp;*.hwpx;*.pdf;*.jpg;*.jpeg;*.png;*.gif)',
                '모든 파일 (*.*)'
            )
            result = window.create_file_dialog(
                webview.OPEN_DIALOG, 
                allow_multiple=True, 
                file_types=file_types
            )
            return list(result) if result else []
        except Exception as e:
            print(f"파일 선택 오류: {e}")
            return []

    def merge_to_pdf(self, window, file_paths):
        try:
            if not file_paths:
                return {"success": False, "message": "선택된 파일이 없습니다."}

            # -----------------------------------------------------------------
            # 💡 [추가] 1. 취합할 원본 파일 중 사용(열려있는) 중인 파일이 있는지 미리 체크
            # -----------------------------------------------------------------
            for path in file_paths:
                # 파일이 존재하고 수정 가능한(읽기 전용이 아닌) 파일인 경우에만 잠금 검사
                if os.path.exists(path) and os.access(path, os.W_OK):
                    try:
                        with open(path, 'a'): pass
                    except IOError:
                        return {
                            "success": False, 
                            "message": f"현재 실행 중이거나 열려 있는 원본 파일이 있습니다.\n파일을 완전히 닫은 후 다시 시도해주세요.\n\n파일명: {os.path.basename(path)}"
                        }

            # 저장 위치 다이얼로그 호출
            save_path = window.create_file_dialog(
                webview.SAVE_DIALOG, 
                save_filename='취합된_문서.pdf',
                file_types=('PDF 파일 (*.pdf)',)
            )
            
            if not save_path:
                return {"success": False, "message": "저장이 취소되었습니다."}
            
            target_path = save_path[0] if isinstance(save_path, (list, tuple)) else save_path

            # -----------------------------------------------------------------
            # 💡 [추가] 2. 덮어쓸 대상(최종 저장될) PDF 파일이 이미 켜져 있는지 미리 체크
            # -----------------------------------------------------------------
            if os.path.exists(target_path):
                try:
                    with open(target_path, 'a'): pass
                except IOError:
                    return {
                        "success": False, 
                        "message": f"저장할 대상 PDF 파일이 현재 PDF 뷰어 등에서 열려있습니다.\n창을 닫은 후 다시 시도해주세요.\n\n파일명: {os.path.basename(target_path)}"
                    }

            import fitz  # PyMuPDF
            import win32com.client as win32
            
            merged_pdf = fitz.open()
            temp_dir = tempfile.mkdtemp()
            hwp = None

            try:
                for i, path in enumerate(file_paths):
                    ext = path.lower().split('.')[-1]
                    
                    if ext in ['hwp', 'hwpx']:
                        if hwp is None:
                            hwp = win32.Dispatch("HWPFrame.HwpObject")
                            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
                        
                        abs_target_path = os.path.abspath(path)
                        temp_pdf_path = os.path.abspath(os.path.join(temp_dir, f"temp_{i}.pdf"))
                        
                        # -----------------------------------------------------------------
                        # 💡 [수정] 3. 매개 변수 개수(Parameter Count) 에러 완벽 차단!
                        # Open, SaveAs에 빈 문자열("")을 3개씩 명시적으로 채워 가장 깐깐한 한글 버전 대응
                        # -----------------------------------------------------------------
                        hwp.Open(abs_target_path, "", "")
                        hwp.SaveAs(temp_pdf_path, "PDF", "")
                        hwp.Clear(1)
                        
                        time.sleep(0.5)
                        
                        if os.path.exists(temp_pdf_path):
                            doc = fitz.open(temp_pdf_path)
                            merged_pdf.insert_pdf(doc)
                            doc.close()
                        else:
                            raise Exception(f"한글 문서를 PDF로 변환하지 못했습니다: {os.path.basename(path)}")
                            
                    elif ext in ['pdf']:
                        doc = fitz.open(path)
                        merged_pdf.insert_pdf(doc)
                        doc.close()
                        
                    elif ext in ['jpg', 'jpeg', 'png', 'gif']:
                        img_doc = fitz.open(path)
                        pdf_bytes = img_doc.convert_to_pdf()
                        img_pdf = fitz.open("pdf", pdf_bytes)
                        merged_pdf.insert_pdf(img_pdf)
                        img_pdf.close()
                        img_doc.close()
                
                merged_pdf.save(target_path)
                
            finally:
                merged_pdf.close()
                if hwp:
                    hwp.Quit()
                shutil.rmtree(temp_dir, ignore_errors=True)
            
            # 생성된 PDF 자동 실행
            try:
                if platform.system() == 'Windows':
                    os.startfile(target_path)
                elif platform.system() == 'Darwin':
                    subprocess.call(('open', target_path))
                else:
                    subprocess.call(('xdg-open', target_path))
            except Exception as e:
                print(f"PDF 자동 실행 실패: {e}")

            return {"success": True, "message": f"모든 파일이 성공적으로 취합되었습니다.\n({os.path.basename(target_path)})"}

        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": f"병합 중 오류가 발생했습니다:\n{str(e)}"}