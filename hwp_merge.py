# -*- coding: utf-8 -*-
# hwp_merge.py

import os
import time
import webview

# [중요] 한글 제어 라이브러리 설정
try:
    import win32com.client
    import pythoncom 
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

class HwpMergeService:
    def select_files(self, window):
        """[파일 열기] 다이얼로그"""
        file_types = ('Hangul Files (*.hwp;*.hwpx)', 'All files (*.*)')
        result = window.create_file_dialog(
            webview.OPEN_DIALOG, 
            allow_multiple=True, 
            file_types=file_types
        )
        return result
    
    def merge_hwp_files(self, file_paths):
        """
        [핵심 로직] 액션(Action) 기반 병합 후 화면 표시
        저장하지 않고 한글 창을 띄웁니다.
        """
        if not file_paths:
            return {"success": False, "message": "취합할 파일이 없습니다."}
        
        if not HAS_WIN32:
            return {
                "success": False, 
                "message": "필수 라이브러리(pywin32)가 없습니다.\n터미널에 'pip install pywin32'를 입력해주세요."
            }

        hwp = None
        try:
            # 1. 스레드 초기화
            pythoncom.CoInitialize()

            # 2. 한글 프로그램 실행 (gencache -> Dispatch 순서로 시도)
            try:
                hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                except Exception as e:
                    return {"success": False, "message": f"한글 프로그램을 실행할 수 없습니다.\n{e}"}

            # 3. 보안 모듈 승인
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            # [핵심] 한글 창을 화면에 보이게 설정
            hwp.XHwpWindows.Item(0).Visible = True
            
            hwp.Clear(1) # 빈 문서로 초기화

            # 4. 파일 순서대로 합치기 (안정적인 Action 방식 사용)
            for index, path in enumerate(file_paths):
                # (1) 액션 생성
                act = hwp.CreateAction("InsertFile")
                pset = act.CreateSet()
                act.GetDefault(pset)
                
                # (2) 파라미터 설정
                pset.SetItem("FileName", path)
                pset.SetItem("KeepSection", 1)   
                pset.SetItem("KeepCharshape", 1) 
                pset.SetItem("KeepParashape", 1) 
                pset.SetItem("KeepStyle", 1)     
                
                # (3) 실행
                act.Execute(pset)
                
                # (4) 문서 끝으로 이동
                hwp.Run("MoveDocEnd")
                
                # (5) 쪽 나누기 (마지막 파일 제외)
                if index < len(file_paths) - 1:
                    hwp.Run("BreakPage")

            # 저장 단계 없이 창을 띄운 상태로 종료
            return {"success": True, "message": "한글 프로그램 창이 열렸습니다."}

        except Exception as e:
            if hwp:
                try: hwp.Quit() 
                except: pass
            return {"success": False, "message": f"작업 중 오류 발생: {str(e)}"}