# -*- coding: utf-8 -*-
import os
import platform
import subprocess
import traceback
import webbrowser
import webview

class ShortcutAPI:
    def __init__(self, data_dir):
        # PenTong.py에서 전달받은 PenTong_Data 경로를 사용합니다.
        self.data_dir = data_dir

    def load_shortcut_data(self):
        file_path = os.path.join(self.data_dir, "shortcuts_data.json")
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            return None
        except Exception as e:
            print(f"바로가기 데이터 로드 에러: {e}")
            return None

    def save_shortcut_data(self, json_data):
        file_path = os.path.join(self.data_dir, "shortcuts_data.json")
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(json_data)
            return True
        except Exception as e:
            print(f"바로가기 데이터 저장 에러: {e}")
            return False

    def execute_shortcut(self, target_path):
        try:
            # 인터넷 URL인 경우 웹 브라우저로 열기
            if target_path.startswith("http://") or target_path.startswith("https://"):
                webbrowser.open(target_path)
                return {"success": True}
            
            # 로컬 파일/폴더인 경우 존재 여부 확인
            if not os.path.exists(target_path):
                return {"success": False, "message": "경로를 찾을 수 없습니다. 파일이나 폴더가 이동/삭제되었을 수 있습니다."}
            
            # 운영체제에 맞는 파일/폴더 실행 (윈도우: os.startfile)
            if platform.system() == 'Windows':
                os.startfile(target_path)
            elif platform.system() == 'Darwin':
                subprocess.call(('open', target_path))
            else:
                subprocess.call(('xdg-open', target_path))
            return {"success": True}
            
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": f"실행 오류: {str(e)}"}

    def select_local_path(self, path_type, window):
        if not window: return None
        try:
            if path_type == 'folder':
                result = window.create_file_dialog(webview.FOLDER_DIALOG)
            else:
                # 파일이나 프로그램
                file_types = ('모든 파일 (*.*)',)
                if path_type == 'program':
                    file_types = ('실행 파일 (*.exe;*.bat;*.cmd;*.lnk)', '모든 파일 (*.*)')
                result = window.create_file_dialog(webview.OPEN_DIALOG, file_types=file_types)
                
            if result and len(result) > 0:
                return result[0]
            return None
        except Exception as e:
            traceback.print_exc()
            return None
