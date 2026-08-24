import os
import webview

class RenameAPI:
    def select_files(self, window):
        """파일 선택 다이얼로그 호출"""
        try:
            result = window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True)
            return result if result else []
        except Exception as e:
            print(f"파일 선택 오류: {e}")
            return []

    def execute_rename(self, payload):
        """파일 이름 실제 변경 실행"""
        success_cnt = 0
        fail_cnt = 0
        errors = []
        results = [] # [추가] 실제 변경된 경로 추적
        
        for item in payload:
            old_path = item.get('oldPath')
            new_path = item.get('newPath')
            
            if not old_path or not new_path or old_path == new_path:
                continue
                
            try:
                # 덮어쓰기 방지: 같은 이름이 있으면 (1), (2) 식으로 자동 변경
                actual_new_path = new_path
                if os.path.exists(actual_new_path):
                    base, ext = os.path.splitext(actual_new_path)
                    idx = 1
                    while os.path.exists(f"{base}({idx}){ext}"):
                        idx += 1
                    actual_new_path = f"{base}({idx}){ext}"
                    
                os.rename(old_path, actual_new_path)
                success_cnt += 1
                
                # [추가] 변경 성공 시 결과 리스트에 기록
                results.append({
                    "oldPath": old_path,
                    "expectedPath": new_path,
                    "actualPath": actual_new_path
                })
                
            except Exception as e:
                fail_cnt += 1
                errors.append(f"{os.path.basename(old_path)} 오류: {str(e)}")
                
        return {
            "success": True, 
            "success_cnt": success_cnt, 
            "fail_cnt": fail_cnt, 
            "errors": errors,
            "results": results # [추가] 프론트엔드로 전달
        }