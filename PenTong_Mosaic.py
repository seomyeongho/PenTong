# -*- coding: utf-8 -*-
import os
import base64
import webview

class MosaicAPI:
    def __init__(self, window):
        self.window = window

    # 1. 폴더 선택 및 이미지 파일 목록 반환
    def open_folder(self):
        folder_path = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        
        if not folder_path:
            return {"status": "cancel"}
            
        folder_path = folder_path[0]
        # 웹 브라우저에서 바로 띄울 수 있는 이미지 확장자들
        image_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.webp')
        image_files = []
        
        for file in os.listdir(folder_path):
            if file.lower().endswith(image_extensions):
                image_files.append({
                    "name": file,
                    "path": os.path.join(folder_path, file)
                })
                
        if not image_files:
            return {"status": "empty"}
            
        return {"status": "success", "files": image_files, "folder": folder_path}
    
    def add_files(self):
        file_types = ('Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.webp)', 'All files (*.*)')
        files = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True, file_types=file_types)
        
        if not files:
            return {"status": "cancel"}
            
        image_files = []
        for file_path in files:
            image_files.append({
                "name": os.path.basename(file_path),
                "path": file_path
            })
        return {"status": "success", "files": image_files}

    def select_save_folder(self):
        folder = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if folder:
            return {"status": "success", "path": folder[0]}
        return {"status": "cancel"}

    # 2. 선택한 이미지를 Base64 문자열(Data URI)로 브라우저에 전달
    def load_image_data(self, file_path):
        try:
            with open(file_path, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                ext = os.path.splitext(file_path)[1].lower().replace('.', '')
                if ext == 'jpg': ext = 'jpeg'
                
                return {"status": "success", "data_uri": f"data:image/{ext};base64,{encoded_string}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    # 3. 작업이 완료된 모자이크 이미지를 '모자이크_완료' 폴더에 저장
    def save_image_data(self, payload):
        try:
            original_path = payload.get("file_path")
            b64_data = payload.get("b64_data")
            
            original_name = os.path.basename(original_path)
            name, ext = os.path.splitext(original_name)
            file_name = f"{name}_Mosaicized{ext}"
            
            save_dir = payload.get("save_dir")
            if not save_dir:
                return {"status": "error", "message": "저장 폴더가 지정되지 않았습니다."}
                
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
                
            save_path = os.path.join(save_dir, file_name)
            
            # HTML 캔버스에서 넘어온 Base64에서 헤더 부분 분리 후 저장
            header, encoded = b64_data.split(",", 1)
            with open(save_path, "wb") as f:
                f.write(base64.b64decode(encoded))
                
            return {"status": "success", "save_path": save_path}
        except Exception as e:
            return {"status": "error", "message": str(e)}