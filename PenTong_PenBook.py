import os
import json
import zipfile
import base64
import mimetypes

class PenBookService:
    def __init__(self, data_dir):
        # PenTong에서 넘겨주는 전용 폴더 경로를 사용합니다.
        self.data_dir = data_dir
        os.makedirs(self.data_dir, exist_ok=True)
        self.active_id = None
        self.meta = {}
        self.pages = {}
        self.images = {}

    def get_registry(self):
        registry = {}
        for f in os.listdir(self.data_dir):
            if f.endswith('.penbook'):
                proj_id = f[:-8]
                registry[proj_id] = {'name': proj_id, 'type': 'note'}
        return registry

    def _save_to_zip(self, project_id):
        import zipfile
        import json
        import uuid
        import os
        
        zip_path = os.path.join(self.data_dir, f'{project_id}.penbook')
        # 💡 [안전조치] 동시 저장 충돌을 막기 위해 매번 고유한 임시 파일(UUID)을 생성합니다.
        tmp_path = zip_path + f'.{uuid.uuid4().hex}.tmp'
        
        try:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.writestr('meta.json', json.dumps(self.meta, ensure_ascii=False, indent=2))
                for node_id, html in self.pages.items():
                    zf.writestr(f'{node_id}.html', html)
                for filename, b_data in self.images.items():
                    zf.writestr(f'images/{filename}', b_data)
            
            # 💡 [안전조치] os.rename 대신 덮어쓰기를 완벽 지원하는 os.replace 사용
            if os.path.exists(tmp_path):
                os.replace(tmp_path, zip_path)
                
        except Exception as e:
            print(f"[💥 치명적 오류] 프로젝트 물리 저장 실패: {e}")
            
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except:
                    pass

    def create_project(self, name, proj_type='note'):
        safe_name = "".join(c for c in name if c.isalnum() or c in " _-가-힣").strip()
        if not safe_name: safe_name = "Project"
        proj_id = safe_name
        i = 1
        while os.path.exists(os.path.join(self.data_dir, f"{proj_id}.penbook")):
            proj_id = f"{safe_name}_{i}"
            i += 1

        self.active_id = proj_id
        self.meta = {'name': proj_id, 'type': proj_type, 'toc': []}
        self.pages = {}
        self.images = {}
        self._save_to_zip(proj_id)
        return {'status': 'success', 'id': proj_id}

    # ── 프로젝트 관리 및 원자적(Atomic) 데이터 철통 방어 보호 메커니즘 ──────────
    def open_project(self, project_id):
        import zipfile
        import json
        
        zip_path = os.path.join(self.data_dir, f'{project_id}.penbook')
        self.active_id = project_id
        self.meta = {'name': project_id, 'type': 'note', 'toc': []}
        self.pages = {}
        self.images = {}
        
        if os.path.exists(zip_path):
            try:
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    # 1. 목차 메타데이터 구조 로드
                    if 'meta.json' in zf.namelist():
                        self.meta = json.loads(zf.read('meta.json').decode('utf-8'))
                    
                    # 💡 [안전조치 1] 데이터 유실 원천 차단!
                    # 압축 파일 내의 모든 HTML 본문과 이미지를 선제적으로 전부 메모리에 완전히 세팅해 둡니다.
                    for name in zf.namelist():
                        if name.endswith('.html'):
                            node_id = name[:-5] # .html 확장자 제거
                            self.pages[node_id] = zf.read(name).decode('utf-8')
                        elif name.startswith('images/') and name != 'images/':
                            filename = name[7:] # images/ 경로 제거
                            self.images[filename] = zf.read(name)
                            
                print(f"[open_project] 보안 로드 완료: {project_id} (본문: {len(self.pages)}개, 이미지: {len(self.images)}개 세이브 가드 작동)")
            except Exception as e:
                print(f"[오류] 프로젝트 로드 중 실패: {e}")
        return self.meta

    def delete_project(self, project_id):
        zip_path = os.path.join(self.data_dir, f'{project_id}.penbook')
        if os.path.exists(zip_path): os.remove(zip_path)
        if self.active_id == project_id:
            self.active_id = None
            self.meta = {}
            self.pages = {}
            self.images = {}
        return {'status': 'success'}

    def get_page(self, project_id, node_id):
        if project_id != self.active_id: self.open_project(project_id)
        # 💡 이미 프로젝트를 열 때 안전 가드 딕셔너리에 다 로드해두었으므로 메모리에서 즉시 안전하게 반환합니다.
        return self.pages.get(node_id, "")

    def save_page(self, project_id, node_id, html):
        if project_id != self.active_id: self.open_project(project_id)
        self.pages[node_id] = html
        self._save_to_zip(project_id)
        return {'status': 'success'}

    def delete_page(self, project_id, node_id):
        if project_id != self.active_id: self.open_project(project_id)
        if node_id in self.pages:
            del self.pages[node_id]
        self._save_to_zip(project_id)
        return {'status': 'success'}

    def save_meta(self, project_id, meta):
        if project_id != self.active_id: self.open_project(project_id)
        self.meta = meta
        self._save_to_zip(project_id)
        return {'status': 'success'}

    def get_images(self, project_id):
        if project_id != self.active_id: self.open_project(project_id)
        return list(self.images.keys())

    def upload_image(self, project_id, filename, base64_data):
        if project_id != self.active_id: self.open_project(project_id)
        if ',' in base64_data: base64_data = base64_data.split(',', 1)[1]
        self.images[filename] = base64.b64decode(base64_data)
        self._save_to_zip(project_id)
        return {'status': 'success', 'path': f'images/{filename}'}

    def get_image_base64(self, project_id, filename):
        if project_id != self.active_id: self.open_project(project_id)
        if filename in self.images:
            data = base64.b64encode(self.images[filename]).decode('utf-8')
            mime, _ = mimetypes.guess_type(filename)
            if not mime: mime = 'image/png'
            return f"data:{mime};base64,{data}"
        return ""

    # ── 이미지 관리 (압축 파일 내부 images 폴더) ──────────
    def delete_image(self, project_id, filename):
        if project_id != self.active_id: self.open_project(project_id)
        if filename in self.images:
            del self.images[filename]
            self._save_to_zip(project_id) # 메모리에서 지우고 즉시 압축 파일 덮어쓰기
        print(f'[delete_image] 삭제됨: images/{filename}')
        return {'status': 'success'}

    def export_image(self, window, project_id, filename):
        import webview
        import os
        if project_id != self.active_id: self.open_project(project_id)
        
        if filename in self.images:
            # 원본 확장자 추출
            ext = os.path.splitext(filename)[1].lower()
            if not ext: ext = '.png'
            
            # 💡 윈도우 기본 '다른 이름으로 저장' 다이얼로그 호출
            file_types = (f'Image Files (*{ext})', 'All files (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename=filename)
            
            if save_path:
                target_path = save_path[0] if isinstance(save_path, (list, tuple)) else save_path
                with open(target_path, 'wb') as f:
                    f.write(self.images[filename])
                return {'status': 'success'}
                
        return {'status': 'canceled'}

    def replace_image(self, project_id, filename, base64_data):
        if project_id != self.active_id: self.open_project(project_id)
        if filename in self.images:
            import os
            # 확장자와 순수 파일명 분리 (예: test.png -> test, .png)
            name, ext = os.path.splitext(filename)
            
            # 💡 기존 파일 백업할 유니크한 이름 조합 생성 (예: test_교체됨.png)
            backup_filename = f"{name}_교체됨{ext}"
            counter = 1
            while backup_filename in self.images:
                backup_filename = f"{name}_교체됨_{counter}{ext}"
                counter += 1
                
            # 1. 기존 이미지 데이터를 백업 키값으로 복사 보관
            self.images[backup_filename] = self.images[filename]
            
            # 2. 새 이미지 베이스64 데이터를 디코딩하여 원래 파일명 자리에 덮어쓰기
            if ',' in base64_data: 
                base64_data = base64_data.split(',', 1)[1]
            self.images[filename] = base64.b64decode(base64_data)
            
            # 3. 변경 사항 메모리 반영 후 즉시 Zip 압축파일 업데이트 저장
            self._save_to_zip(project_id)
            print(f'[replace_image] 완료: {filename} -> 백업생성: {backup_filename}')
            return {'status': 'success'}
            
        return {'status': 'error', 'message': '원본 파일을 찾을 수 없습니다.'}

    def load_settings(self):
        path = os.path.join(self.data_dir, 'settings.json')
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f: return json.load(f)
            except: pass
        return {}

    def save_settings(self, settings):
        path = os.path.join(self.data_dir, 'settings.json')
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
        return {'status': 'success'}

    # ── [추가] 학생 배포용 단일 HTML 바탕화면 내보내기 엔진 ──────────
    def export_standalone_html(self, window, filename, content):
        import webview
        import os
        try:
            # 1. 선생님의 윈도우 바탕화면 경로를 자동으로 찾아냅니다.
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            
            # 2. PyWebView 표준 윈도우 [다른 이름으로 저장] 창을 띄웁니다!
            file_path = window.create_file_dialog(
                webview.SAVE_DIALOG, 
                directory=desktop_path, 
                save_filename=filename,
                file_types=('HTML Files (*.html;*.htm)', 'All files (*.*)')
            )
            
            if not file_path:
                return {"status": "cancel"} # 사용자가 창을 그냥 닫거나 취소한 경우
                
            # 3. 지정한 경로에 한글 깨짐 없이(utf-8) 완벽하게 파일로 구워냅니다.
            target_path = file_path[0] if isinstance(file_path, (list, tuple)) else file_path
            with open(target_path, "w", encoding="utf-8") as f:
                f.write(content)
                
            return {"status": "success"}
            
        except Exception as e:
            print(f"[Export Error] 단일 HTML 추출 중 치명적 오류 발생: {e}")
            return {"status": "error", "message": str(e)}


    # ── 궁극의 클립보드 스캐너 (브라우저 보안 팝업 원천 차단) ──────────
    def read_clipboard(self):
        import base64
        import io
        res = {'text': None, 'html': None, 'image': None, 'status': 'success'}
        
        # 1. 텍스트 및 HTML 읽기 (win32clipboard 활용)
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            
            # HTML 코드가 복사되었는지 확인
            try:
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                if win32clipboard.IsClipboardFormatAvailable(cf_html):
                    data = win32clipboard.GetClipboardData(cf_html)
                    html_str = data.decode('utf-8', 'ignore')
                    
                    # 윈도우 클립보드의 불필요한 메타 헤더 자르기
                    if "StartHTML:" in html_str:
                        import re
                        match = re.search(r'StartHTML:(\d+)', html_str)
                        if match:
                            res['html'] = html_str[int(match.group(1)):]
                    else:
                        res['html'] = html_str
            except: pass
            
            # 순수 텍스트 확인
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    res['text'] = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            except: pass
            
            win32clipboard.CloseClipboard()
        except Exception as e:
            # 실패 시 파이썬 내장 라이브러리로 텍스트 강제 획득
            try:
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                res['text'] = root.clipboard_get()
                root.destroy()
            except: pass

        # 2. 이미지 읽기 (Pillow 활용)
        try:
            from PIL import ImageGrab, Image
            import os
            
            img = ImageGrab.grabclipboard()
            
            # 이미지 파일 자체를 복사(파일 복사)했을 때 이미지로 읽어들이기
            if isinstance(img, list) and len(img) > 0 and isinstance(img[0], str):
                if os.path.exists(img[0]) and img[0].lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                    img = Image.open(img[0])
                    
            if img and hasattr(img, 'save'):
                buffer = io.BytesIO()
                # 투명도 및 포맷 유지 처리 후 Base64 인코딩
                if img.mode != 'RGB' and img.mode != 'RGBA':
                    img = img.convert('RGBA')
                img.save(buffer, format='PNG')
                res['image'] = base64.b64encode(buffer.getvalue()).decode('utf-8')
        except: pass

        return res

    # ── 프로젝트 공유 (가져오기 / 내보내기) 엔진 ──────────
    def export_project(self, window, project_id):
        import shutil
        import os
        import webview
        
        # 1. 내보내기 전, 현재 켜둔 프로젝트라면 최신본으로 강제 디스크 저장
        if project_id != self.active_id:
            self.open_project(project_id)
        else:
            self._save_to_zip(project_id)
            
        # 2. 파일명 예쁘게 정제
        project_name = self.meta.get('name', project_id)
        safe_name = "".join(c for c in project_name if c.isalnum() or c in " _-가-힣").strip()
        if not safe_name: safe_name = "Project"
        default_filename = f"{safe_name}.penbook"
        
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        
        try:
            # 바탕화면 기본 경로로 '다른 이름으로 저장' 팝업 띄우기
            file_path = window.create_file_dialog(
                webview.SAVE_DIALOG, 
                directory=desktop_path, 
                save_filename=default_filename,
                file_types=('PenBook Project (*.penbook)', 'All files (*.*)')
            )
            
            if not file_path:
                return {"status": "cancel"}
                
            target_path = file_path[0] if isinstance(file_path, (list, tuple)) else file_path
            src_path = os.path.join(self.data_dir, f'{project_id}.penbook')
            
            if os.path.exists(src_path):
                shutil.copy2(src_path, target_path) # 안전하게 파일 복제
                return {"status": "success"}
            else:
                return {"status": "error", "message": "원본 프로젝트 파일을 찾을 수 없습니다."}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def import_project(self, window):
        import shutil
        import os
        import webview
        import zipfile
        import json
        
        try:
            # 파일 선택 팝업 띄우기
            file_path = window.create_file_dialog(
                webview.OPEN_DIALOG, 
                file_types=('PenBook Project (*.penbook)', 'All files (*.*)')
            )
            
            if not file_path:
                return {"status": "cancel"}
                
            src_path = file_path[0] if isinstance(file_path, (list, tuple)) else file_path
            
            # 유효한 프로젝트 파일인지 압축 풀기 전 1차 검사
            if not zipfile.is_zipfile(src_path):
                return {"status": "error", "message": "올바른 PenBook 프로젝트 파일이 아닙니다."}
                
            with zipfile.ZipFile(src_path, 'r') as zf:
                if 'meta.json' not in zf.namelist():
                    return {"status": "error", "message": "메타데이터가 없는 손상된 프로젝트입니다."}
                meta = json.loads(zf.read('meta.json').decode('utf-8'))
                
            # 💡 [핵심] 기존에 있는 프로젝트와 이름이 충돌하지 않도록 안전한 새 ID 발급
            base_name = meta.get('name', '가져온 프로젝트')
            safe_name = "".join(c for c in base_name if c.isalnum() or c in " _-가-힣").strip()
            if not safe_name: safe_name = "Project"
            
            proj_id = safe_name
            i = 1
            while os.path.exists(os.path.join(self.data_dir, f"{proj_id}.penbook")):
                proj_id = f"{safe_name}_{i}"
                i += 1
                
            dest_path = os.path.join(self.data_dir, f"{proj_id}.penbook")
            shutil.copy2(src_path, dest_path)
            
            # 💡 [안전장치] 압축 파일 내부의 meta.json을 뜯어서 내부 ID도 새 ID로 통일시켜 충돌을 완벽 차단!
            with zipfile.ZipFile(dest_path, 'r') as zin:
                item_dict = {item.filename: zin.read(item.filename) for item in zin.filelist}
            
            meta_dict = json.loads(item_dict['meta.json'].decode('utf-8'))
            meta_dict['name'] = proj_id
            item_dict['meta.json'] = json.dumps(meta_dict, ensure_ascii=False, indent=2).encode('utf-8')
            
            with zipfile.ZipFile(dest_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for fname, bdata in item_dict.items():
                    zout.writestr(fname, bdata)
            
            # 원본 프로젝트의 이름표(name)를 프론트엔드로 전달
            return {"status": "success", "id": proj_id, "name": base_name}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def rename_project(self, old_id, new_name):
        import os
        import json
        import zipfile
        
        # 파일명으로 쓸 수 없는 특수문자 제거
        safe_name = "".join(c for c in new_name if c.isalnum() or c in " _-가-힣").strip()
        if not safe_name: safe_name = "Project"
        
        new_id = safe_name
        i = 1
        # 충돌 방지: 새로운 이름이 이미 존재하면 (2), (3) 번호 자동 부여
        while os.path.exists(os.path.join(self.data_dir, f"{new_id}.penbook")) and new_id != old_id:
            new_id = f"{safe_name}_{i}"
            i += 1
            
        old_path = os.path.join(self.data_dir, f"{old_id}.penbook")
        new_path = os.path.join(self.data_dir, f"{new_id}.penbook")
        
        try:
            # 1. 파일 이름 바꾸기 (이름이 달라진 경우만 물리적 이동)
            if new_id != old_id:
                # 현재 열려있는 프로젝트라면 안전하게 메모리 상태를 디스크에 먼저 저장
                if self.active_id == old_id:
                    self._save_to_zip(old_id)
                os.rename(old_path, new_path)
            
            # 2. 파일명과 관계없이 압축파일(zip) 내부의 meta.json 도 완벽하게 업데이트!
            with zipfile.ZipFile(new_path, 'r') as zin:
                item_dict = {item.filename: zin.read(item.filename) for item in zin.filelist}
                
            if 'meta.json' in item_dict:
                meta_dict = json.loads(item_dict['meta.json'].decode('utf-8'))
                meta_dict['name'] = new_name
                item_dict['meta.json'] = json.dumps(meta_dict, ensure_ascii=False, indent=2).encode('utf-8')
                
                with zipfile.ZipFile(new_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for fname, bdata in item_dict.items():
                        zout.writestr(fname, bdata)
            
            # 3. 백엔드 메모리 상태 갱신
            if self.active_id == old_id:
                if new_id != old_id:
                    self.active_id = new_id
                self.meta['name'] = new_name
                
            return {"status": "success", "new_id": new_id}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def transfer_node(self, src_proj, dst_proj, node_id, mode):
        import copy
        import string
        import random
        
        try:
            # 원본 프로젝트 열기
            if self.active_id != src_proj:
                self.open_project(src_proj)
                
            # 선택한 항목과 그 아래에 달린 하위 폴더들까지 싹쓸이 추적
            def get_descendants(n_id):
                desc = [n_id]
                for child in [n for n in self.meta.get('toc', []) if n.get('parentId') == n_id]:
                    desc.extend(get_descendants(child['id']))
                return desc
            
            nodes_to_transfer = get_descendants(node_id)
            toc_items = [n for n in self.meta.get('toc', []) if n['id'] in nodes_to_transfer]
            
            # 충돌 방지를 위해 타겟 프로젝트용 새로운 고유 ID 발급
            def uid(): return ''.join(random.choices(string.ascii_lowercase + string.digits, k=8))
            id_map = {n: uid() for n in nodes_to_transfer}
            
            new_toc_items = []
            for item in toc_items:
                new_item = copy.deepcopy(item)
                new_item['id'] = id_map[item['id']]
                if new_item.get('parentId') in id_map:
                    new_item['parentId'] = id_map[new_item['parentId']]
                elif item['id'] == node_id:
                    new_item['parentId'] = None # 타겟 프로젝트의 최상위 레벨(맨 아래)로 들어감
                new_toc_items.append(new_item)
            
            pages_data = {id_map[n]: self.pages.get(n, "") for n in nodes_to_transfer}
            images_data = copy.deepcopy(self.images)
            
            # 이동(Move)인 경우, 원본 프로젝트에서 삭제하고 저장
            if mode == 'move':
                self.meta['toc'] = [n for n in self.meta['toc'] if n['id'] not in nodes_to_transfer]
                for n in nodes_to_transfer:
                    if n in self.pages:
                        del self.pages[n]
                self._save_to_zip(src_proj)
                
            # 타겟 프로젝트를 열어서 데이터 주입
            self.open_project(dst_proj)
            if 'toc' not in self.meta: self.meta['toc'] = []
            self.meta['toc'].extend(new_toc_items)
            
            for n, content in pages_data.items():
                self.pages[n] = content
            for img_name, img_bytes in images_data.items():
                if img_name not in self.images:
                    self.images[img_name] = img_bytes
                    
            self._save_to_zip(dst_proj)
            
            # 작업이 끝났으니 원래 화면에 켜져 있던 프로젝트로 조용히 복귀
            self.open_project(src_proj)
            
            return {"status": "success"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def import_project_from_data(self, filename, base64_data):
        import shutil
        import os
        import zipfile
        import json
        import base64
        import tempfile
        
        try:
            # 브라우저에서 넘어온 Base64 파일 데이터에서 헤더를 떼어내고 디코딩합니다.
            if ',' in base64_data:
                base64_data = base64_data.split(',', 1)[1]
                
            temp_dir = tempfile.gettempdir()
            src_path = os.path.join(temp_dir, f"temp_{filename}")
            
            with open(src_path, 'wb') as f:
                f.write(base64.b64decode(base64_data))
            
            # 유효한 프로젝트 파일인지 압축 풀기 전 1차 검사
            if not zipfile.is_zipfile(src_path):
                return {"status": "error", "message": "올바른 PenBook 프로젝트 파일이 아닙니다."}
                
            with zipfile.ZipFile(src_path, 'r') as zf:
                if 'meta.json' not in zf.namelist():
                    return {"status": "error", "message": "메타데이터가 없는 손상된 프로젝트입니다."}
                meta = json.loads(zf.read('meta.json').decode('utf-8'))
                
            # 이름 중복 충돌 방지를 위해 안전한 새 ID 발급
            base_name = meta.get('name', filename.replace('.penbook', ''))
            safe_name = "".join(c for c in base_name if c.isalnum() or c in " _-가-힣").strip()
            if not safe_name: safe_name = "Project"
            
            proj_id = safe_name
            i = 1
            while os.path.exists(os.path.join(self.data_dir, f"{proj_id}.penbook")):
                proj_id = f"{safe_name}_{i}"
                i += 1
                
            dest_path = os.path.join(self.data_dir, f"{proj_id}.penbook")
            shutil.copy2(src_path, dest_path)
            
            # 압축 파일 내부의 meta.json 도 새 ID로 통일시켜 충돌 원천 차단!
            with zipfile.ZipFile(dest_path, 'r') as zin:
                item_dict = {item.filename: zin.read(item.filename) for item in zin.filelist}
            
            meta_dict = json.loads(item_dict['meta.json'].decode('utf-8'))
            meta_dict['name'] = proj_id
            item_dict['meta.json'] = json.dumps(meta_dict, ensure_ascii=False, indent=2).encode('utf-8')
            
            with zipfile.ZipFile(dest_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for fname, bdata in item_dict.items():
                    zout.writestr(fname, bdata)
            
            try: os.remove(src_path)
            except: pass
            
            return {"status": "success", "id": proj_id, "name": base_name}
        except Exception as e:
            return {"status": "error", "message": str(e)}