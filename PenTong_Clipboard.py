# -*- coding: utf-8 -*-
import os
import json
import threading
import time
import uuid
import base64
import hashlib
import datetime
from io import BytesIO

try:
    import win32clipboard
    import win32con
except ImportError:
    print("[경고] pywin32 라이브러리가 없습니다.")

try:
    from PIL import ImageGrab, Image
except ImportError:
    print("[경고] Pillow 모듈이 없습니다.")

class ClipboardService:
    # window=None 매개변수를 추가하고 내부 변수로 저장합니다.
    def __init__(self, data_dir, window=None): # ◀ window 추가
        self.data_dir = data_dir
        self.window = window # ◀ 추가
        self.settings_file = os.path.join(self.data_dir, "clipboard_settings.json")
        self.history_file = os.path.join(self.data_dir, "clipboard_history.json") # ◀ 🛑 이 줄을 추가해 주세요!

        self.max_items = 100 
        
        self.is_running = False
        self.monitor_thread = None
        self.last_seq = 0
        self.last_hash = None
        
        self.settings = self._load_settings()
        self._cleanup_expired_items() # 시작 시 기간 지난 항목 청소
        
        if self.settings.get("enabled", False):
            self.start_monitoring()

    def _load_settings(self):
        default_settings = {"enabled": False, "retention_days": 1}
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    default_settings.update(data)
            except: pass
        return default_settings

    def _save_settings(self):
        with open(self.settings_file, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f)

    def get_status(self):
        return {"enabled": self.is_running, "retention_days": self.settings.get("retention_days", 1)}

    def toggle(self, state):
        self.settings["enabled"] = state
        self._save_settings()
        if state and not self.is_running:
            self.start_monitoring()
        elif not state and self.is_running:
            self.is_running = False
        return {"success": True, "enabled": self.is_running}

    def set_retention(self, days):
        self.settings["retention_days"] = int(days)
        self._save_settings()
        self._cleanup_expired_items()
        return {"success": True}

    def start_monitoring(self):
        if self.is_running: return
        self.is_running = True
        try: self.last_seq = win32clipboard.GetClipboardSequenceNumber()
        except: pass
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()

    def _monitor_loop(self):
        while self.is_running:
            try:
                current_seq = win32clipboard.GetClipboardSequenceNumber()
                if current_seq != self.last_seq:
                    self.last_seq = current_seq
                    time.sleep(0.3) 
                    self._process_clipboard()
            except Exception: pass
            time.sleep(0.5)

    def _process_clipboard(self):
        raw_data_dict = {}
        preview_type = "unknown"
        preview_content = "복합 개체 (PPT, 엑셀, 원본 이미지 등)"

        try: win32clipboard.OpenClipboard()
        except Exception: return

        try:
            fmt = 0
            while True:
                fmt = win32clipboard.EnumClipboardFormats(fmt)
                if fmt == 0: break
                try:
                    data = win32clipboard.GetClipboardData(fmt)
                    if data is None: continue
                    if isinstance(data, str):
                        raw_data_dict[str(fmt)] = {"type": "string", "data": data}
                    elif isinstance(data, tuple):
                        raw_data_dict[str(fmt)] = {"type": "tuple", "data": list(data)}
                    else:
                        raw_data_dict[str(fmt)] = {"type": "bytes", "data": base64.b64encode(data).decode('utf-8')}
                except Exception: pass

            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                try:
                    text_data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    if text_data and text_data.strip():
                        preview_type = "text"
                        preview_content = text_data
                except: pass
            elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                try:
                    files = win32clipboard.GetClipboardData(win32con.CF_HDROP)
                    preview_type = "file"
                    preview_content = "📁 복사된 파일/폴더:\n" + "\n".join(files)
                except: pass
        finally:
            win32clipboard.CloseClipboard()

        if preview_type in ["unknown", "text"]:
            try:
                img = ImageGrab.grabclipboard()
                if img and isinstance(img, Image.Image):
                    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P': img = img.convert('RGBA')
                        bg.paste(img, mask=img.split()[-1]) 
                        img = bg
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                    buffered = BytesIO()
                    img.save(buffered, format="JPEG", quality=85)
                    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
                    preview_type = "image"
                    preview_content = f"data:image/jpeg;base64,{img_str}"
            except Exception: pass

        if raw_data_dict:
            # 완벽한 중복 검사를 위한 MD5 해시 생성
            json_str = json.dumps(raw_data_dict, sort_keys=True)
            data_hash = hashlib.md5(json_str.encode('utf-8')).hexdigest()
            
            if data_hash == self.last_hash: return
            self.last_hash = data_hash
            self._add_item(preview_type, preview_content, raw_data_dict, data_hash)
            # --- [이 부분 추가] 새로운 복사가 감지되면 UI iframe을 즉시 원격 새로고침 ---
            if self.window:
                try:
                    js_code = """
                    var frame = document.getElementById('frame-다중클립보드');
                    if (frame && frame.contentWindow && typeof frame.contentWindow.loadHistory === 'function') {
                        frame.contentWindow.loadHistory();
                    }
                    """
                    self.window.evaluate_js(js_code)
                except Exception as e:
                    print(f"[경고] UI 실시간 갱신 실패: {e}")

    def _add_item(self, p_type, p_content, raw_data, data_hash):
        history = self._read_full_history(skip_cleanup=True)
        
        # [수정] 동일한 해시(내용)가 있는지 검사하여 기존 것 삭제 (영구보관은 건드리지 않음)
        is_already_permanent = any(item.get('hash') == data_hash and item.get('is_permanent') for item in history)
        if is_already_permanent:
            return # 이미 영구보관함에 있는 내용이면 최근 목록에 추가하지 않음
            
        # 기존 최근 목록에 동일한 데이터가 있다면 삭제 (최상단으로 끌어올리기 위함)
        history = [item for item in history if item.get('hash') != data_hash]

        history.insert(0, {
            "id": str(uuid.uuid4()),
            "type": p_type,
            "content": p_content,
            "raw_data": raw_data,
            "hash": data_hash,
            "is_permanent": False,
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        history = history[:self.max_items]
        self._save_history(history)

    def _cleanup_expired_items(self):
        history = self._read_full_history(skip_cleanup=True)
        retention_days = self.settings.get("retention_days", 1)
        cutoff = datetime.datetime.now() - datetime.timedelta(days=retention_days)
        
        cleaned = []
        changed = False
        for item in history:
            if item.get("is_permanent", False):
                cleaned.append(item)
            else:
                try:
                    item_time = datetime.datetime.strptime(item["timestamp"], "%Y-%m-%d %H:%M:%S")
                    if item_time >= cutoff:
                        cleaned.append(item)
                    else:
                        changed = True
                except:
                    cleaned.append(item)
                    
        if changed: self._save_history(cleaned)
        return cleaned

    def _read_full_history(self, skip_cleanup=False):
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return []

    def _save_history(self, history):
        with open(self.history_file, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False)

    def get_history(self):
        full_history = self._cleanup_expired_items()
        return [{
            "id": item["id"], 
            "type": item["type"], 
            "content": item["content"], 
            "is_permanent": item.get("is_permanent", False),
            "timestamp": item["timestamp"]
        } for item in full_history]

    def make_permanent(self, item_id):
        history = self._read_full_history(skip_cleanup=True)
        for item in history:
            if item['id'] == item_id:
                item['is_permanent'] = True
                item['timestamp'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                break
        self._save_history(history)
        return {"success": True}

    def copy_item(self, item_id):
        full_history = self._read_full_history(skip_cleanup=True)
        for item in full_history:
            if item['id'] == item_id:
                raw_data = item.get('raw_data', {})
                if not raw_data: return {"success": False, "message": "원본 데이터 손실"}

                try:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                except Exception as e:
                    return {"success": False, "message": f"클립보드 접근 거부: {str(e)}"}

                success_count = 0
                errors = []
                for fmt_str, obj in raw_data.items():
                    fmt = int(fmt_str)
                    dtype = obj.get("type")
                    val = obj.get("data")
                    try:
                        if dtype == "string": win32clipboard.SetClipboardData(fmt, val)
                        elif dtype == "tuple": win32clipboard.SetClipboardData(fmt, tuple(val))
                        else: win32clipboard.SetClipboardData(fmt, base64.b64decode(val))
                        success_count += 1
                    except Exception as e:
                        errors.append(f"Format {fmt} 에러: {str(e)}")
                win32clipboard.CloseClipboard()
                
                try:
                    self.last_seq = win32clipboard.GetClipboardSequenceNumber()
                    self.last_hash = item.get('hash')
                except: pass
                
                if success_count > 0:
                    msg = "복사되었습니다!" if not errors else f"복사 완료 (일부 포맷 제외됨)"
                    return {"success": True, "message": msg, "errors": errors}
                else:
                    return {"success": False, "message": f"복원 실패:\n{chr(10).join(errors)[:200]}"}
        return {"success": False, "message": "항목을 찾을 수 없습니다."}

    def delete_item(self, item_id):
        history = self._read_full_history(skip_cleanup=True)
        history = [item for item in history if item['id'] != item_id]
        self._save_history(history)
        return {"success": True}

    def clear_recent(self):
        history = self._read_full_history(skip_cleanup=True)
        # 영구보관된 것만 남기고 모두 삭제
        history = [item for item in history if item.get('is_permanent')]
        self._save_history(history)
        return {"success": True}