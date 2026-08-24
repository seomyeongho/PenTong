import os
import subprocess
import fnmatch
import time
import json
import base64
import string

class FileSearchService:
    def __init__(self):
        self._stop_flag = False

    def get_drives(self):
        r"""윈도우의 연결된 모든 물리/논리 드라이브(C:\, D:\ 등) 목록을 가져옵니다."""
        drives = []
        if os.name == 'nt':
            for letter in string.ascii_uppercase:
                path = f"{letter}:\\"
                if os.path.exists(path):
                    drives.append(path)
        else:
            drives = ['/']
        return drives

    def stop_search(self):
        """진행 중인 검색을 강제로 중단시킵니다."""
        self._stop_flag = True

    def search_files(self, keyword, base_paths, window=None, exclude_exts=None, exclude_no_ext=True):
        """
        다중 경로 및 모든 드라이브 지원, 실시간 고속 파일 검색 로직
        """
        self._stop_flag = False
        chunk = []
        max_results = 5000
        total_found = 0
        
        # UI 업데이트용 타이머
        last_flush_time = time.time()
        loop_counter = 0

        keyword_lower = keyword.lower()
        has_wildcard = '*' in keyword_lower or '?' in keyword_lower
        
        # 스마트 와일드카드 처리
        search_pattern = keyword_lower
        if has_wildcard:
            if not search_pattern.startswith('*'): search_pattern = '*' + search_pattern
            if not search_pattern.endswith('*'): search_pattern = search_pattern + '*'
            search_pattern = search_pattern.replace('**', '*') 

        if exclude_exts is None: exclude_exts = []
        clean_exclude_exts = set(ext.lower() if ext.startswith('.') else f'.{ext.lower()}' for ext in exclude_exts)

        # 다중 경로 해석 (단일 문자열, 리스트, "ALL_DRIVES" 모두 처리)
        if isinstance(base_paths, str):
            base_paths = [base_paths]
            
        actual_paths = []
        for p in base_paths:
            if p == "ALL_DRIVES":
                actual_paths.extend(self.get_drives())
            else:
                actual_paths.append(p)
        actual_paths = list(set(actual_paths)) # 중복 경로 제거

        def flush_chunk(is_last=False):
            nonlocal chunk, last_flush_time
            if window and (chunk or is_last):
                try:
                    json_str = json.dumps(chunk)
                    b64_data = base64.b64encode(json_str.encode('utf-8')).decode('utf-8')
                    js_is_last = 'true' if is_last else 'false'
                    
                    js_code = f"""
                    (function() {{
                        var data = '{b64_data}';
                        var isLast = {js_is_last};
                        if(window.receiveChunk) window.receiveChunk(data, isLast);
                        for(var i=0; i<window.frames.length; i++) {{
                            try {{ if(window.frames[i].receiveChunk) window.frames[i].receiveChunk(data, isLast); }} catch(e) {{}}
                        }}
                    }})();
                    """
                    window.evaluate_js(js_code)
                except Exception as e:
                    print(f"UI 전송 실패: {e}")
            chunk = []
            last_flush_time = time.time()

        def scan_dir(path):
            nonlocal total_found, loop_counter
            if self._stop_flag or total_found >= max_results:
                return
                
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        loop_counter += 1
                        if self._stop_flag or total_found >= max_results:
                            break
                        
                        is_dir = entry.is_dir()
                        
                        # --- 파일 확장자 필터링 ---
                        if not is_dir:
                            _, ext = os.path.splitext(entry.name)
                            ext_lower = ext.lower()
                            
                            if exclude_no_ext and not ext_lower: continue
                            if ext_lower in clean_exclude_exts: continue
                        # -------------------------

                        name_lower = entry.name.lower()
                        is_match = False
                        
                        if has_wildcard:
                            is_match = fnmatch.fnmatch(name_lower, search_pattern)
                        else:
                            is_match = keyword_lower in name_lower

                        if is_match:
                            try:
                                stat = entry.stat()
                                item = {
                                    "name": entry.name,
                                    "path": entry.path,
                                    "size": stat.st_size,
                                    "is_dir": is_dir,
                                    "modified": stat.st_mtime
                                }
                                chunk.append(item)
                                total_found += 1
                                
                            except OSError: pass

                        # 100번 루프마다 시간 체크, 0.5초 경과 시 전송 (프리징 방지)
                        if loop_counter % 100 == 0:
                            if len(chunk) >= 50 or (chunk and time.time() - last_flush_time >= 0.5):
                                flush_chunk()
                                time.sleep(0.005)

                        if is_dir:
                            scan_dir(entry.path)
            except (PermissionError, FileNotFoundError, OSError):
                pass

        # 모든 드라이브 및 다중 경로 순차 스캔
        for path in actual_paths:
            if self._stop_flag or total_found >= max_results:
                break
            if os.path.exists(path):
                scan_dir(path)
            
        flush_chunk(is_last=True)
        return True 

    def open_file_location(self, filepath):
        try:
            filepath = os.path.normpath(filepath)
            if os.path.isdir(filepath): subprocess.run(['explorer', filepath])
            else: subprocess.run(['explorer', '/select,', filepath])
        except: pass

    def execute_file(self, filepath):
        try: os.startfile(os.path.normpath(filepath))
        except: pass