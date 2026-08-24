# -*- coding: utf-8 -*-
# python -m PyInstaller --noconsole --onefile --add-data "*.html;." PenTong.py
import sys
import os

# [추가] 에디터(VS Code 등)에서 실행 경로가 꼬여도 무조건 현재 폴더의 모듈(Pdf.py 등)을 찾도록 경로 강제 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import logging
import webview
import json
import traceback
import threading
import time
import base64
import tempfile
import subprocess
import platform
import keyboard

# ==========================================================
# [추가] 프로그램 버전 정의
# ==========================================================
APP_VERSION = "ver2026.08.05"

# [추가] 신구대조표 한글 생성 모듈 로드 (에러 방지 구조 적용)
try:
    import hwp_utils
except Exception as e:
    print(f"[경고] hwp_utils 로드 실패: {e}")
    class HwpUtilsDummy:
        def open_hwp_and_insert(self, html_content): return False, "hwp_utils 모듈 없음"
    hwp_utils = type('obj', (object,), {'HwpController': HwpUtilsDummy})

# 모듈 가져오기를 각각 분리하여, 하나의 에러가 다른 기능에 영향을 주지 않도록 수정
try:
    import hwp_merge
except Exception as e:
    print(f"[경고] hwp_merge 로드 실패: {e}")
    class HwpDummy:
        def select_files(self, window): return []
        def merge_hwp_files(self, files): return "모듈 없음"
    hwp_merge = type('obj', (object,), {'HwpMergeService': HwpDummy})

try:
    import excel_merge
except Exception as e:
    print(f"[경고] excel_merge 로드 실패: {e}")
    class ExcelDummy:
        def select_files(self, window): return []
        def merge_excel_sheets(self, data): return "모듈 없음"
        def get_sheet_names(self, path): return []
        def open_file(self, path): return "모듈 없음"
        def split_excel_by_column(self, params): return {"success": False, "message": "모듈 없음"}
        def analyze_excel_sheet(self, params): return {"success": False, "message": "모듈 없음"}
        def get_excel_row_data(self, params): return {"success": False, "message": "모듈 없음"}
        def detect_data_end(self, params): return {"success": False, "message": "모듈 없음"}
    excel_merge = type('obj', (object,), {'ExcelMergeService': ExcelDummy})

try:
    import Pdf
except Exception as e:
    print(f"[경고] Pdf 로드 실패: {e}")
    pdf_err = str(e)
    class PdfDummy:
        def select_pdf_files(self, window, multiple): raise Exception(f"Pdf 오류 발생:\n{pdf_err}")
        def process_pdf_merge(self, window, file_paths): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_split(self, window, file_path, pages_range): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_split_all(self, window, file_path): raise Exception(f"Pdf 오류: {pdf_err}") 
        def process_pdf_compress(self, window, file_path): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_protect(self, window, file_path, password): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_unlock(self, window, file_path, password): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_rotate(self, window, file_path, pages_range, angle): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_extract_text(self, window, file_path): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_delete_pages(self, window, file_path, pages_range): raise Exception(f"Pdf 오류: {pdf_err}")
        def process_pdf_watermark(self, window, target_path, watermark_path): raise Exception(f"Pdf 오류: {pdf_err}")
    Pdf = type('obj', (object,), {'PdfService': PdfDummy})

try:
    import PenTong_Schedule
except Exception as e:
    print(f"[경고] PenTong_Schedule 로드 실패: {e}")
    class ScheduleDummy:
        def save_schedule(self, data): return {"status": "error", "message": "일정 모듈 없음"}
    PenTong_Schedule = type('obj', (object,), {
        'ScheduleAPI': ScheduleDummy,
        'start_background_service': lambda: None,
        'start_tray_and_alarm_only': lambda: None
    })

# ==========================================================
# [추가] 이미지 및 GIF 애니메이션 전용 모듈 로드
# ==========================================================
try:
    import PenTong_Image
except Exception as e:
    print(f"[경고] PenTong_Image 로드 실패: {e}")
    class ImageDummy:
        def generate_gif(self, payload): return {"success": False, "error": f"이미지 모듈 없음: {str(e)}"}
        def save_gif_dialog(self, window, b64): return {"success": False, "message": "모듈 없음"}
        def save_pengif_project(self, window, json_str): return {"success": False, "message": "모듈 없음"}
        def load_pengif_project(self, window): return {"success": False, "message": "모듈 없음"}
    PenTong_Image = type('obj', (object,), {'ImageAPI': ImageDummy})

# ==========================================================
# [추가] 컴퓨터 정보 관리 모듈 로드 (에러 방지 구조)
# ==========================================================
try:
    import pc_manager
except Exception as e:
    print(f"[경고] pc_manager 로드 실패: {e}")
    class PCDummy:
        def get_all_system_info(self): return {"error": "컴퓨터 관리 모듈 없음"}
        def run_management_tool(self, tool_id): return {"status": "error", "message": "모듈 없음"}
        def change_pc_name(self, new_name): return {"status": "error", "message": "모듈 없음"}
    pc_manager = type('obj', (object,), {'PCManager': PCDummy})


# ==========================================================
# [추가] 폴더/파일 목록 추출기 전용 모듈 로드
# ==========================================================
try:
    import folder_extractor
except Exception as e:
    print(f"[경고] folder_extractor 로드 실패: {e}")
    class FolderExtractorDummy:
        def __init__(self, window): pass
        def scan_folder(self, payload): return {"success": False, "message": "모듈 없음"}
        def export_excel(self, payload): return {"success": False, "message": "모듈 없음"}
    folder_extractor = type('obj', (object,), {'FolderExtractorAPI': FolderExtractorDummy})

# ==========================================================
# [추가] 파일 이름 일괄 변경 전용 모듈 로드
# ==========================================================
try:
    import PenTong_Rename
except Exception as e:
    print(f"[경고] PenTong_Rename 로드 실패: {e}")
    class RenameDummy:
        def select_files(self, window): return []
        def execute_rename(self, payload): return {"success": False, "errors": ["모듈 없음"]}
    PenTong_Rename = type('obj', (object,), {'RenameAPI': RenameDummy})
# ==========================================================
# [추가] 파일 찾기 전용 모듈 로드
# ==========================================================

try:
    import file_search
except Exception as e:
    print(f"[경고] file_search 로드 실패: {e}")
    class FileSearchDummy:
        def search_files(self, keyword, base_path): return []
        def open_file_location(self, filepath): pass
    file_search = type('obj', (object,), {'FileSearchService': FileSearchDummy})

# ==========================================================
# [추가] 다중 클립보드 전용 모듈 로드
# ==========================================================
try:
    import PenTong_Clipboard
except Exception as e:
    print(f"[경고] PenTong_Clipboard 로드 실패: {e}")
    class ClipboardDummy:
        def get_status(self): return {"enabled": False, "retention_days": 1}
        def toggle(self, s): return {"success": False}
        def set_retention(self, d): return {"success": False}
        def get_history(self): return []
        def make_permanent(self, id): return {"success": False}
        def copy_item(self, id): return {"success": False, "message": "모듈 없음"}
        def delete_item(self, id): return {"success": False}
        def clear_recent(self): return {"success": False}
    PenTong_Clipboard = type('obj', (object,), {'ClipboardService': ClipboardDummy})

# ==========================================================
# [추가] 엑셀 자동 취합 2.0 (다중 시트/드래그앤드롭) 모듈 로드
# ==========================================================
try:
    import excel_merger_api
except Exception as e:
    print(f"[경고] excel_merger_api 로드 실패: {e}")
    class ExcelMergerAPIDummy:
        def __init__(self, window, data_dir=None): pass
        def select_reference_file(self): return {"success": False, "message": "모듈 없음"}
        def refresh_reference_file(self, payload): return {"success": False, "message": "모듈 없음"} 
        def scan_folder(self): return {"success": False, "message": "모듈 없음"}
        def save_project(self, data): return {"success": False, "message": "모듈 없음"}
        def execute_merge(self, data): return {"success": False, "message": "모듈 없음"}
        def get_merged_data(self, data): return {"success": False, "message": "모듈 없음"}
        def export_excel(self, data): return {"success": False, "message": "모듈 없음"}
    excel_merger_api = type('obj', (object,), {'ExcelMergerAPI': ExcelMergerAPIDummy})

# ==========================================================
# [추가] 만능 즐겨찾기 전용 모듈 로드
# ==========================================================
try:
    import PenTong_Shortcut
except Exception as e:
    print(f"[경고] PenTong_Shortcut 로드 실패: {e}")
    class ShortcutDummy:
        def __init__(self, data_dir): pass
        def load_shortcut_data(self): return None
        def save_shortcut_data(self, json_data): return False
        def execute_shortcut(self, target_path): return {"success": False, "message": "모듈 없음"}
        def select_local_path(self, path_type, window): return None
    PenTong_Shortcut = type('obj', (object,), {'ShortcutAPI': ShortcutDummy})

# ==========================================================
# [추가] 다중 문서 PDF 병합 전용 모듈 로드
# ==========================================================
try:
    import multi_pdf_merger
except Exception as e:
    print(f"[경고] multi_pdf_merger 로드 실패: {e}")
    class MultiPdfDummy:
        def select_mixed_files(self, window): return []
        def merge_to_pdf(self, window, file_paths): return {"success": False, "message": "모듈 없음"}
    multi_pdf_merger = type('obj', (object,), {'MultiPdfMergerService': MultiPdfDummy})

try:
    import PenTong_PenBook
except Exception as e:
    print(f"[경고] PenTong_PenBook 로드 실패: {e}")
    class PenBookDummy:
        def __init__(self, data_dir): pass
    PenTong_PenBook = type('obj', (object,), {'PenBookService': PenBookDummy})


# [신규 추가] 화면 판서 (Zoom) 모듈 로드
try:
    import PenTong_Zoom
except Exception as e:
    print(f"[경고] PenTong_Zoom 로드 실패: {e}")
    class ZoomDummy:
        def start_zoom_manual(self): return {"success": False, "message": "모듈이 없습니다."}
        def clean_up_and_exit(self): pass
    PenTong_Zoom = type('obj', (object,), {'ZoomService': lambda main_api: ZoomDummy()})

logger = logging.getLogger('pywebview')
logger.setLevel(logging.CRITICAL)

def resource_path(relative_path):
    """ PyInstaller 등 리소스 경로 해결 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==========================================================
# [추가] 모자이크 기능 전용 모듈 로드
# ==========================================================
try:
    import PenTong_Mosaic
except Exception as e:
    print(f"[경고] PenTong_Mosaic 로드 실패: {e}")
    class MosaicDummy:
        def __init__(self, window): pass
        def open_folder(self): return {"status": "error", "message": "모듈 없음"}
        def load_image_data(self, path): return {"status": "error", "message": "모듈 없음"}
        def save_image_data(self, payload): return {"status": "error", "message": "모듈 없음"}
    PenTong_Mosaic = type('obj', (object,), {'MosaicAPI': MosaicDummy})


class Api:
    def __init__(self):
        self.window = None
        self.has_clicked_ribbon = False
        
        # -----------------------------------------------------------
        # [설정] 경로 및 폴더 설정
        # -----------------------------------------------------------
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.data_dir = os.path.join(self.base_dir, "PenTong_Data")
        
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
            except Exception as e:
                print(f"폴더 생성 실패: {e}")

        self.settings_path = os.path.join(self.data_dir, "settings.json")
        self.saved_apps_path = os.path.join(self.data_dir, "saved_apps.json")
        self.memos_path = os.path.join(self.data_dir, "quick_memos.json")

        self.settings = self._load_settings()

        # ==========================================================
        # 🛑 [수정] 서비스 초기화 - 응답없음 방지를 위한 지연 로딩(Lazy Loading)
        # pywebview가 스캔하지 못하게 언더바(_) 처리
        # ==========================================================
        self._hwp_service = None
        self._excel_service = None
        self._pdf_service = None
        self._hwp_ctrl = None
        self._schedule_service = None
        self._image_service = None
        self._shortcut_service = None
        self._zoom_service = None
        self._file_search_service = None
        self._pc_service = None
        self._excel_merger_api2 = None
        self._multi_pdf_service = None

        # 기존 코드 아래에 추가
        self._multi_pdf_service = None
        self._clipboard_service = None # [추가] 다중 클립보드 지연 로딩
        self._folder_extractor_service = None # [추가] 폴더 추출기 지연 로딩

        self._folder_extractor_service = None # [추가] 폴더 추출기 지연 로딩
        self._mosaic_service = None           # [추가] 모자이크 지연 로딩

        self._rename_service = None           # [추가] 파일 이름 변경 지연 로딩

        # 화면 맵핑
        self.VIEW_MAP = {
                'hello1': '시작화면_몬스터세상.html',
                'hello2': '시작화면_펜통.html',
                'hello3': '시작화면_펜통힐링.html',
                'settings': 'settings.html',
                'free_sheet': 'free_sheet.html',
                '추첨1': '추첨1.html',
                '사다리1': '사다리1.html',
                '달리기1': '달리기1.html',
                'qrcode1': 'qrcode1.html',
                '스마트계산기': '스마트계산기.html',
                '산출기초계산기': '산출기초계산기.html',
                '아날로그시계': '아날로그시계.html',
                '디지털시계': '디지털시계.html',
                '타이머': '타이머.html',
                '스톱워치': '스톱워치.html',
                '단위변환기': '단위변환기.html',
                '텍스트정리': '텍스트정리.html',
                '텍스트비교': '텍스트비교.html',
                '환율계산기': '환율계산기.html',
                '우편번호검색': '우편번호검색.html',
                '태극기': '태극기.html',
                '한글파일취합': '한글파일취합.html',
                '엑셀파일취합': '엑셀파일취합.html',
                '엑셀파일취합2': '엑셀파일취합2.html',
                '엑셀분할': '엑셀분할.html',  
                '중복데이터': '중복데이터.html',
                '데이터정제전화날짜시간': '데이터정제전화날짜시간.html',
                '이메일정제': '이메일정제.html',
                '도로명주소정제': '도로명주소정제.html',
                '데이터요약': '데이터요약.html',
                '차트': '차트.html',
                '자동필터': '자동필터.html',
                '주민번호검증': '주민번호사업자번호검증.html',
                '함수그래프': '함수그래프.html',
                '모둠활동도우미': '모둠활동도우미.html',
                '날짜계산기': '날짜계산기.html',
                '점수판': '점수판.html',
                '전자칠판': '전자칠판.html',
                'PDF도구': 'pdf_tool.html',  
                '대진표': '대진표.html',
                '시계놀이': '시계놀이.html',
                '좌석배치': '좌석배치.html',    
                '신구대조표': '신구대조표.html',
                '신구대조표HWPX': '신구대조표HWPX.html',
                '구슬추첨기': '구슬추첨기.html',
                '펜통정보': '펜통정보.html',
                '아이콘변환기': '아이콘변환기.html',
                '선형회귀': '선형회귀.html',
                '단일퍼셉트론': '단일퍼셉트론.html',
                '다층퍼셉트론': '다층퍼셉트론.html',
                '경사하강법': '경사하강법.html',
                '언어모델원리': '언어모델원리.html',
                '인공신경망': '인공신경망.html',
                '일정관리': '일정관리.html',
                '이미지투명배경편집기': '이미지투명배경편집기.html',
                'GIF애니메이션': 'GIF애니메이션.html',
                '워드클라우드': '워드클라우드.html',
                '부산블록맵': '부산블록맵.html',
                '컴퓨터관리': '컴퓨터관리.html',
                '개인정보마스킹': '개인정보마스킹.html',
                '나이스바이트계산기': '나이스바이트계산기.html',
                '파일찾기': '파일찾기.html',
                '자주찾는사이트': '자주찾는사이트.html',
                '화면판서': '화면판서.html',
                'quick_memo': 'quick_memo.html',
                '세계시계': '세계시계.html',
                '이모지모음': '이모지모음.html',
                '색상팔레트': '색상팔레트.html',
                '다중문서PDF병합': '다중문서PDF병합.html',
                '도형생성기':'도형생성기.html',
                'PenBook': 'PenBook.html',
                '한글용량줄이기': '한글용량줄이기.html',
                '다중클립보드': '다중클립보드.html',
                '폴더파일목록추출기': '폴더파일목록추출기.html',
                '모자이크': '모자이크.html',
                '파일이름일괄변경': '파일이름일괄변경.html',
                '사진반듯하게펴기': '사진반듯하게펴기.html',
                '증명사진보정': '증명사진보정.html',
                '한글테스트': '한글테스트.html',

        }

    # ==========================================================
    # [추가] 모듈 지연 로딩을 위한 Getter 함수들
    # ==========================================================
    def _get_hwp_service(self):
        if self._hwp_service is None: self._hwp_service = hwp_merge.HwpMergeService()
        return self._hwp_service

    def _get_excel_service(self):
        if self._excel_service is None: self._excel_service = excel_merge.ExcelMergeService()
        return self._excel_service

    def _get_pdf_service(self):
        if self._pdf_service is None: self._pdf_service = Pdf.PdfService()
        return self._pdf_service

    def _get_hwp_ctrl(self):
        if self._hwp_ctrl is None: self._hwp_ctrl = hwp_utils.HwpController()
        return self._hwp_ctrl

    def _get_schedule_service(self):
        if self._schedule_service is None: self._schedule_service = PenTong_Schedule.ScheduleAPI()
        return self._schedule_service

    def _get_image_service(self):
        if self._image_service is None: self._image_service = PenTong_Image.ImageAPI()
        return self._image_service

    def _get_shortcut_service(self):
        if self._shortcut_service is None: self._shortcut_service = PenTong_Shortcut.ShortcutAPI(self.data_dir)
        return self._shortcut_service

    def _get_zoom_service(self):
        if self._zoom_service is None: self._zoom_service = PenTong_Zoom.ZoomService(self)
        return self._zoom_service

    def _get_file_search_service(self):
        if self._file_search_service is None: self._file_search_service = file_search.FileSearchService()
        return self._file_search_service

    def _get_pc_service(self):
        if self._pc_service is None: self._pc_service = pc_manager.PCManager()
        return self._pc_service

    def _get_excel_merger_api2(self):
        if self._excel_merger_api2 is None: self._excel_merger_api2 = excel_merger_api.ExcelMergerAPI(self.window, self.data_dir)
        return self._excel_merger_api2

    def _get_multi_pdf_service(self):
        if self._multi_pdf_service is None: self._multi_pdf_service = multi_pdf_merger.MultiPdfMergerService()
        return self._multi_pdf_service

    def _get_penbook_service(self):
        if not hasattr(self, '_penbook_service') or self._penbook_service is None:
            # PenTong_Data 폴더 안의 'PenBook_Data' 폴더를 독립적으로 사용합니다.
            pb_dir = os.path.join(self.data_dir, 'PenBook_Data')
            self._penbook_service = PenTong_PenBook.PenBookService(pb_dir)
        return self._penbook_service

    # ==========================================================
    # [추가] 다중 클립보드 전용 API 통로들
    # ==========================================================
    
    def _get_clipboard_service(self):
        if self._clipboard_service is None:
            self._clipboard_service = PenTong_Clipboard.ClipboardService(self.data_dir, self.window) # ◀ self.window 추가
        return self._clipboard_service

    def cb_get_status(self): return self._get_clipboard_service().get_status()
    def cb_toggle(self, state): return self._get_clipboard_service().toggle(state)
    def cb_set_retention(self, days): return self._get_clipboard_service().set_retention(days)
    def cb_get_history(self): return self._get_clipboard_service().get_history()
    def cb_make_permanent(self, item_id): return self._get_clipboard_service().make_permanent(item_id)
    def cb_copy_item(self, item_id): return self._get_clipboard_service().copy_item(item_id)
    def cb_delete_item(self, item_id): return self._get_clipboard_service().delete_item(item_id)
    def cb_clear_recent(self): return self._get_clipboard_service().clear_recent()


    # ==========================================================
    # [추가] 폴더/파일 목록 추출기 전용 API 통로들
    # ==========================================================
    def _get_folder_extractor_service(self):
        if self._folder_extractor_service is None:
            self._folder_extractor_service = folder_extractor.FolderExtractorAPI(self.window)
        return self._folder_extractor_service

    def scan_folder_for_extractor(self, payload):
        return self._get_folder_extractor_service().scan_folder(payload)

    def export_folder_list_to_excel(self, payload):
        return self._get_folder_extractor_service().export_excel(payload)

    # ==========================================================
    # [추가] 모자이크 전용 API 통로들
    # ==========================================================
    def _get_mosaic_service(self):
        if self._mosaic_service is None:
            self._mosaic_service = PenTong_Mosaic.MosaicAPI(self.window)
        return self._mosaic_service

    def mosaic_open_folder(self):
        return self._get_mosaic_service().open_folder()
    def mosaic_add_files(self):
        return self._get_mosaic_service().add_files()

    def mosaic_select_save_folder(self):
        return self._get_mosaic_service().select_save_folder()
    def mosaic_load_image(self, filepath):
        return self._get_mosaic_service().load_image_data(filepath)

    def mosaic_save_image(self, payload):
        return self._get_mosaic_service().save_image_data(payload)


    # ==========================================================
    # 워드클라우드용 텍스트 키워드 추출 API (외부 라이브러리, Java 설치 X)
    # ==========================================================
    def extract_keywords(self, text, options):
        import re
        from collections import Counter
        
        try:
            exclude_particles = options.get('exclude_particles', True)
            max_words = options.get('max_words', 500)
            
            # 1. 특수문자 제거 및 공백 기준 분리
            clean_text = re.sub(r'[-_=+!@#$%^&*()\[\]{};:\'",.<>/?~`\n\r]', ' ', text)
            raw_words = clean_text.split()
            
            words = []
            if exclude_particles:
                # 2. 잘라낼 한국어 조사/어미 목록 (긴 것부터 순서대로 매칭되게 배치)
                particles = [
                    '에서는', '에서도', '으로는', '까지도', '부터는', '입니다', '합니다',
                    '에서', '으로', '부터', '까지', '에게', '보다', '처럼', '마다', '하고',
                    '은', '는', '이', '가', '을', '를', '의', '와', '과', '도', '로', '에', '다'
                ]
                
                for w in raw_words:
                    stripped = False
                    for p in particles:
                        # 단어가 해당 조사로 끝나는지 확인
                        if w.endswith(p):
                            stem = w[:-len(p)]
                            # [핵심] 조사를 떼어내고 남은 '명사'가 최소 2글자 이상일 때만 분리 인정!
                            # 예: 학교는 -> 학교(2글자 O), 도로 -> 도(1글자 X -> '도로' 유지)
                            if len(stem) >= 2:
                                words.append(stem)
                                stripped = True
                                break # 조사를 찾았으므로 더 이상 찾지 않음
                    
                    # 매칭되는 조사가 없었거나, 떼어냈더니 1글자라서 기각된 경우 원본 단어 유지
                    if not stripped:
                        words.append(w)
            else:
                words = raw_words

            # 3. 최종적으로 2글자 이상인 의미 있는 단어만 필터링
            final_words = [w for w in words if len(w) >= 2]
            
            # 4. 빈도수 카운트 및 정렬
            counts = Counter(final_words)
            result = [{"word": word, "count": count} for word, count in counts.most_common(max_words)]
            
            return {"keywords": result}
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"error": f"키워드 추출 중 오류: {str(e)}"}

    # ==========================================================
    # [추가] 버전 정보 반환 API
    # ==========================================================
    def get_app_version(self):
        return APP_VERSION
    # ==========================================================
    # [추가] 사진 반듯하게 펴기 전용 저장 API
    # ==========================================================
    def save_scanned_image(self, payload):
        try:
            import webview
            import base64
            
            b64_data = payload.get('image_data', '')
            default_name = payload.get('default_name', '스캔결과_펴기.png')
            
            # Base64 헤더 제거
            if ',' in b64_data:
                b64_data = b64_data.split(',')[1]
                
            file_types = ('PNG Image (*.png)', 'JPEG Image (*.jpg;*.jpeg)', 'All files (*.*)')
            
            # 다른 이름으로 저장 다이얼로그 호출
            save_paths = self.window.create_file_dialog(
                webview.SAVE_DIALOG, 
                save_filename=default_name,
                file_types=file_types
            )
            
            if save_paths:
                target_path = save_paths[0] if isinstance(save_paths, (list, tuple)) else save_paths
                with open(target_path, 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "이미지가 저장되었습니다."}
            
            return {"success": False, "message": "저장이 취소되었습니다.", "cancelled": True}
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "message": f"저장 중 오류 발생: {str(e)}"}

    # ==========================================================
    # [신규 추가] 클립보드 이미지 복사 API (사진 펴기 연동)
    # ==========================================================
    def copy_to_clipboard(self, payload):
        import platform
        import tempfile
        import os
        import base64
        import subprocess

        try:
            b64_data = payload.get('image_data', '')
            if ',' in b64_data:
                b64_data = b64_data.split(',')[1]
                
            img_data = base64.b64decode(b64_data)
            
            # Windows 환경: 임시 파일 생성 후 PowerShell로 클립보드에 이미지 복사 (추가 모듈 불필요)
            if platform.system() == 'Windows':
                # 1. 임시 경로에 이미지 저장
                tmp_path = os.path.join(tempfile.gettempdir(), 'pentong_clip_temp.png')
                with open(tmp_path, 'wb') as f:
                    f.write(img_data)
                    
                # 2. PowerShell 스크립트를 통해 클립보드로 전송
                ps_script = f"""
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing
                $img = [System.Drawing.Image]::FromFile('{tmp_path}')
                [System.Windows.Forms.Clipboard]::SetImage($img)
                $img.Dispose()
                """
                # 창이 번쩍이는 것을 막기 위해 CREATE_NO_WINDOW(0x08000000) 플래그 사용
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_script], 
                    creationflags=0x08000000
                )
                return {"success": True}
            else:
                return {"success": False, "message": "이 OS에서는 지원하지 않습니다."}
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "message": f"복사 실패: {str(e)}"}

    def _load_settings(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: return self._get_default_settings()
        return self._get_default_settings()

    def _get_default_settings(self):
        return {
            "user_name": "",
            "affiliation": "", 
            "last_path": os.path.join(os.path.expanduser("~"), "Documents"),
            "theme": "none"     
        }

    def get_settings(self, params=None):
        return self.settings

    def get_user_info(self):
        return {
            "name": self.settings.get('user_name', ''),
            "affiliation": self.settings.get('affiliation', ''),
            "theme": self.settings.get('theme', 'none')
        }

    
    # [신규 추가] 화면 판서 API 연결
    def start_zoom_manual(self, payload='right'):
        return self._get_zoom_service().start_zoom_manual(payload)

    def move_zoom_panel(self, payload):
        return self._get_zoom_service().move_panel(payload)

    def change_zoom_pin_state(self, payload):
        return self._get_zoom_service().change_pin_state(payload)

    def save_settings(self, new_settings):
        try:
            self.settings.update(new_settings)
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
            
            if self.window:
                current_name = self.settings.get('user_name', '')
                current_aff = self.settings.get('affiliation', '')
                if current_name:
                     self.window.set_title(f"PenTong - {current_name} ({current_aff})")
            
            return {"status": "success", "message": "설정이 저장되었습니다."}
        except Exception as e:
            return {"status": "error", "message": f"저장 실패: {str(e)}"}
    
    def save_apps(self, apps_list):
        try:
            if isinstance(apps_list, list):
                apps_list = apps_list[:10]
                
            with open(self.saved_apps_path, 'w', encoding='utf-8') as f:
                json.dump(apps_list, f, ensure_ascii=False, indent=4)
            return {"status": "success"}
        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def get_saved_apps(self):
        if os.path.exists(self.saved_apps_path):
            try:
                with open(self.saved_apps_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                traceback.print_exc()
                return []
        return []

    def get_memos(self):
        """메모장을 열 때 기존 데이터를 불러와서 HTML로 넘겨줍니다."""
        if os.path.exists(self.memos_path):
            try:
                with open(self.memos_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                traceback.print_exc()
                return []
        return []

    def save_memos(self, memos):
        """JS에서 자동 저장 요청이 올 때마다 파일에 덮어씁니다."""
        try:
            with open(self.memos_path, 'w', encoding='utf-8') as f:
                json.dump(memos, f, ensure_ascii=False, indent=4)
            return {"status": "success"}
        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    # ==========================================================
    # [추가] 인터넷 바로가기 통로 모음 (내부는 PenTong_Shortcut에서 처리)
    # ==========================================================
    def load_shortcut_data(self):
        return self._get_shortcut_service().load_shortcut_data()

    def save_shortcut_data(self, json_data):
        return self._get_shortcut_service().save_shortcut_data(json_data)

    def execute_shortcut(self, target_path):
        return self._get_shortcut_service().execute_shortcut(target_path)

    def execute_file(self, filepath):
        """특정 파일(주로 엑셀)을 OS 기본 프로그램으로 엽니다."""
        import os, platform, subprocess
        try:
            if not os.path.exists(filepath):
                return {"success": False, "message": "파일을 찾을 수 없습니다."}
                
            if platform.system() == 'Windows':
                os.startfile(filepath)
            elif platform.system() == 'Darwin':
                subprocess.call(('open', filepath))
            else:
                subprocess.call(('xdg-open', filepath))
            return {"success": True}
        except Exception as e:
            return {"success": False, "message": f"파일 열기 실패: {str(e)}"}

    def select_local_path(self, path_type):
        return self._get_shortcut_service().select_local_path(path_type, self.window)

    def save_and_open_excel(self, data, filename):
        try:
            file_bytes = base64.b64decode(data)
            temp_dir = tempfile.gettempdir()
            file_path = os.path.join(temp_dir, filename)
            
            with open(file_path, 'wb') as f:
                f.write(file_bytes)
                
            if platform.system() == 'Windows':
                os.startfile(file_path)
            elif platform.system() == 'Darwin':
                subprocess.call(('open', file_path))
            else:
                subprocess.call(('xdg-open', file_path))
                
            return {"status": "success", "message": "파일이 열렸습니다."}
        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def initialize_user(self, name, affiliation, theme='none'):
        try:
            print(f"사용자 등록: 이름={name}, 소속={affiliation}, 테마={theme}")
            self.settings['user_name'] = name
            self.settings['affiliation'] = affiliation
            self.settings['theme'] = theme
            
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
            
            if self.window:
                self.window.set_title(f"PenTong - {name} ({affiliation})")

            def switch_page():
                time.sleep(0.2) 
                main_page_path = resource_path('index.html')
                if self.window:
                    self.window.load_url(main_page_path)

            threading.Thread(target=switch_page, daemon=True).start()
            
            return {"status": "success"}
        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    # --- 기존 API (지연 로딩 적용) ---
    def select_files(self, params=None): return self._get_hwp_service().select_files(self.window)
    def merge_hwp_files(self, file_paths): return self._get_hwp_service().merge_hwp_files(file_paths)
    def enable_hwp_drag_drop(self): return self._get_hwp_service().enable_native_drag_drop(self.window)
     # [추가] 다중 문서 PDF 병합 API 통로
    def select_mixed_files(self, params=None): return self._get_multi_pdf_service().select_mixed_files(self.window)
    def merge_to_pdf(self, file_paths): return self._get_multi_pdf_service().merge_to_pdf(self.window, file_paths)

    def select_excel_files(self, params=None): return self._get_excel_service().select_files(self.window)
    def get_excel_sheets(self, file_path): return self._get_excel_service().get_sheet_names(file_path)
    def merge_excel_files(self, merge_data): return self._get_excel_service().merge_excel_sheets(merge_data)
    def open_excel_file(self, file_path): return self._get_excel_service().open_file(file_path)
    def open_in_excel(self, data): return getattr(self._get_excel_service(), 'open_in_excel', lambda x: {"success": False})(data)
    
    def select_folder(self, params=None):
        if self.window:
            result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
            return result
        return None

    def split_excel_by_column(self, params): return self._get_excel_service().split_excel_by_column(params)
    def analyze_excel_sheet(self, params): return self._get_excel_service().analyze_excel_sheet(params)
    def get_excel_row_data(self, params): return self._get_excel_service().get_excel_row_data(params)
    def detect_data_end(self, params): return self._get_excel_service().detect_data_end(params)
    
    def save_seat_data(self, json_string):
        try:
            file_types = ('PenTong Seat Files (*.ptseat)', 'JSON Files (*.json)', 'All files (*.*)')
            save_path = self.window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='좌석배치.ptseat')
            if save_path:
                with open(save_path[0], 'w', encoding='utf-8') as f:
                    f.write(json_string)
                return {"success": True, "message": "배치 상태가 저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def load_seat_data(self):
        try:
            file_types = ('PenTong Seat Files (*.ptseat)', 'JSON Files (*.json)', 'All files (*.*)')
            open_path = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=file_types)
            if open_path:
                with open(open_path[0], 'r', encoding='utf-8') as f:
                    data = f.read()
                return {"success": True, "data": data}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def save_modum_data(self, json_string, filename="모둠활동.modum"):
        try:
            file_types = ('PenTong Modum Files (*.modum)', 'All files (*.*)')
            save_path = self.window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename=filename)
            if save_path:
                with open(save_path[0], 'w', encoding='utf-8') as f:
                    f.write(json_string)
                return {"success": True, "message": "모둠 데이터가 파일로 저장되었습니다."}
            return {"success": False, "message": "저장이 취소되었습니다."}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def load_modum_data(self):
        try:
            file_types = ('PenTong Modum Files (*.modum)', 'All files (*.*)')
            open_path = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=file_types)
            if open_path:
                with open(open_path[0], 'r', encoding='utf-8') as f:
                    data = f.read()
                return {"success": True, "data": data}
            return {"success": False, "message": "불러오기가 취소되었습니다."}
        except Exception as e:
            return {"success": False, "message": str(e)}


    # --- PDF API (지연 로딩 적용) ---
    def select_pdf_files(self, multiple=False): return self._get_pdf_service().select_pdf_files(self.window, multiple)
    def process_pdf_merge(self, file_paths): return self._get_pdf_service().process_pdf_merge(self.window, file_paths)
    def process_pdf_split(self, file_path, pages_range): return self._get_pdf_service().process_pdf_split(self.window, file_path, pages_range)
    def process_pdf_split_all(self, file_path): return self._get_pdf_service().process_pdf_split_all(self.window, file_path) 
    def process_pdf_compress(self, file_path): return self._get_pdf_service().process_pdf_compress(self.window, file_path)
    def process_pdf_protect(self, file_path, password): return self._get_pdf_service().process_pdf_protect(self.window, file_path, password)
    def process_pdf_unlock(self, file_path, password): return self._get_pdf_service().process_pdf_unlock(self.window, file_path, password)
    def process_pdf_rotate(self, file_path, pages_range, angle): return self._get_pdf_service().process_pdf_rotate(self.window, file_path, pages_range, angle)
    def process_pdf_extract_text(self, file_path): return self._get_pdf_service().process_pdf_extract_text(self.window, file_path)
    def process_pdf_delete_pages(self, file_path, pages_range): return self._get_pdf_service().process_pdf_delete_pages(self.window, file_path, pages_range)
    def process_pdf_watermark(self, target_path, watermark_path): return self._get_pdf_service().process_pdf_watermark(self.window, target_path, watermark_path)

    # ==========================================================
    # [추가] PenBook 전용 API 통로들 
    # (PenTong의 기존 함수들과 이름이 충돌하지 않도록 'pb_' 접두사 부착)
    # ==========================================================
    def pb_get_registry(self): return self._get_penbook_service().get_registry()
    def pb_create_project(self, name, ptype): return self._get_penbook_service().create_project(name, ptype)
    def pb_open_project(self, pid): return self._get_penbook_service().open_project(pid)
    def pb_delete_project(self, pid): return self._get_penbook_service().delete_project(pid)
    def pb_get_page(self, pid, nid): return self._get_penbook_service().get_page(pid, nid)
    def pb_save_page(self, pid, nid, html): return self._get_penbook_service().save_page(pid, nid, html)
    def pb_delete_page(self, pid, nid): return self._get_penbook_service().delete_page(pid, nid)
    def pb_save_meta(self, pid, meta): return self._get_penbook_service().save_meta(pid, meta)
    def pb_get_images(self, pid): return self._get_penbook_service().get_images(pid)
    def pb_upload_image(self, pid, fname, b64): return self._get_penbook_service().upload_image(pid, fname, b64)
    def pb_get_image_base64(self, pid, fname): return self._get_penbook_service().get_image_base64(pid, fname)
    def pb_load_settings(self): return self._get_penbook_service().load_settings()
    def pb_save_settings(self, settings): return self._get_penbook_service().save_settings(settings)
    def pb_delete_image(self, pid, fname): return self._get_penbook_service().delete_image(pid, fname)
    def pb_export_image(self, pid, fname): return self._get_penbook_service().export_image(self.window, pid, fname)
    def pb_replace_image(self, pid, fname, b64): return self._get_penbook_service().replace_image(pid, fname, b64)
    def pb_read_clipboard(self): return self._get_penbook_service().read_clipboard()
    def pb_export_standalone_html(self, filename, content): return self._get_penbook_service().export_standalone_html(self.window, filename, content)
    def pb_export_project(self, pid): return self._get_penbook_service().export_project(self.window, pid)
    def pb_import_project(self): return self._get_penbook_service().import_project(self.window)
    def pb_rename_project(self, old_id, new_name): return self._get_penbook_service().rename_project(old_id, new_name)
    def pb_transfer_node(self, src, dst, node, mode): return self._get_penbook_service().transfer_node(src, dst, node, mode)
    def pb_import_project_from_data(self, filename, b64): return self._get_penbook_service().import_project_from_data(filename, b64)


    # ==========================================================
    # [추가] 한글 파일 용량 줄이기 API
    # ==========================================================
    def select_hwp_files(self):
        """JS에서 호출: 압축할 한글 파일 선택 창 띄우기"""
        return self._get_hwp_ctrl().select_hwp_files(self.window)
    def get_hwp_info(self, payload):
        """[추가] JS에서 호출: 드래그된 파일의 용량 및 메타데이터 추출"""
        return self._get_hwp_ctrl().get_hwp_info(payload)
    # [수정된 부분] 
    def compress_hwp_files(self, payload):
        """JS에서 호출: 폴더 선택 후 HWPX 파일 내부 이미지 직접 압축/저장"""
        return self._get_hwp_ctrl().compress_hwp_files(payload, window=self.window)

    def create_hwp_diff(self, html_content):
        success, message = self._get_hwp_ctrl().open_hwp_and_insert(html_content)
        return [success, message]

    def toggle_fullscreen(self):
        if self.window:
            try: self.window.toggle_fullscreen()
            except: pass

    def handle_ribbon_click(self, action):
        if not action.startswith("hello"):
            self.has_clicked_ribbon = True

        print(f"클릭 요청: {action}")
        try:
            if action in self.VIEW_MAP:
                file_path = resource_path(self.VIEW_MAP[action])
                if not os.path.exists(file_path):
                    return {"type": "error", "content": f"파일 없음: {self.VIEW_MAP[action]}"}
                with open(file_path, 'r', encoding='utf-8') as f:
                    return {"type": "view", "content": f.read()}
            elif action == 'save':
                return {"type": "alert", "content": "저장 완료!"}
            else:
                return {"type": "alert", "content": f"명령: {action}"}
        except Exception as e:
            traceback.print_exc()
            return {"type": "error", "content": str(e)}

    # ==========================================================
    # [일정관리 전용 API 통로들]
    # ==========================================================
    def save_schedule(self, data): 
        return self._get_schedule_service().save_schedule(data)

    def get_schedules(self, start_date_str, end_date_str): 
        return self._get_schedule_service().get_schedules(start_date_str, end_date_str)
        
    def delete_schedule(self, sch_id):
        return self._get_schedule_service().delete_schedule(sch_id)
        
    def snooze_alarm_api(self, data):
        return self._get_schedule_service().snooze_alarm_api(data)
        
    def get_upcoming_alarms(self):
        return self._get_schedule_service().get_upcoming_alarms()

    # ==========================================================
    # [추가] 파일 찾기 API 통로들
    # ==========================================================
    def search_files(self, keyword, base_path, exclude_exts=[], exclude_no_ext=True):
        return self._get_file_search_service().search_files(keyword, base_path, self.window, exclude_exts, exclude_no_ext)
        
    def stop_search(self):
        return self._get_file_search_service().stop_search()

    def open_file_location(self, filepath):
        return self._get_file_search_service().open_file_location(filepath)

    def execute_file_search(self, filepath): # 이름 겹침 방지(기존 execute_file 유지)
        return self._get_file_search_service().execute_file(filepath)

    # ==========================================================
    # [추가] 이미지/GIF 처리 및 PenGif 프로젝트 API 통로
    # ==========================================================
    def generate_gif(self, payload):
        return self._get_image_service().generate_gif(payload)

    def save_gif_dialog(self, b64_data):
        return self._get_image_service().save_gif_dialog(self.window, b64_data)

    def save_pengif_project(self, json_string):
        return self._get_image_service().save_pengif_project(self.window, json_string)

    def load_pengif_project(self):
        return self._get_image_service().load_pengif_project(self.window)

    def generate_webp(self, payload):
        return self._get_image_service().generate_webp(payload)

    def save_webp_dialog(self, b64_data):
        return self._get_image_service().save_webp_dialog(self.window, b64_data)

    def generate_apng(self, payload):
        return self._get_image_service().generate_apng(payload)

    def save_apng_dialog(self, b64_data):
        return self._get_image_service().save_apng_dialog(self.window, b64_data)

    def check_ffmpeg(self):
        return self._get_image_service().check_ffmpeg()

    def download_ffmpeg(self):
        return self._get_image_service().download_ffmpeg()

    def generate_webm(self, payload):
        return self._get_image_service().generate_webm(payload)

    def save_webm_dialog(self, b64_data):
        return self._get_image_service().save_webm_dialog(self.window, b64_data)

    def generate_mp4(self, payload):
        return self._get_image_service().generate_mp4(payload)

    def save_mp4_dialog(self, b64_data):
        return self._get_image_service().save_mp4_dialog(self.window, b64_data)

    # ==========================================================
    # [추가] PC 관리 및 정보 조회 API 통로
    # ==========================================================
    def get_all_system_info(self):
        return self._get_pc_service().get_all_system_info()

    def run_management_tool(self, tool_id):
        return self._get_pc_service().run_management_tool(tool_id)

    def change_pc_name(self, new_name):
        return self._get_pc_service().change_pc_name(new_name)

    # ==========================================================
    # [추가] 엑셀 자동 취합 2.0 전용 API 통로들
    # ==========================================================
    def select_reference_file(self):
        return self._get_excel_merger_api2().select_reference_file()

    def refresh_reference_file(self, payload):
        return self._get_excel_merger_api2().refresh_reference_file(payload)

    def scan_folder(self):
        return self._get_excel_merger_api2().scan_folder()

    def refresh_folder(self, payload):
        return self._get_excel_merger_api2().refresh_folder(payload)

    def save_project(self, payload):
        return self._get_excel_merger_api2().save_project(payload)
    
    def get_projects(self): return self._get_excel_merger_api2().get_projects()        
    def delete_project(self, payload): return self._get_excel_merger_api2().delete_project(payload)
    def rename_project(self, payload): return self._get_excel_merger_api2().rename_project(payload)
    
    # [추가] 취합 실행 및 저장 통로
    def execute_merge(self, payload): return self._get_excel_merger_api2().execute_merge(payload)

    def get_merged_data(self, payload):
        return self._get_excel_merger_api2().get_merged_data(payload)
        
    def export_excel(self, payload):
        return self._get_excel_merger_api2().export_excel(payload)

    def export_settings(self, payload):
        return self._get_excel_merger_api2().export_settings(payload)

    def import_settings(self, payload):
        return self._get_excel_merger_api2().import_settings(payload)
    def export_reference_file(self, payload):
        return self._get_excel_merger_api2().export_reference_file(payload)


    # 1. 기존 함수 수정: 'filepath' 항목을 추가로 반환하도록 딱 1줄만 변경합니다.
    def read_hwpx_file(self, file_path=None):
        # 1. 파일 경로가 파라미터로 넘어오지 않은 경우 (기존 방식: 버튼 클릭 시)
        if file_path is None:
            import webview
            file_types = ('HWPX Files (*.hwpx)', 'All files (*.*)')
            result = self.window.create_file_dialog(
                webview.OPEN_DIALOG, 
                allow_multiple=False, 
                file_types=file_types
            )
            if not result:
                return None
            file_path = result[0]

        # 2. 전달받은 경로(또는 위에서 선택한 경로)를 가지고 바로 텍스트를 추출합니다.
        # (여기에 있던 중복된 window.create_file_dialog 코드는 삭제했습니다!)
        import os
        from hwpx_parser import HwpXParser
        
        try:
            parser = HwpXParser()
            text = parser.extract_text(file_path)
            filename = os.path.basename(file_path)
            
            return {"text": text, "filename": filename, "filepath": file_path}
        except Exception as e:
            print(f"HWPX 파싱 오류: {e}")
            return None

    # 2. 신규 함수 추가: 파일 저장 다이얼로그를 띄우고 HWPX를 생성합니다.
    def save_diff_hwpx(self, old_path, new_path):
        import webview
        window = webview.windows[0]
        
        # 저장 다이얼로그 오픈 (원하는 경로와 파일명 지정)
        save_paths = window.create_file_dialog(
            webview.SAVE_DIALOG, 
            directory='', 
            save_filename='최종_신구대조표.hwpx'
        )
        
        if save_paths:
            # pywebview 버전에 따라 문자열 튜플/리스트로 반환될 수 있음
            target_path = save_paths[0] if isinstance(save_paths, (list, tuple)) else save_paths
            
            from hwpx_parser import HwpXParser
            parser = HwpXParser()
            
            # hwpx_parser.py에 만들어둔 직접 생성 함수 호출!
            success = parser.build_diff_hwpx(old_path, new_path, target_path)
            return {"success": success, "save_path": target_path}
            
        return {"success": False}

    # ==========================================================
    # 색상 팔레트 데이터 저장/불러오기 (즐겨찾기, 최근항목)
    # ==========================================================
    def save_color_data(self, json_string):
        try:
            filepath = os.path.join(self.data_dir, 'color_palette_data.json')
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(json_string)
            return {"success": True}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}

    def load_color_data(self):
        try:
            filepath = os.path.join(self.data_dir, 'color_palette_data.json')
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            traceback.print_exc()
            return None
        return None

    # VIEW_MAP 딕셔너리 안에 아래 1줄 추가 (적당한 위치에 콤마 주의)
    # '파일이름변경': '파일이름일괄변경.html',

    # ==========================================================
    # [추가] 파일 이름 일괄 변경 전용 API 통로들
    # ==========================================================
    def _get_rename_service(self):
        if self._rename_service is None:
            self._rename_service = PenTong_Rename.RenameAPI()
        return self._rename_service

    def rename_select_files(self):
        return self._get_rename_service().select_files(self.window)

    def execute_rename(self, payload):
        return self._get_rename_service().execute_rename(payload)

    def open_github(self):
        import webbrowser
        # 펜통 깃허브 주소로 연결
        webbrowser.open("https://github.com/seomyeongho/PenTong")
        return {"success": True}


if __name__ == '__main__':

    if "--tray" in sys.argv:
        print("백그라운드(트레이) 모드로 실행합니다.")
        PenTong_Schedule.start_tray_and_alarm_only()
        sys.exit(0)
    
    api = Api()
    PenTong_Schedule.start_background_service() 
    
    user_settings = api.get_settings()
    user_name = user_settings.get('user_name', '')
    affiliation = user_settings.get('affiliation', '')
    
    initial_title = 'PenTong'
    
    if (not user_name or user_name.strip() == "") or (not affiliation or affiliation.strip() == ""):
        start_url = resource_path('startup.html')
    else:
        start_url = resource_path('index.html')
        initial_title = f'PenTong - {user_name} ({affiliation})'

    if not os.path.exists(start_url) and 'startup.html' not in start_url:
         start_url = resource_path('index.html')

    start_url = f"file://{start_url}"

    # 🛑 [수정] js_api=api 주석을 해제하여 자바스크립트 통신 활성화
    window = webview.create_window(
        initial_title,
        url=start_url, 
        width=1300,
        height=800,
        js_api=api,
        maximized=False, 
        background_color='#F3F3F3'
        
    )
    
    api.window = window

   
   # ==========================================
    # 창 닫기 이벤트 핸들러 정의 및 연결
    # 🛑 [수정] zoom_service 호출 시 _get_zoom_service() 사용
    # ==========================================
    def on_closing():
        if api.has_clicked_ribbon:
            try:
                result = window.create_confirmation_dialog(
                    '프로그램 종료', 
                    '저장하지 않은 작업 내용은 모두 사라집니다.\n\n정말 프로그램을 종료하시겠습니까?'
                )
                if result:
                    api._get_zoom_service().clean_up_and_exit()
                return result
            except:
                api._get_zoom_service().clean_up_and_exit()
                return True
        else:
            api._get_zoom_service().clean_up_and_exit()
            return True

   # ... (기존 창 닫기 이벤트 on_closing 연결 코드 바로 아래)
    window.events.closing += on_closing

    # ==========================================
    # [최종 수정된 드래그 앤 드롭 로직 - PenTong.py]
    # ==========================================
    def bind_drag_and_drop():
        try:
            from webview.dom import DOMEventHandler
            
            def on_drop(e):
                # 파일 정보 추출
                transfer = e.get('dataTransfer') or {}
                files = transfer.get('files', [])
                
                paths = [f.get('pywebviewFullPath') for f in files if f.get('pywebviewFullPath')]
                
                if paths:
                    import json
                    paths_json = json.dumps(paths)
                    
                    # [핵심 수정] 무조건 현재 사용자 눈에 띄워진(active) 프로그램만 찾아서 파일을 넘겨줍니다!
                    js_code = f"""
                        (function() {{
                            // class에 'active'가 있는 현재 열린 창(iframe)만 정확하게 골라냅니다.
                            var activeIframes = document.querySelectorAll('iframe.active');
                            for (var i = 0; i < activeIframes.length; i++) {{
                                if (activeIframes[i].contentWindow && activeIframes[i].contentWindow.addDroppedFiles) {{
                                    activeIframes[i].contentWindow.addDroppedFiles({paths_json});
                                    return;
                                }}
                            }}
                            // iframe이 없을 경우를 대비한 부모 창 폴백
                            if (window.addDroppedFiles) window.addDroppedFiles({paths_json});
                        }})();
                    """
                    window.evaluate_js(js_code)

            # 전역 요소에 직접 이벤트 바인딩
            drop_zone = window.dom.get_element('#global-drop-zone')
            if drop_zone:
                drop_zone.events.drop += DOMEventHandler(on_drop)
            
        except Exception as ex:
            print(f"드래그 앤 드롭 바인딩 오류: {ex}")

    def on_loaded():
        if not getattr(window, '_dnd_bound', False):
            bind_drag_and_drop()
            window._dnd_bound = True

    window.events.loaded += on_loaded
    
 
    webview.start(debug=False, gui='edgechromium')
