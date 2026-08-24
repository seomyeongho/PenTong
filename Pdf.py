# -*- coding: utf-8 -*-
import os
import webview
from PyPDF2 import PdfReader, PdfWriter

class PdfService:
    def select_pdf_files(self, window, multiple=False):
        if not window: return []
        file_types = ('PDF Files (*.pdf)', 'All files (*.*)')
        result = window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=multiple, file_types=file_types)
        if result:
            return [{"path": p, "name": os.path.basename(p)} for p in result]
        return []

    def _save_dialog(self, window, default_filename, is_txt=False):
        if not window: return None
        file_types = ('Text Files (*.txt)', 'All files (*.*)') if is_txt else ('PDF Files (*.pdf)', 'All files (*.*)')
        result = window.create_file_dialog(webview.SAVE_DIALOG, save_filename=default_filename, file_types=file_types)
        return result[0] if result else None

    # [추가] 폴더 선택 다이얼로그
    def _folder_dialog(self, window):
        if not window: return None
        result = window.create_file_dialog(webview.FOLDER_DIALOG)
        return result[0] if result else None

    def _parse_range(self, pages_range, total_pages):
        if not pages_range or str(pages_range).strip().lower() in ['all', '전체', '']:
            return set(range(total_pages))
        pages_to_extract = set()
        parts = [p.strip() for p in str(pages_range).split(',')]
        for part in parts:
            if '-' in part:
                start, end = map(int, part.split('-'))
                for p in range(start, end + 1):
                    pages_to_extract.add(p - 1)
            else:
                pages_to_extract.add(int(part) - 1)
        return pages_to_extract

    # 1. PDF 합치기
    def process_pdf_merge(self, window, file_paths):
        try:
            save_path = self._save_dialog(window, "merged_document.pdf")
            if not save_path: return {"success": False, "error": "cancelled"}
            merger = PdfWriter()
            for path in file_paths: merger.append(path)
            with open(save_path, "wb") as f: merger.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 2. PDF 자르기 (일부 추출)
    def process_pdf_split(self, window, file_path, pages_range):
        try:
            save_path = self._save_dialog(window, f"split_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            writer = PdfWriter()
            pages_to_extract = self._parse_range(pages_range, len(reader.pages))
            for i in range(len(reader.pages)):
                if i in pages_to_extract: writer.add_page(reader.pages[i])
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except ValueError: return {"success": False, "error": "페이지 범위 형식이 잘못되었습니다. (숫자와 하이픈, 쉼표만 사용)"}
        except Exception as e: return {"success": False, "error": str(e)}

    # [신규] 2-1. PDF 1장씩 낱장으로 모두 분할하여 폴더에 저장
    def process_pdf_split_all(self, window, file_path):
        try:
            save_folder = self._folder_dialog(window)
            if not save_folder: return {"success": False, "error": "cancelled"}
            
            reader = PdfReader(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                # 원본파일명_1페이지.pdf 형식으로 저장
                output_filename = f"{base_name}_{i+1}페이지.pdf"
                output_path = os.path.join(save_folder, output_filename)
                with open(output_path, "wb") as f:
                    writer.write(f)
                    
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 3. PDF 압축
    def process_pdf_compress(self, window, file_path):
        try:
            save_path = self._save_dialog(window, f"compressed_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            writer = PdfWriter()
            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)
            writer.add_metadata({})
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 4. PDF 비밀번호 설정
    def process_pdf_protect(self, window, file_path, password):
        try:
            save_path = self._save_dialog(window, f"protected_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            writer = PdfWriter()
            for page in reader.pages: writer.add_page(page)
            writer.encrypt(password)
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 5. PDF 비밀번호 해제
    def process_pdf_unlock(self, window, file_path, password):
        try:
            reader = PdfReader(file_path)
            if not reader.is_encrypted: return {"success": False, "error": "이 파일은 비밀번호가 설정되어 있지 않습니다."}
            if not reader.decrypt(password): return {"success": False, "error": "입력하신 비밀번호가 틀렸습니다."}
            save_path = self._save_dialog(window, f"unlocked_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            writer = PdfWriter()
            for page in reader.pages: writer.add_page(page)
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 6. PDF 회전
    def process_pdf_rotate(self, window, file_path, pages_range, angle):
        try:
            save_path = self._save_dialog(window, f"rotated_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            writer = PdfWriter()
            to_rotate = self._parse_range(pages_range, len(reader.pages))
            for i in range(len(reader.pages)):
                page = reader.pages[i]
                if i in to_rotate:
                    try: page.rotate(int(angle))
                    except AttributeError: page.rotateClockwise(int(angle))
                writer.add_page(page)
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 7. PDF 텍스트 추출
    def process_pdf_extract_text(self, window, file_path):
        try:
            default_name = os.path.basename(file_path).replace('.pdf', '.txt')
            save_path = self._save_dialog(window, default_name, is_txt=True)
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            text_content = ""
            for i, page in enumerate(reader.pages):
                text_content += f"--- {i+1} 페이지 ---\n"
                extracted = page.extract_text()
                if extracted: text_content += extracted + "\n\n"
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(text_content)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    # 8. PDF 페이지 삭제
    def process_pdf_delete_pages(self, window, file_path, pages_range):
        try:
            save_path = self._save_dialog(window, f"deleted_{os.path.basename(file_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            reader = PdfReader(file_path)
            writer = PdfWriter()
            to_delete = self._parse_range(pages_range, len(reader.pages))
            for i in range(len(reader.pages)):
                if i not in to_delete:
                    writer.add_page(reader.pages[i])
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except ValueError: return {"success": False, "error": "페이지 범위 형식이 잘못되었습니다."}
        except Exception as e: return {"success": False, "error": str(e)}

    # 9. PDF 워터마크 병합
    def process_pdf_watermark(self, window, target_path, watermark_path):
        try:
            save_path = self._save_dialog(window, f"watermarked_{os.path.basename(target_path)}")
            if not save_path: return {"success": False, "error": "cancelled"}
            target_reader = PdfReader(target_path)
            watermark_reader = PdfReader(watermark_path)
            watermark_page = watermark_reader.pages[0]
            writer = PdfWriter()
            for page in target_reader.pages:
                page.merge_page(watermark_page)
                writer.add_page(page)
            with open(save_path, "wb") as f: writer.write(f)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}