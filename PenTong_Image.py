# -*- coding: utf-8 -*-
import base64
import io
import traceback
import webview  # [추가] 윈도우 다이얼로그 창 제어용
from PIL import Image

# 파일 상단에 필요한 모듈 추가
import os
import urllib.request
import zipfile
import tempfile
import subprocess
import shutil
import time

class ImageAPI:
    def __init__(self):
        # [신규] FFmpeg 엔진이 저장될 경로 설정 (프로그램 옆 PenTong_Data/engine 폴더)
        self.engine_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "PenTong_Data", "engine")
        os.makedirs(self.engine_dir, exist_ok=True)
        self.ffmpeg_path = os.path.join(self.engine_dir, "ffmpeg.exe")

    def check_ffmpeg(self):
        """ 엔진 설치 여부 확인 """
        return {"exists": os.path.exists(self.ffmpeg_path)}

    def download_ffmpeg(self):
        """ 백그라운드에서 FFmpeg 윈도우용 가벼운 빌드를 다운로드하고 압축을 풉니다. """
        try:
            if os.path.exists(self.ffmpeg_path):
                return {"success": True}
            
            # FFmpeg 윈도우용 빌드 다운로드 URL (약 30~40MB)
            url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
            zip_path = os.path.join(self.engine_dir, "ffmpeg.zip")
            
            urllib.request.urlretrieve(url, zip_path) # 다운로드 실행
            
            # 압축 해제 후 ffmpeg.exe만 빼내기
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                for member in zip_ref.namelist():
                    if member.endswith("ffmpeg.exe"):
                        # [핵심 1] source 파일 객체도 확실하게 닫아주어 락(Lock)을 해제합니다.
                        with zip_ref.open(member) as source, open(self.ffmpeg_path, "wb") as target:
                            shutil.copyfileobj(source, target)
                        break
                        
            # [핵심 2] 백신 검사 등으로 파일이 잠겨있을 수 있으므로 안전하게 삭제 처리합니다.
            try:
                import time
                time.sleep(0.5) # 잠시 대기
                os.remove(zip_path) # 다 쓴 압축파일 삭제
            except Exception:
                pass # 삭제에 실패해도 엔진 설치는 완료된 것이므로 무시하고 계속 진행
                
            return {"success": True}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": str(e)}

    def generate_webm(self, payload):
        """ 크롬에서 완벽하게 재생되는 투명 WebM 인코딩 """
        try:
            frames_data = payload.get('frames', [])
            if not frames_data: return {"success": False, "error": "데이터가 없습니다."}

            fps = round(1000 / max(int(frames_data[0].get('delay', 100)), 10), 2)

            with tempfile.TemporaryDirectory() as temp_dir:
                for i, frame in enumerate(frames_data):
                    img_bytes = base64.b64decode(frame['base64'])
                    with open(os.path.join(temp_dir, f"frame_{i:04d}.png"), "wb") as f:
                        f.write(img_bytes)

                out_path = os.path.join(temp_dir, "output.webm")

                # [수정] pad 옵션에 color=black@0을 추가하여 깨짐을 막고 투명도를 완벽히 보존합니다.
                cmd = [
                    self.ffmpeg_path, "-framerate", str(fps),
                    "-i", os.path.join(temp_dir, "frame_%04d.png"),
                    "-vf", "pad=ceil(iw/2)*2:ceil(ih/2)*2:0:0:color=black@0",
                    "-c:v", "libvpx-vp9", "-pix_fmt", "yuva420p",
                    "-auto-alt-ref", "0", "-b:v", "2M", "-y", out_path
                ]

                startupinfo = None
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                subprocess.run(cmd, startupinfo=startupinfo, check=True)

                with open(out_path, "rb") as f:
                    webm_bytes = f.read()

            return {"success": True, "data": base64.b64encode(webm_bytes).decode('utf-8')}
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "error": str(e)}

    def save_webm_dialog(self, window, b64_data):
        """ 생성된 WebM 데이터를 PC에 저장합니다. """
        try:
            file_types = ('투명 비디오 (*.webm)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='pentong_animation.webm')
            if save_path:
                with open(save_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ... (기존 save_webm_dialog 끝) ...
    
    def generate_mp4(self, payload):
        """ FFmpeg를 이용해 가장 호환성이 높은 MP4 비디오(H.264)로 인코딩합니다. """
        try:
            frames_data = payload.get('frames', [])
            if not frames_data: return {"success": False, "error": "데이터가 없습니다."}
            
            # 지연시간을 FPS로 변환
            fps = round(1000 / max(int(frames_data[0].get('delay', 100)), 10), 2)
            
            with tempfile.TemporaryDirectory() as temp_dir:
                for i, frame in enumerate(frames_data):
                    img_bytes = base64.b64decode(frame['base64'])
                    with open(os.path.join(temp_dir, f"frame_{i:04d}.png"), "wb") as f:
                        f.write(img_bytes)
                        
                out_path = os.path.join(temp_dir, "output.mp4")
                
                # [설정] 가장 표준적인 MP4(H.264, YUV420p) 코덱 적용
                # 홀수 해상도 방지 옵션포함
                cmd = [
                    self.ffmpeg_path, "-framerate", str(fps),
                    "-i", os.path.join(temp_dir, "frame_%04d.png"),
                    "-vf", "pad=ceil(iw/2)*2:ceil(ih/2)*2:0:0:color=black@0",
                    "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", "23", "-y", out_path
                ]
                
                startupinfo = None
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    
                subprocess.run(cmd, startupinfo=startupinfo, check=True)
                
                with open(out_path, "rb") as f:
                    mp4_bytes = f.read()
                    
            return {"success": True, "data": base64.b64encode(mp4_bytes).decode('utf-8')}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": str(e)}

    def save_mp4_dialog(self, window, b64_data):
        """ 생성된 MP4 데이터를 저장합니다. """
        try:
            file_types = ('MP4 비디오 (*.mp4)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='animation.mp4')
            if save_path:
                with open(save_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def generate_gif(self, payload):
        """ HTML에서 전달받은 Base64 프레임 데이터를 모아 투명 배경 GIF로 병합하여 반환합니다. """
        try:
            frames_data = payload.get('frames', [])
            settings = payload.get('settings', {})
            loop_val = int(settings.get('loop', 0))

            if not frames_data:
                return {"success": False, "error": "프레임 데이터가 전달되지 않았습니다."}

            images = []
            durations = []

            for frame in frames_data:
                img_bytes = base64.b64decode(frame['base64'])
                img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                
                alpha = img.getchannel('A')
                img_rgb = img.convert('RGB')
                
                img_p = img_rgb.convert('P', palette=Image.ADAPTIVE, colors=255)
                mask = Image.eval(alpha, lambda a: 255 if a <= 128 else 0)
                img_p.paste(255, mask)
                img_p.info['transparency'] = 255

                images.append(img_p)
                durations.append(int(frame.get('delay', 200)))

            output = io.BytesIO()
            images[0].save(
                output,
                format='GIF',
                save_all=True,
                append_images=images[1:],
                duration=durations,
                loop=loop_val,
                disposal=2,
                transparency=255
            )

            gif_base64 = base64.b64encode(output.getvalue()).decode('utf-8')
            
            return {
                "success": True,
                "data": gif_base64,
                "filename": "pentong_animation.gif"
            }
            
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": f"백엔드 GIF 변환 실패: {str(e)}"}

    def save_gif_dialog(self, window, b64_data):
        """ 생성된 GIF 데이터를 받아 '다른 이름으로 저장' 대화상자를 띄웁니다. """
        try:
            if not window:
                return {"success": False, "message": "창 객체가 없습니다."}
            file_types = ('GIF 애니메이션 (*.gif)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='pentong_animation.gif')
            
            if save_path and len(save_path) > 0:
                with open(save_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}

    # ==========================================================
    # [신규] PenGif 프로젝트 파일 저장/불러오기 로직
    # ==========================================================
    def save_pengif_project(self, window, json_string):
        """ 편집 중인 레이어, 텍스트 등 모든 속성 정보를 .PenGif 확장자로 저장합니다. """
        try:
            if not window: return {"success": False, "message": "창 객체가 없습니다."}
            file_types = ('PenGif 프로젝트 파일 (*.PenGif)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='내작품.PenGif')
            
            if save_path and len(save_path) > 0:
                with open(save_path[0], 'w', encoding='utf-8') as f:
                    f.write(json_string)
                return {"success": True, "message": "프로젝트가 저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}

    def load_pengif_project(self, window):
        """ 사용자가 저장해둔 .PenGif 파일을 선택하여 텍스트로 읽어옵니다. """
        try:
            if not window: return {"success": False, "message": "창 객체가 없습니다."}
            file_types = ('PenGif 프로젝트 파일 (*.PenGif)', '모든 파일 (*.*)')
            open_path = window.create_file_dialog(webview.OPEN_DIALOG, file_types=file_types)
            
            if open_path and len(open_path) > 0:
                with open(open_path[0], 'r', encoding='utf-8') as f:
                    data = f.read()
                return {"success": True, "data": data}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}

    def generate_webp(self, payload):
        """ HTML에서 전달받은 Base64 프레임 데이터를 모아 투명 배경 WebP 애니메이션으로 병합하여 반환합니다. """
        try:
            frames_data = payload.get('frames', [])
            settings = payload.get('settings', {})
            loop_val = int(settings.get('loop', 0))

            if not frames_data:
                return {"success": False, "error": "프레임 데이터가 전달되지 않았습니다."}

            images = []
            durations = []

            for frame in frames_data:
                img_bytes = base64.b64decode(frame['base64'])
                # WebP는 RGBA(투명도 포함 풀컬러)를 완벽하게 지원하므로 복잡한 마스킹 없이 바로 사용 가능합니다.
                img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                images.append(img)
                durations.append(int(frame.get('delay', 200)))

            output = io.BytesIO()
            images[0].save(
                output,
                format='WEBP',
                save_all=True,
                append_images=images[1:],
                duration=durations,
                loop=loop_val,
                background=(0, 0, 0, 0) # 배경 투명 처리
            )

            webp_base64 = base64.b64encode(output.getvalue()).decode('utf-8')
            
            return {
                "success": True,
                "data": webp_base64,
                "filename": "pentong_animation.webp"
            }
            
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": f"백엔드 WebP 변환 실패: {str(e)}"}

    def save_webp_dialog(self, window, b64_data):
        """ 생성된 WebP 데이터를 받아 '다른 이름으로 저장' 대화상자를 띄웁니다. """
        try:
            if not window:
                return {"success": False, "message": "창 객체가 없습니다."}
            file_types = ('WebP 애니메이션 (*.webp)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='pentong_animation.webp')
            
            if save_path and len(save_path) > 0:
                with open(save_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}

    def generate_apng(self, payload):
        """ HTML에서 전달받은 Base64 프레임 데이터를 모아 투명 배경 APNG로 병합하여 반환합니다. """
        try:
            frames_data = payload.get('frames', [])
            settings = payload.get('settings', {})
            loop_val = int(settings.get('loop', 0))

            if not frames_data:
                return {"success": False, "error": "프레임 데이터가 전달되지 않았습니다."}

            images = []
            durations = []

            for frame in frames_data:
                img_bytes = base64.b64decode(frame['base64'])
                # PNG는 RGBA를 완벽 지원합니다.
                img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                images.append(img)
                durations.append(int(frame.get('delay', 200)))

            output = io.BytesIO()
            # Pillow의 PNG 저장 시 save_all=True를 쓰면 APNG가 됩니다.
            images[0].save(
                output,
                format='PNG',
                save_all=True,
                append_images=images[1:],
                duration=durations,
                loop=loop_val
            )

            apng_base64 = base64.b64encode(output.getvalue()).decode('utf-8')
            
            return {
                "success": True,
                "data": apng_base64,
                "filename": "pentong_animation.png"
            }
            
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": f"백엔드 APNG 변환 실패: {str(e)}"}

    def save_apng_dialog(self, window, b64_data):
        """ 생성된 APNG 데이터를 받아 '다른 이름으로 저장' 대화상자를 띄웁니다. """
        try:
            if not window:
                return {"success": False, "message": "창 객체가 없습니다."}
            file_types = ('APNG 애니메이션 (*.png)', '모든 파일 (*.*)')
            save_path = window.create_file_dialog(webview.SAVE_DIALOG, file_types=file_types, save_filename='pentong_animation.png')
            
            if save_path and len(save_path) > 0:
                with open(save_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return {"success": True, "message": "저장되었습니다."}
            return {"success": False, "message": "취소됨"}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "message": str(e)}