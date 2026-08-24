import platform
import socket
import psutil
import subprocess
import datetime
import uuid
import re

class PCManager:
    def __init__(self):
        pass

    def get_all_system_info(self):
        """
        한 번의 호출로 OS, CPU, RAM, 디스크, 네트워크 등 모든 정보를 딕셔너리로 반환합니다.
        """
        info = {
            "os": {},
            "cpu": {},
            "ram": {},
            "disk": [],
            "network": {}
        }

        try:
            # 1. OS 정보
            info["os"]["system"] = f"{platform.system()} {platform.release()}"
            info["os"]["version"] = platform.version()
            info["os"]["architecture"] = platform.machine()
            
            boot_time_timestamp = psutil.boot_time()
            bt = datetime.datetime.fromtimestamp(boot_time_timestamp)
            info["os"]["boot_time"] = f"{bt.year}-{bt.month:02d}-{bt.day:02d} {bt.hour:02d}:{bt.minute:02d}:{bt.second:02d}"

            # 2. CPU 정보
            info["cpu"]["name"] = platform.processor()
            info["cpu"]["physical_cores"] = psutil.cpu_count(logical=False)
            info["cpu"]["total_cores"] = psutil.cpu_count(logical=True)
            info["cpu"]["usage"] = f"{psutil.cpu_percent(interval=0.1)}%"

            # [추가] GPU 및 메인보드 정보 (콘솔창 깜빡임 방지 처리)
            try:
                creation_flags = 0x08000000 if platform.system() == "Windows" else 0
                
                gpu_cmd = 'powershell "Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name"'
                gpu_result = subprocess.run(gpu_cmd, shell=True, capture_output=True, text=True, creationflags=creation_flags)
                gpu_name = gpu_result.stdout.strip().replace('\n', ', ')
                info["cpu"]["gpu"] = gpu_name if gpu_name else "알 수 없음"
                
                mb_cmd = 'powershell "Get-CimInstance Win32_BaseBoard | Select-Object -ExpandProperty Product"'
                mb_result = subprocess.run(mb_cmd, shell=True, capture_output=True, text=True, creationflags=creation_flags)
                mb_name = mb_result.stdout.strip()
                info["cpu"]["motherboard"] = mb_name if mb_name else "알 수 없음"
            except:
                info["cpu"]["gpu"] = "정보 없음"
                info["cpu"]["motherboard"] = "정보 없음"

            # 3. RAM 정보
            svmem = psutil.virtual_memory()
            info["ram"]["total"] = self._get_size(svmem.total)
            info["ram"]["used"] = self._get_size(svmem.used)
            info["ram"]["available"] = self._get_size(svmem.available)
            info["ram"]["percentage"] = f"{svmem.percent}%"

            # 4. 디스크 정보 (모든 파티션)
            partitions = psutil.disk_partitions()
            for partition in partitions:
                # CD-ROM 등의 드라이브는 용량 확인 시 에러가 발생하므로 예외 처리
                try:
                    partition_usage = psutil.disk_usage(partition.mountpoint)
                    info["disk"].append({
                        "device": partition.device,
                        "total": self._get_size(partition_usage.total),
                        "used": self._get_size(partition_usage.used),
                        "free": self._get_size(partition_usage.free),
                        "percentage": f"{partition_usage.percent}%"
                    })
                except PermissionError:
                    continue

            # 5. 네트워크 정보
            info["network"]["hostname"] = socket.gethostname()
            info["network"]["ip"] = socket.gethostbyname(socket.gethostname())
            info["network"]["mac"] = ':'.join(re.findall('..', '%012x' % uuid.getnode())).upper()

        except Exception as e:
            info["error"] = str(e)

        return info

    def run_management_tool(self, tool_id):
        """
        지정된 관리 도구를 실행합니다.
        """
        tools = {
            "control_panel": "control",
            "device_manager": "devmgmt.msc",
            "network_connections": "ncpa.cpl",
            "printers": "control printers",
            "programs": "appwiz.cpl",
            "sys_properties": "sysdm.cpl",
            "disk_management": "diskmgmt.msc",
            "services": "services.msc",
            "regedit": "regedit",
            "taskmgr": "taskmgr",
            "cmd": "start cmd",
            "calc": "calc",
            "notepad": "notepad",
            "mspaint": "mspaint",
            "soundrecorder": "explorer ms-soundrecorder:",
            "snippingtool": "snippingtool",
            "explorer": "explorer",
            "magnify": "magnify",
            "osk": "osk",
            "dxdiag": "dxdiag",
            "eventvwr": "eventvwr.msc"
        }

        command = tools.get(tool_id)
        if command:
            try:
                # 백그라운드에서 비동기로 실행하여 프로그램이 멈추지 않게 함
                subprocess.Popen(command, shell=True)
                return {"status": "success", "message": f"{tool_id} 실행 완료"}
            except Exception as e:
                return {"status": "error", "message": str(e)}
        else:
            return {"status": "error", "message": "알 수 없는 도구입니다."}

    def change_pc_name(self, new_name):
        """
        컴퓨터 이름을 변경합니다. (관리자 권한 필요)
        """
        if not new_name or len(new_name.strip()) == 0:
            return {"status": "error", "message": "새 컴퓨터 이름을 입력해주세요."}
            
        try:
            # PowerShell 명령어로 컴퓨터 이름 변경 시도
            command = f'powershell.exe -Command "Rename-Computer -NewName {new_name.strip()}"'
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                return {"status": "success", "message": f"컴퓨터 이름이 '{new_name}'(으)로 변경되었습니다.\n적용을 위해 컴퓨터를 재부팅해야 합니다."}
            else:
                return {"status": "error", "message": f"이름 변경 실패. 관리자 권한으로 실행되었는지 확인해주세요.\n오류: {result.stderr}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def _get_size(self, bytes, suffix="B"):
        """바이트 단위를 읽기 쉬운 포맷(KB, MB, GB 등)으로 변환"""
        factor = 1024
        for unit in ["", "K", "M", "G", "T", "P"]:
            if bytes < factor:
                return f"{bytes:.2f}{unit}{suffix}"
            bytes /= factor