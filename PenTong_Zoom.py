# -*- coding: utf-8 -*-
import os
import sys
import base64
import threading
import time
import webview
from io import BytesIO
import tkinter as tk
import ctypes
import keyboard

# 윈도우 화면 배율(DPI) 강제 인식
try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ToolTip:
    def __init__(self, widget, text, cvs=None, item=None):
        self.widget = widget
        self.cvs = cvs
        self.item = item
        self.text = text
        self.tipwindow = None
        self.id = None
        if self.cvs and self.item:
            self.cvs.tag_bind(self.item, "<Enter>", self.enter)
            self.cvs.tag_bind(self.item, "<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(300, self.showtip)

    def unschedule(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def showtip(self, event=None):
        if self.tipwindow or not self.text: return
        x = self.widget.winfo_pointerx() + 15
        y = self.widget.winfo_pointery() + 15

        sw = self.widget.winfo_screenwidth()
        if x + 150 > sw:  
            x = self.widget.winfo_pointerx() - 160  

        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#2b2b2b", foreground="#ffffff", relief=tk.FLAT,
                         font=("Malgun Gothic", 9), padx=8, pady=4)
        label.pack()

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw: tw.destroy()

class TkToolbar:
    def __init__(self, zoom_service):
        self.zoom_service = zoom_service
        self.root = None
        self.is_capturing = False
        self.active_color = 'blue'
        self.active_tool = 'pen'
        self.is_fade = False
        self.is_timer = False
        self.tool_inds = {}
        self.color_inds = {}
        self.is_visible = False

    def run(self):
        self.root = tk.Tk()
        self.root.title("PenTongTkToolbar")
        self.root.overrideredirect(True)
        self.root.wm_attributes("-topmost", True)
        self.root.wm_attributes("-transparentcolor", "#FF00FF") 
        self.root.config(bg="#FF00FF")

        self.root.wm_attributes("-alpha", 0.35) 
        self.root.bind("<Enter>", lambda e: self.root.wm_attributes("-alpha", 1.0))
        self.root.bind("<Leave>", lambda e: self.root.wm_attributes("-alpha", 0.35))

        w = 56 
        self.cvs = tk.Canvas(self.root, width=w, bg="#FF00FF", highlightthickness=0)
        self.cvs.pack(fill=tk.BOTH, expand=True)

        x = 28 
        y = 10
        def add_space(d): nonlocal y; y += d
        def add_divider():
            nonlocal y
            self.cvs.create_line(12, y, w-12, y, fill="#ddd")
            add_space(10)

        for i in range(3):
            self.cvs.create_oval(22, y+i*5, 24, y+2+i*5, fill="#aaa", outline="")
            self.cvs.create_oval(32, y+i*5, 34, y+2+i*5, fill="#aaa", outline="")
        b_handle = self._bind_box(10, y-5, w-10, y+20, self._start_drag, self._do_drag)
        ToolTip(self.root, "패널 이동", self.cvs, b_handle)
        add_space(25)

        self.btn_cap_out = self.cvs.create_oval(12, y, 44, y+32, outline="#009688", width=2, fill="#f1f8e9")
        self.btn_cap_txt = self.cvs.create_text(x, y+16, text="📷", font=("Segoe UI Emoji", 14), fill="#333")
        b_cap = self._bind_box(12, y, 44, y+32, lambda e: self._toggle_capture())
        self.tt_cap = ToolTip(self.root, "판서 시작 [F1]", self.cvs, b_cap)
        add_space(42)
        add_divider()

        self.timer_ind = self._create_icon(x, y+16, '⏱️', '타이머 모드 [T]', lambda e: self._toggle_timer())
        add_space(32)
        add_divider()

        for col, hex_c, desc in [('blue', '#4A90E2', '파란색 [B]'), ('green', '#34A853', '초록색 [G]'), 
                                 ('yellow', '#F8E71C', '노란색 [Y]'), ('red', '#FF5252', '빨간색 [R]')]:
            self._create_color(x, y+16, col, hex_c, desc)
            add_space(32)
        add_divider()

        for tool, icon, desc in [
            ('pen', '🖊️', '일반 펜 [P]'), ('highlighter', '🖍️', '형광펜 [H]'),
            ('eraser', '🧽', '부분 지우개 [E]'),
            ('rect', '🔲', '네모 상자'), ('circle', '⭕', '동그라미'),
            ('arrow', '↗️', '화살표'), ('text', '🔤', '텍스트 입력')
        ]:
            self._create_icon(x, y+16, icon, desc, lambda e, t=tool: self._set_tool(t), is_tool=True, tool_name=tool)
            add_space(32)
        add_divider()

        self.fade_ind = self._create_icon(x, y+16, '🪄', '사라지는 펜 효과 [F]', lambda e: self._toggle_fade())
        add_space(32)
        self.pin_ind = self._create_icon(x, y+16, '📌', '필기 내용 유지 (토글)', lambda e: self._toggle_pin())
        add_space(32)
        add_divider()

        self._create_icon(x, y+16, '↩️', '되돌리기 [Ctrl+Z]', lambda e: self.zoom_service.send_command('undo'))
        add_space(32)
        self._create_icon(x, y+16, '🗑️', '전체 지우기 [C]', lambda e: self.zoom_service.send_command('clearAll'))
        add_space(32)
        self._create_icon(x, y+16, '💾', '이미지로 저장', lambda e: self.zoom_service.send_command('triggerSave'))
        add_space(32)
        self._create_icon(x, y+16, '📋', '클립보드 복사', lambda e: self.zoom_service.send_command('triggerCopy'))
        add_space(32)
        add_divider()

        self.cvs.create_oval(12, y, 44, y+32, outline="#ffcdd2", width=2, fill="#ffebee")
        self.cvs.create_text(x, y+16, text="❌", font=("Segoe UI Emoji", 12))
        b_close = self._bind_box(12, y, 44, y+32, lambda e: self.zoom_service.close_panel())
        ToolTip(self.root, "패널 완전히 닫기", self.cvs, b_close)
        add_space(42)

        h = y 
        self.cvs.config(height=h)
        
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
       
        pos = getattr(self.zoom_service, 'panel_position', 'right')
        x_pos = 20 if pos == 'left' else sw - w - 20 
        y_pos = int((sh - h) / 2) 
        self.root.geometry(f"{w}x{h}+{x_pos}+{y_pos}")

        self.cvs.lower(self._draw_round_rect(5, 5, w-5, h-5, 23, fill="#ffffff", outline="#cccccc", width=1))

        self._update_color_selection()
        self._update_tool_selection()
        
        # 🚨 [여기 추가!] 초기 핀 상태를 화면에 적용하고 자바스크립트에 알림
        self.is_pinned = getattr(self.zoom_service, 'is_pinned', True)
        self._toggle_pin(force_state=self.is_pinned)

        self.root.update_idletasks()
        self._make_uncapturable()
        
        self.is_visible = True
        self.root.mainloop()
    
    def _toggle_pin(self, force_state=None):
        if force_state is not None: self.is_pinned = force_state
        else: self.is_pinned = not self.is_pinned
        
        # 버튼 색상 활성화/비활성화
        self.cvs.itemconfig(self.pin_ind, outline="#009688" if self.is_pinned else "", fill="#e0f2f1" if self.is_pinned else "")
        # 자바스크립트(도화지)로 핀 상태 전송
        self.zoom_service.send_command('togglePin', str(self.is_pinned).lower())

    def _make_uncapturable(self):
        try:
            hwnd = ctypes.windll.user32.FindWindowW(None, "PenTongTkToolbar")
            if hwnd: ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, 0)
        except: pass

    def _draw_round_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2, x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1]
        return self.cvs.create_polygon(points, smooth=True, **kwargs)

    def _create_color(self, x, y, name, color, tt_text):
        ind = self.cvs.create_oval(x-14, y-14, x+14, y+14, outline="", width=2)
        self.color_inds[name] = ind
        self.cvs.create_oval(x-9, y-9, x+9, y+9, fill=color, outline="")
        box = self._bind_box(x-15, y-15, x+15, y+15, lambda e, n=name: self._set_color(n))
        ToolTip(self.root, tt_text, self.cvs, box)

    def _create_icon(self, x, y, icon, tt_text, cmd, is_tool=False, tool_name=None):
        ind = self.cvs.create_oval(x-16, y-16, x+16, y+16, outline="", width=2)
        self.cvs.create_text(x, y, text=icon, font=("Segoe UI Emoji", 14))
        box = self._bind_box(x-16, y-16, x+16, y+16, cmd)
        ToolTip(self.root, tt_text, self.cvs, box)
        if is_tool and tool_name:
            self.tool_inds[tool_name] = ind
        return ind

    def _bind_box(self, x1, y1, x2, y2, cmd, cmd2=None):
        box = self.cvs.create_rectangle(x1, y1, x2, y2, fill="", outline="")
        self.cvs.tag_bind(box, "<Button-1>", cmd)
        if cmd2: self.cvs.tag_bind(box, "<B1-Motion>", cmd2)
        return box

    def _set_color(self, mode):
        self.active_color = mode
        self._update_color_selection()
        cm = {'blue':'#4A90E2', 'green':'#34A853', 'yellow':'#F8E71C', 'red':'#FF5252'}
        self.zoom_service.send_command('setColor', cm[mode])

    def _set_tool(self, tool):
        self.active_tool = tool
        self._update_tool_selection()
        self.zoom_service.send_command('setTool', tool)

    def _update_color_selection(self):
        for n, item in self.color_inds.items(): self.cvs.itemconfig(item, outline="")
        if self.active_color in self.color_inds:
            self.cvs.itemconfig(self.color_inds[self.active_color], outline="#009688")

    def _update_tool_selection(self):
        for n, item in self.tool_inds.items(): self.cvs.itemconfig(item, outline="", fill="")
        if self.active_tool in self.tool_inds:
            self.cvs.itemconfig(self.tool_inds[self.active_tool], outline="#009688", fill="#e0f2f1")

    def _toggle_timer(self, force_state=None):
        if force_state is not None: self.is_timer = force_state
        else: self.is_timer = not self.is_timer
        self.cvs.itemconfig(self.timer_ind, outline="#009688" if self.is_timer else "", fill="#e0f2f1" if self.is_timer else "")
        self.zoom_service.send_command('toggleTimer', str(self.is_timer).lower())

    def _toggle_fade(self, force_state=None):
        if force_state is not None: self.is_fade = force_state
        else: self.is_fade = not self.is_fade
        self.cvs.itemconfig(self.fade_ind, outline="#009688" if self.is_fade else "", fill="#e0f2f1" if self.is_fade else "")
        self.zoom_service.send_command('toggleFade', str(self.is_fade).lower())

    def _toggle_capture(self, state=None):
        # 🚨 판서 시작 버튼을 누르는 즉시 툴팁을 강제로 숨겨서 캡처에 찍히지 않게 함
        try: self.tt_cap.hidetip()
        except: pass
        
        if state is not None: self.is_capturing = state
        else: self.is_capturing = not self.is_capturing
        self.zoom_service.toggle_capture(self.is_capturing)

    def update_capture_ui(self, state):
        self.is_capturing = state
        def _update():
            if self.is_capturing:
                self.cvs.itemconfig(self.btn_cap_out, outline="#f44336", fill="#ffebee")
                self.cvs.itemconfig(self.btn_cap_txt, text="⏹️")
                self.tt_cap.text = "판서 종료 [ESC]"
            else:
                self.cvs.itemconfig(self.btn_cap_out, outline="#009688", fill="#f1f8e9")
                self.cvs.itemconfig(self.btn_cap_txt, text="📷")
                self.tt_cap.text = "판서 시작 [F1]"
        self.root.after(0, _update)

    def set_color_ui(self, mode):
        self.root.after(0, lambda: self._set_color(mode))
        
    def set_tool_ui(self, tool):
        self.root.after(0, lambda: self._set_tool(tool))

    def set_feature_ui(self, feature, state):
        if feature == 'fade': self.root.after(0, lambda: self._toggle_fade(force_state=state))
        elif feature == 'timer': self.root.after(0, lambda: self._toggle_timer(force_state=state))

    def _start_drag(self, event):
        self._x = event.x
        self._y = event.y

    def _do_drag(self, event):
        x = self.root.winfo_x() - self._x + event.x
        y = self.root.winfo_y() - self._y + event.y
        self.root.geometry(f"+{x}+{y}")

    def destroy(self):
        self.is_visible = False
        if self.root:
            try: self.root.after(0, self.root.withdraw)
            except: pass

    def show_panel(self):
        self.is_visible = True
        if self.root:
            try: 
                self.root.wm_attributes("-alpha", 0.35)
                self.root.after(0, self.root.deiconify)
            except: pass

    def set_position(self, pos):
        """패널이 재사용될 때 방향이 바뀌었으면 즉시 위치를 이동시키는 함수"""
        if not self.root: return
        def _update():
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            w = 56
            self.root.update_idletasks()
            h = self.root.winfo_height()
            x_pos = 20 if pos == 'left' else sw - w - 20
            y_pos = int((sh - h) / 2)
            self.root.geometry(f"{w}x{h}+{x_pos}+{y_pos}")
        self.root.after(0, _update) # 스레드 충돌 방지를 위해 after 사용

    def reset_to_default(self):
        self.active_color = 'blue'
        self.active_tool = 'pen'
        self.is_fade = False
        self.is_timer = False
        self._update_color_selection()
        self._update_tool_selection()
        self._toggle_fade(force_state=False)
        self._toggle_timer(force_state=False)
        self.update_capture_ui(False)
        self.zoom_service.send_command('setColor', '#4A90E2')
        self.zoom_service.send_command('setTool', 'pen')
        self.zoom_service.send_command('toggleFade', 'false')
        self.zoom_service.send_command('toggleTimer', 'false')
        self.zoom_service.send_command('clearAll')
        self.is_pinned = getattr(self.zoom_service, 'is_pinned', True)
        self._toggle_pin(force_state=self.is_pinned)

class ZoomService:
    def __init__(self, main_api):
        self.main_api = main_api
        self.toolbar = None
        self.canvas_window = None
        self.is_capturing = False
        self.target_monitor = None
        self.f1_hook = None  

    def _on_f1(self):
        if self.toolbar:
            self.toggle_capture(not self.is_capturing)

    def start_zoom_manual(self, payload='right'):
        pos = 'right'
        pin = True
        
        # HTML에서 딕셔너리로 묶어서 보낸 데이터 파싱
        if isinstance(payload, dict):
            pos = payload.get('pos', 'right')
            pin = payload.get('pin', True)
        elif isinstance(payload, str):
            pos = payload
            
        self.panel_position = pos
        self.is_pinned = pin

        # 패널이 이미 실행 중이면 위치와 핀 상태를 동시 갱신
        if self.toolbar is not None and getattr(self.toolbar, 'is_visible', True):
            self.toolbar.set_position(pos)
            self.toolbar._toggle_pin(force_state=pin)
            return {"success": True, "message": "이미 실행 중이므로 위치만 변경합니다."}
        
        if self.main_api.window and not hasattr(self, '_close_bound'):
            try:
                self.main_api.window.events.closed += self.clean_up_and_exit
                self._close_bound = True
            except: pass

        import threading
        if self.canvas_window is None:
            threading.Thread(target=self._init_system, daemon=True).start()
        else:
            threading.Thread(target=self._reuse_system, daemon=True).start()
            
        return {"success": True, "message": "판서 패널을 띄웁니다."}

    # 🚨 [새로 추가] HTML에서 위치 버튼을 클릭할 때마다 불리는 함수
    def move_panel(self, pos):
        self.panel_position = pos
        if self.toolbar and getattr(self.toolbar, 'is_visible', False):
            self.toolbar.set_position(pos)
        return {"success": True}

    def change_pin_state(self, state):
        # 자바스크립트에서 넘어온 true/false 값을 파이썬 형식으로 변환하여 저장
        self.is_pinned = state if isinstance(state, bool) else str(state).lower() == 'true'
        
        # 패널이 켜져 있다면 즉시 패널의 📌 버튼 불을 켜거나 끄고, 도화지에도 알림
        if self.toolbar and getattr(self.toolbar, 'is_visible', False):
            if hasattr(self.toolbar, 'root'):
                self.toolbar.root.after(0, lambda: self.toolbar._toggle_pin(force_state=self.is_pinned))
        return {"success": True}

    def _reuse_system(self):
        import ctypes
        import mss
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        pt = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))

        with mss.mss() as sct:
            monitors = sct.monitors
            self.target_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
            if len(monitors) > 1:
                for m in monitors[1:]:
                    if m["left"] <= pt.x < m["left"] + m["width"] and m["top"] <= pt.y < m["top"] + m["height"]:
                        self.target_monitor = m
                        break

        # 창의 크기만 미리 잡아두고 이동은 시키지 않음 (계속 우주 밖 대기)
        m = self.target_monitor
        try:
            # 👇 파이썬 공식 명령어로 창을 우주 밖에서 준비시킵니다.
            self.canvas_window.move(-30000, -30000)
            self.canvas_window.resize(m["width"], m["height"])
        except: pass

        if self.toolbar:
            self.toolbar.set_position(self.panel_position) 
            self.toolbar.show_panel()
            self.toolbar.reset_to_default()
            

        try:
            if self.f1_hook: keyboard.remove_hotkey(self.f1_hook)
            self.f1_hook = keyboard.add_hotkey('f1', self._on_f1, suppress=True)
        except: pass

    def _init_system(self):
        import ctypes
        import mss
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        pt = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))

        with mss.mss() as sct:
            monitors = sct.monitors
            self.target_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
            if len(monitors) > 1:
                for m in monitors[1:]:
                    if m["left"] <= pt.x < m["left"] + m["width"] and m["top"] <= pt.y < m["top"] + m["height"]:
                        self.target_monitor = m
                        break

        m = self.target_monitor
        
        # 완전 클린 로직: 창을 생성할 때부터 사용자 눈에 안 띄는 우주 밖(-30000)에 고정
        self.canvas_window = webview.create_window(
            '화면 판서 오버레이',
            url=resource_path('zoom_overlay.html'),
            js_api=self,
            x=-30000, y=-30000, width=m["width"], height=m["height"], 
            frameless=True, transparent=True,  
            easy_drag=False ,
            on_top=True  
        )
        self.canvas_window.events.closed += self._on_canvas_closed

        # 🚨 [여기서부터 핵심 추가!] 
        # 웹뷰 엔진이 창을 완전히 만들 때까지 기다렸다가 패널을 생성하도록 강제합니다.
        import time
        for _ in range(50):  # 최대 5초 대기
            if ctypes.windll.user32.FindWindowW(None, '화면 판서 오버레이'):
                time.sleep(0.1)  # 윈도우 시스템이 창 생성을 완전히 인식할 0.1초의 틈을 줍니다.
                break
            time.sleep(0.1)

        # 👉 판서창(webview) 생성이 끝난 것을 확인한 뒤에야 패널(Tkinter)을 생성합니다.
        # 이렇게 하면 무조건 [도화지 -> 패널] 순서가 되어 자연스러운 최상단 Z-order가 형성됩니다.
        self.toolbar = TkToolbar(self)

        try:
            if self.f1_hook: keyboard.remove_hotkey(self.f1_hook)
            self.f1_hook = keyboard.add_hotkey('f1', self._on_f1, suppress=True)
        except: pass

        self.toolbar.run()

    def canvas_initialized(self):
        import ctypes
        hwnd = ctypes.windll.user32.FindWindowW(None, '화면 판서 오버레이')
        if hwnd:
            ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, 0)
            ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
            
            # 1. 속성 세팅 (작업표시줄 숨김 + 투명 레이어)
            new_style = (ex_style | 0x80000 | 0x80) & ~0x40000
            ctypes.windll.user32.SetWindowLongW(hwnd, -20, new_style)
            
            # 🚨 2. [가장 핵심!] 창이 생성되자마자 우주 밖에서 '영구 최상단(Topmost)' 권한을 획득합니다!
            # 0x0001(크기 변경 안함) | 0x0002(이동 안함) | 0x0010(포커스 뺏지 않음) = 0x0013
            ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0013)

    def get_initial_background(self):
        return None

    def sync_toolbar_state(self, action, arg):
        if not self.toolbar: return
        if action == 'color': self.toolbar.set_color_ui(arg)
        elif action == 'tool': self.toolbar.set_tool_ui(arg)
        elif action == 'feature':
            state = arg.split('|')[1] == 'true'
            feature = arg.split('|')[0]
            self.toolbar.set_feature_ui(feature, state)

    def toggle_capture(self, state):
        self.is_capturing = state
        if self.toolbar: 
            self.toolbar.update_capture_ui(state) 
        threading.Thread(target=self._process_capture, args=(state,), daemon=True).start()

    def _process_capture(self, state):
        if state:
            # 🚨 [확실한 숨김 처리] 투명도가 아니라 창 자체를 윈도우에서 잠깐 뽑아냅니다.
            if self.toolbar and hasattr(self.toolbar, 'root') and self.toolbar.root:
                try:
                    self.toolbar.root.withdraw()  # 화면과 작업표시줄에서 완전히 소멸시킴
                    self.toolbar.root.update()
                    time.sleep(0.1)  # OS가 화면을 확실히 지울 수 있도록 0.1초의 시간을 줍니다.
                except: pass

            try:
                import mss
                from PIL import Image
                import ctypes

                class POINT(ctypes.Structure):
                    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
                pt = POINT()
                ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))

                with mss.mss() as sct:
                    monitors = sct.monitors
                    self.target_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
                    if len(monitors) > 1:
                        for m in monitors[1:]:
                            if m["left"] <= pt.x < m["left"] + m["width"] and m["top"] <= pt.y < m["top"] + m["height"]:
                                self.target_monitor = m
                                break

                    m = self.target_monitor
                    
                    sct_img = sct.grab(self.target_monitor) 
                    img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                    
                    extrema = img.convert("L").getextrema()
                    if extrema[0] == extrema[1]:
                        raise Exception("보안 프로그램 차단 (블랙 스크린 감지)")
                    
                    buffer = BytesIO()
                    img.save(buffer, format="PNG")
                    b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
                    self.bg_data = f"data:image/png;base64,{b64}"
                
                if self.canvas_window:
                    try: self.canvas_window.evaluate_js(f"updateBackground('{self.bg_data}')")
                    except: pass

            except Exception as e:
                print(f"캡처 오류: {e}")
                self.bg_data = "ERROR" 
                if self.canvas_window:
                    try: self.canvas_window.evaluate_js("updateBackground('ERROR')")
                    except: pass
            finally:
                # 🚨 [무조건 복구] 캡처가 끝나면 정상/에러 상관없이 패널을 다시 화면에 띄웁니다.
                if self.toolbar and hasattr(self.toolbar, 'root') and self.toolbar.root:
                    try:
                        self.toolbar.root.deiconify()
                    except: pass

        else:
            if self.canvas_window:
                try: self.canvas_window.evaluate_js("endCapture()")
                except: pass
                
                # 파이썬 공식 명령어로 창을 우주 밖으로 치웁니다.
                try: self.canvas_window.move(-30000, -30000)
                except: pass

    def lift_panel(self):
        """JS에서 마우스를 뗄 때 호출하여 패널을 강제로 최상단으로 올림"""
        if self.toolbar and hasattr(self.toolbar, 'root'):
            def _lift():
                try:
                    self.toolbar.root.attributes('-topmost', True)
                    self.toolbar.root.lift()
                except: pass
            # Tkinter 스레드 안전성을 위해 after 사용
            self.toolbar.root.after(0, _lift)
                
    def on_canvas_ready(self):
        if not self.is_capturing: return
        if self.canvas_window:
            m = self.target_monitor
            try:
                import ctypes, time
                hwnd_cvs = ctypes.windll.user32.FindWindowW(None, '화면 판서 오버레이')
                hwnd_tk = ctypes.windll.user32.FindWindowW(None, 'PenTongTkToolbar')
                
                if hwnd_cvs:
                    # 1. 투명도 0으로 가림
                    ctypes.windll.user32.SetLayeredWindowAttributes(hwnd_cvs, 0, 0, 0x2)
                    
                    # 🚨 엔진 명령어(resize, move)를 지우고 API로 '조용히' 좌표 이동
                    # 0x0610 = NOACTIVATE(16) + NOOWNERZORDER(512) + NOSENDCHANGING(1024)
                    target_z = hwnd_tk if hwnd_tk else -1
                    ctypes.windll.user32.SetWindowPos(
                        hwnd_cvs, target_z, 
                        m["left"], m["top"], m["width"], m["height"], 
                        0x0610 | 0x0040 # 0x0040 = SHOWWINDOW
                    )
                
                    # 2. 선생님의 0.05초 정답 로직 유지
                    time.sleep(0.05)
                    ctypes.windll.user32.SetLayeredWindowAttributes(hwnd_cvs, 0, 255, 0x2)
            except: pass

    def save_image(self, b64_data):
        def _save():
            import tkinter.filedialog as fd
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            path = fd.asksaveasfilename(defaultextension=".png", filetypes=[("PNG 이미지", "*.png")], title="판서 저장")
            root.destroy()
            
            if path:
                try:
                    header, encoded = b64_data.split(",", 1)
                    with open(path, "wb") as f:
                        f.write(base64.b64decode(encoded))
                    self.send_command('showToast', '💾 이미지가 저장되었습니다.')
                except Exception as e: print("Save error:", e)
        threading.Thread(target=_save, daemon=True).start()
        return True

    def copy_image(self, b64_data):
        def _copy():
            import tempfile
            import subprocess
            try:
                header, encoded = b64_data.split(",", 1)
                data = base64.b64decode(encoded)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(data)
                    tmp_name = tmp.name
                cmd = f'Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile("{tmp_name}"))'
                subprocess.run(["powershell", "-command", cmd], creationflags=subprocess.CREATE_NO_WINDOW)
                self.send_command('showToast', '📋 클립보드에 복사되었습니다. (Ctrl+V)')
            except Exception as e: print("Copy error:", e)
        threading.Thread(target=_copy, daemon=True).start()
        return True

    def send_command(self, cmd, arg=None):
        if self.canvas_window:
            try:
                if arg is not None: self.canvas_window.evaluate_js(f"{cmd}('{arg}')")
                else: self.canvas_window.evaluate_js(f"{cmd}()")
            except Exception:
                pass 

    def close_panel(self):
        self.is_capturing = False

        if self.canvas_window: 
            try: self.canvas_window.evaluate_js("endCapture()")
            except: pass
            
            # 👇 패널 닫을 때도 파이썬 공식 명령어로 우주 밖으로 치웁니다.
            try: self.canvas_window.move(-30000, -30000)
            except: pass

        if self.toolbar: 
            self.toolbar.reset_to_default()
            self.toolbar.destroy()
            
        if self.main_api.window:
            try:
                self.main_api.window.evaluate_js("setZoomStatus(false)")
            except: pass

        try:
            if getattr(self, 'f1_hook', None):
                keyboard.remove_hotkey(self.f1_hook)
                self.f1_hook = None
        except: pass

    def clean_up_and_exit(self):
        import ctypes
        try:
            if getattr(self, 'f1_hook', None):
                keyboard.remove_hotkey(self.f1_hook)
                self.f1_hook = None
            keyboard.unhook_all()
        except: pass

        try:
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
        except: pass

        if self.toolbar and hasattr(self.toolbar, 'root') and self.toolbar.root:
            try:
                def _kill_tk():
                    try: self.toolbar.root.quit()
                    except: pass
                    try: self.toolbar.root.destroy()
                    except: pass
                self.toolbar.root.after(0, _kill_tk) 
            except: pass

        if self.canvas_window:
            def _kill_canvas():
                import time
                time.sleep(0.1) 
                try: self.canvas_window.destroy()
                except: pass
            threading.Thread(target=_kill_canvas, daemon=True).start()

    def _on_canvas_closed(self):
        self.canvas_window = None