import os
import sqlite3
import threading
import time
import sys
import json
from datetime import datetime, timedelta

main_window = None

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "PenTong_Data")
DB_PATH = os.path.join(DATA_DIR, "schedule.db")

def init_db():
    if not os.path.exists(DATA_DIR):
        try: os.makedirs(DATA_DIR)
        except: return
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, content TEXT, category TEXT,
            is_all_day BOOLEAN, start_date TEXT, start_time TEXT, end_date TEXT,
            end_time TEXT, use_sound BOOLEAN DEFAULT 1
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS schedule_alarms (
            id INTEGER PRIMARY KEY AUTOINCREMENT, schedule_id INTEGER, alarm_minutes INTEGER, is_triggered BOOLEAN DEFAULT 0
        )
    ''')
    cursor.execute("PRAGMA table_info(schedules)")
    if 'color' not in [info[1] for info in cursor.fetchall()]:
        cursor.execute("ALTER TABLE schedules ADD COLUMN color TEXT DEFAULT '#e3f2fd'")

    conn.commit()
    conn.close()

init_db()

def snooze_alarm_db(alarm_id, snooze_mins, s_date, s_time):
    try:
        now = datetime.now()
        snooze_target = now + timedelta(minutes=int(snooze_mins))
        sch_datetime = datetime.strptime(f"{s_date} {s_time}", "%Y-%m-%d %H:%M")
        new_alarm_mins = int((sch_datetime - snooze_target).total_seconds() / 60)
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("UPDATE schedule_alarms SET is_triggered=0, alarm_minutes=? WHERE id=?", (new_alarm_mins, alarm_id))
        conn.commit(); conn.close()
    except Exception as e: print(f"Snooze error: {e}")

class ScheduleAPI:
    def save_schedule(self, data):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            is_all_day = 1 if data['isAllDay'] else 0
            use_sound = 1 if data.get('useSound', True) else 0
            sch_id = data.get('id')
            content = data.get('content', '')
            alarms = data.get('alarms', []) 
            color = data.get('color', '#e3f2fd')

            if sch_id:
                cursor.execute('''UPDATE schedules SET title=?, content=?, category=?, is_all_day=?, start_date=?, start_time=?, end_date=?, end_time=?, use_sound=?, color=? WHERE id=?''', 
                               (data['title'], content, data['category'], is_all_day, data['startDate'], data['startTime'], data['endDate'], data['endTime'], use_sound, color, sch_id))
                cursor.execute("DELETE FROM schedule_alarms WHERE schedule_id=?", (sch_id,))
            else:
                cursor.execute('''INSERT INTO schedules (title, content, category, is_all_day, start_date, start_time, end_date, end_time, use_sound, color) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                               (data['title'], content, data['category'], is_all_day, data['startDate'], data['startTime'], data['endDate'], data['endTime'], use_sound, color))
                sch_id = cursor.lastrowid

            for mins in alarms:
                cursor.execute("INSERT INTO schedule_alarms (schedule_id, alarm_minutes, is_triggered) VALUES (?, ?, 0)", (sch_id, int(mins)))
                
            conn.commit(); conn.close()
            return {"status": "success"}
        except Exception as e: return {"status": "error", "message": str(e)}

    def get_schedules(self, start_date_str, end_date_str):
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''SELECT * FROM schedules WHERE start_date <= ? AND end_date >= ? ORDER BY start_date ASC, start_time ASC''', (end_date_str, start_date_str))
            rows = cursor.fetchall()
            schedules = [dict(row) for row in rows]
            
            for sch in schedules:
                cursor.execute("SELECT alarm_minutes FROM schedule_alarms WHERE schedule_id=? ORDER BY alarm_minutes DESC", (sch['id'],))
                sch['alarms'] = [r['alarm_minutes'] for r in cursor.fetchall()]
                
            conn.close()
            return {"status": "success", "data": schedules}
        except Exception as e: return {"status": "error", "message": str(e)}

    def snooze_alarm_api(self, data):
        snooze_alarm_db(data['alarm_id'], data['mins'], data['s_date'], data['s_time'])
        return {"status": "success"}
        
    def delete_schedule(self, sch_id):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM schedules WHERE id=?", (sch_id,))
            cursor.execute("DELETE FROM schedule_alarms WHERE schedule_id=?", (sch_id,))
            conn.commit(); conn.close()
            return {"status": "success"}
        except Exception as e: return {"status": "error", "message": str(e)}

    def get_upcoming_alarms(self):
        try:
            now = datetime.now()
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT s.title, s.start_date, s.start_time, a.alarm_minutes 
                FROM schedule_alarms a JOIN schedules s ON a.schedule_id = s.id 
                WHERE a.is_triggered = 0 AND s.is_all_day = 0
            ''')
            rows = cursor.fetchall()
            upcoming = []
            
            for row in rows:
                title, s_date, s_time, alarm_min = row
                if s_date and s_time:
                    try:
                        sch_dt = datetime.strptime(f"{s_date} {s_time}", "%Y-%m-%d %H:%M")
                        alarm_time = sch_dt - timedelta(minutes=alarm_min)
                        diff_mins = (alarm_time - now).total_seconds() / 60
                        if 0 < diff_mins <= 20:
                            upcoming.append({"title": title, "mins_left": int(diff_mins) + 1})
                    except: pass
            conn.close()
            
            if upcoming:
                upcoming.sort(key=lambda x: x["mins_left"])
                return {"status": "success", "data": upcoming[0]}
            return {"status": "empty"}
        except Exception as e: return {"status": "error", "message": str(e)}


def trigger_persistent_alarm(alarm_id, sch_id, title, content, s_date, s_time, use_sound):
    global main_window
    success = False
    
    if main_window:
        try:
            safe_aid = json.dumps(alarm_id); safe_sid = json.dumps(sch_id); safe_title = json.dumps(title)
            safe_content = json.dumps(content or '내용 없음'); safe_sdate = json.dumps(s_date); safe_stime = json.dumps(s_time)
            
            sound_code = ""
            if use_sound == 1:
                sound_code = "try {var ctx=new(window.AudioContext||window.webkitAudioContext)();var osc=ctx.createOscillator();osc.type='square';osc.frequency.setValueAtTime(800,ctx.currentTime);osc.frequency.setValueAtTime(1200,ctx.currentTime+0.1);var gain=ctx.createGain();gain.gain.setValueAtTime(0.1,ctx.currentTime);osc.connect(gain);gain.connect(ctx.destination);osc.start();osc.stop(ctx.currentTime+0.3);}catch(e){}"

            refresh_calendar_code = "try{var f=document.getElementsByTagName('iframe');for(var i=0;i<f.length;i++){if(f[i].contentWindow&&typeof f[i].contentWindow.renderCalendar==='function'){f[i].contentWindow.renderCalendar();}}}catch(e){}"

            js_code = f"""
            (function() {{
                var overlay = document.createElement('div');
                overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.4);z-index:2147483647;display:flex;justify-content:center;align-items:center;backdrop-filter:blur(3px);';
                
                var box = document.createElement('div');
                box.style.cssText = 'background:#fff;padding:25px 20px 20px 20px;border-radius:15px;text-align:center;width:280px;display:flex;flex-direction:column;box-shadow:0 8px 30px rgba(0,0,0,0.3);border:3px solid #ff9f43;font-family:sans-serif;position:relative;';
                
                var dragHandle = document.createElement('div');
                dragHandle.style.cssText = 'position:absolute; top:0; left:0; width:100%; height:30px; cursor:move; border-radius:15px 15px 0 0; background:transparent; z-index:10;';
                box.appendChild(dragHandle);

                box.innerHTML += `
                    <div style="font-size:35px;margin-bottom:5px;animation:shake 0.5s infinite;pointer-events:none;">⏰</div>
                    <style>@keyframes shake {{ 0% {{transform:rotate(-10deg);}} 50% {{transform:rotate(10deg);}} 100% {{transform:rotate(-10deg);}} }}</style>
                    <h3 style="color:#333;margin:0 0 10px 0;word-break:break-all;font-size:15px;">${{{safe_title}}}</h3>
                    <div style="color:#666;font-size:12px;white-space:pre-wrap;margin-bottom:15px;background:#fff8e1;padding:10px;border-radius:8px;text-align:left;overflow-y:auto;max-height:120px;border:1px solid #ffe082;">${{{safe_content}}}</div>
                    
                    <div style="background:#f1f3f5; padding:10px; border-radius:8px; margin-bottom:15px; display:flex; justify-content:center; align-items:center; gap:5px; font-size:12px;">
                        <input type="number" id="snoozeMins" value="5" min="1" style="width:45px; text-align:center; padding:4px; border:1px solid #ccc; border-radius:4px; font-size:12px; outline:none;">
                        <span style="font-weight:bold; color:#555;">분 뒤 다시 알림</span>
                        <button id="snooze-btn-x" style="background:#2e86de; color:white; border:none; padding:6px 12px; border-radius:4px; cursor:pointer; font-weight:bold; font-size:12px; transition:0.2s;">적용</button>
                    </div>
                    
                    <button id="close-alarm-btn-x" style="background:#ff9f43;color:white;border:none;padding:10px 0;border-radius:20px;font-size:13px;cursor:pointer;font-weight:bold;box-shadow:0 4px 10px rgba(255,159,67,0.3);width:100%; transition:0.2s;">확인 (알림 끄기)</button>
                `;
                overlay.appendChild(box); document.body.appendChild(overlay);
                
                var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
                dragHandle.onmousedown = function(e) {{
                    e.preventDefault();
                    pos3 = e.clientX; pos4 = e.clientY;
                    document.onmouseup = function() {{ document.onmouseup = null; document.onmousemove = null; }};
                    document.onmousemove = function(e) {{
                        e.preventDefault();
                        pos1 = pos3 - e.clientX; pos2 = pos4 - e.clientY;
                        pos3 = e.clientX; pos4 = e.clientY;
                        box.style.position = 'absolute';
                        box.style.top = (box.offsetTop - pos2) + "px";
                        box.style.left = (box.offsetLeft - pos1) + "px";
                        box.style.margin = "0"; 
                    }};
                }};

                document.getElementById('close-alarm-btn-x').onclick = function() {{ 
                    this.innerText = '알림이 꺼졌습니다'; this.style.background = '#ced4da'; this.style.color = '#495057'; this.style.boxShadow = 'none';
                    setTimeout(() => {{ document.body.removeChild(overlay); {refresh_calendar_code} }}, 800);
                }};
                document.getElementById('snooze-btn-x').onclick = function() {{
                    var mins = document.getElementById('snoozeMins').value;
                    if(window.pywebview && window.pywebview.api) {{ window.pywebview.api.snooze_alarm_api({{alarm_id: {{{safe_aid}}}, mins: mins, s_date: {{{safe_sdate}}}, s_time: {{{safe_stime}}}}}); }}
                    this.innerText = '적용됨!'; this.style.background = '#51cf66';
                    setTimeout(() => {{ document.body.removeChild(overlay); {refresh_calendar_code} }}, 800);
                }};
                {sound_code}
            }})();
            """
            main_window.show()
            main_window.restore()
            main_window.evaluate_js(js_code)
            success = True
        except: pass

    # 🛑 [수정] 윈도우 창에서는 마우스 커서 분리 적용
    if not success:
        def show_tk():
            import tkinter as tk
            root = tk.Tk()
            
            root.overrideredirect(True) 
            root.attributes('-topmost', True)
            root.configure(bg="#ff9f43") 
            
            # 전체 화면은 일반 커서
            main_frame = tk.Frame(root, bg="white", bd=0)
            main_frame.pack(fill="both", expand=True, padx=3, pady=3) 

            # 드래그 앤 드롭 함수
            def start_drag(e):
                root.x = e.x
                root.y = e.y

            def do_drag(e):
                x = root.winfo_x() - root.x + e.x
                y = root.winfo_y() - root.y + e.y
                root.geometry(f"+{x}+{y}")

            # 🛑 상단에 드래그 전용 투명 핸들 배치
            drag_bar = tk.Frame(main_frame, bg="white", height=15, cursor="fleur")
            drag_bar.pack(fill="x")
            drag_bar.bind("<Button-1>", start_drag)
            drag_bar.bind("<B1-Motion>", do_drag)

            if use_sound == 1:
                try: print('\a')
                except: pass

            # 아이콘 및 타이틀에만 십자 커서(fleur) 적용
            alarm_icon = tk.Label(main_frame, text="⏰", font=("Arial", 28), bg="white", cursor="fleur")
            alarm_icon.pack(pady=(0, 5))
            alarm_icon.bind("<Button-1>", start_drag); alarm_icon.bind("<B1-Motion>", do_drag)
            
            title_label = tk.Label(main_frame, text=title, font=("Malgun Gothic", 12, "bold"), bg="white", wraplength=230, cursor="fleur")
            title_label.pack(pady=0, padx=10)
            title_label.bind("<Button-1>", start_drag); title_label.bind("<B1-Motion>", do_drag)
            
            # 나머지 요소는 기본 커서 및 버튼은 손가락(hand2) 커서
            text_frame = tk.Frame(main_frame, bg="white")
            text_frame.pack(pady=10, fill="both", expand=True, padx=15)
            text_widget = tk.Text(text_frame, font=("Malgun Gothic", 10), bg="#fff8e1", wrap="word", height=4, width=25, relief="flat", padx=8, pady=8)
            scrollbar = tk.Scrollbar(text_frame, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            text_widget.insert("1.0", content or '내용 없음')
            text_widget.config(state="disabled")
            text_widget.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            frame = tk.Frame(main_frame, bg="#f1f3f5", padx=5, pady=5)
            frame.pack(pady=5, fill="x", padx=15)
            
            snooze_var = tk.StringVar(value="5")
            tk.Entry(frame, textvariable=snooze_var, width=5, font=("Arial", 11), justify="center").pack(side=tk.LEFT, padx=5)
            tk.Label(frame, text="분 뒤 다시 알림", bg="#f1f3f5", font=("Malgun Gothic", 9, "bold")).pack(side=tk.LEFT)
            
            def on_snooze():
                snooze_alarm_db(alarm_id, int(snooze_var.get() or 5), s_date, s_time)
                snooze_btn.config(text="적용됨!", bg="#51cf66")
                root.after(800, root.destroy)
                
            def on_close():
                close_btn.config(text="알림이 꺼졌습니다", bg="#ced4da", fg="#495057")
                root.after(800, root.destroy)

            # 🛑 버튼에 손가락 모양(hand2) 추가
            snooze_btn = tk.Button(frame, text="적용", command=on_snooze, bg="#2e86de", fg="white", font=("Malgun Gothic", 9, "bold"), relief="flat", cursor="hand2")
            snooze_btn.pack(side=tk.RIGHT, padx=5)
            
            close_btn = tk.Button(main_frame, text="확인 (알림 끄기)", command=on_close, bg="#ff9f43", fg="white", font=("Malgun Gothic", 11, "bold"), relief="flat", pady=5, cursor="hand2")
            close_btn.pack(pady=10, fill="x", padx=20)
            
            root.update_idletasks()
            width = 280
            height = int(root.winfo_reqheight())
            x = (root.winfo_screenwidth() // 2) - (width // 2)
            y = (root.winfo_screenheight() // 2) - (height // 2)
            root.geometry(f'{width}x{height}+{x}+{y}')
            
            root.mainloop()
            
        threading.Thread(target=show_tk, daemon=True).start()

def alarm_checker():
    time.sleep(5) 
    while True:
        try:
            now = datetime.now()
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT a.id, s.id, s.title, s.content, s.start_date, s.start_time, a.alarm_minutes, s.use_sound, s.category
                FROM schedule_alarms a JOIN schedules s ON a.schedule_id = s.id 
                WHERE a.is_triggered = 0 AND s.is_all_day = 0
            ''')
            rows = cursor.fetchall()
            
            for row in rows:
                a_id, s_id, title, content, s_date, s_time, alarm_min, use_sound, category = row
                
                if s_date and s_time:
                    try:
                        sch_datetime = datetime.strptime(f"{s_date} {s_time}", "%Y-%m-%d %H:%M")
                        alarm_time = sch_datetime - timedelta(minutes=alarm_min)
                        
                        if now >= alarm_time:
                            display_title = title
                            if (now - alarm_time).total_seconds() > 300: 
                                display_title = "[놓친 알림] " + title

                            trigger_persistent_alarm(a_id, s_id, display_title, content, s_date, s_time, use_sound)
                            
                            if category == 'quick':
                                cursor.execute("DELETE FROM schedules WHERE id = ?", (s_id,))
                                cursor.execute("DELETE FROM schedule_alarms WHERE schedule_id = ?", (s_id,))
                            else:
                                cursor.execute("UPDATE schedule_alarms SET is_triggered = 1 WHERE id = ?", (a_id,))
                            conn.commit()
                    except ValueError: pass
            conn.close()
        except: pass
        time.sleep(15) 

def start_background_service(): 
    threading.Thread(target=alarm_checker, daemon=True).start()