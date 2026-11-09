# -*- coding: utf-8 -*-
import os
import sys
import re
import gc
import json
import time
import ctypes
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta

# PyInstaller windowed 모드 스트림 보호
class NullWriter:
    def write(self, s):
        pass
    def flush(self):
        pass

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()

try:
    import cv2
    import numpy as np
    from PIL import Image, ImageTk
    CV_AVAILABLE = True
except ImportError:
    CV_AVAILABLE = False

try:
    import winsound
except ImportError:
    winsound = None

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. 중복 실행 방지 (Single Instance Lock)
# ==========================================
SINGLE_INSTANCE_MUTEX = None
if os.name == 'nt':
    try:
        kernel32 = ctypes.windll.kernel32
        MUTEX_NAME = "Aluko_S_Structure_Inspector_Mutex"
        SINGLE_INSTANCE_MUTEX = kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if kernel32.GetLastError() == 183:
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showwarning("중복 실행 경고", "이미 Aluko S-Structure 외관 검사 프로그램이 실행 중입니다.")
            sys.exit(0)
    except Exception:
        pass

# ==========================================
# 2. S-Structure 모델 및 단일 파일 경로 설정
# ==========================================
STATION_TITLE = "Aluko_S-Structure_Visual_Inspector"
MODEL_NAME = "S-Structure"
TARGET_BARCODE_PREFIX = "MPL03079AB"
DEFAULT_PASSWORD = "123456"
FOOT_SWITCH_COOLDOWN = 0.8  # 풋스위치 연속 오입력 방지 (초)

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
DATA_DIR = os.path.join(BASE_DIR, "data")
SINGLE_EXCEL_PATH = os.path.join(DATA_DIR, "S_Structure_Inspection_History.xlsx")
COUNT_FILE = os.path.join(BASE_DIR, "counts_s_structure.json")
CFG_FILE = os.path.join(BASE_DIR, "config_s_structure.json")
CAPTURE_BASE_DIR = os.path.join(BASE_DIR, "captures_s_structure")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(CAPTURE_BASE_DIR, exist_ok=True)

FILE_ATTRIBUTE_NORMAL = 0x80
FILE_ATTRIBUTE_HIDDEN = 0x02

def unhide_file(filepath):
    try:
        if os.name == 'nt' and os.path.exists(filepath):
            ctypes.windll.kernel32.SetFileAttributesW(str(filepath), FILE_ATTRIBUTE_NORMAL)
    except Exception:
        pass

def hide_file(filepath):
    try:
        if os.name == 'nt' and os.path.exists(filepath):
            ctypes.windll.kernel32.SetFileAttributesW(str(filepath), FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass

def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|;]', '_', name).strip()

DEFAULT_APP_CFG = {
    "foot_switch_keys": ["space", "F2", "F4", "KP_Add"],
    "allow_empty_enter_as_footswitch": True,
    "save_mode": "merged_single",  # "merged_single", "separate", "both"
}

def load_app_cfg():
    cfg = dict(DEFAULT_APP_CFG)
    if os.path.exists(CFG_FILE):
        try:
            with open(CFG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg

def save_app_cfg(cfg):
    try:
        with open(CFG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ==========================================
# 3. 다국어 패키지 딕셔너리
# ==========================================
LANG_PACK = {
    "한국어": {
        "title": STATION_TITLE,
        "pw_setting": "⚙ 비밀번호 설정",
        "tab_scan": "  외관 검사 (Inspection)  ",
        "tab_recode": "  이력 조회 (Re-code)  ",
        "model_label": "모델:",
        "code_label": "인식코드:",
        "last_scan": "현재/마지막 바코드",
        "input_guide": "바코드 스캔 입력 (스캔 후 풋스위치로 1차·2차 촬영)",
        "reset_btn": "RESET (카운터 초기화)",
        "manager_btn": "MANAGER MODE",
        "manager_btn_on": "MANAGER [DMC 재검사 대기]",
        "manager_btn_rework": "MANAGER [RE-WORK 진행 중]",
        "record_header": "S-Structure 외관 검사 기록 (최근 7일 표시 / 전체 중복 검사 적용)",
        "th_no": "No",
        "th_day": "DAY",
        "th_time": "TIME",
        "th_dmc": "Barcode (DMC)",
        "th_shot1": "1차 촬영",
        "th_shot2": "2차 촬영",
        "th_judgment": "JUDGMENT",
        "th_photo": "사진 파일",
        "th_content": "Content",
        "filter_day": "조회 일자:",
        "filter_time": "시간대:",
        "search_btn": "🔍 검색",
        "save_btn": "💾 Save (엑셀 추출)",
        "dup_scan_tag": "[중복 스캔]",
        "sorting_title": "⚠️ Sorting 필요 제품 경고",
        "sorting_msg": "[알림: Sorting 필요 제품]\n\nBarcode: {code}\n\n해당 제품은 Sorting 대상 리스트에 등록되어 있습니다.\n제품을 별도로 격리한 뒤 [Enter] 키를 누르세요.",
        "ng_model_title": "⚠️ NG - 코드 불일치",
        "ng_model_msg": "[NG: S-Structure 바코드가 아닙니다]\n\n기준 인식코드: {target}\n스캔된 코드: {code}\n\n관리자 비밀번호 6자리를 입력하여 해제하세요.",
        "ng_dup_title": "🚫 QR NG - 중복 바코드 감지",
        "ng_dup_msg": "[QR NG 발생: 이미 검사 완료된 바코드입니다]\n\n스캔 바코드: {code}\n중복 스캔 이력이 엑셀에 NG로 기록되었습니다.\n\n관리자 비밀번호 6자리를 입력하여 해제하세요.",
        "ng_incomplete_title": "⚠️ NG - 이전 제품 촬영 미완료",
        "ng_incomplete_msg": "[NG: 촬영이 완료되지 않은 상태에서 새 바코드가 스캔되었습니다]\n\n현재 진행 바코드: {curr_code}\n현재 단계: {step_desc}\n\n1차 및 2차 촬영을 먼저 완료하거나 [✕ 취소] 버튼을 누르세요.\n관리자 비밀번호 6자리를 입력하여 해제하세요.",
        "ng_mgr_err_title": "⚠️ NG - 관리자 모드 오류",
        "ng_mgr_err_msg": "[NG: 신규 바코드는 일반 모드에서 스캔해야 합니다]\n\n스캔 바코드: {code}\n재검사 모드에서는 이미 등록된 바코드만 재스캔 가능합니다.\n\n관리자 비밀번호 6자리를 입력하여 해제하세요.",
        "unlock_btn": "확인 및 잠금 해제",
        "confirm_btn": "확인 (Enter)",
        "pw_err": "비밀번호가 올바르지 않습니다.",
        "cam_ng_title": "⚠️ CAM NG - 웹캠 연결 오류",
        "cam_ng_msg": "[CAM NG 발생]\n\n외장 USB 웹캠 연결이 끊어졌거나 영상이 수신되지 않습니다.\nUSB 케이블을 다시 연결하세요."
    },
    "English": {
        "title": STATION_TITLE,
        "pw_setting": "⚙ Password Setting",
        "tab_scan": "  Inspection  ",
        "tab_recode": "  Re-code  ",
        "model_label": "Model:",
        "code_label": "Code:",
        "last_scan": "Current/Last Barcode",
        "input_guide": "Scan Barcode -> Press Foot Switch (1st & 2nd Shot)",
        "reset_btn": "RESET (Clear Counter)",
        "manager_btn": "MANAGER MODE",
        "manager_btn_on": "MANAGER [Waiting Rescan]",
        "manager_btn_rework": "MANAGER [RE-WORK Active]",
        "record_header": "S-Structure Inspection Records (Last 7 Days)",
        "th_no": "No",
        "th_day": "DAY",
        "th_time": "TIME",
        "th_dmc": "Barcode (DMC)",
        "th_shot1": "1st Shot",
        "th_shot2": "2nd Shot",
        "th_judgment": "JUDGMENT",
        "th_photo": "Photo File",
        "th_content": "Content",
        "filter_day": "Search Date:",
        "filter_time": "Time Range:",
        "search_btn": "🔍 Search",
        "save_btn": "💾 Save (Excel)",
        "dup_scan_tag": "[Duplicate Scan]",
        "sorting_title": "⚠️ Sorting Required Warning",
        "sorting_msg": "[Notice: Sorting Item]\n\nBarcode: {code}\n\nThis item is on the Sorting target list.\nIsolate the product and press [Enter].",
        "ng_model_title": "⚠️ NG - Code Mismatch",
        "ng_model_msg": "[NG: Barcode does not match S-Structure]\n\nTarget Code: {target}\nScanned Code: {code}\n\nEnter 6-digit admin password to unlock.",
        "ng_dup_title": "🚫 QR NG - Duplicate Barcode Detected",
        "ng_dup_msg": "[QR NG: Already inspected barcode]\n\nScanned Barcode: {code}\nRecorded as NG.\n\nEnter 6-digit admin password to unlock.",
        "ng_incomplete_title": "⚠️ NG - Previous Capture Incomplete",
        "ng_incomplete_msg": "[NG: New barcode scanned before completing 2 shots]\n\nCurrent Barcode: {curr_code}\nStage: {step_desc}\n\nEnter 6-digit admin password to unlock.",
        "ng_mgr_err_title": "⚠️ NG - Manager Mode Error",
        "ng_mgr_err_msg": "[NG: New barcodes must be scanned in normal mode]\n\nScanned: {code}\n\nEnter 6-digit admin password to unlock.",
        "unlock_btn": "Confirm & Unlock",
        "confirm_btn": "Confirm (Enter)",
        "pw_err": "Incorrect password.",
        "cam_ng_title": "⚠️ CAM NG - Webcam Connection Error",
        "cam_ng_msg": "[CAM NG Occurred]\n\nExternal USB webcam is disconnected.\nPlease reconnect the USB cable."
    },
    "Polski": {
        "title": STATION_TITLE,
        "pw_setting": "⚙ Ustawienia hasła",
        "tab_scan": "  Inspekcja  ",
        "tab_recode": "  Re-code  ",
        "model_label": "Model:",
        "code_label": "Kod:",
        "last_scan": "Obecny/Ostatni skan",
        "input_guide": "Zeskanuj kod -> Użyj przełącznika nożnego (1 i 2 zdjęcie)",
        "reset_btn": "RESET (Wyczyść licznik)",
        "manager_btn": "MANAGER MODE",
        "manager_btn_on": "MANAGER [Oczekiwanie]",
        "manager_btn_rework": "MANAGER [RE-WORK Aktywny]",
        "record_header": "Zapisy S-Structure (Ostatnie 7 dni)",
        "th_no": "No",
        "th_day": "DAY",
        "th_time": "TIME",
        "th_dmc": "Barcode (DMC)",
        "th_shot1": "Zdjęcie 1",
        "th_shot2": "Zdjęcie 2",
        "th_judgment": "JUDGMENT",
        "th_photo": "Plik zdjęcia",
        "th_content": "Content",
        "filter_day": "Data:",
        "filter_time": "Zakres czasu:",
        "search_btn": "🔍 Szukaj",
        "save_btn": "💾 Zapisz (Excel)",
        "dup_scan_tag": "[Duplikat skanu]",
        "sorting_title": "⚠️ Ostrzeżenie - Wymagane sortowanie",
        "sorting_msg": "[Uwaga: Przedmiot do sortowania]\n\nBarcode: {code}\n\nOdizoluj produkt i naciśnij [Enter].",
        "ng_model_title": "⚠️ NG - Niezgodność kodu",
        "ng_model_msg": "[NG: Kod kreskowy nie pasuje do S-Structure]\n\nKod docelowy: {target}\nZeskanowany Kod: {code}\n\nWprowadź 6-cyfrowe hasło administratora.",
        "ng_dup_title": "🚫 QR NG - Wykryto zduplikowany kod",
        "ng_dup_msg": "[QR NG: Ten kod został już zeskanowany]\n\nZeskanowany kod: {code}\n\nWprowadź hasło administratora.",
        "ng_incomplete_title": "⚠️ NG - Nieukończone zdjęcia",
        "ng_incomplete_msg": "[NG: Zeskanowano nowy kod przed wykonaniem 2 zdjęć]\n\nObecny kod: {curr_code}\nEtap: {step_desc}\n\nWprowadź hasło administratora.",
        "ng_mgr_err_title": "⚠️ NG - Błąd trybu menedżera",
        "ng_mgr_err_msg": "[NG: Nowe kody muszą być skanowane w trybie ogólnym]\n\nZeskanowany kod: {code}\n\nWprowadź hasło administratora.",
        "unlock_btn": "Potwierdź i odblokuj",
        "confirm_btn": "Potwierdź (Enter)",
        "pw_err": "Nieprawidłowe hasło.",
        "cam_ng_title": "⚠️ CAM NG - Błąd kamery USB",
        "cam_ng_msg": "[Błąd CAM NG]\n\nZewnętrzna kamera USB została odłączona.\nPodłącz ponownie kabel USB."
    }
}

BG_MAIN = "#1a1f26"
BG_PANEL = "#222731"
BG_INPUT = "#15181e"
TEXT_COLOR = "#e1e4ea"
TEXT_MUTED = "#8b949e"
ACCENT_YELLOW = "#f59f00"
ACCENT_BLUE = "#38bdf8"


class SStructureInspectorApp:
    STATE_WAIT_SCAN = 0   # 1단계: 바코드 스캔 대기
    STATE_WAIT_SHOT1 = 1  # 2단계: 1차 촬영 대기 (풋스위치)
    STATE_WAIT_SHOT2 = 2  # 3단계: 제품 뒤집기 후 2차 촬영 대기 (풋스위치)

    def __init__(self, root):
        self.root = root
        self.root.title(STATION_TITLE)
        self.root.geometry("1460x900")
        self.root.minsize(1280, 800)
        self.root.configure(bg=BG_MAIN)

        self.current_lang = tk.StringVar(value="한국어")
        self.admin_password = DEFAULT_PASSWORD
        self.app_cfg = load_app_cfg()

        self.is_manager_mode = False
        self.is_rework_mode = False

        self.inspection_state = self.STATE_WAIT_SCAN
        self.active_barcode = ""
        self.shot1_frame = None
        self.shot1_time_str = ""
        self.shot2_frame = None
        self.shot2_time_str = ""
        self.last_footswitch_time = 0.0

        self.cam_photo = None
        self.shot1_photo = None
        self.shot2_photo = None

        self.cap = None
        self.current_webcam_frame = None
        self.is_camera_ready = False
        self.cam_thread_running = True
        self.cam_index_used = -1
        self.cam_ng_popup_active = False
        self.cam_ng_dialog = None

        self.active_popup = None
        self.alarm_thread_running = False
        self.is_learning_footswitch = False

        self.last_scanned_code = ""
        self.last_scanned_time = 0.0
        self.auto_submit_timer = None
        self.global_scan_buffer = []

        self.counts = self.load_counts()
        self.scanned_history = set()
        self.sorting_list = set()
        self.total_excel_seq = 0
        self.file_lock = threading.Lock()

        self.setup_custom_styles()
        self.setup_ui()
        self.setup_global_key_listener()

        self.update_stat_cards()
        self.update_step_indicator()
        self.load_history_from_single_excel()

        self.start_usb_webcam_stream()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def t(self, key, **kwargs):
        pack = LANG_PACK.get(self.current_lang.get(), LANG_PACK["한국어"])
        text = pack.get(key, "")
        if kwargs:
            return text.format(**kwargs)
        return text

    # ==========================================
    # 팝업 중앙 배치 및 포커스 헬퍼 (누락 수정 완료)
    # ==========================================
    def center_popup(self, dialog, width, height):
        self.root.update_idletasks()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()

        x = rx + (rw - width) // 2
        y = ry + (rh - height) // 2
        dialog.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")

    def clear_scan_buffers(self):
        if self.auto_submit_timer:
            try:
                self.root.after_cancel(self.auto_submit_timer)
            except Exception:
                pass
            self.auto_submit_timer = None
        self.global_scan_buffer.clear()
        if hasattr(self, "scan_entry"):
            self.scan_entry.delete(0, tk.END)

    def close_active_popup(self, dialog=None):
        target = dialog or self.active_popup
        if target:
            try:
                target.grab_release()
            except Exception:
                pass
            try:
                target.destroy()
            except Exception:
                pass
        self.active_popup = None
        self.clear_scan_buffers()

    # ==========================================
    # 사운드 알람 제어
    # ==========================================
    def play_ok_beep(self, freq=1000, dur=120):
        def _beep():
            if winsound:
                try:
                    winsound.Beep(freq, dur)
                except Exception:
                    pass
        threading.Thread(target=_beep, daemon=True).start()

    def play_triple_beep(self):
        def _beep():
            if winsound:
                for _ in range(3):
                    try:
                        winsound.Beep(1200, 250)
                        time.sleep(0.08)
                    except Exception:
                        pass
        threading.Thread(target=_beep, daemon=True).start()

    def start_looping_ng_alarm(self):
        self.alarm_thread_running = True
        def _loop():
            while self.alarm_thread_running:
                if winsound:
                    for _ in range(3):
                        if not self.alarm_thread_running:
                            break
                        try:
                            winsound.Beep(1200, 250)
                            time.sleep(0.08)
                        except Exception:
                            pass
                for _ in range(15):
                    if not self.alarm_thread_running:
                        break
                    time.sleep(0.1)
        threading.Thread(target=_loop, daemon=True).start()

    def stop_looping_ng_alarm(self):
        self.alarm_thread_running = False

    # ==========================================
    # 웹캠 스트리밍 및 좌측 하단 캡처 표시
    # ==========================================
    def find_and_open_camera(self):
        search_order = [1, 2, 0, 3]
        for idx in search_order:
            for api_backend in [cv2.CAP_DSHOW, cv2.CAP_ANY]:
                try:
                    cap = cv2.VideoCapture(idx, api_backend)
                    if cap.isOpened():
                        ret, test_frame = cap.read()
                        if ret and test_frame is not None and test_frame.size > 0:
                            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
                            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
                            self.cam_index_used = idx
                            return cap
                        cap.release()
                except Exception:
                    pass
        return None

    def start_usb_webcam_stream(self):
        def _webcam_worker():
            consecutive_fail_count = 0
            while self.cam_thread_running:
                if CV_AVAILABLE:
                    if self.cap is None or not self.cap.isOpened():
                        self.cap = self.find_and_open_camera()
                        if self.cap and self.cap.isOpened():
                            self.is_camera_ready = True
                            consecutive_fail_count = 0
                            cam_type_str = f"USB 웹캠 (CAM {self.cam_index_used})"
                            self.root.after(0, lambda: self.lbl_cam_status.config(
                                text=f"● LIVE ({cam_type_str})", fg="#22c55e"
                            ))
                            self.root.after(0, self.close_cam_ng_popup)
                        else:
                            self.is_camera_ready = False
                            consecutive_fail_count += 1
                            self.root.after(0, lambda: self.lbl_cam_status.config(
                                text="✕ 웹캠 연결 끊김 (재연결 중...)", fg="#f87171"
                            ))
                            if consecutive_fail_count >= 3:
                                self.root.after(0, self.open_cam_ng_popup)
                            time.sleep(1.0)
                            continue

                    ret, frame = self.cap.read()
                    if ret and frame is not None and frame.size > 0:
                        self.current_webcam_frame = frame
                        consecutive_fail_count = 0
                        if self.cam_ng_popup_active:
                            self.root.after(0, self.close_cam_ng_popup)
                    else:
                        consecutive_fail_count += 1
                        if self.cap:
                            try:
                                self.cap.release()
                            except Exception:
                                pass
                        self.cap = None
                        self.is_camera_ready = False
                        gc.collect()
                        if consecutive_fail_count >= 3:
                            self.root.after(0, self.open_cam_ng_popup)
                        time.sleep(0.5)
                else:
                    time.sleep(0.5)

        threading.Thread(target=_webcam_worker, daemon=True).start()
        self.root.after(200, self.update_camera_canvas)

    def update_camera_canvas(self):
        if self.current_webcam_frame is not None and self.is_camera_ready:
            frame_to_show = self.current_webcam_frame.copy()
            ch, cw = frame_to_show.shape[:2]

            target_w, target_h = 410, 155
            scale = max(target_w / cw, target_h / ch)
            scaled_w = int(cw * scale)
            scaled_h = int(ch * scale)

            resized = cv2.resize(frame_to_show, (scaled_w, scaled_h))
            start_x = max(0, (scaled_w - target_w) // 2)
            start_y = max(0, (scaled_h - target_h) // 2)
            final_crop = resized[start_y:start_y + target_h, start_x:start_x + target_w]

            rgb = cv2.cvtColor(final_crop, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(rgb)
            self.cam_photo = ImageTk.PhotoImage(pil_img)

            self.cam_canvas.delete("all")
            self.cam_canvas.create_image(0, 0, anchor="nw", image=self.cam_photo)
        else:
            self.cam_canvas.delete("all")
            self.cam_canvas.create_text(205, 77, text="[ CAM DISCONNECTED ]", fill="#ef4444", font=("Consolas", 11, "bold"))

        self.root.after(33, self.update_camera_canvas)

    def render_captured_thumbnail(self, canvas_widget, frame_bgr, slot_num, timestamp_text=""):
        canvas_widget.delete("all")
        target_w, target_h = 198, 145

        if frame_bgr is None or not CV_AVAILABLE:
            placeholder = "1차 촬영 대기 중\n(전면)" if slot_num == 1 else "2차 촬영 대기 중\n(뒤집기 후 후면)"
            canvas_widget.config(highlightbackground="#334155", highlightthickness=1)
            canvas_widget.create_text(
                target_w // 2, target_h // 2,
                text=placeholder, fill="#64748b", font=("맑은 고딕", 9, "bold"), justify=tk.CENTER
            )
            return

        ch, cw = frame_bgr.shape[:2]
        scale = max(target_w / cw, target_h / ch)
        scaled_w, scaled_h = int(cw * scale), int(ch * scale)
        resized = cv2.resize(frame_bgr, (scaled_w, scaled_h))
        sx = max(0, (scaled_w - target_w) // 2)
        sy = max(0, (scaled_h - target_h) // 2)
        crop = resized[sy:sy + target_h, sx:sx + target_w].copy()

        cv2.rectangle(crop, (0, 0), (target_w, 24), (20, 83, 45), -1)
        label_str = f"SHOT {slot_num} OK ({timestamp_text})"
        cv2.putText(crop, label_str, (6, 17), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 255, 255), 1, cv2.LINE_AA)

        rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
        tk_img = ImageTk.PhotoImage(Image.fromarray(rgb))
        if slot_num == 1:
            self.shot1_photo = tk_img
        else:
            self.shot2_photo = tk_img

        canvas_widget.config(highlightbackground="#22c55e", highlightthickness=2)
        canvas_widget.create_image(0, 0, anchor="nw", image=tk_img)

    def open_photo_zoom_popup(self, slot_num):
        frame = self.shot1_frame if slot_num == 1 else self.shot2_frame
        if frame is None or not CV_AVAILABLE:
            return
        win = tk.Toplevel(self.root)
        win.title(f"{self.active_barcode or self.last_scanned_code} - {slot_num}차 촬영 원본 확인")
        win.configure(bg=BG_MAIN)
        win.transient(self.root)
        self.center_popup(win, 840, 540)

        ch, cw = frame.shape[:2]
        scale = min(800 / cw, 480 / ch)
        disp = cv2.resize(frame, (int(cw * scale), int(ch * scale)))
        rgb = cv2.cvtColor(disp, cv2.COLOR_BGR2RGB)
        photo = ImageTk.PhotoImage(Image.fromarray(rgb))

        lbl = tk.Label(win, image=photo, bg=BG_MAIN)
        lbl.image = photo
        lbl.pack(expand=True, pady=10)

    # ==========================================
    # 일자별 폴더 사진 저장
    # ==========================================
    def save_inspection_photos(self, barcode_str, frame1, frame2, t1_str, t2_str):
        if not CV_AVAILABLE or frame1 is None or frame2 is None:
            return "-"

        now = datetime.now()
        yy_mm = now.strftime("%y.%m")
        today_str = now.strftime("%Y-%m-%d")

        base_day_dir = os.path.join(CAPTURE_BASE_DIR, f"{yy_mm} captures", today_str)
        os.makedirs(base_day_dir, exist_ok=True)

        safe_code = sanitize_filename(barcode_str)
        if not safe_code:
            safe_code = f"MPL03079AB_{now.strftime('%H%M%S')}"

        merged_filename = f"{safe_code}.jpg"
        merged_path = os.path.join(base_day_dir, merged_filename)
        if os.path.exists(merged_path):
            suffix = now.strftime("%H%M%S")
            merged_filename = f"{safe_code}_{suffix}.jpg"
            merged_path = os.path.join(base_day_dir, merged_filename)
            sep1_path = os.path.join(base_day_dir, f"{safe_code}_{suffix}_1차.jpg")
            sep2_path = os.path.join(base_day_dir, f"{safe_code}_{suffix}_2차.jpg")
        else:
            sep1_path = os.path.join(base_day_dir, f"{safe_code}_1차.jpg")
            sep2_path = os.path.join(base_day_dir, f"{safe_code}_2차.jpg")

        save_mode = self.app_cfg.get("save_mode", "merged_single")

        def _write_jpg(path, img):
            ok, buf = cv2.imencode(".jpg", img, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
            if ok:
                with open(path, "wb") as f:
                    f.write(buf)

        def _save_worker(f1, f2):
            try:
                h1, w1 = f1.shape[:2]
                h2, w2 = f2.shape[:2]
                if (h1, w1) != (h2, w2):
                    f2 = cv2.resize(f2, (w1, h1))

                f1_annotated = f1.copy()
                f2_annotated = f2.copy()
                cv2.rectangle(f1_annotated, (0, 0), (w1, 40), (0, 0, 0), -1)
                cv2.putText(f1_annotated, f"[1st SHOT] {barcode_str} ({today_str} {t1_str})",
                            (12, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2, cv2.LINE_AA)

                cv2.rectangle(f2_annotated, (0, 0), (w1, 40), (0, 0, 0), -1)
                cv2.putText(f2_annotated, f"[2nd SHOT - FLIPPED] {barcode_str} ({today_str} {t2_str})",
                            (12, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2, cv2.LINE_AA)

                if save_mode in ("merged_single", "both"):
                    combined = np.hstack((f1_annotated, f2_annotated))
                    _write_jpg(merged_path, combined)

                if save_mode in ("separate", "both"):
                    _write_jpg(sep1_path, f1_annotated)
                    _write_jpg(sep2_path, f2_annotated)
            except Exception:
                pass

        threading.Thread(target=_save_worker, args=(frame1.copy(), frame2.copy()), daemon=True).start()

        if save_mode == "separate":
            return f"{os.path.basename(sep1_path)}, {os.path.basename(sep2_path)}"
        return merged_filename

    # ==========================================
    # 웹캠 연결 끊김 팝업
    # ==========================================
    def open_cam_ng_popup(self):
        if self.cam_ng_popup_active or self.active_popup:
            return
        self.cam_ng_popup_active = True
        self.play_triple_beep()

        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("cam_ng_title"))
        dialog.resizable(False, False)
        dialog.configure(bg="#3a1c1f")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_popup(dialog, 460, 220)
        self.cam_ng_dialog = dialog

        tk.Label(dialog, text="📷 CAM NG - 외장 웹캠 신호 없음", font=("맑은 고딕", 12, "bold"), bg="#3a1c1f", fg="#ff6b6b").pack(pady=(20, 10))
        tk.Label(dialog, text=self.t("cam_ng_msg"), font=("맑은 고딕", 10), bg="#3a1c1f", fg="#e1e4ea", justify=tk.CENTER).pack(pady=10)

        tk.Button(dialog, text="닫기 (재연결 대기)", command=self.close_cam_ng_popup, font=("맑은 고딕", 10, "bold"),
                  bg="#dc3545", fg="#ffffff", relief="flat", padx=15, pady=4).pack(pady=10)

    def close_cam_ng_popup(self):
        if self.cam_ng_dialog:
            try:
                self.cam_ng_dialog.grab_release()
                self.cam_ng_dialog.destroy()
            except Exception:
                pass
            self.cam_ng_dialog = None
        self.cam_ng_popup_active = False

    def on_closing(self):
        self.stop_looping_ng_alarm()
        self.cam_thread_running = False
        if self.cap and self.cap.isOpened():
            self.cap.release()
        self.root.destroy()

    # ==========================================
    # 카운터 저장/로드
    # ==========================================
    def load_counts(self):
        default_c = {"total": 0, "ok": 0, "ng": 0}
        if os.path.exists(COUNT_FILE):
            try:
                with open(COUNT_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    default_c.update(data)
            except Exception:
                pass
        return default_c

    def save_counts(self):
        try:
            with open(COUNT_FILE, "w", encoding="utf-8") as f:
                json.dump(self.counts, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def update_stat_cards(self):
        self.lbl_total_val.config(text=str(self.counts["total"]))
        self.lbl_ok_val.config(text=str(self.counts["ok"]))
        self.lbl_ng_val.config(text=str(self.counts["ng"]))

    # ==========================================
    # UI 구성
    # ==========================================
    def setup_custom_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Dark.TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("Dark.TNotebook.Tab", background="#2a2f3a", foreground=TEXT_MUTED,
                        font=("맑은 고딕", 10, "bold"), padding=[15, 5])
        style.map("Dark.TNotebook.Tab",
                  background=[("selected", BG_PANEL)],
                  foreground=[("selected", "#ffffff")])

        style.configure("Dark.TCombobox",
                        fieldbackground=BG_INPUT,
                        background="#303642",
                        foreground="#ffffff",
                        arrowcolor="#ffffff",
                        darkcolor=BG_INPUT,
                        lightcolor=BG_INPUT)

        style.configure("Dark.Treeview",
                        background="#1e232d",
                        foreground=TEXT_COLOR,
                        fieldbackground="#1e232d",
                        rowheight=26,
                        font=("맑은 고딕", 9))
        style.configure("Dark.Treeview.Heading",
                        background="#282e3a",
                        foreground="#adb5bd",
                        font=("맑은 고딕", 9, "bold"),
                        relief="flat")
        style.map("Dark.Treeview",
                  background=[("selected", "#2b5278")],
                  foreground=[("selected", "#ffffff")])

    def setup_ui(self):
        header_frame = tk.Frame(self.root, bg=BG_MAIN, height=45)
        header_frame.pack(fill=tk.X, padx=20, pady=(10, 4))

        tk.Label(header_frame, text=STATION_TITLE, font=("Arial", 13, "bold"),
                 fg=TEXT_COLOR, bg=BG_MAIN).pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(header_frame, text=f"ITEM: {MODEL_NAME} ({TARGET_BARCODE_PREFIX})",
                 font=("맑은 고딕", 10, "bold"), fg=ACCENT_YELLOW, bg="#282e3a", padx=10, pady=3).pack(side=tk.LEFT)

        self.lang_combo = ttk.Combobox(
            header_frame,
            textvariable=self.current_lang,
            values=["한국어", "English", "Polski"],
            state="readonly",
            font=("맑은 고딕", 9, "bold"),
            width=9,
            style="Dark.TCombobox"
        )
        self.lang_combo.pack(side=tk.RIGHT, padx=(10, 0))
        self.lang_combo.bind("<<ComboboxSelected>>", self.on_language_changed)

        tk.Label(header_frame, text="🌐", font=("맑은 고딕", 12), bg=BG_MAIN, fg=TEXT_COLOR).pack(side=tk.RIGHT)

        self.btn_pw = tk.Button(
            header_frame, text=self.t("pw_setting"), command=self.change_password_dialog,
            bg="#2c323d", fg=TEXT_COLOR, activebackground="#3a4250", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 9), padx=10, pady=3, cursor="hand2"
        )
        self.btn_pw.pack(side=tk.RIGHT, padx=(0, 12))

        self.btn_footswitch_cfg = tk.Button(
            header_frame, text=f"🦶 풋스위치 키 설정 ({'/'.join(self.app_cfg['foot_switch_keys'][:2])})",
            command=self.open_footswitch_learn_dialog,
            bg="#1e3a5f", fg="#93c5fd", activebackground="#2b5278", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 9, "bold"), padx=10, pady=3, cursor="hand2"
        )
        self.btn_footswitch_cfg.pack(side=tk.RIGHT, padx=(0, 10))

        self.btn_hidden_delete = tk.Button(
            header_frame, text="", command=self.trigger_hidden_delete_action,
            bg=BG_MAIN, activebackground=BG_MAIN, bd=0, relief="flat", cursor="arrow", width=3
        )
        self.btn_hidden_delete.pack(side=tk.RIGHT, padx=(0, 2))

        self.notebook = ttk.Notebook(self.root, style="Dark.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

        self.tab_scan = tk.Frame(self.notebook, bg=BG_MAIN)
        self.notebook.add(self.tab_scan, text=self.t("tab_scan"))

        self.tab_recode = tk.Frame(self.notebook, bg=BG_MAIN)
        self.notebook.add(self.tab_recode, text=self.t("tab_recode"))

        self.build_scan_tab()
        self.build_recode_tab()

    def build_scan_tab(self):
        main_frame = tk.Frame(self.tab_scan, bg=BG_MAIN)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=8)

        left_panel = tk.Frame(main_frame, bg=BG_PANEL, width=450)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_panel.pack_propagate(False)

        top_header = tk.Frame(left_panel, bg=BG_PANEL)
        top_header.pack(fill=tk.X, padx=18, pady=(8, 4))

        self.lbl_model_info = tk.Label(
            top_header,
            text=f"{self.t('model_label')} {MODEL_NAME}   |   {self.t('code_label')} {TARGET_BARCODE_PREFIX}",
            font=("맑은 고딕", 10, "bold"), fg=ACCENT_YELLOW, bg=BG_PANEL, anchor="w"
        )
        self.lbl_model_info.pack(fill=tk.X)

        step_frame = tk.Frame(left_panel, bg="#181c24", pady=4, padx=4)
        step_frame.pack(fill=tk.X, padx=18, pady=(2, 4))
        step_frame.columnconfigure((0, 1, 2), weight=1)

        self.lbl_step1 = tk.Label(step_frame, text="① 바코드 스캔", font=("맑은 고딕", 8, "bold"), bg="#2563eb", fg="#ffffff", pady=3)
        self.lbl_step1.grid(row=0, column=0, padx=2, sticky="ew")

        self.lbl_step2 = tk.Label(step_frame, text="② 1차 촬영(전면)", font=("맑은 고딕", 8, "bold"), bg="#222731", fg=TEXT_MUTED, pady=3)
        self.lbl_step2.grid(row=0, column=1, padx=2, sticky="ew")

        self.lbl_step3 = tk.Label(step_frame, text="③ 뒤집어 2차 촬영", font=("맑은 고딕", 8, "bold"), bg="#222731", fg=TEXT_MUTED, pady=3)
        self.lbl_step3.grid(row=0, column=2, padx=2, sticky="ew")

        self.status_box = tk.Label(
            left_panel, text="READY", font=("Arial", 28, "bold"),
            fg="#adb5bd", bg="#2a2e37", height=2, relief="flat"
        )
        self.status_box.pack(fill=tk.X, padx=18, pady=3)

        self.lbl_last_scan = tk.Label(
            left_panel, text=f"{self.t('last_scan')}: -", font=("맑은 고딕", 9, "bold"),
            fg=ACCENT_BLUE, bg=BG_PANEL, anchor="w"
        )
        self.lbl_last_scan.pack(fill=tk.X, padx=18, pady=(2, 1))

        self.lbl_input_guide = tk.Label(
            left_panel, text=self.t("input_guide"), font=("맑은 고딕", 8),
            fg=TEXT_MUTED, bg=BG_PANEL, anchor="w"
        )
        self.lbl_input_guide.pack(fill=tk.X, padx=18)

        entry_row = tk.Frame(left_panel, bg=BG_PANEL)
        entry_row.pack(fill=tk.X, padx=18, pady=(3, 6))

        self.scan_entry = tk.Entry(
            entry_row, font=("Consolas", 11), bg=BG_INPUT, fg="#ffffff",
            insertbackground="#ffffff", relief="flat", highlightthickness=1,
            highlightbackground="#343c4c", highlightcolor="#3b82f6"
        )
        self.scan_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        self.scan_entry.bind("<Return>", self.on_scan_entry_return)
        self.scan_entry.bind("<KeyRelease>", self.on_entry_key_release)

        self.btn_manual_shot = tk.Button(
            entry_row, text="📸 촬영(F2)", command=self.trigger_footswitch_capture,
            bg="#0284c7", fg="#ffffff", activebackground="#0369a1", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 8, "bold"), padx=8, pady=4, cursor="hand2"
        )
        self.btn_manual_shot.pack(side=tk.LEFT, padx=(5, 2))

        self.btn_cancel_cycle = tk.Button(
            entry_row, text="✕ 취소", command=self.cancel_current_inspection,
            bg="#475569", fg="#ffffff", activebackground="#334155", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 8, "bold"), padx=6, pady=4, cursor="hand2"
        )
        self.btn_cancel_cycle.pack(side=tk.LEFT, padx=(2, 0))

        stats_frame = tk.Frame(left_panel, bg=BG_PANEL)
        stats_frame.pack(fill=tk.X, padx=18, pady=2)
        stats_frame.columnconfigure((0, 1, 2), weight=1)

        card_total = tk.Frame(stats_frame, bg="#1a1e26", pady=3)
        card_total.grid(row=0, column=0, padx=2, sticky="nsew")
        self.lbl_total_val = tk.Label(card_total, text="0", font=("Arial", 13, "bold"), fg=TEXT_COLOR, bg="#1a1e26")
        self.lbl_total_val.pack()
        tk.Label(card_total, text="TOTAL", font=("Arial", 7, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        card_ok = tk.Frame(stats_frame, bg="#1a1e26", pady=3)
        card_ok.grid(row=0, column=1, padx=2, sticky="nsew")
        self.lbl_ok_val = tk.Label(card_ok, text="0", font=("Arial", 13, "bold"), fg="#28a745", bg="#1a1e26")
        self.lbl_ok_val.pack()
        tk.Label(card_ok, text="OK", font=("Arial", 7, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        card_ng = tk.Frame(stats_frame, bg="#1a1e26", pady=3)
        card_ng.grid(row=0, column=2, padx=2, sticky="nsew")
        self.lbl_ng_val = tk.Label(card_ng, text="0", font=("Arial", 13, "bold"), fg="#dc3545", bg="#1a1e26")
        self.lbl_ng_val.pack()
        tk.Label(card_ng, text="NG", font=("Arial", 7, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        btn_row = tk.Frame(left_panel, bg=BG_PANEL)
        btn_row.pack(fill=tk.X, padx=18, pady=(4, 4))
        self.btn_reset = tk.Button(
            btn_row, text=self.t("reset_btn"), command=self.open_reset_dialog,
            bg="#2c323d", fg="#ff8787", activebackground="#3d2729", activeforeground="#ff6b6b",
            relief="flat", font=("맑은 고딕", 8, "bold"), pady=3, cursor="hand2"
        )
        self.btn_reset.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

        self.btn_manager = tk.Button(
            btn_row, text=self.t("manager_btn"), command=self.toggle_manager_mode,
            bg="#2c323d", fg="#adb5bd", activebackground="#303642", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 8, "bold"), pady=3, cursor="hand2"
        )
        self.btn_manager.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(3, 0))

        cam_panel = tk.Frame(left_panel, bg=BG_PANEL)
        cam_panel.pack(fill=tk.X, padx=18, pady=(2, 4))

        top_info = tk.Frame(cam_panel, bg=BG_PANEL)
        top_info.pack(fill=tk.X)
        tk.Label(top_info, text="📹 실시간 웹캠 (LIVE)", font=("맑은 고딕", 8, "bold"), fg="#38bdf8", bg=BG_PANEL).pack(side=tk.LEFT)
        self.lbl_cam_status = tk.Label(top_info, text="○ 연결 확인 중...", font=("맑은 고딕", 8, "bold"), fg="#f87171", bg=BG_PANEL)
        self.lbl_cam_status.pack(side=tk.RIGHT)

        self.cam_canvas = tk.Canvas(cam_panel, width=410, height=155, bg="#0f172a", highlightthickness=1, highlightbackground="#334155")
        self.cam_canvas.pack(fill=tk.X, pady=(2, 0))

        captured_panel = tk.Frame(left_panel, bg=BG_PANEL)
        captured_panel.pack(fill=tk.BOTH, expand=True, padx=18, pady=(4, 8))

        cap_header = tk.Frame(captured_panel, bg=BG_PANEL)
        cap_header.pack(fill=tk.X, pady=(0, 3))
        tk.Label(
            cap_header, text="📸 풋스위치 촬영 사진 확인 (더블클릭 시 확대)",
            font=("맑은 고딕", 9, "bold"), fg=ACCENT_YELLOW, bg=BG_PANEL
        ).pack(side=tk.LEFT)

        slots_frame = tk.Frame(captured_panel, bg=BG_PANEL)
        slots_frame.pack(fill=tk.BOTH, expand=True)
        slots_frame.columnconfigure((0, 1), weight=1)

        s1_box = tk.Frame(slots_frame, bg=BG_PANEL)
        s1_box.grid(row=0, column=0, padx=(0, 4), sticky="nsew")
        tk.Label(s1_box, text="[1차 촬영 - 전면]", font=("맑은 고딕", 8, "bold"), fg="#93c5fd", bg=BG_PANEL).pack(anchor="w")
        self.canvas_shot1 = tk.Canvas(s1_box, width=198, height=145, bg="#0f172a", highlightthickness=1, highlightbackground="#334155", cursor="hand2")
        self.canvas_shot1.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        self.canvas_shot1.bind("<Double-Button-1>", lambda e: self.open_photo_zoom_popup(1))

        s2_box = tk.Frame(slots_frame, bg=BG_PANEL)
        s2_box.grid(row=0, column=1, padx=(4, 0), sticky="nsew")
        tk.Label(s2_box, text="[2차 촬영 - 후면(뒤집기)]", font=("맑은 고딕", 8, "bold"), fg="#fde047", bg=BG_PANEL).pack(anchor="w")
        self.canvas_shot2 = tk.Canvas(s2_box, width=198, height=145, bg="#0f172a", highlightthickness=1, highlightbackground="#334155", cursor="hand2")
        self.canvas_shot2.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        self.canvas_shot2.bind("<Double-Button-1>", lambda e: self.open_photo_zoom_popup(2))

        self.render_captured_thumbnail(self.canvas_shot1, None, 1)
        self.render_captured_thumbnail(self.canvas_shot2, None, 2)

        right_panel = tk.Frame(main_frame, bg=BG_MAIN)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.lbl_right_header = tk.Label(
            right_panel, text=self.t("record_header"), font=("맑은 고딕", 10, "bold"), fg="#9aa0a6", bg=BG_MAIN, anchor="w"
        )
        self.lbl_right_header.pack(fill=tk.X, pady=(0, 6))

        columns = ("No", "DAY", "TIME", "DMC", "Shot1", "Shot2", "JUDGMENT", "Photo", "Content")
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings", style="Dark.Treeview")
        self.tree.tag_configure("ng_row", background="#3a1c1f", foreground="#ff6b6b")

        self.tree.heading("No", text=self.t("th_no"))
        self.tree.heading("DAY", text=self.t("th_day"))
        self.tree.heading("TIME", text=self.t("th_time"))
        self.tree.heading("DMC", text=self.t("th_dmc"))
        self.tree.heading("Shot1", text=self.t("th_shot1"))
        self.tree.heading("Shot2", text=self.t("th_shot2"))
        self.tree.heading("JUDGMENT", text=self.t("th_judgment"))
        self.tree.heading("Photo", text=self.t("th_photo"))
        self.tree.heading("Content", text=self.t("th_content"))

        self.tree.column("No", width=55, anchor="center")
        self.tree.column("DAY", width=90, anchor="center")
        self.tree.column("TIME", width=80, anchor="center")
        self.tree.column("DMC", width=210, anchor="w")
        self.tree.column("Shot1", width=80, anchor="center")
        self.tree.column("Shot2", width=80, anchor="center")
        self.tree.column("JUDGMENT", width=80, anchor="center")
        self.tree.column("Photo", width=180, anchor="w")
        self.tree.column("Content", width=110, anchor="center")

        tree_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def build_recode_tab(self):
        recode_frame = tk.Frame(self.tab_recode, bg=BG_MAIN)
        recode_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        filter_bar = tk.Frame(recode_frame, bg=BG_PANEL, pady=8, padx=15)
        filter_bar.pack(fill=tk.X, pady=(0, 10))

        today_str = datetime.now().strftime("%Y-%m-%d")

        self.lbl_filter_day = tk.Label(filter_bar, text=self.t("filter_day"), font=("맑은 고딕", 9, "bold"), fg=TEXT_COLOR, bg=BG_PANEL)
        self.lbl_filter_day.pack(side=tk.LEFT)
        self.entry_start_day = tk.Entry(filter_bar, width=11, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_start_day.insert(0, today_str)
        self.entry_start_day.pack(side=tk.LEFT, padx=5)

        tk.Label(filter_bar, text="~", fg=TEXT_MUTED, bg=BG_PANEL).pack(side=tk.LEFT)
        self.entry_end_day = tk.Entry(filter_bar, width=11, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_end_day.insert(0, today_str)
        self.entry_end_day.pack(side=tk.LEFT, padx=5)

        self.lbl_filter_time = tk.Label(filter_bar, text=self.t("filter_time"), font=("맑은 고딕", 9, "bold"), fg=TEXT_COLOR, bg=BG_PANEL)
        self.lbl_filter_time.pack(side=tk.LEFT, padx=(15, 0))
        self.entry_start_time = tk.Entry(filter_bar, width=9, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_start_time.insert(0, "00:00:00")
        self.entry_start_time.pack(side=tk.LEFT, padx=5)

        tk.Label(filter_bar, text="~", fg=TEXT_MUTED, bg=BG_PANEL).pack(side=tk.LEFT)
        self.entry_end_time = tk.Entry(filter_bar, width=9, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_end_time.insert(0, "23:59:59")
        self.entry_end_time.pack(side=tk.LEFT, padx=5)

        self.btn_search = tk.Button(
            filter_bar, text=self.t("search_btn"), command=self.apply_recode_filter,
            bg="#2b5278", fg="#ffffff", relief="flat", font=("맑은 고딕", 9, "bold"), padx=12, pady=2, cursor="hand2"
        )
        self.btn_search.pack(side=tk.LEFT, padx=15)

        self.btn_save = tk.Button(
            filter_bar, text=self.t("save_btn"), command=self.save_recode_to_excel,
            bg="#198754", fg="#ffffff", relief="flat", font=("맑은 고딕", 9, "bold"), padx=12, pady=2, cursor="hand2"
        )
        self.btn_save.pack(side=tk.RIGHT)

        cols = ("No", "DAY", "TIME", "DMC", "Shot1", "Shot2", "JUDGMENT", "Photo", "Content")
        self.tree_recode = ttk.Treeview(recode_frame, columns=cols, show="headings", style="Dark.Treeview")
        self.tree_recode.tag_configure("ng_row", background="#3a1c1f", foreground="#ff6b6b")

        for c, h_key, w, anc in [
            ("No", "th_no", 55, "center"),
            ("DAY", "th_day", 90, "center"),
            ("TIME", "th_time", 80, "center"),
            ("DMC", "th_dmc", 220, "w"),
            ("Shot1", "th_shot1", 80, "center"),
            ("Shot2", "th_shot2", 80, "center"),
            ("JUDGMENT", "th_judgment", 80, "center"),
            ("Photo", "th_photo", 190, "w"),
            ("Content", "th_content", 110, "center"),
        ]:
            self.tree_recode.heading(c, text=self.t(h_key))
            self.tree_recode.column(c, width=w, anchor=anc)

        scroll_r = ttk.Scrollbar(recode_frame, orient=tk.VERTICAL, command=self.tree_recode.yview)
        self.tree_recode.configure(yscroll=scroll_r.set)

        self.tree_recode.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_r.pack(side=tk.RIGHT, fill=tk.Y)

    # ==========================================
    # 단계 표시줄 및 상태창 업데이트
    # ==========================================
    def update_step_indicator(self):
        if self.inspection_state == self.STATE_WAIT_SCAN:
            self.lbl_step1.config(bg="#2563eb", fg="#ffffff")
            self.lbl_step2.config(bg="#222731", fg=TEXT_MUTED)
            self.lbl_step3.config(bg="#222731", fg=TEXT_MUTED)
        elif self.inspection_state == self.STATE_WAIT_SHOT1:
            self.lbl_step1.config(bg="#15803d", fg="#ffffff")
            self.lbl_step2.config(bg="#d97706", fg="#ffffff")
            self.lbl_step3.config(bg="#222731", fg=TEXT_MUTED)
        elif self.inspection_state == self.STATE_WAIT_SHOT2:
            self.lbl_step1.config(bg="#15803d", fg="#ffffff")
            self.lbl_step2.config(bg="#15803d", fg="#ffffff")
            self.lbl_step3.config(bg="#d97706", fg="#ffffff")

    def set_status(self, text, fg_color, bg_color):
        if len(text) <= 4:
            font_size = 34
        elif len(text) <= 10:
            font_size = 24
        else:
            font_size = 18
        self.status_box.config(text=text, fg=fg_color, bg=bg_color, font=("맑은 고딕", font_size, "bold"))

    def cancel_current_inspection(self):
        if self.inspection_state == self.STATE_WAIT_SCAN:
            return
        self.inspection_state = self.STATE_WAIT_SCAN
        self.active_barcode = ""
        self.shot1_frame = None
        self.shot2_frame = None
        self.render_captured_thumbnail(self.canvas_shot1, None, 1)
        self.render_captured_thumbnail(self.canvas_shot2, None, 2)
        self.update_step_indicator()
        self.set_status("READY", "#adb5bd", "#2a2e37")
        self.lbl_last_scan.config(text=f"{self.t('last_scan')}: - (취소됨)")
        self.scan_entry.focus_set()

    # ==========================================
    # USB 풋스위치 및 바코드 스캐너 키 리스너 분리
    # ==========================================
    def open_footswitch_learn_dialog(self):
        if self.active_popup:
            return
        self.clear_scan_buffers()
        self.is_learning_footswitch = True

        win = tk.Toplevel(self.root)
        win.title("USB 풋스위치 키 설정")
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()
        self.active_popup = win
        self.center_popup(win, 440, 240)

        tk.Label(win, text="[ 🦶 USB 풋스위치 입력 키 감지 ]", font=("맑은 고딕", 12, "bold"), fg=ACCENT_BLUE, bg=BG_PANEL).pack(pady=(18, 8))
        tk.Label(
            win,
            text="지금 연결된 USB 풋스위치 페달을 1회 밟아주세요.\n(또는 촬영 단축키로 사용할 키보드 키를 누르세요)",
            font=("맑은 고딕", 9), fg=TEXT_COLOR, bg=BG_PANEL, justify=tk.CENTER
        ).pack(pady=6)

        lbl_cur = tk.Label(
            win, text=f"현재 등록된 키: {', '.join(self.app_cfg['foot_switch_keys'])}",
            font=("Consolas", 10, "bold"), fg=ACCENT_YELLOW, bg=BG_INPUT, padx=10, pady=6
        )
        lbl_cur.pack(pady=8)

        def close_learn():
            self.is_learning_footswitch = False
            self.close_active_popup(win)
            self.scan_entry.focus_set()

        def on_key_captured(event):
            key_name = event.keysym
            if key_name in ("Escape",):
                close_learn()
                return
            new_keys = [key_name]
            if "F2" not in new_keys:
                new_keys.append("F2")
            self.app_cfg["foot_switch_keys"] = new_keys
            save_app_cfg(self.app_cfg)
            self.btn_footswitch_cfg.config(text=f"🦶 풋스위치 키 설정 ({'/'.join(new_keys[:2])})")
            close_learn()
            messagebox.showinfo("설정 완료", f"풋스위치 작동 키가 [{key_name}] 키로 설정되었습니다.", parent=self.root)
            self.scan_entry.focus_set()

        win.bind("<KeyPress>", on_key_captured)
        win.protocol("WM_DELETE_WINDOW", close_learn)
        win.focus_force()

    def on_scan_entry_return(self, event):
        if self.active_popup or self.is_learning_footswitch:
            return
        text = self.scan_entry.get().strip()
        if text:
            self.process_scan(text.upper())
        elif self.app_cfg.get("allow_empty_enter_as_footswitch", True) and self.inspection_state in (self.STATE_WAIT_SHOT1, self.STATE_WAIT_SHOT2):
            self.trigger_footswitch_capture()

    def on_entry_key_release(self, event):
        if self.active_popup or self.is_learning_footswitch:
            return
        if event.keysym in self.app_cfg.get("foot_switch_keys", []):
            return
        if event.keysym in ("Return", "KP_Enter"):
            if self.auto_submit_timer:
                self.root.after_cancel(self.auto_submit_timer)
                self.auto_submit_timer = None
            return

        text = self.scan_entry.get().strip()
        if len(text) >= 10:
            if self.auto_submit_timer:
                self.root.after_cancel(self.auto_submit_timer)
            self.auto_submit_timer = self.root.after(150, self._trigger_auto_submit_entry)

    def _trigger_auto_submit_entry(self):
        if self.active_popup or self.is_learning_footswitch:
            return
        text = self.scan_entry.get().strip()
        if text:
            self.process_scan(text.upper())

    def setup_global_key_listener(self):
        def _on_key_press(event):
            if self.is_learning_footswitch or self.active_popup:
                return

            if event.keysym in self.app_cfg.get("foot_switch_keys", []):
                if event.keysym == "space":
                    cur_txt = self.scan_entry.get().replace(" ", "")
                    self.scan_entry.delete(0, tk.END)
                    self.scan_entry.insert(0, cur_txt)
                self.trigger_footswitch_capture()
                return "break"

            focused = self.root.focus_get()
            if isinstance(focused, tk.Entry) and focused != self.scan_entry:
                return

            if event.keysym in ("Return", "KP_Enter"):
                if self.auto_submit_timer:
                    self.root.after_cancel(self.auto_submit_timer)
                    self.auto_submit_timer = None

                if self.global_scan_buffer:
                    scanned_text = "".join(self.global_scan_buffer).strip()
                    self.global_scan_buffer.clear()
                    self.scan_entry.delete(0, tk.END)
                    if scanned_text:
                        self.process_scan(scanned_text.upper())
                elif not self.scan_entry.get().strip():
                    if self.app_cfg.get("allow_empty_enter_as_footswitch", True) and self.inspection_state in (self.STATE_WAIT_SHOT1, self.STATE_WAIT_SHOT2):
                        self.trigger_footswitch_capture()
            elif event.char and event.char.isprintable() and focused != self.scan_entry:
                self.global_scan_buffer.append(event.char)
                if len(self.global_scan_buffer) >= 10:
                    if self.auto_submit_timer:
                        self.root.after_cancel(self.auto_submit_timer)
                    self.auto_submit_timer = self.root.after(150, self._trigger_auto_submit_global)

        self.root.bind_all("<KeyPress>", _on_key_press)

    def _trigger_auto_submit_global(self):
        if self.active_popup or self.is_learning_footswitch:
            self.global_scan_buffer.clear()
            return
        if self.global_scan_buffer:
            scanned_text = "".join(self.global_scan_buffer).strip()
            self.global_scan_buffer.clear()
            self.scan_entry.delete(0, tk.END)
            if scanned_text:
                self.process_scan(scanned_text.upper())

    # ==========================================
    # 핵심 동작 1: 바코드 스캔 처리
    # ==========================================
    def process_scan(self, raw_code):
        self.clear_scan_buffers()
        raw_code = raw_code.strip().upper()

        if not raw_code:
            return

        if self.active_popup:
            self.play_triple_beep()
            return

        now_ts = time.time()
        if raw_code == self.last_scanned_code and (now_ts - self.last_scanned_time) < 2.0:
            return
        self.last_scanned_code = raw_code
        self.last_scanned_time = now_ts

        if self.inspection_state != self.STATE_WAIT_SCAN:
            step_desc = "1차 촬영 대기 중" if self.inspection_state == self.STATE_WAIT_SHOT1 else "2차 촬영(뒤집기) 대기 중"
            self.set_status("STEP NG", "#dc3545", "#3a1c1f")
            self.open_lock_popup(
                title_text=self.t("ng_incomplete_title"),
                msg=self.t("ng_incomplete_msg", curr_code=self.active_barcode, step_desc=step_desc),
                header_bg="#2d1d20", header_fg="#f87171"
            )
            return

        scanned_prefix = raw_code[:10].upper()
        self.lbl_last_scan.config(text=f"{self.t('last_scan')}: {raw_code}")

        if scanned_prefix != TARGET_BARCODE_PREFIX:
            self.set_status("MODEL NG", "#dc3545", "#3a1c1f")
            self.open_lock_popup(
                title_text=self.t("ng_model_title"),
                msg=self.t("ng_model_msg", target=TARGET_BARCODE_PREFIX, code=raw_code[:15]),
                header_bg="#2d1d20", header_fg="#f87171"
            )
            return

        if raw_code in self.sorting_list:
            self.set_status("SORTING", "#f59f00", "#3d2716")
            self.open_sorting_popup(raw_code)
            return

        is_already_scanned = (raw_code in self.scanned_history)

        if self.is_manager_mode:
            if not is_already_scanned:
                self.set_status("MGR NG", "#dc3545", "#3a1c1f")
                self.open_lock_popup(
                    title_text=self.t("ng_mgr_err_title"),
                    msg=self.t("ng_mgr_err_msg", code=raw_code),
                    header_bg="#2d1d20", header_fg="#f87171"
                )
                return
        elif not self.is_rework_mode:
            if is_already_scanned:
                now = datetime.now()
                day_str = now.strftime("%Y-%m-%d")
                time_str = now.strftime("%H:%M:%S")

                self.counts["ng"] += 1
                self.counts["total"] += 1
                self.save_counts()
                self.update_stat_cards()

                self.set_status("DUP NG", "#fd7e14", "#3d2716")
                dup_text = self.t("dup_scan_tag")

                self.handle_duplicate_in_excel_and_tree(raw_code, day_str, time_str, dup_text)

                self.open_lock_popup(
                    title_text=self.t("ng_dup_title"),
                    msg=self.t("ng_dup_msg", code=raw_code),
                    header_bg="#352316", header_fg="#fb923c"
                )
                return

        self.active_barcode = raw_code
        self.shot1_frame = None
        self.shot2_frame = None
        self.shot1_time_str = ""
        self.shot2_time_str = ""
        self.render_captured_thumbnail(self.canvas_shot1, None, 1)
        self.render_captured_thumbnail(self.canvas_shot2, None, 2)

        self.inspection_state = self.STATE_WAIT_SHOT1
        self.update_step_indicator()
        self.set_status("1차 촬영 대기", "#38bdf8", "#1e3a5f")
        self.play_ok_beep(900, 100)
        self.scan_entry.focus_set()

    # ==========================================
    # 핵심 동작 2: 풋스위치 1차 촬영 & 2차 촬영 처리
    # ==========================================
    def trigger_footswitch_capture(self):
        if self.active_popup:
            return

        now_ts = time.time()
        if (now_ts - self.last_footswitch_time) < FOOT_SWITCH_COOLDOWN:
            return
        self.last_footswitch_time = now_ts

        if self.inspection_state == self.STATE_WAIT_SCAN:
            self.play_triple_beep()
            self.set_status("바코드 선스캔!", "#f59f00", "#3d2716")
            return

        if not self.is_camera_ready or self.current_webcam_frame is None:
            self.open_cam_ng_popup()
            return

        captured = self.current_webcam_frame.copy()
        now_dt = datetime.now()
        time_str = now_dt.strftime("%H:%M:%S")
        day_str = now_dt.strftime("%Y-%m-%d")

        if self.inspection_state == self.STATE_WAIT_SHOT1:
            self.shot1_frame = captured
            self.shot1_time_str = time_str
            self.render_captured_thumbnail(self.canvas_shot1, self.shot1_frame, 1, time_str)

            self.inspection_state = self.STATE_WAIT_SHOT2
            self.update_step_indicator()
            self.set_status("뒤집어 2차 촬영", "#fde047", "#422006")
            self.play_ok_beep(1100, 130)

        elif self.inspection_state == self.STATE_WAIT_SHOT2:
            self.shot2_frame = captured
            self.shot2_time_str = time_str
            self.render_captured_thumbnail(self.canvas_shot2, self.shot2_frame, 2, time_str)

            saved_photo_name = self.save_inspection_photos(
                self.active_barcode, self.shot1_frame, self.shot2_frame,
                self.shot1_time_str, self.shot2_time_str
            )

            content_tag = ""
            if self.is_rework_mode:
                content_tag = "RE-WORK"
            elif self.is_manager_mode:
                content_tag = "[DMC 재스캔]"

            self.counts["ok"] += 1
            self.counts["total"] += 1
            self.save_counts()
            self.update_stat_cards()

            self.scanned_history.add(self.active_barcode)
            self.total_excel_seq += 1
            seq_no = self.total_excel_seq

            row_tuple = (
                seq_no, day_str, time_str, self.active_barcode,
                self.shot1_time_str, self.shot2_time_str, "OK", saved_photo_name, content_tag
            )
            self.tree.insert("", 0, values=row_tuple)

            self.append_record_to_single_excel(
                seq_no, day_str, time_str, self.active_barcode,
                f"{day_str} {self.shot1_time_str}", f"{day_str} {self.shot2_time_str}",
                "OK", saved_photo_name, content_tag
            )

            if self.is_manager_mode or self.is_rework_mode:
                self.is_manager_mode = False
                self.is_rework_mode = False
                self.btn_manager.config(bg="#2c323d", fg="#adb5bd", text=self.t("manager_btn"))

            self.inspection_state = self.STATE_WAIT_SCAN
            self.active_barcode = ""
            self.update_step_indicator()
            self.set_status("OK (검사 완료)", "#28a745", "#193322")
            self.play_ok_beep(1400, 180)
            self.scan_entry.focus_set()

    # ==========================================
    # 단일 엑셀 파일 관리
    # ==========================================
    def open_or_init_single_workbook(self):
        unhide_file(SINGLE_EXCEL_PATH)
        if os.path.exists(SINGLE_EXCEL_PATH):
            try:
                wb = openpyxl.load_workbook(SINGLE_EXCEL_PATH)
                ws = wb["스캔실적"] if "스캔실적" in wb.sheetnames else wb.active
                return wb, ws
            except Exception:
                pass

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "스캔실적"

        headers = [
            "순번", "검사일자", "완료시간", "제품 바코드 (DMC)",
            "1차 촬영일시", "2차 촬영일시", "판정", "사진 파일명", "Content"
        ]
        ws.append(headers)

        header_fill = PatternFill(start_color="1F242D", end_color="1F242D", fill_type="solid")
        header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")

        for col in range(1, 10):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        col_widths = {'A': 10, 'B': 14, 'C': 14, 'D': 34, 'E': 22, 'F': 22, 'G': 12, 'H': 36, 'I': 18}
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        ws_sort = wb.create_sheet(title="sorting")
        ws_sort.cell(row=1, column=2, value="Sorting 대상 DMC Code")
        ws_sort.cell(row=1, column=3, value="판정")
        ws_sort.column_dimensions['B'].width = 35
        ws_sort.column_dimensions['C'].width = 12

        return wb, ws

    def load_history_from_single_excel(self):
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(SINGLE_EXCEL_PATH):
            return

        today_dt = datetime.now().date()
        cutoff_date_str = (today_dt - timedelta(days=7)).strftime("%Y-%m-%d")

        def _loader():
            try:
                rows_to_insert = []
                self.scanned_history.clear()
                self.sorting_list.clear()
                max_seq = 0

                with self.file_lock:
                    unhide_file(SINGLE_EXCEL_PATH)
                    wb = openpyxl.load_workbook(SINGLE_EXCEL_PATH, data_only=True)

                    if "sorting" in wb.sheetnames:
                        ws_sort = wb["sorting"]
                        for r in range(2, 2001):
                            cell_v = ws_sort.cell(row=r, column=2).value
                            if cell_v:
                                self.sorting_list.add(str(cell_v).strip().upper())

                    ws = wb["스캔실적"] if "스캔실적" in wb.sheetnames else wb.active
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not row or len(row) < 7:
                            continue

                        seq_val = row[0]
                        day_val = str(row[1] or "").strip()
                        time_val = str(row[2] or "").strip()
                        dmc_val = str(row[3] or "").strip().upper()
                        s1_full = str(row[4] or "").strip()
                        s2_full = str(row[5] or "").strip()
                        res_val = str(row[6] or "OK").strip()
                        photo_val = str(row[7] or "").strip() if len(row) >= 8 and row[7] else ""
                        content_val = str(row[8] or "").strip() if len(row) >= 9 and row[8] else ""

                        if isinstance(seq_val, int) and seq_val > max_seq:
                            max_seq = seq_val

                        if dmc_val.startswith(TARGET_BARCODE_PREFIX):
                            self.scanned_history.add(dmc_val)

                        s1_short = s1_full.split()[-1] if " " in s1_full else s1_full
                        s2_short = s2_full.split()[-1] if " " in s2_full else s2_full

                        if day_val >= cutoff_date_str:
                            rows_to_insert.append((
                                seq_val, day_val, time_val, dmc_val,
                                s1_short, s2_short, res_val, photo_val, content_val
                            ))

                    hide_file(SINGLE_EXCEL_PATH)

                self.total_excel_seq = max_seq

                def _populate():
                    for r in reversed(rows_to_insert):
                        tag = "ng_row" if r[6] == "NG" else ""
                        self.tree.insert("", tk.END, values=r, tags=(tag,) if tag else ())

                self.root.after(0, _populate)
            except Exception:
                pass

        threading.Thread(target=_loader, daemon=True).start()

    def append_record_to_single_excel(self, seq_no, day_str, time_str, dmc_code, s1_ts, s2_ts, judgment, photo_name, content_text):
        with self.file_lock:
            try:
                wb, ws = self.open_or_init_single_workbook()

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
                )
                ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                row_data = [seq_no, day_str, time_str, dmc_code, s1_ts, s2_ts, judgment, photo_name, content_text]
                ws.append(row_data)
                r_idx = ws.max_row

                for col in range(1, 10):
                    c = ws.cell(row=r_idx, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in [1, 2, 3, 5, 6, 7, 9] else "left", vertical="center")
                    if col == 7 and judgment == "OK":
                        c.fill = ok_fill

                wb.save(SINGLE_EXCEL_PATH)
                hide_file(SINGLE_EXCEL_PATH)
            except Exception:
                pass

    def handle_duplicate_in_excel_and_tree(self, raw_code, day_str, time_str, dup_text):
        self.total_excel_seq += 1
        seq_no = self.total_excel_seq

        self.tree.insert("", 0, values=(seq_no, day_str, time_str, raw_code, "-", "-", "NG", "-", dup_text), tags=("ng_row",))
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            if vals and str(vals[3]).strip().upper() == raw_code and vals[6] != "NG":
                vals[6] = "NG"
                vals[8] = dup_text
                self.tree.item(item_id, values=vals, tags=("ng_row",))

        with self.file_lock:
            try:
                wb, ws = self.open_or_init_single_workbook()
                ng_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                ng_font = Font(name="맑은 고딕", size=10, bold=True, color="C00000")
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
                )

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    dmc_cell = row[3]
                    res_cell = row[6]
                    content_cell = row[8]
                    if dmc_cell.value and str(dmc_cell.value).strip().upper() == raw_code and res_cell.value != "NG":
                        res_cell.value = "NG"
                        content_cell.value = dup_text
                        for c in row[:9]:
                            c.fill = ng_fill
                            c.font = ng_font

                ws.append([seq_no, day_str, time_str, raw_code, "-", "-", "NG", "-", dup_text])
                h_idx = ws.max_row
                for col in range(1, 10):
                    c = ws.cell(row=h_idx, column=col)
                    c.border = thin_border
                    c.fill = ng_fill
                    c.font = ng_font
                    c.alignment = Alignment(horizontal="center" if col in [1, 2, 3, 5, 6, 7, 9] else "left", vertical="center")

                wb.save(SINGLE_EXCEL_PATH)
                hide_file(SINGLE_EXCEL_PATH)
            except Exception:
                pass

    # ==========================================
    # Re-code 탭 검색 및 엑셀 저장
    # ==========================================
    def apply_recode_filter(self):
        self.tree_recode.delete(*self.tree_recode.get_children())
        if not os.path.exists(SINGLE_EXCEL_PATH):
            return

        start_day = self.entry_start_day.get().strip()
        end_day = self.entry_end_day.get().strip()
        start_time = self.entry_start_time.get().strip()
        end_time = self.entry_end_time.get().strip()

        try:
            matched = []
            with self.file_lock:
                unhide_file(SINGLE_EXCEL_PATH)
                wb = openpyxl.load_workbook(SINGLE_EXCEL_PATH, data_only=True)
                ws = wb["스캔실적"] if "스캔실적" in wb.sheetnames else wb.active

                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or len(row) < 7:
                        continue
                    seq_val = row[0]
                    r_day = str(row[1] or "").strip()
                    r_time = str(row[2] or "00:00:00").strip()
                    dmc_val = str(row[3] or "").strip()
                    s1_val = str(row[4] or "").strip()
                    s2_val = str(row[5] or "").strip()
                    res_val = str(row[6] or "OK").strip()
                    photo_val = str(row[7] or "").strip() if len(row) >= 8 and row[7] else ""
                    content_val = str(row[8] or "").strip() if len(row) >= 9 and row[8] else ""

                    if start_day and r_day < start_day:
                        continue
                    if end_day and r_day > end_day:
                        continue
                    if start_time and r_time < start_time:
                        continue
                    if end_time and r_time > end_time:
                        continue

                    s1_short = s1_val.split()[-1] if " " in s1_val else s1_val
                    s2_short = s2_val.split()[-1] if " " in s2_val else s2_val
                    matched.append((seq_val, r_day, r_time, dmc_val, s1_short, s2_short, res_val, photo_val, content_val))

                hide_file(SINGLE_EXCEL_PATH)

            for m in reversed(matched):
                tag = "ng_row" if m[6] == "NG" else ""
                self.tree_recode.insert("", tk.END, values=m, tags=(tag,) if tag else ())
        except Exception:
            pass

    def save_recode_to_excel(self):
        items = self.tree_recode.get_children()
        if not items:
            return

        now_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_name = f"S_Structure_Recode_{now_ts}.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=initial_name,
            title="Save Records"
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Re-code"

            headers = ["순번", "DAY", "TIME", "Barcode (DMC)", "1차 촬영", "2차 촬영", "판정", "사진 파일명", "Content"]
            ws.append(headers)

            header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )

            for col in range(1, 10):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, item_id in enumerate(items, start=2):
                row_vals = self.tree_recode.item(item_id, "values")
                ws.append(list(row_vals))
                for col in range(1, 10):
                    c = ws.cell(row=idx, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in [1, 2, 3, 5, 6, 7, 9] else "left", vertical="center")

            wb.save(save_path)
            messagebox.showinfo("Success", f"Saved successfully:\n{save_path}", parent=self.root)
        except Exception:
            pass

    # ==========================================
    # 관리자 모드 / 행 삭제 / 잠금 팝업
    # ==========================================
    def toggle_manager_mode(self):
        if self.active_popup:
            return
        self.clear_scan_buffers()

        win = tk.Toplevel(self.root)
        win.title("MANAGER AUTH")
        win.resizable(False, False)
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()
        self.active_popup = win
        self.center_popup(win, 360, 210)

        tk.Label(win, text="MANAGER MODE\n관리자 비밀번호 6자리를 입력하세요.",
                 font=("맑은 고딕", 10, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 8))

        pw_entry = tk.Entry(win, show="*", font=("Arial", 14), justify="center", bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        pw_entry.pack(pady=5)

        lbl_err = tk.Label(win, text="", font=("맑은 고딕", 9), fg="#ff6b6b", bg=BG_PANEL)
        lbl_err.pack()

        def close_mgr_auth():
            self.close_active_popup(win)
            self.scan_entry.focus_set()

        def verify(event=None):
            if pw_entry.get() == self.admin_password:
                self.close_active_popup(win)
                self.open_manager_actions_dialog()
            else:
                lbl_err.config(text=self.t("pw_err"))
                pw_entry.delete(0, tk.END)
                pw_entry.focus_force()

        pw_entry.bind("<Return>", verify)
        win.protocol("WM_DELETE_WINDOW", close_mgr_auth)
        tk.Button(win, text="로그인", command=verify, bg="#28a745", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=4, cursor="hand2").pack(pady=10)
        pw_entry.focus_force()

    def open_manager_actions_dialog(self):
        self.clear_scan_buffers()
        dialog = tk.Toplevel(self.root)
        dialog.title("MANAGER CONTROL PANEL")
        dialog.resizable(False, False)
        dialog.configure(bg=BG_PANEL)
        dialog.transient(self.root)
        dialog.grab_set()
        self.active_popup = dialog
        self.center_popup(dialog, 480, 320)

        tk.Label(dialog, text="[ 관리자 작업 선택 ]", font=("맑은 고딕", 12, "bold"), fg="#38bdf8", bg=BG_PANEL).pack(pady=(15, 10))

        def close_panel():
            self.close_active_popup(dialog)
            self.scan_entry.focus_set()

        def act_rescan_dmc():
            self.is_manager_mode = True
            self.is_rework_mode = False
            self.btn_manager.config(bg="#f59f00", fg="#000000", text=self.t("manager_btn_on"))
            close_panel()
            messagebox.showinfo("안내", "중복 바코드 재검사 모드가 활성화되었습니다.\n1회 검사(1차·2차 촬영) 완료 시 자동으로 일반 모드로 전환됩니다.", parent=self.root)
            self.scan_entry.focus_set()

        tk.Button(dialog, text="🔄 중복 바코드 DMC 재검사(재촬영) 활성화", command=act_rescan_dmc,
                  bg="#1e3a5f", fg="#93c5fd", font=("맑은 고딕", 10, "bold"), relief="flat", pady=7, cursor="hand2").pack(fill=tk.X, padx=30, pady=5)

        def act_activate_rework():
            self.is_rework_mode = True
            self.is_manager_mode = False
            self.btn_manager.config(bg="#d97706", fg="#ffffff", text=self.t("manager_btn_rework"))
            close_panel()
            messagebox.showinfo("RE-WORK 모드 활성화", "🛠 RE-WORK 모드가 활성화되었습니다.\n중복 여부와 무관하게 재촬영 및 기록이 가능합니다.", parent=self.root)
            self.scan_entry.focus_set()

        tk.Button(dialog, text="🛠 RE-WORK 모드 활성화 (반송품 재검사)", command=act_activate_rework,
                  bg="#78350f", fg="#fde68a", font=("맑은 고딕", 10, "bold"), relief="flat", pady=7, cursor="hand2").pack(fill=tk.X, padx=30, pady=5)

        mode_names = {
            "merged_single": "현재: 1차+2차 좌우 병합 1개 파일 (바코드.jpg)",
            "separate": "현재: 1차/2차 개별 파일 (바코드_1차.jpg, 2차.jpg)",
            "both": "현재: 병합 파일 + 개별 파일 모두 저장"
        }
        cur_m = self.app_cfg.get("save_mode", "merged_single")

        def toggle_save_mode():
            order = ["merged_single", "both", "separate"]
            nxt = order[(order.index(self.app_cfg.get("save_mode", "merged_single")) + 1) % len(order)]
            self.app_cfg["save_mode"] = nxt
            save_app_cfg(self.app_cfg)
            btn_mode.config(text=f"🖼 사진 저장 형식 변경 [{mode_names[nxt]}]")

        btn_mode = tk.Button(dialog, text=f"🖼 사진 저장 형식 변경 [{mode_names.get(cur_m)}]", command=toggle_save_mode,
                             bg="#065f46", fg="#a7f3d0", font=("맑은 고딕", 9, "bold"), relief="flat", pady=7, cursor="hand2")
        btn_mode.pack(fill=tk.X, padx=30, pady=5)

        tk.Button(dialog, text="닫기", command=close_panel, bg="#374151", fg="#ffffff",
                  font=("맑은 고딕", 9, "bold"), relief="flat", padx=20, pady=4, cursor="hand2").pack(pady=(10, 5))
        dialog.protocol("WM_DELETE_WINDOW", close_panel)

    def trigger_hidden_delete_action(self):
        if self.active_popup:
            return
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 없음", "삭제할 기록 행을 우측 테이블에서 마우스로 먼저 선택해 주세요.", parent=self.root)
            self.scan_entry.focus_set()
            return

        self.clear_scan_buffers()
        win = tk.Toplevel(self.root)
        win.title("MANAGER AUTH [DELETE]")
        win.resizable(False, False)
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()
        self.active_popup = win
        self.center_popup(win, 360, 210)

        tk.Label(win, text="선택 기록 삭제\n관리자 비밀번호 6자리를 입력하세요.",
                 font=("맑은 고딕", 10, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 8))

        pw_entry = tk.Entry(win, show="*", font=("Arial", 14), justify="center", bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        pw_entry.pack(pady=5)

        lbl_err = tk.Label(win, text="", font=("맑은 고딕", 9), fg="#ff6b6b", bg=BG_PANEL)
        lbl_err.pack()

        def close_del():
            self.close_active_popup(win)
            self.scan_entry.focus_set()

        def verify_and_delete(event=None):
            if pw_entry.get() == self.admin_password:
                self.close_active_popup(win)
                self.delete_selected_rows(selected_items)
            else:
                lbl_err.config(text=self.t("pw_err"))
                pw_entry.delete(0, tk.END)
                pw_entry.focus_force()

        pw_entry.bind("<Return>", verify_and_delete)
        win.protocol("WM_DELETE_WINDOW", close_del)
        tk.Button(win, text="삭제 실행", command=verify_and_delete, bg="#dc3545", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=4, cursor="hand2").pack(pady=10)
        pw_entry.focus_force()

    def delete_selected_rows(self, selected_items):
        if not os.path.exists(SINGLE_EXCEL_PATH):
            return
        deleted_count = 0
        ok_removed = 0
        ng_removed = 0

        with self.file_lock:
            unhide_file(SINGLE_EXCEL_PATH)
            wb = openpyxl.load_workbook(SINGLE_EXCEL_PATH)
            ws = wb["스캔실적"] if "스캔실적" in wb.sheetnames else wb.active

            for item_id in selected_items:
                row_vals = self.tree.item(item_id, "values")
                if not row_vals or len(row_vals) < 7:
                    continue
                target_dmc = str(row_vals[3]).strip().upper()
                target_time = str(row_vals[2]).strip()
                res_val = str(row_vals[6]).strip()

                for r_idx in range(ws.max_row, 1, -1):
                    r_dmc = str(ws.cell(row=r_idx, column=4).value or "").strip().upper()
                    r_time = str(ws.cell(row=r_idx, column=3).value or "").strip()
                    if r_dmc == target_dmc and r_time == target_time:
                        ws.delete_rows(r_idx)
                        deleted_count += 1
                        if res_val == "NG":
                            ng_removed += 1
                        else:
                            ok_removed += 1
                        break

                self.tree.delete(item_id)
                if target_dmc in self.scanned_history:
                    self.scanned_history.discard(target_dmc)

            if deleted_count > 0:
                wb.save(SINGLE_EXCEL_PATH)
            hide_file(SINGLE_EXCEL_PATH)

        if deleted_count > 0:
            self.counts["ok"] = max(0, self.counts["ok"] - ok_removed)
            self.counts["ng"] = max(0, self.counts["ng"] - ng_removed)
            self.counts["total"] = max(0, self.counts["total"] - (ok_removed + ng_removed))
            self.save_counts()
            self.update_stat_cards()
            messagebox.showinfo("삭제 완료", f"선택한 기록 {deleted_count}건이 삭제되었습니다.", parent=self.root)

        self.scan_entry.focus_set()

    def open_sorting_popup(self, dmc_code):
        if self.active_popup:
            self.close_active_popup()
        self.clear_scan_buffers()

        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("sorting_title"))
        dialog.resizable(False, False)
        dialog.configure(bg="#3a1c1f")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        self.active_popup = dialog
        self.center_popup(dialog, 520, 290)
        self.play_triple_beep()

        msg = self.t("sorting_msg", code=dmc_code)
        tk.Label(dialog, text=msg, font=("맑은 고딕", 11, "bold"), bg="#3a1c1f", fg="#ff6b6b", justify=tk.LEFT).pack(pady=25, padx=20)

        def close_dialog(event=None):
            self.close_active_popup(dialog)
            self.set_status("READY", "#adb5bd", "#2a2e37")
            self.scan_entry.focus_set()

        dialog.bind("<Return>", close_dialog)
        btn = tk.Button(dialog, text=self.t("confirm_btn"), command=close_dialog,
                        font=("맑은 고딕", 11, "bold"), bg="#dc3545", fg="#ffffff", relief="flat", padx=20, pady=6)
        btn.pack(pady=10)
        btn.focus_force()

    def open_lock_popup(self, title_text, msg, header_bg, header_fg):
        if self.active_popup:
            self.close_active_popup()
        self.clear_scan_buffers()
        self.start_looping_ng_alarm()

        dialog = tk.Toplevel(self.root)
        dialog.title(title_text)
        dialog.resizable(False, False)
        dialog.configure(bg=header_bg)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        self.active_popup = dialog
        self.center_popup(dialog, 540, 330)

        def force_close_lock():
            self.stop_looping_ng_alarm()
            self.close_active_popup(dialog)
            if self.inspection_state == self.STATE_WAIT_SCAN:
                self.set_status("READY", "#adb5bd", "#2a2e37")
            self.scan_entry.focus_set()

        btn_hidden_close = tk.Button(
            dialog, text="", command=force_close_lock,
            bg=header_bg, activebackground=header_bg, bd=0, relief="flat", width=4, height=1, cursor="arrow"
        )
        btn_hidden_close.place(relx=1.0, x=-2, y=2, anchor="ne")

        tk.Label(dialog, text=msg, font=("맑은 고딕", 10, "bold"), bg=header_bg, fg=header_fg, justify=tk.LEFT).pack(pady=(20, 12), padx=20)

        pw_entry = tk.Entry(dialog, show="*", font=("Arial", 16), justify="center", width=14,
                            bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        pw_entry.pack(pady=5)

        lbl_err = tk.Label(dialog, text="", font=("맑은 고딕", 10, "bold"), fg="#ff6b6b", bg=header_bg)
        lbl_err.pack()

        def unlock(event=None):
            if pw_entry.get() == self.admin_password:
                force_close_lock()
            else:
                self.play_triple_beep()
                lbl_err.config(text=self.t("pw_err"))
                pw_entry.delete(0, tk.END)
                pw_entry.focus_force()

        pw_entry.bind("<Return>", unlock)
        tk.Button(dialog, text=self.t("unlock_btn"), command=unlock,
                  font=("맑은 고딕", 11, "bold"), bg="#dc3545", fg="#ffffff",
                  relief="flat", padx=16, pady=5, cursor="hand2").pack(pady=10)
        pw_entry.focus_force()

    def open_reset_dialog(self):
        if self.active_popup:
            return
        self.clear_scan_buffers()

        win = tk.Toplevel(self.root)
        win.title(self.t("reset_btn"))
        win.resizable(False, False)
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()
        self.active_popup = win
        self.center_popup(win, 360, 210)

        tk.Label(win, text="[S-Structure] 카운터를 초기화하려면\n관리자 비밀번호를 입력하세요.",
                 font=("맑은 고딕", 10, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 8))

        pw_entry = tk.Entry(win, show="*", font=("Arial", 14), justify="center", bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        pw_entry.pack(pady=5)

        lbl_err = tk.Label(win, text="", font=("맑은 고딕", 9), fg="#ff6b6b", bg=BG_PANEL)
        lbl_err.pack()

        def close_reset():
            self.close_active_popup(win)
            self.scan_entry.focus_set()

        def do_reset(event=None):
            if pw_entry.get() == self.admin_password:
                self.counts = {"total": 0, "ok": 0, "ng": 0}
                self.save_counts()
                self.update_stat_cards()
                close_reset()
            else:
                lbl_err.config(text=self.t("pw_err"))
                pw_entry.delete(0, tk.END)
                pw_entry.focus_force()

        pw_entry.bind("<Return>", do_reset)
        win.protocol("WM_DELETE_WINDOW", close_reset)
        tk.Button(win, text=self.t("unlock_btn"), command=do_reset, bg="#dc3545", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=4, cursor="hand2").pack(pady=10)
        pw_entry.focus_force()

    def change_password_dialog(self):
        if self.active_popup:
            return
        self.clear_scan_buffers()

        win = tk.Toplevel(self.root)
        win.title(self.t("pw_setting"))
        win.resizable(False, False)
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()
        self.active_popup = win
        self.center_popup(win, 340, 230)

        tk.Label(win, text="현재 비밀번호", font=("맑은 고딕", 9), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 2))
        curr_entry = tk.Entry(win, show="*", font=("Arial", 11), justify="center", bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        curr_entry.pack()

        tk.Label(win, text="새 6자리 숫자 비밀번호", font=("맑은 고딕", 9), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(10, 2))
        new_entry = tk.Entry(win, show="*", font=("Arial", 11), justify="center", bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        new_entry.pack()

        def close_pw():
            self.close_active_popup(win)
            self.scan_entry.focus_set()

        def apply_pw():
            if curr_entry.get() != self.admin_password:
                messagebox.showerror("Error", self.t("pw_err"), parent=win)
                return
            new_val = new_entry.get()
            if len(new_val) != 6 or not new_val.isdigit():
                messagebox.showerror("Error", "6자리 숫자로 입력하세요.", parent=win)
                return
            self.admin_password = new_val
            messagebox.showinfo("Success", "비밀번호가 변경되었습니다.", parent=win)
            close_pw()

        win.protocol("WM_DELETE_WINDOW", close_pw)
        tk.Button(win, text=self.t("unlock_btn"), command=apply_pw, bg="#2b5278", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=4, cursor="hand2").pack(pady=15)
        curr_entry.focus_force()

    def on_language_changed(self, event=None):
        self.root.title(self.t("title"))
        self.notebook.tab(0, text=self.t("tab_scan"))
        self.notebook.tab(1, text=self.t("tab_recode"))
        self.btn_pw.config(text=self.t("pw_setting"))
        self.btn_reset.config(text=self.t("reset_btn"))
        self.lbl_input_guide.config(text=self.t("input_guide"))
        self.lbl_right_header.config(text=self.t("record_header"))
        self.lbl_filter_day.config(text=self.t("filter_day"))
        self.lbl_filter_time.config(text=self.t("filter_time"))
        self.btn_search.config(text=self.t("search_btn"))
        self.btn_save.config(text=self.t("save_btn"))
        self.scan_entry.focus_set()


if __name__ == "__main__":
    root = tk.Tk()
    app = SStructureInspectorApp(root)
    root.mainloop()
