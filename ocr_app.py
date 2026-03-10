import os
import sys
import cv2
import mss
import json
import time
import numpy as np
import pytesseract
import pygetwindow as gw
import ctypes
from ctypes import wintypes
import concurrent.futures

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                             QComboBox, QFrame, QFileDialog, QScrollArea, QSlider, 
                             QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QGridLayout, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRect
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPen, QColor, QShortcut, QKeySequence, QIcon

# --- DPI Scaling Fixes ---
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTOSCREENSCALEFACTOR"] = "1"

# ==========================================================
# SET TESSERACT PATH (PyInstaller Compatible)
# ==========================================================
# This checks if we are running as a compiled .exe or as a Python script
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

local_tess = os.path.join(current_dir, "Tesseract-OCR", "tesseract.exe")

if os.path.exists(local_tess):
    pytesseract.pytesseract.tesseract_cmd = local_tess
    os.environ["TESSDATA_PREFIX"] = os.path.join(current_dir, "Tesseract-OCR", "tessdata")
else:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]

class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]

class ROIOverlayWidget(QWidget):
    rois_changed = pyqtSignal(list)
    roi_selected = pyqtSignal(int) 

    def __init__(self, scroll_area):
        super().__init__()
        self.scroll_area = scroll_area
        self.current_pixmap = QPixmap()
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus) 
        
        self.zoom_factor = 1.0
        self.is_new_source = True
        
        self.rois = []  
        self.area_counter = 1
        self.selected_id = None
        self.drag_state = None
        self.last_mouse_pos = None

    def add_field(self):
        self.rois.append({
            'id': self.area_counter,
            'name': f"Area {self.area_counter}", 
            'rect': [0.4, 0.4, 0.2, 0.2],
            'type': 'General Text',
            'threshold': 1,  
            'thickness': 5,  
            'confidence': 6, 
            'is_on_scene': False 
        })
        self.selected_id = self.area_counter
        self.area_counter += 1
        self.rois_changed.emit(self.rois)
        self.roi_selected.emit(self.selected_id)

    def remove_selected_field(self):
        if self.selected_id is not None:
            self.rois = [r for r in self.rois if r['id'] != self.selected_id]
            self.selected_id = None
            self.update()
            self.rois_changed.emit(self.rois)
            self.roi_selected.emit(-1)

    def add_to_scene(self):
        if self.selected_id is not None:
            for roi in self.rois:
                if roi['id'] == self.selected_id:
                    roi['is_on_scene'] = True
                    break
            self.update()
            self.rois_changed.emit(self.rois)

    def remove_from_scene(self):
        if self.selected_id is not None:
            for roi in self.rois:
                if roi['id'] == self.selected_id:
                    roi['is_on_scene'] = False
                    break
            self.update()
            self.rois_changed.emit(self.rois)

    def select_roi_by_id(self, roi_id):
        self.selected_id = roi_id
        self.update()

    def set_frame(self, pixmap):
        self.current_pixmap = pixmap
        if self.is_new_source:
            self.fit_to_view()
            self.is_new_source = False
        else:
            self.update_size()
        self.update() 

    def update_size(self):
        if self.current_pixmap.isNull(): return
        new_w = int(self.current_pixmap.width() * self.zoom_factor)
        new_h = int(self.current_pixmap.height() * self.zoom_factor)
        if self.width() != new_w or self.height() != new_h:
            self.setFixedSize(new_w, new_h)

    def fit_to_view(self):
        if self.current_pixmap.isNull(): return
        vw = self.scroll_area.viewport().width()
        vh = self.scroll_area.viewport().height()
        pw = self.current_pixmap.width()
        ph = self.current_pixmap.height()
        if pw > 0 and ph > 0:
            scale_w = vw / pw
            scale_h = vh / ph
            self.zoom_factor = min(scale_w, scale_h) * 0.95 
            self.update_size()

    def reset_view(self):
        self.fit_to_view()

    def wheelEvent(self, event):
        delta_y = event.angleDelta().y()
        delta_x = event.angleDelta().x()
        delta = delta_y if delta_y != 0 else delta_x
        if delta == 0: return

        modifiers = event.modifiers()
        if modifiers == Qt.KeyboardModifier.ControlModifier:
            if delta > 0: self.zoom_factor *= 1.15
            else: self.zoom_factor *= 0.85
            self.zoom_factor = max(0.2, min(self.zoom_factor, 10.0))
            self.update_size()
            event.accept()
        elif modifiers == Qt.KeyboardModifier.AltModifier:
            hbar = self.scroll_area.horizontalScrollBar()
            hbar.setValue(hbar.value() - int(delta / 2))
            event.accept()
        else:
            vbar = self.scroll_area.verticalScrollBar()
            vbar.setValue(vbar.value() - int(delta / 2))
            event.accept()

    def paintEvent(self, event):
        if self.current_pixmap.isNull(): return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

        w, h = self.width(), self.height()
        painter.drawPixmap(0, 0, w, h, self.current_pixmap)

        for roi in self.rois:
            if not roi['is_on_scene']: 
                continue

            nx, ny, nw, nh = roi['rect']
            rx, ry, rw, rh = int(nx * w), int(ny * h), int(nw * w), int(nh * h)
            is_selected = (roi['id'] == self.selected_id)

            if is_selected:
                painter.setPen(QPen(QColor(0, 255, 0), 3))
                painter.setBrush(QColor(0, 255, 0, 40)) 
            else:
                painter.setPen(QPen(QColor(255, 50, 50), 2))
                painter.setBrush(Qt.BrushStyle.NoBrush)

            painter.drawRect(rx, ry, rw, rh)
            
            name_text = roi['name'] if roi['name'] else "Unnamed"
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRect(rx, ry - 20, max(80, len(name_text)*8), 20)
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rx + 5, ry - 5, name_text)

            if is_selected:
                painter.setBrush(QColor(0, 255, 0))
                painter.drawRect(rx + rw - 12, ry + rh - 12, 12, 12)

    def mousePressEvent(self, event):
        mx, my = event.pos().x(), event.pos().y()
        w, h = self.width(), self.height()
        
        clicked_roi = None
        for roi in reversed(self.rois):
            if not roi['is_on_scene']: continue

            nx, ny, nw, nh = roi['rect']
            rx, ry, rw, rh = nx * w, ny * h, nw * w, nh * h
            
            resize_handle = QRect(int(rx + rw - 15), int(ry + rh - 15), 15, 15)
            full_box = QRect(int(rx), int(ry), int(rw), int(rh))
            
            if resize_handle.contains(mx, my):
                self.selected_id = roi['id']
                self.drag_state = 'resize'
                self.last_mouse_pos = (mx, my)
                clicked_roi = roi
                break
            elif full_box.contains(mx, my):
                self.selected_id = roi['id']
                self.drag_state = 'move'
                self.last_mouse_pos = (mx, my)
                clicked_roi = roi
                break
                
        if not clicked_roi: 
            self.selected_id = None
            self.roi_selected.emit(-1)
        else:
            self.roi_selected.emit(self.selected_id)
            
        self.update()

    def mouseMoveEvent(self, event):
        if self.selected_id is None or self.drag_state is None: return
        mx, my = event.pos().x(), event.pos().y()
        dx, dy = mx - self.last_mouse_pos[0], my - self.last_mouse_pos[1]
        self.last_mouse_pos = (mx, my)
        
        dnx, dny = dx / self.width(), dy / self.height()
        
        for roi in self.rois:
            if roi['id'] == self.selected_id and roi['is_on_scene']:
                nx, ny, nw, nh = roi['rect']
                if self.drag_state == 'move':
                    nx = max(0.0, min(nx + dnx, 1.0 - nw))
                    ny = max(0.0, min(ny + dny, 1.0 - nh))
                    roi['rect'] = [nx, ny, nw, nh]
                elif self.drag_state == 'resize':
                    nw = max(0.02, min(nw + dnx, 1.0 - nx)) 
                    nh = max(0.02, min(nh + dny, 1.0 - ny)) 
                    roi['rect'] = [nx, ny, nw, nh]
                break
                
        self.update()
        self.rois_changed.emit(self.rois)

    def mouseReleaseEvent(self, event):
        self.drag_state = None
        self.last_mouse_pos = None

class CaptureEngine(QThread):
    frame_signal = pyqtSignal(np.ndarray)
    ocr_signal = pyqtSignal(dict, dict) 
    previews_signal = pyqtSignal(dict) 

    def __init__(self):
        super().__init__()
        self.running = False
        self.ocr_enabled = False
        self.source_type = None 
        self.source_path = None  
        self.ocr_counter = 0
        self._new_source_requested = False
        self.active_rois =[]
        self.thread_count = 2
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=self.thread_count)
        self.active_futures =[]

    def set_thread_count(self, count):
        self.thread_count = count
        if self.executor:
            self.executor.shutdown(wait=False)
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=self.thread_count)
        self.active_futures =[] 

    def set_source_video(self, path):
        self.source_type = "video"
        self.source_path = path
        self._new_source_requested = True

    def set_source_screen(self, title):
        self.source_type = "screen"
        self.source_path = title
        self._new_source_requested = True

    def update_rois(self, rois):
        self.active_rois = rois

    def capture_window_direct(self, hwnd):
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        w, h = rect.right - rect.left, rect.bottom - rect.top
        if w <= 0 or h <= 0: return None
            
        hwndDC = user32.GetWindowDC(hwnd)
        mfcDC = gdi32.CreateCompatibleDC(hwndDC)
        saveBitMap = gdi32.CreateCompatibleBitmap(hwndDC, w, h)
        gdi32.SelectObject(mfcDC, saveBitMap)
        
        result = user32.PrintWindow(hwnd, mfcDC, 2)
        if result == 0:
            user32.ReleaseDC(hwnd, hwndDC); gdi32.DeleteDC(mfcDC); gdi32.DeleteObject(saveBitMap)
            return None
            
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = w
        bmi.bmiHeader.biHeight = -h 
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0
        
        buffer = ctypes.create_string_buffer(w * h * 4)
        gdi32.GetDIBits(mfcDC, saveBitMap, 0, h, buffer, ctypes.byref(bmi), 0)
        
        user32.ReleaseDC(hwnd, hwndDC); gdi32.DeleteDC(mfcDC); gdi32.DeleteObject(saveBitMap)
        img = np.frombuffer(buffer, dtype=np.uint8).reshape((h, w, 4)).copy()
        img[:, :, 3] = 255 
        return img

    def preprocess_image(self, crop, threshold_val, thickness_val):
        gray = cv2.cvtColor(crop, cv2.COLOR_BGRA2GRAY)
        if threshold_val <= 1:
            _, processed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        else:
            thresh_calc = int(255 * (threshold_val / 10.0))
            _, processed = cv2.threshold(gray, thresh_calc, 255, cv2.THRESH_BINARY)
            
        thick_calc = thickness_val - 5
        if thick_calc != 0:
            k_size = abs(thick_calc) + 1
            kernel = np.ones((k_size, k_size), np.uint8)
            if thick_calc > 0:
                processed = cv2.erode(processed, kernel, iterations=1)
            else:
                processed = cv2.dilate(processed, kernel, iterations=1)
        return processed

    def _process_single_roi(self, roi, processed_img):
        """Helper function to process a single ROI in its own thread."""
        # Using --psm 7 (Treat as a single text line) instead of 6 for massive speed boost
        config = "--psm 7"
        if roi['type'] == 'Numbers Only':
            config = "-c tessedit_char_whitelist=0123456789 --psm 7"
        elif roi['type'] == 'Time Format':
            config = "-c tessedit_char_whitelist=0123456789:. --psm 7"
            
        data = pytesseract.image_to_data(processed_img, config=config, output_type=pytesseract.Output.DICT)
        words = []
        req_conf = roi['confidence'] * 10 
        
        for i in range(len(data['text'])):
            conf = int(data['conf'][i])
            word = data['text'][i].strip()
            if word and conf >= req_conf:
                words.append(word)
                
        text = " ".join(words)
        return roi, text

    def _run_ocr_background(self, rois_copy, previews_copy, ocr_start_time):
        """Processes all ROIs for the current frame concurrently."""
        extracted_data = {}
        areas_scanned = 0
        
        # Filter to only the ROIs that have a processed preview image
        active_rois = [roi for roi in rois_copy if roi['id'] in previews_copy]
        
        if active_rois:
            # Spawn a thread pool specifically to process all these active ROIs at the exact same time
            with concurrent.futures.ThreadPoolExecutor(max_workers=len(active_rois)) as executor:
                futures =[]
                for roi in active_rois:
                    areas_scanned += 1
                    processed = previews_copy[roi['id']]
                    futures.append(executor.submit(self._process_single_roi, roi, processed))
                
                # Gather the results as they finish processing
                for future in concurrent.futures.as_completed(futures):
                    roi, text = future.result()
                    if text:
                        safe_name = roi['name'] if roi['name'] else f"Area_{roi['id']}"
                        extracted_data[safe_name] = text
                        
        ocr_end_time = time.time()
        if extracted_data:
            process_ms = int((ocr_end_time - ocr_start_time) * 1000)
            metadata = {"process_time_ms": process_ms, "areas_scanned": areas_scanned}
            self.ocr_signal.emit(metadata, extracted_data)

    def run(self):
        self.running = True
        with mss.mss() as sct:
            cap = None
            target_fps = 30.0 
            video_start_time = 0
            total_frames = 0

            while self.running:
                loop_start = time.time() 

                if self._new_source_requested:
                    if cap is not None: cap.release(); cap = None
                    self._new_source_requested = False
                    self.ocr_counter = 0

                frame = None

                if self.source_type == "video" and self.source_path:
                    if cap is None: 
                        cap = cv2.VideoCapture(self.source_path)
                        target_fps = cap.get(cv2.CAP_PROP_FPS)
                        if target_fps <= 0 or np.isnan(target_fps): target_fps = 30.0
                        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                        video_start_time = time.time()

                    elapsed = time.time() - video_start_time
                    target_frame = int(elapsed * target_fps)

                    if target_frame >= total_frames and total_frames > 0:
                        video_start_time = time.time()
                        target_frame = 0

                    cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
                    ret, v_frame = cap.read()
                    
                    if not ret:
                        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                        video_start_time = time.time()
                        continue
                    frame = cv2.cvtColor(v_frame, cv2.COLOR_BGR2BGRA)

                elif self.source_type == "screen" and self.source_path:
                    target_fps = 30.0 
                    try:
                        wins = gw.getWindowsWithTitle(self.source_path)
                        if wins and not wins[0].isMinimized:
                            frame = self.capture_window_direct(wins[0]._hWnd)
                            if frame is None:
                                win = wins[0]
                                screenshot = sct.grab({"top": win.top, "left": win.left, "width": win.width, "height": win.height})
                                frame = np.array(screenshot)
                    except:
                        pass

                if frame is not None:
                    self.frame_signal.emit(frame)
                    
                    previews = {}
                    fh, fw = frame.shape[:2]
                    scene_rois =[r for r in self.active_rois if r['is_on_scene']]
                    
                    for roi in scene_rois:
                        nx, ny, nw, nh = roi['rect']
                        x, y, w, h = int(nx * fw), int(ny * fh), int(nw * fw), int(nh * fh)
                        crop = frame[max(0, y):y+h, max(0, x):x+w]
                        if crop.size > 0:
                            processed = self.preprocess_image(crop, roi['threshold'], roi['thickness'])
                            previews[roi['id']] = processed
                    self.previews_signal.emit(previews)

                    if self.ocr_enabled:
                        self.ocr_counter += 1
                        trigger_frames = max(1, int(target_fps / 10)) 
                        
                        if self.ocr_counter >= trigger_frames: 
                            self.active_futures = [f for f in self.active_futures if not f.done()]
                            if len(self.active_futures) < self.thread_count:
                                if scene_rois:
                                    ocr_start_time = time.time() 
                                    rois_copy =[r.copy() for r in scene_rois]
                                    prev_copy = {k: v.copy() for k, v in previews.items()}
                                    future = self.executor.submit(self._run_ocr_background, rois_copy, prev_copy, ocr_start_time)
                                    self.active_futures.append(future)
                            self.ocr_counter = 0
                
                elapsed_time_ms = (time.time() - loop_start) * 1000
                target_delay_ms = 1000.0 / target_fps
                sleep_time = int(max(1, target_delay_ms - elapsed_time_ms))
                self.msleep(sleep_time)

            if cap: cap.release()
            self.executor.shutdown(wait=False)

    def stop(self):
        self.running = False


class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pro OCR Application")

        icon_path = os.path.join(current_dir, "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            

        self.setMinimumSize(1400, 900)
        self.setStyleSheet("""
            QMainWindow { background-color: #1e1e1e; } 
            QLabel { color: #ccc; font-size: 13px; }
            QPushButton { background-color: #383838; color: white; border-radius: 4px; padding: 6px; border: 1px solid #555; font-size: 13px; }
            QPushButton:hover { background-color: #4a4a4a; }
            QComboBox, QLineEdit { background-color: #2a2a2a; color: white; border: 1px solid #444; padding: 5px; border-radius: 3px; font-size: 13px; }
        """)
        
        self.engine = CaptureEngine()
        self.engine.frame_signal.connect(self.update_preview)
        self.engine.ocr_signal.connect(self.update_ocr_text)
        self.engine.previews_signal.connect(self.update_roi_preview)
        
        self.internal_update = False 
        self.init_ui()
        
        self.shortcut_reset = QShortcut(QKeySequence("Ctrl+R"), self)
        self.shortcut_reset.activated.connect(self.preview_overlay.reset_view)
        
        # Look for default workspace
        self.load_default_workspace()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # --- LEFT PANEL ---
        left_container = QWidget()
        left_container.setFixedWidth(440) 
        left_container.setStyleSheet("background-color: #242424; border-right: 1px solid #333;")
        left_master_layout = QVBoxLayout(left_container)
        left_master_layout.setContentsMargins(8, 8, 8, 8)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab { background: #2a2a2a; color: #888; padding: 8px 15px; border: 1px solid #333; border-bottom: none; font-size: 13px; }
            QTabBar::tab:selected { background: #3a3a3a; color: white; font-weight: bold; }
            QTabWidget::pane { border: 1px solid #333; background: #2a2a2a; }
        """)

        # ----- TAB 1: CONFIGURATION -----
        tab_config = QWidget()
        config_layout = QVBoxLayout(tab_config)
        config_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        source_header_layout = QHBoxLayout()
        source_header_layout.addWidget(QLabel("<b>Source Selection</b>"))
        source_header_layout.addStretch()
        
        source_header_layout.addWidget(QLabel("CPU Threads:"))
        self.combo_threads = QComboBox()
        max_cores = os.cpu_count() or 4
        self.combo_threads.addItems([str(i) for i in range(1, max_cores + 1)])
        self.combo_threads.setCurrentText("2") 
        self.combo_threads.currentTextChanged.connect(self.handle_thread_change)
        source_header_layout.addWidget(self.combo_threads)
        
        config_layout.addLayout(source_header_layout)

        self.combo_source = QComboBox()
        self.combo_source.addItems(["Select Source", "Open a Video File", "Screen Capture"])
        self.combo_source.currentIndexChanged.connect(self.handle_source_change)
        config_layout.addWidget(self.combo_source)

        self.screen_widget = QWidget()
        screen_layout = QHBoxLayout(self.screen_widget)
        screen_layout.setContentsMargins(0, 0, 0, 5)
        
        self.combo_windows = QComboBox()
        self.combo_windows.currentTextChanged.connect(self.handle_window_pick)
        
        self.btn_refresh_windows = QPushButton("Refresh")
        self.btn_refresh_windows.setFixedWidth(70)
        self.btn_refresh_windows.clicked.connect(self.refresh_window_list)
        
        screen_layout.addWidget(self.combo_windows, stretch=1)
        screen_layout.addWidget(self.btn_refresh_windows)
        
        self.screen_widget.hide()
        config_layout.addWidget(self.screen_widget)

        # --- PROFILE MANAGEMENT (NEW) ---
        profile_layout = QHBoxLayout()
        self.btn_load_prof = QPushButton("Load Profile")
        self.btn_save_prof = QPushButton("Save Profile")
        self.btn_save_def = QPushButton("Set as Default")
        
        self.btn_load_prof.clicked.connect(self.load_profile_dialog)
        self.btn_save_prof.clicked.connect(self.save_profile_dialog)
        self.btn_save_def.clicked.connect(self.save_default_workspace)
        
        profile_layout.addWidget(self.btn_load_prof)
        profile_layout.addWidget(self.btn_save_prof)
        profile_layout.addWidget(self.btn_save_def)
        config_layout.addLayout(profile_layout)

        # --- TABLE WITH 3 COLUMNS AND SIDE BUTTONS ---
        table_layout = QHBoxLayout()
        table_layout.setSpacing(5) 
        
        self.roi_table = QTableWidget(0, 3) 
        self.roi_table.setHorizontalHeaderLabels(["", "Field", "Value"])
        self.roi_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.roi_table.setColumnWidth(0, 25)
        self.roi_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.roi_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.roi_table.verticalHeader().setVisible(False)
        self.roi_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.roi_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.roi_table.setFixedHeight(160) 
        self.roi_table.setStyleSheet("""
            QTableWidget { background-color: #1e1e1e; color: white; border: 1px solid #444; gridline-color: #333; font-size: 13px; }
            QHeaderView::section { background-color: #333; color: white; border: none; padding: 4px; font-weight: bold; }
            QTableWidget::item:selected { background-color: #2980b9; }
        """)
        self.roi_table.itemSelectionChanged.connect(self.on_table_selection)
        self.roi_table.itemChanged.connect(self.on_table_item_changed) 
        
        table_layout.addWidget(self.roi_table)

        btn_layout = QVBoxLayout()
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(5)
        btn_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        self.btn_add_roi = QPushButton("+")
        self.btn_add_roi.setFixedSize(30, 30)
        self.btn_remove_roi = QPushButton("-")
        self.btn_remove_roi.setFixedSize(30, 30)
        
        btn_layout.addWidget(self.btn_add_roi)
        btn_layout.addWidget(self.btn_remove_roi)
        
        table_layout.addLayout(btn_layout)
        config_layout.addLayout(table_layout)

        scene_btn_layout = QHBoxLayout()
        self.btn_add_scene = QPushButton("Add to Scene ->")
        self.btn_remove_scene = QPushButton("Remove Selected")
        scene_btn_layout.addWidget(self.btn_add_scene)
        scene_btn_layout.addWidget(self.btn_remove_scene)
        config_layout.addLayout(scene_btn_layout)

        self.props_frame = QFrame()
        self.props_frame.setStyleSheet("QFrame { background-color: #2e2e2e; border-radius: 3px; border: 1px solid #444; margin-top: 10px; }")
        props_main_layout = QVBoxLayout(self.props_frame)
        props_main_layout.setContentsMargins(10, 10, 10, 10)
        props_main_layout.setSpacing(8)
        
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("Target:"))
        self.lbl_target = QLabel("Select an item above")
        self.lbl_target.setStyleSheet("color: #00d2ff; font-weight: bold; border: none;")
        header_layout.addWidget(self.lbl_target)
        header_layout.addStretch()
        
        self.btn_defaults = QPushButton("Defaults")
        self.btn_defaults.clicked.connect(self.reset_to_defaults)
        header_layout.addWidget(self.btn_defaults)
        
        props_main_layout.addLayout(header_layout)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(QLabel("Format"), 0, 0)
        self.combo_type = QComboBox()
        self.combo_type.addItems(["General Text", "Numbers Only", "Time Format"])
        self.combo_type.currentIndexChanged.connect(self.sync_properties)
        grid.addWidget(self.combo_type, 0, 1)

        self.sl_thresh = QSlider(Qt.Orientation.Horizontal)
        self.sl_thick = QSlider(Qt.Orientation.Horizontal)
        self.sl_conf = QSlider(Qt.Orientation.Horizontal)

        for sl in [self.sl_thresh, self.sl_thick, self.sl_conf]:
            sl.setRange(1, 10)
            sl.setTickPosition(QSlider.TickPosition.TicksBelow)
            sl.setTickInterval(1)
            sl.valueChanged.connect(self.sync_properties)
            sl.setStyleSheet("QSlider::handle:horizontal { background: #888; width: 12px; border-radius: 6px; }")

        grid.addWidget(QLabel("Binarize"), 1, 0)
        grid.addWidget(self.sl_thresh, 1, 1)

        grid.addWidget(QLabel("Cleanup/Dilate"), 2, 0)
        grid.addWidget(self.sl_thick, 2, 1)

        grid.addWidget(QLabel("Conf. Th"), 3, 0)
        grid.addWidget(self.sl_conf, 3, 1)

        props_main_layout.addLayout(grid)

        self.lbl_crop_preview = QLabel()
        self.lbl_crop_preview.setFixedSize(360, 60) 
        self.lbl_crop_preview.setStyleSheet("background-color: #000; border: 1px solid #555;")
        self.lbl_crop_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        preview_layout = QHBoxLayout()
        preview_layout.addStretch()
        preview_layout.addWidget(self.lbl_crop_preview)
        preview_layout.addStretch()
        props_main_layout.addLayout(preview_layout)

        config_layout.addWidget(self.props_frame)
        config_layout.addStretch() 

        # ----- TAB 2: LIVE OUTPUT -----
        tab_output = QWidget()
        output_layout = QVBoxLayout(tab_output)
        output_layout.setContentsMargins(8, 8, 8, 8)
        
        self.meta_card = QFrame()
        self.meta_card.setStyleSheet("background-color: #222; border-radius: 4px; border: 1px solid #444; padding: 4px;")
        meta_layout = QHBoxLayout(self.meta_card)
        
        self.lbl_meta_ping = QLabel("Processing: -- ms")
        self.lbl_meta_areas = QLabel("Scanned: 0")
        
        for lbl in [self.lbl_meta_ping, self.lbl_meta_areas]:
            lbl.setStyleSheet("color: #aaa; font-weight: bold; font-size: 14px; border: none;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            meta_layout.addWidget(lbl)
            
        output_layout.addWidget(self.meta_card)

        self.ocr_output = QTextEdit()
        self.ocr_output.setReadOnly(True)
        self.ocr_output.setStyleSheet("background-color: #0d0d0d; color: #00ff41; font-family: Consolas; font-size: 15px; border: 1px solid #333;")
        output_layout.addWidget(self.ocr_output)

        self.tabs.addTab(tab_config, "Configuration")
        self.tabs.addTab(tab_output, "Live Data")
        
        left_master_layout.addWidget(self.tabs)

        self.btn_ocr = QPushButton("START OCR DETECTION")
        self.btn_ocr.setCheckable(True)
        self.btn_ocr.setFixedHeight(50)
        self.btn_ocr.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; font-size: 15px; border: none; border-radius: 4px;")
        self.btn_ocr.clicked.connect(self.toggle_ocr_logic)
        left_master_layout.addWidget(self.btn_ocr)


        # --- RIGHT PANEL ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setStyleSheet("background-color: #000; border: none;")
        
        self.preview_overlay = ROIOverlayWidget(self.scroll_area)
        self.scroll_area.setWidget(self.preview_overlay)
        
        self.btn_add_roi.clicked.connect(self.preview_overlay.add_field)
        self.btn_remove_roi.clicked.connect(self.preview_overlay.remove_selected_field)
        self.btn_add_scene.clicked.connect(self.preview_overlay.add_to_scene)
        self.btn_remove_scene.clicked.connect(self.preview_overlay.remove_from_scene)
        
        self.preview_overlay.rois_changed.connect(self.sync_table_to_rois)
        self.preview_overlay.roi_selected.connect(self.populate_properties_panel)

        layout.addWidget(left_container)
        layout.addWidget(self.scroll_area, stretch=1)

        self.enable_properties_panel(False)

    # --- PROFILE MANAGEMENT LOGIC ---
    def save_profile_dialog(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Profile", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'w') as f:
                    json.dump(self.preview_overlay.rois, f, indent=4)
                QMessageBox.information(self, "Success", "Profile saved successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save profile: {e}")

    def load_profile_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Profile", "", "JSON Files (*.json)")
        if path:
            self.load_profile(path)

    def save_default_workspace(self):
        default_path = os.path.join(current_dir, "default_workspace.json")
        try:
            with open(default_path, 'w') as f:
                json.dump(self.preview_overlay.rois, f, indent=4)
            self.btn_save_def.setStyleSheet("background-color: #2ecc71; color: white;")
            self.btn_save_def.setText("Saved!")
            QApplication.processEvents()
            time.sleep(0.5)
            self.btn_save_def.setStyleSheet("background-color: #383838; color: white;")
            self.btn_save_def.setText("Set as Default")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to set default: {e}")

    def load_default_workspace(self):
        default_path = os.path.join(current_dir, "default_workspace.json")
        if os.path.exists(default_path):
            self.load_profile(default_path)

    def load_profile(self, path):
        try:
            with open(path, 'r') as f:
                data = json.load(f)
            
            # Reset UI States
            self.preview_overlay.selected_id = None
            self.enable_properties_panel(False)
            
            # Inject Data
            self.preview_overlay.rois = data
            
            # Fix area counter so new boxes don't overlap IDs
            if data:
                max_id = max(roi['id'] for roi in data)
                self.preview_overlay.area_counter = max_id + 1
            else:
                self.preview_overlay.area_counter = 1
                
            self.sync_table_to_rois(data)
            self.preview_overlay.update()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load profile: {e}")


    # --- EXISTING LOGIC ---
    def handle_thread_change(self, value):
        self.engine.set_thread_count(int(value))

    def reset_to_defaults(self):
        if self.preview_overlay.selected_id is None: return
        self.combo_type.setCurrentText("General Text")
        self.sl_thresh.setValue(1)  
        self.sl_thick.setValue(5)   
        self.sl_conf.setValue(6)    

    def on_table_item_changed(self, item):
        if self.internal_update: return
        if item.column() == 1:
            roi_id = item.data(Qt.ItemDataRole.UserRole)
            new_name = item.text().strip()
            for roi in self.preview_overlay.rois:
                if roi['id'] == roi_id:
                    roi['name'] = new_name
                    if self.preview_overlay.selected_id == roi_id:
                        self.lbl_target.setText(new_name if new_name else "Unnamed Field")
                    break
            self.preview_overlay.update()
            self.engine.update_rois(self.preview_overlay.rois)

    def on_table_selection(self):
        if self.internal_update: return
        row = self.roi_table.currentRow()
        if row >= 0:
            item = self.roi_table.item(row, 1)
            roi_id = item.data(Qt.ItemDataRole.UserRole)
            self.preview_overlay.select_roi_by_id(roi_id)
            self.populate_properties_panel(roi_id)

    def sync_table_to_rois(self, rois):
        self.internal_update = True
        self.roi_table.setRowCount(len(rois))
        
        for i, roi in enumerate(rois):
            status_char = "✓" if roi['is_on_scene'] else "✖"
            item_status = QTableWidgetItem(status_char)
            item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_status.setFlags(item_status.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if roi['is_on_scene']:
                item_status.setForeground(QColor("#2ecc71")) 
            else:
                item_status.setForeground(QColor("#888888")) 
                
            item_name = QTableWidgetItem(roi['name'])
            item_name.setData(Qt.ItemDataRole.UserRole, roi['id']) 
            item_name.setFlags(item_name.flags() | Qt.ItemFlag.ItemIsEditable)
            
            curr_val = self.roi_table.item(i, 2)
            val_text = curr_val.text() if curr_val else ""
            item_val = QTableWidgetItem(val_text)
            item_val.setFlags(item_val.flags() & ~Qt.ItemFlag.ItemIsEditable)
            
            self.roi_table.setItem(i, 0, item_status)
            self.roi_table.setItem(i, 1, item_name)
            self.roi_table.setItem(i, 2, item_val)
            
            if roi['id'] == self.preview_overlay.selected_id:
                self.roi_table.selectRow(i)
                
        self.internal_update = False
        self.engine.update_rois(rois)

    def enable_properties_panel(self, enabled):
        self.combo_type.setEnabled(enabled)
        self.sl_thresh.setEnabled(enabled)
        self.sl_thick.setEnabled(enabled)
        self.sl_conf.setEnabled(enabled)
        self.btn_defaults.setEnabled(enabled)
        
        if not enabled:
            self.lbl_target.setText("Select an item above")
            self.lbl_crop_preview.clear()

    def handle_source_change(self, index):
        source = self.combo_source.currentText()
        self.screen_widget.hide()
        if source == "Open a Video File":
            path, _ = QFileDialog.getOpenFileName(self, "Select Video", "", "Video Files (*.mp4 *.avi *.mkv)")
            if path:
                self.preview_overlay.is_new_source = True
                self.engine.set_source_video(path)
                if not self.engine.isRunning(): self.engine.start()
        elif source == "Screen Capture":
            self.screen_widget.show()
            self.refresh_window_list()

    def refresh_window_list(self):
        self.combo_windows.blockSignals(True)
        self.combo_windows.clear()
        self.combo_windows.addItem("--- Select Window ---")
        titles = sorted([w.title for w in gw.getAllWindows() if w.title.strip()])
        self.combo_windows.addItems(titles)
        self.combo_windows.blockSignals(False)

    def handle_window_pick(self, title):
        if title and title != "--- Select Window ---":
            self.preview_overlay.is_new_source = True
            self.engine.set_source_screen(title)
            if not self.engine.isRunning(): self.engine.start()
        else:
            self.engine.stop()

    def populate_properties_panel(self, roi_id):
        if roi_id == -1:
            self.enable_properties_panel(False)
            self.roi_table.clearSelection()
            return
            
        roi = next((r for r in self.preview_overlay.rois if r['id'] == roi_id), None)
        if roi:
            self.enable_properties_panel(True)
            self.lbl_target.setText(roi['name'] if roi['name'] else "Unnamed Field")
            
            self.combo_type.blockSignals(True)
            self.sl_thresh.blockSignals(True)
            self.sl_thick.blockSignals(True)
            self.sl_conf.blockSignals(True)

            self.combo_type.setCurrentText(roi['type'])
            self.sl_thresh.setValue(roi['threshold'])
            self.sl_thick.setValue(roi['thickness'])
            self.sl_conf.setValue(roi['confidence'])

            self.combo_type.blockSignals(False)
            self.sl_thresh.blockSignals(False)
            self.sl_thick.blockSignals(False)
            self.sl_conf.blockSignals(False)
            
            self.internal_update = True
            for i in range(self.roi_table.rowCount()):
                if self.roi_table.item(i, 1).data(Qt.ItemDataRole.UserRole) == roi_id:
                    self.roi_table.selectRow(i)
                    break
            self.internal_update = False

    def sync_properties(self):
        roi_id = self.preview_overlay.selected_id
        if roi_id is None: return
        
        for roi in self.preview_overlay.rois:
            if roi['id'] == roi_id:
                roi['type'] = self.combo_type.currentText()
                roi['threshold'] = self.sl_thresh.value()
                roi['thickness'] = self.sl_thick.value()
                roi['confidence'] = self.sl_conf.value()
                break
                
        self.engine.update_rois(self.preview_overlay.rois)

    def toggle_ocr_logic(self):
        state = self.btn_ocr.isChecked()
        self.engine.ocr_enabled = state
        self.btn_ocr.setText("STOP OCR DETECTION" if state else "START OCR DETECTION")
        self.btn_ocr.setStyleSheet(f"background-color: {'#c0392b' if state else '#2980b9'}; color: white; font-weight: bold; font-size: 15px; border: none; border-radius: 4px;")
        
        if state:
            self.tabs.setCurrentIndex(1)

    def update_preview(self, frame):
        h, w, c = frame.shape
        q_img = QImage(frame.data, w, h, w*c, QImage.Format.Format_RGBA8888).rgbSwapped()
        self.preview_overlay.set_frame(QPixmap.fromImage(q_img))

    def update_roi_preview(self, previews_dict):
        roi_id = self.preview_overlay.selected_id
        if roi_id is not None and roi_id in previews_dict:
            processed = previews_dict[roi_id]
            h, w = processed.shape
            q_img = QImage(processed.data, w, h, w, QImage.Format.Format_Grayscale8)
            pixmap = QPixmap.fromImage(q_img)
            scaled = pixmap.scaled(340, 60, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.lbl_crop_preview.setPixmap(scaled)

    def update_ocr_text(self, metadata, data_dict):
        self.lbl_meta_ping.setText(f"Processing: {metadata.get('process_time_ms', '--')} ms")
        self.lbl_meta_areas.setText(f"Scanned: {metadata.get('areas_scanned', 0)}")
        
        self.internal_update = True
        
        display_dict = {}
        
        for i in range(self.roi_table.rowCount()):
            field_name = self.roi_table.item(i, 1).text()
            roi_id = self.roi_table.item(i, 1).data(Qt.ItemDataRole.UserRole)
            safe_name = field_name if field_name else f"Area_{roi_id}"
            
            current_value = self.roi_table.item(i, 2).text()
            
            if safe_name in data_dict:
                new_val = data_dict[safe_name]
                self.roi_table.item(i, 2).setText(new_val)
                display_dict[safe_name] = new_val
            else:
                display_dict[safe_name] = current_value
                
        self.internal_update = False
        
        formatted_json = json.dumps(display_dict, indent=4)
        self.ocr_output.setPlainText(formatted_json)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = OCRApp()
    win.show()
    sys.exit(app.exec())