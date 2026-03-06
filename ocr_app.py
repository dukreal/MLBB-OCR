import os
import sys
import cv2
import mss
import json
import numpy as np
import pytesseract
import pygetwindow as gw
import ctypes
from ctypes import wintypes
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                             QComboBox, QFrame, QSplitter, QFileDialog, QScrollArea, QSlider, QLineEdit)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRect
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPen, QColor, QShortcut, QKeySequence

# --- DPI Scaling Fixes ---
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTOSCREENSCALEFACTOR"] = "1"

# ==========================================================
# SET YOUR TESSERACT PATH HERE
# ==========================================================
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

    def add_roi(self):
        self.rois.append({
            'id': self.area_counter,
            'name': f'Area_{self.area_counter}',
            'rect': [0.4, 0.4, 0.2, 0.2],
            'type': 'General Text',
            'threshold': 0, 
            'thickness': 0, 
            'confidence': 60 
        })
        self.selected_id = self.area_counter
        self.area_counter += 1
        self.update()
        self.rois_changed.emit(self.rois)
        self.roi_selected.emit(self.selected_id)

    def remove_selected_roi(self):
        if self.selected_id is not None:
            self.rois = [r for r in self.rois if r['id'] != self.selected_id]
            self.selected_id = None
            self.update()
            self.rois_changed.emit(self.rois)
            self.roi_selected.emit(-1)

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
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRect(rx, ry - 20, max(80, len(roi['name'])*8), 20)
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rx + 5, ry - 5, roi['name'])

            if is_selected:
                painter.setBrush(QColor(0, 255, 0))
                painter.drawRect(rx + rw - 12, ry + rh - 12, 12, 12)

    def mousePressEvent(self, event):
        mx, my = event.pos().x(), event.pos().y()
        w, h = self.width(), self.height()
        
        clicked_roi = None
        for roi in reversed(self.rois):
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
            if roi['id'] == self.selected_id:
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
    ocr_signal = pyqtSignal(str)
    previews_signal = pyqtSignal(dict) 

    def __init__(self):
        super().__init__()
        self.running = False
        self.ocr_enabled = False
        self.source_type = None 
        self.source_path = None  
        self.ocr_counter = 0
        self._new_source_requested = False
        self.active_rois = []

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

    def preprocess_image(self, crop, threshold, thickness):
        gray = cv2.cvtColor(crop, cv2.COLOR_BGRA2GRAY)
        if threshold == 0:
            _, processed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        else:
            _, processed = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)
            
        if thickness != 0:
            k_size = abs(thickness) + 1
            kernel = np.ones((k_size, k_size), np.uint8)
            if thickness > 0:
                processed = cv2.erode(processed, kernel, iterations=1)
            else:
                processed = cv2.dilate(processed, kernel, iterations=1)
        return processed

    def run(self):
        self.running = True
        with mss.mss() as sct:
            cap = None
            while self.running:
                if self._new_source_requested:
                    if cap is not None: cap.release(); cap = None
                    self._new_source_requested = False

                frame = None

                if self.source_type == "video" and self.source_path:
                    if cap is None: cap = cv2.VideoCapture(self.source_path)
                    ret, v_frame = cap.read()
                    if not ret:
                        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                        continue
                    frame = cv2.cvtColor(v_frame, cv2.COLOR_BGR2BGRA)

                elif self.source_type == "screen" and self.source_path:
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
                    for roi in self.active_rois:
                        nx, ny, nw, nh = roi['rect']
                        x, y, w, h = int(nx * fw), int(ny * fh), int(nw * fw), int(nh * fh)
                        crop = frame[max(0, y):y+h, max(0, x):x+w]
                        if crop.size > 0:
                            processed = self.preprocess_image(crop, roi['threshold'], roi['thickness'])
                            previews[roi['id']] = processed
                    self.previews_signal.emit(previews)

                    if self.ocr_enabled:
                        self.ocr_counter += 1
                        if self.ocr_counter >= 30: 
                            json_results = {}
                            if not self.active_rois:
                                gray = cv2.cvtColor(frame, cv2.COLOR_BGRA2GRAY)
                                text = pytesseract.image_to_string(gray).strip()
                                if text: json_results["Full_Screen"] = text
                            else:
                                for roi in self.active_rois:
                                    if roi['id'] in previews:
                                        processed = previews[roi['id']]
                                        config = "--psm 6"
                                        if roi['type'] == 'Numbers Only':
                                            config = "-c tessedit_char_whitelist=0123456789 --psm 6"
                                        elif roi['type'] == 'Time Format':
                                            config = "-c tessedit_char_whitelist=0123456789:. --psm 6"
                                            
                                        data = pytesseract.image_to_data(processed, config=config, output_type=pytesseract.Output.DICT)
                                        words = []
                                        for i in range(len(data['text'])):
                                            conf = int(data['conf'][i])
                                            word = data['text'][i].strip()
                                            if word and conf >= roi['confidence']:
                                                words.append(word)
                                                
                                        text = " ".join(words)
                                        if text:
                                            json_results[roi['name']] = text
                            
                            if json_results:
                                self.ocr_signal.emit(json.dumps(json_results, indent=4))
                            self.ocr_counter = 0
                
                self.msleep(30)

            if cap: cap.release()

    def stop(self):
        self.running = False


class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pro OCR - Fixed Properties Panel")
        self.setMinimumSize(1400, 900)
        self.setStyleSheet("QMainWindow { background-color: #1a1a1a; } QLabel { color: #eee; }")
        
        self.engine = CaptureEngine()
        self.engine.frame_signal.connect(self.update_preview)
        self.engine.ocr_signal.connect(self.update_ocr_text)
        self.engine.previews_signal.connect(self.update_roi_preview)
        
        self.init_ui()
        
        self.shortcut_reset = QShortcut(QKeySequence("Ctrl+R"), self)
        self.shortcut_reset.activated.connect(self.preview_overlay.reset_view)

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- LEFT PANEL ---
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setStyleSheet("QScrollArea { border: none; }")
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Source
        left_layout.addWidget(QLabel("<b>SOURCE SELECTION</b>"))
        self.combo_source = QComboBox()
        self.combo_source.addItems(["Select Source", "Open a Video File", "Screen Capture"])
        self.combo_source.currentIndexChanged.connect(self.handle_source_change)
        left_layout.addWidget(self.combo_source)

        self.screen_widget = QWidget()
        screen_layout = QVBoxLayout(self.screen_widget)
        screen_layout.setContentsMargins(0, 5, 0, 5)
        screen_layout.addWidget(QLabel("<b>SELECT WINDOW</b>"))
        self.combo_windows = QComboBox()
        self.combo_windows.currentTextChanged.connect(self.handle_window_pick)
        screen_layout.addWidget(self.combo_windows)
        self.screen_widget.hide()
        left_layout.addWidget(self.screen_widget)

        # Region Buttons
        roi_buttons_layout = QHBoxLayout()
        self.btn_add_roi = QPushButton("+ Add Area")
        self.btn_add_roi.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 8px;")
        roi_buttons_layout.addWidget(self.btn_add_roi)
        self.btn_remove_roi = QPushButton("- Remove Selected")
        self.btn_remove_roi.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold; padding: 8px;")
        roi_buttons_layout.addWidget(self.btn_remove_roi)
        left_layout.addLayout(roi_buttons_layout)

        # --- FIXED ROI PROPERTIES PANEL ---
        self.props_frame = QFrame()
        self.props_frame.setStyleSheet("background-color: #2a2a2a; border-radius: 5px; padding: 5px;")
        props_layout = QVBoxLayout(self.props_frame)
        props_layout.addWidget(QLabel("<b>REGION PROPERTIES</b>"))
        
        props_layout.addWidget(QLabel("Region Name:"))
        self.inp_name = QLineEdit()
        self.inp_name.textChanged.connect(self.sync_properties)
        props_layout.addWidget(self.inp_name)

        props_layout.addWidget(QLabel("Text Type:"))
        self.combo_type = QComboBox()
        self.combo_type.addItems(["General Text", "Numbers Only", "Time Format"])
        self.combo_type.currentIndexChanged.connect(self.sync_properties)
        props_layout.addWidget(self.combo_type)

        self.lbl_thresh = QLabel("Binarization (Auto)")
        props_layout.addWidget(self.lbl_thresh)
        self.sl_thresh = QSlider(Qt.Orientation.Horizontal)
        self.sl_thresh.setRange(0, 255)
        self.sl_thresh.valueChanged.connect(self.sync_properties)
        props_layout.addWidget(self.sl_thresh)

        self.lbl_thick = QLabel("Thickness (0)")
        props_layout.addWidget(self.lbl_thick)
        self.sl_thick = QSlider(Qt.Orientation.Horizontal)
        self.sl_thick.setRange(-3, 3)
        self.sl_thick.setValue(0)
        self.sl_thick.valueChanged.connect(self.sync_properties)
        props_layout.addWidget(self.sl_thick)

        self.lbl_conf = QLabel("Confidence Filter (60%)")
        props_layout.addWidget(self.lbl_conf)
        self.sl_conf = QSlider(Qt.Orientation.Horizontal)
        self.sl_conf.setRange(0, 100)
        self.sl_conf.setValue(60)
        self.sl_conf.valueChanged.connect(self.sync_properties)
        props_layout.addWidget(self.sl_conf)

        props_layout.addWidget(QLabel("<i>Live Processed Crop:</i>"))
        self.lbl_crop_preview = QLabel("No Region Selected")
        self.lbl_crop_preview.setMinimumHeight(60)
        self.lbl_crop_preview.setStyleSheet("background-color: #000; border: 1px solid #555;")
        self.lbl_crop_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        props_layout.addWidget(self.lbl_crop_preview)

        left_layout.addWidget(self.props_frame)
        
        # Disable properties by default
        self.enable_properties_panel(False)

        # OCR Output
        self.btn_ocr = QPushButton("START OCR DETECTION")
        self.btn_ocr.setCheckable(True)
        self.btn_ocr.setFixedHeight(50)
        self.btn_ocr.setStyleSheet("background-color: #333; color: white; font-weight: bold; margin-top: 15px;")
        self.btn_ocr.clicked.connect(self.toggle_ocr_logic)
        left_layout.addWidget(self.btn_ocr)

        left_layout.addWidget(QLabel("<b>JSON OUTPUT</b>"))
        self.ocr_output = QTextEdit()
        self.ocr_output.setReadOnly(True)
        self.ocr_output.setStyleSheet("background-color: #0d0d0d; color: #00ff41; font-family: Consolas; font-size: 13px; border: 1px solid #333;")
        left_layout.addWidget(self.ocr_output)

        left_scroll.setWidget(left_panel)

        # --- RIGHT PANEL ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setStyleSheet("background-color: #000; border-left: 2px solid #333;")
        
        self.preview_overlay = ROIOverlayWidget(self.scroll_area)
        self.scroll_area.setWidget(self.preview_overlay)
        
        self.btn_add_roi.clicked.connect(self.preview_overlay.add_roi)
        self.btn_remove_roi.clicked.connect(self.preview_overlay.remove_selected_roi)
        self.preview_overlay.rois_changed.connect(self.engine.update_rois)
        self.preview_overlay.roi_selected.connect(self.populate_properties_panel)

        splitter.addWidget(left_scroll)
        splitter.addWidget(self.scroll_area)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter)

    def enable_properties_panel(self, enabled):
        """Grays out or enables the properties panel without hiding it."""
        self.inp_name.setEnabled(enabled)
        self.combo_type.setEnabled(enabled)
        self.sl_thresh.setEnabled(enabled)
        self.sl_thick.setEnabled(enabled)
        self.sl_conf.setEnabled(enabled)
        
        if not enabled:
            self.inp_name.blockSignals(True)
            self.inp_name.clear()
            self.inp_name.blockSignals(False)
            self.lbl_crop_preview.clear()
            self.lbl_crop_preview.setText("No Region Selected")
            self.lbl_thresh.setText("Binarization")
            self.lbl_thick.setText("Thickness")
            self.lbl_conf.setText("Confidence Filter")

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
        self.combo_windows.clear()
        titles = sorted([w.title for w in gw.getAllWindows() if w.title.strip()])
        self.combo_windows.addItems(titles)

    def handle_window_pick(self, title):
        if title:
            self.preview_overlay.is_new_source = True
            self.engine.set_source_screen(title)
            if not self.engine.isRunning(): self.engine.start()

    def populate_properties_panel(self, roi_id):
        if roi_id == -1:
            self.enable_properties_panel(False)
            return
            
        roi = next((r for r in self.preview_overlay.rois if r['id'] == roi_id), None)
        if roi:
            self.enable_properties_panel(True)
            
            self.inp_name.blockSignals(True)
            self.combo_type.blockSignals(True)
            self.sl_thresh.blockSignals(True)
            self.sl_thick.blockSignals(True)
            self.sl_conf.blockSignals(True)

            self.inp_name.setText(roi['name'])
            self.combo_type.setCurrentText(roi['type'])
            
            self.sl_thresh.setValue(roi['threshold'])
            self.lbl_thresh.setText(f"Binarization ({'Auto' if roi['threshold']==0 else roi['threshold']})")
            
            self.sl_thick.setValue(roi['thickness'])
            self.lbl_thick.setText(f"Thickness ({roi['thickness']})")
            
            self.sl_conf.setValue(roi['confidence'])
            self.lbl_conf.setText(f"Confidence Filter ({roi['confidence']}%)")

            self.inp_name.blockSignals(False)
            self.combo_type.blockSignals(False)
            self.sl_thresh.blockSignals(False)
            self.sl_thick.blockSignals(False)
            self.sl_conf.blockSignals(False)

    def sync_properties(self):
        roi_id = self.preview_overlay.selected_id
        if roi_id is None: return
        
        for roi in self.preview_overlay.rois:
            if roi['id'] == roi_id:
                roi['name'] = self.inp_name.text()
                roi['type'] = self.combo_type.currentText()
                roi['threshold'] = self.sl_thresh.value()
                roi['thickness'] = self.sl_thick.value()
                roi['confidence'] = self.sl_conf.value()
                
                self.lbl_thresh.setText(f"Binarization ({'Auto' if roi['threshold']==0 else roi['threshold']})")
                self.lbl_thick.setText(f"Thickness ({roi['thickness']})")
                self.lbl_conf.setText(f"Confidence Filter ({roi['confidence']}%)")
                break
                
        self.preview_overlay.update()
        self.engine.update_rois(self.preview_overlay.rois)

    def toggle_ocr_logic(self):
        state = self.btn_ocr.isChecked()
        self.engine.ocr_enabled = state
        self.btn_ocr.setText("STOP OCR DETECTION" if state else "START OCR DETECTION")
        self.btn_ocr.setStyleSheet(f"background-color: {'#c0392b' if state else '#333'}; color: white; font-weight: bold; margin-top: 15px;")

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
            scaled = pixmap.scaled(self.lbl_crop_preview.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.lbl_crop_preview.setPixmap(scaled)

    def update_ocr_text(self, text):
        self.ocr_output.append(f"{text}\n")
        self.ocr_output.verticalScrollBar().setValue(self.ocr_output.verticalScrollBar().maximum())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = OCRApp()
    win.show()
    sys.exit(app.exec())