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
                             QComboBox, QFrame, QSplitter, QFileDialog)
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
    """Custom Widget to draw the video feed, handle ROIs, Zoom, and Pan."""
    rois_changed = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.current_pixmap = QPixmap()
        self.pixmap_rect = QRect()
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus) # Required to capture key events properly
        
        # Viewport (Zoom/Pan) Settings
        self.zoom_factor = 1.0
        self.pan_x = 0
        self.pan_y = 0
        
        # ROI Data
        self.rois = []  
        self.area_counter = 1
        self.selected_id = None
        
        # Interaction States
        self.drag_state = None
        self.last_mouse_pos = None

    def add_roi(self):
        self.rois.append({
            'id': self.area_counter,
            'name': f'Area_{self.area_counter}',
            'rect': [0.4, 0.4, 0.2, 0.2]  # Normalized coordinates
        })
        self.selected_id = self.area_counter
        self.area_counter += 1
        self.update()
        self.rois_changed.emit(self.rois)

    def remove_selected_roi(self):
        if self.selected_id is not None:
            self.rois = [r for r in self.rois if r['id'] != self.selected_id]
            self.selected_id = None
            self.update()
            self.rois_changed.emit(self.rois)

    def set_frame(self, pixmap):
        self.current_pixmap = pixmap
        self.update() 

    def reset_view(self):
        """Triggered by Ctrl+R to reset zoom and pan"""
        self.zoom_factor = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.update()

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        modifiers = event.modifiers()

        if modifiers == Qt.KeyboardModifier.ControlModifier:
            # Zoom In/Out
            if delta > 0: self.zoom_factor += 0.1
            else: self.zoom_factor -= 0.1
            # Clamp zoom between 50% and 500%
            self.zoom_factor = max(0.5, min(self.zoom_factor, 5.0))
            
        elif modifiers == Qt.KeyboardModifier.AltModifier:
            # Pan Left/Right
            if delta > 0: self.pan_x += 40
            else: self.pan_x -= 40
            
        elif modifiers == Qt.KeyboardModifier.ShiftModifier:
            # Pan Up/Down
            if delta > 0: self.pan_y += 40
            else: self.pan_y -= 40

        self.update()

    def paintEvent(self, event):
        if self.current_pixmap.isNull():
            return
            
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

        # Calculate base dimensions preserving aspect ratio
        widget_w, widget_h = self.width(), self.height()
        pix_w, pix_h = self.current_pixmap.width(), self.current_pixmap.height()
        
        scale = min(widget_w / pix_w, widget_h / pix_h)
        base_w, base_h = pix_w * scale, pix_h * scale

        # Apply Zoom Factor
        current_w = base_w * self.zoom_factor
        current_h = base_h * self.zoom_factor

        # Center Anchor + Apply Pan offsets
        x_offset = (widget_w - current_w) / 2 + self.pan_x
        y_offset = (widget_h - current_h) / 2 + self.pan_y

        self.pixmap_rect = QRect(int(x_offset), int(y_offset), int(current_w), int(current_h))
        
        # 1. Draw the Video Frame scaled and positioned
        painter.drawPixmap(self.pixmap_rect, self.current_pixmap)

        # 2. Draw the ROI Boxes mapping to the zoomed coordinates
        for roi in self.rois:
            nx, ny, nw, nh = roi['rect']
            rx = int(x_offset + nx * current_w)
            ry = int(y_offset + ny * current_h)
            rw = int(nw * current_w)
            rh = int(nh * current_h)

            is_selected = (roi['id'] == self.selected_id)

            if is_selected:
                painter.setPen(QPen(QColor(0, 255, 0), 3))
                painter.setBrush(QColor(0, 255, 0, 40)) 
            else:
                painter.setPen(QPen(QColor(255, 50, 50), 2))
                painter.setBrush(Qt.BrushStyle.NoBrush)

            painter.drawRect(rx, ry, rw, rh)

            # Name Tag Background
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRect(rx, ry - 20, 80, 20)

            # Name Tag Text
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rx + 5, ry - 5, roi['name'])

            # Draw "Resize Handle" if selected
            if is_selected:
                painter.setBrush(QColor(0, 255, 0))
                painter.drawRect(rx + rw - 12, ry + rh - 12, 12, 12)

    def mousePressEvent(self, event):
        if self.pixmap_rect.isNull() or not self.pixmap_rect.contains(event.pos()):
            self.selected_id = None
            self.update()
            return
            
        mx, my = event.pos().x(), event.pos().y()
        px, py, pw, ph = self.pixmap_rect.x(), self.pixmap_rect.y(), self.pixmap_rect.width(), self.pixmap_rect.height()
        
        clicked_roi = None
        for roi in reversed(self.rois):
            nx, ny, nw, nh = roi['rect']
            rx, ry = px + nx * pw, py + ny * ph
            rw, rh = nw * pw, nh * ph
            
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
                
        if not clicked_roi: self.selected_id = None
        self.update()

    def mouseMoveEvent(self, event):
        if self.selected_id is None or self.drag_state is None:
            return
            
        mx, my = event.pos().x(), event.pos().y()
        dx = mx - self.last_mouse_pos[0]
        dy = my - self.last_mouse_pos[1]
        self.last_mouse_pos = (mx, my)
        
        pw, ph = self.pixmap_rect.width(), self.pixmap_rect.height()
        dnx, dny = dx / pw, dy / ph
        
        for roi in self.rois:
            if roi['id'] == self.selected_id:
                nx, ny, nw, nh = roi['rect']
                if self.drag_state == 'move':
                    nx = max(0.0, min(nx + dnx, 1.0 - nw))
                    ny = max(0.0, min(ny + dny, 1.0 - nh))
                    roi['rect'] = [nx, ny, nw, nh]
                elif self.drag_state == 'resize':
                    nw = max(0.05, min(nw + dnx, 1.0 - nx)) 
                    nh = max(0.05, min(nh + dny, 1.0 - ny)) 
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

    def run(self):
        self.running = True
        with mss.mss() as sct:
            cap = None
            while self.running:
                if self._new_source_requested:
                    if cap is not None:
                        cap.release()
                        cap = None
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
                    
                    if self.ocr_enabled:
                        self.ocr_counter += 1
                        if self.ocr_counter >= 30: 
                            json_results = {}
                            if not self.active_rois:
                                gray = cv2.cvtColor(frame, cv2.COLOR_BGRA2GRAY)
                                text = pytesseract.image_to_string(gray).strip()
                                if text: json_results["Full_Screen"] = text
                            else:
                                fh, fw = frame.shape[:2]
                                for roi in self.active_rois:
                                    nx, ny, nw, nh = roi['rect']
                                    x, y = int(nx * fw), int(ny * fh)
                                    w, h = int(nw * fw), int(nh * fh)
                                    
                                    crop = frame[y:y+h, x:x+w]
                                    if crop.size > 0:
                                        gray = cv2.cvtColor(crop, cv2.COLOR_BGRA2GRAY)
                                        text = pytesseract.image_to_string(gray).strip()
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
        self.setWindowTitle("Pro OCR - Zoom & Pan Canvas")
        self.setMinimumSize(1300, 850)
        self.setStyleSheet("QMainWindow { background-color: #1a1a1a; } QLabel { color: #eee; }")
        
        self.engine = CaptureEngine()
        self.engine.frame_signal.connect(self.update_preview)
        self.engine.ocr_signal.connect(self.update_ocr_text)
        
        self.init_ui()
        
        # --- Keyboard Shortcuts ---
        self.shortcut_reset = QShortcut(QKeySequence("Ctrl+R"), self)
        self.shortcut_reset.activated.connect(self.preview_overlay.reset_view)

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- LEFT PANEL ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

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

        # ROI Controls
        left_layout.addWidget(QLabel("<b>REGION CONTROLS (Ctrl+Scroll = Zoom | Alt+Scroll = Pan)</b>"))
        roi_buttons_layout = QHBoxLayout()
        
        self.btn_add_roi = QPushButton("+ Add Area")
        self.btn_add_roi.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 8px;")
        roi_buttons_layout.addWidget(self.btn_add_roi)
        
        self.btn_remove_roi = QPushButton("- Remove Selected")
        self.btn_remove_roi.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold; padding: 8px;")
        roi_buttons_layout.addWidget(self.btn_remove_roi)
        
        left_layout.addLayout(roi_buttons_layout)

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

        # --- RIGHT PANEL ---
        self.preview_overlay = ROIOverlayWidget()
        self.preview_overlay.setStyleSheet("background-color: #000; border-left: 2px solid #333;")
        
        self.btn_add_roi.clicked.connect(self.preview_overlay.add_roi)
        self.btn_remove_roi.clicked.connect(self.preview_overlay.remove_selected_roi)
        self.preview_overlay.rois_changed.connect(self.engine.update_rois)

        splitter.addWidget(left_panel)
        splitter.addWidget(self.preview_overlay)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter)

    def handle_source_change(self, index):
        source = self.combo_source.currentText()
        self.screen_widget.hide()

        if source == "Open a Video File":
            path, _ = QFileDialog.getOpenFileName(self, "Select Video", "", "Video Files (*.mp4 *.avi *.mkv)")
            if path:
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
            self.engine.set_source_screen(title)
            if not self.engine.isRunning(): self.engine.start()

    def toggle_ocr_logic(self):
        state = self.btn_ocr.isChecked()
        self.engine.ocr_enabled = state
        self.btn_ocr.setText("STOP OCR DETECTION" if state else "START OCR DETECTION")
        self.btn_ocr.setStyleSheet(f"background-color: {'#c0392b' if state else '#333'}; color: white; font-weight: bold; margin-top: 15px;")

    def update_preview(self, frame):
        h, w, c = frame.shape
        q_img = QImage(frame.data, w, h, w*c, QImage.Format.Format_RGBA8888).rgbSwapped()
        self.preview_overlay.set_frame(QPixmap.fromImage(q_img))

    def update_ocr_text(self, text):
        self.ocr_output.append(f"{text}\n")
        self.ocr_output.verticalScrollBar().setValue(self.ocr_output.verticalScrollBar().maximum())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = OCRApp()
    win.show()
    sys.exit(app.exec())