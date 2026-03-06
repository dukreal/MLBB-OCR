import os
# Set DPI awareness before any imports to ensure sharp scaling
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTOSCREENSCALEFACTOR"] = "1"

import sys
import cv2
import mss
import numpy as np
import pytesseract
import pygetwindow as gw
import ctypes
from ctypes import wintypes
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                             QComboBox, QFrame, QSplitter, QFileDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QImage, QPixmap

# ==========================================================
# SET YOUR TESSERACT PATH HERE
# ==========================================================
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Windows Native GDI/User32 bindings for TRUE Window Capture
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

    def set_source_video(self, path):
        self.source_type = "video"
        self.source_path = path
        self._new_source_requested = True

    def set_source_screen(self, title):
        self.source_type = "screen"
        self.source_path = title
        self._new_source_requested = True

    def capture_window_direct(self, hwnd):
        """
        True Window Capture: Grabs the window's internal render buffer.
        This ignores any windows sitting on top of the target window.
        """
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        
        if w <= 0 or h <= 0:
            return None
            
        # Create Device Contexts
        hwndDC = user32.GetWindowDC(hwnd)
        mfcDC = gdi32.CreateCompatibleDC(hwndDC)
        saveBitMap = gdi32.CreateCompatibleBitmap(hwndDC, w, h)
        gdi32.SelectObject(mfcDC, saveBitMap)
        
        # PW_RENDERFULLCONTENT = 2 (Forces hardware-accelerated windows to render)
        result = user32.PrintWindow(hwnd, mfcDC, 2)
        
        if result == 0:
            user32.ReleaseDC(hwnd, hwndDC)
            gdi32.DeleteDC(mfcDC)
            gdi32.DeleteObject(saveBitMap)
            return None
            
        # Extract the Bitmap
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = w
        bmi.bmiHeader.biHeight = -h # Negative height to prevent upside-down image
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0
        
        buffer = ctypes.create_string_buffer(w * h * 4)
        gdi32.GetDIBits(mfcDC, saveBitMap, 0, h, buffer, ctypes.byref(bmi), 0)
        
        # Cleanup Memory
        user32.ReleaseDC(hwnd, hwndDC)
        gdi32.DeleteDC(mfcDC)
        gdi32.DeleteObject(saveBitMap)
        
        # Convert to Numpy Array and force alpha channel to be opaque
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

                # --- MODE 1: VIDEO ---
                if self.source_type == "video" and self.source_path:
                    if cap is None:
                        cap = cv2.VideoCapture(self.source_path)
                    
                    ret, v_frame = cap.read()
                    if not ret:
                        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                        continue
                    frame = cv2.cvtColor(v_frame, cv2.COLOR_BGR2BGRA)

                # --- MODE 2: SCREEN / WINDOW ---
                elif self.source_type == "screen" and self.source_path:
                    try:
                        wins = gw.getWindowsWithTitle(self.source_path)
                        if wins:
                            win = wins[0]
                            if not win.isMinimized:
                                # 1. True Window Capture (No Mirror Effect)
                                frame = self.capture_window_direct(win._hWnd)
                                
                                # 2. Fallback to Screen Capture if PrintWindow blocked
                                if frame is None:
                                    monitor = {"top": win.top, "left": win.left, "width": win.width, "height": win.height}
                                    screenshot = sct.grab(monitor)
                                    frame = np.array(screenshot)
                    except:
                        pass

                # --- OUTPUT AND OCR ---
                if frame is not None:
                    self.frame_signal.emit(frame)
                    
                    if self.ocr_enabled:
                        self.ocr_counter += 1
                        if self.ocr_counter >= 30: # OCR every ~30 frames (1 sec)
                            gray = cv2.cvtColor(frame, cv2.COLOR_BGRA2GRAY)
                            text = pytesseract.image_to_string(gray)
                            if text.strip():
                                self.ocr_signal.emit(text.strip())
                            self.ocr_counter = 0
                
                self.msleep(30) # ~30 FPS

            if cap: cap.release()

    def stop(self):
        self.running = False

class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pro OCR - Unified Live Scene")
        self.setMinimumSize(1200, 800)
        self.setStyleSheet("QMainWindow { background-color: #1a1a1a; } QLabel { color: #eee; }")
        
        self.engine = CaptureEngine()
        self.engine.frame_signal.connect(self.update_preview)
        self.engine.ocr_signal.connect(self.update_ocr_text)
        
        self.init_ui()

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

        self.btn_ocr = QPushButton("START OCR DETECTION")
        self.btn_ocr.setCheckable(True)
        self.btn_ocr.setFixedHeight(50)
        self.btn_ocr.setStyleSheet("background-color: #333; color: white; font-weight: bold;")
        self.btn_ocr.clicked.connect(self.toggle_ocr_logic)
        left_layout.addWidget(self.btn_ocr)

        left_layout.addWidget(QLabel("<b>EXTRACTED TEXT</b>"))
        self.ocr_output = QTextEdit()
        self.ocr_output.setReadOnly(True)
        self.ocr_output.setStyleSheet("background-color: #000; color: #00ff41; font-family: Consolas; font-size: 13px;")
        left_layout.addWidget(self.ocr_output)

        # --- RIGHT PANEL ---
        self.right_container = QFrame()
        self.right_container.setStyleSheet("background-color: #000; border-left: 2px solid #333;")
        right_layout = QVBoxLayout(self.right_container)
        
        self.preview_label = QLabel("LIVE SCENE DISPLAY")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet("color: #444; font-size: 20px; font-weight: bold;")
        right_layout.addWidget(self.preview_label)

        splitter.addWidget(left_panel)
        splitter.addWidget(self.right_container)
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
        self.btn_ocr.setStyleSheet(f"background-color: {'#c0392b' if state else '#333'}; color: white; font-weight: bold;")

    def update_preview(self, frame):
        h, w, c = frame.shape
        q_img = QImage(frame.data, w, h, w*c, QImage.Format.Format_RGBA8888).rgbSwapped()
        pixmap = QPixmap.fromImage(q_img)
        scaled = pixmap.scaled(self.preview_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.preview_label.setPixmap(scaled)

    def update_ocr_text(self, text):
        self.ocr_output.append(f"> {text}")
        self.ocr_output.verticalScrollBar().setValue(self.ocr_output.verticalScrollBar().maximum())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = OCRApp()
    win.show()
    sys.exit(app.exec())