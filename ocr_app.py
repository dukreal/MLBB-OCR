import os
import re
os.environ["OMP_THREAD_LIMIT"] = "1"

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
import threading

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                             QComboBox, QFrame, QFileDialog, QScrollArea, QSlider, 
                             QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QGridLayout, QMessageBox,
                             QLineEdit)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRect, QPoint
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

# ==========================================================
# OCR VALIDATION LAYER
# ==========================================================
class OCRValidator:
    """
    Validates OCR-read values for specific game fields before they are written
    to the UI or the live_overlay_data.json export.

    Rules per field (matched by substring, case-insensitive):
        min_val        - absolute floor; values below this are rejected
        max_val        - absolute ceiling; values above this are rejected
        max_jump       - maximum single-step INCREASE allowed from the last
                         accepted value (None = no jump limit)
        max_drop       - maximum single-step DECREASE allowed (default: same
                         as max_jump). Gold should never drop much — set tight.
        min_floor_pct  - if set (0.0–1.0), a new value below this fraction of
                         the last accepted value is rejected outright as a
                         catastrophic OCR drop (e.g. 5084 → 114).
        confirm_needed - how many consecutive OCR readings of a suspicious
                         (large-jump) value are required before it is accepted
                         as a real change.  Defaults to CONFIRM_NEEDED.
        block_m_suffix - if True, any raw value ending in 'm'/'M' is rejected
                         outright (millions impossible in-match for gold).

    Values with a trailing 'k'/'K' (e.g. "20k") are expanded to their numeric
    equivalent (20000) before all checks.

    HOW THE CONFIRMATION WINDOW WORKS
    ----------------------------------
    When a new value exceeds max_jump it is NOT immediately accepted or
    permanently rejected.  Instead it enters a "pending" state and a counter
    starts.  Every subsequent OCR cycle is checked:

      - Same suspicious value again?  Counter increments.
        Once counter reaches confirm_needed the value is accepted and becomes
        the new baseline.  This handles genuine large jumps (e.g. player
        suddenly earns a lot of gold after a teamfight) that the OCR reads
        consistently.

      - Different value?  Counter resets.  A one-off OCR spike (noise) will
        only appear for one or two frames, so it never accumulates enough
        confirmations to be accepted.

    IMPROVEMENTS OVER V1
    ---------------------
    - Separate max_drop limit: catches the 5084→114 type crash instantly
      without needing 3 confirmations (drops don't need confirmation — a
      real gold drop in MLBB is always small: death penalty ~300 gold max).
    - min_floor_pct: secondary hard block — if value drops below X% of last
      accepted, reject immediately regardless of max_drop (catches partial
      digit reads like "114" when real value is "5084").
    - Pending window now uses a tolerance band instead of exact match,
      so small OCR jitter on a genuinely large jump still confirms correctly.
    - Confirmed large jumps update the baseline immediately so subsequent
      normal increments don't re-trigger the jump check.
    - All rejections logged with reason tag for easier debugging.
    """

    # How many consecutive identical suspicious readings are needed
    # before a large-jump INCREASE is accepted as a real change.
    CONFIRM_NEEDED = 3

    # Field rules: key is a substring that must appear in the ROI name
    # (case-insensitive).  First matching rule wins.
    FIELD_RULES = {
        # Kills — max_jump=8 allows catching up after missed frames,
        # confirm_needed=3 confirms quickly, confirm_drop handles bad baselines
        "blue-kill": {"min_val": 0, "max_val": 40, "max_jump": 8, "max_drop": 3, "confirm_first": True, "confirm_needed": 3, "confirm_drop": True},
        "red-kill":  {"min_val": 0, "max_val": 40, "max_jump": 8, "max_drop": 3, "confirm_first": True, "confirm_needed": 3, "confirm_drop": True},

        # Gold — millions impossible in-match; large drops are always OCR noise
        # Death penalty in MLBB is ~300 gold max, so max_drop is tight.
        # min_floor_pct=0.5 means: if new value < 50% of last, reject instantly.
        "blue-gold": {
            "min_val": 0, "max_val": 999_999,
            "max_jump": 2_500, "max_drop": 400,
            "min_floor_pct": 0.1,  # only block if drops below 10% of last value
            "block_m_suffix": True,
            "confirm_needed": 5,
            "confirm_drop": True,
        },
        "red-gold": {
            "min_val": 0, "max_val": 999_999,
            "max_jump": 2_500, "max_drop": 400,
            "min_floor_pct": 0.1,  # only block if drops below 10% of last value
            "block_m_suffix": True,
            "confirm_needed": 5,
            "confirm_drop": True,
        },

        # Lord kills — confirm_first prevents misread locking baseline from 0
        "lord-blue":  {"min_val": 0, "max_val": 10, "max_jump": 10, "max_drop": 10, "confirm_first": True, "confirm_needed": 8},
        "lord-red":   {"min_val": 0, "max_val": 10, "max_jump": 10, "max_drop": 10, "confirm_first": True, "confirm_needed": 8},

        # Tower kills — max 9 in MLBB, only go up, confirm_first prevents misread locking baseline
        "tower-blue": {"min_val": 0, "max_val": 9, "max_jump": 9, "max_drop": 0, "confirm_first": True, "confirm_needed": 8},
        "tower-red":  {"min_val": 0, "max_val": 9, "max_jump": 9, "max_drop": 0, "confirm_first": True, "confirm_needed": 8},
    }

    def __init__(self):
        # Last accepted numeric value per field
        self._last_values:   dict[str, float] = {}
        # Last accepted raw string per field (returned while pending)
        self._last_raw:      dict[str, str]   = {}
        # Pending suspicious value waiting for confirmation  {field: numeric}
        self._pending_value: dict[str, float] = {}
        # Pending raw string (so we can return it once confirmed)
        self._pending_raw:   dict[str, str]   = {}
        # How many times in a row the pending value has been seen
        self._pending_count: dict[str, int]   = {}
        # Persistent correction: tracks how many frames a consistent lower
        # value has been seen — if it exceeds CORRECTION_OVERRIDE it forces
        # the baseline down even if max_drop would normally block it
        self._correction_value: dict[str, float] = {}
        self._correction_count: dict[str, int]   = {}

        # KDA-specific state: last accepted (kills, deaths, assists) triple,
        # and pending-decrease tracking (mirrors the numeric pending logic above).
        self._last_kda:         dict[str, tuple[int, int, int]] = {}
        self._kda_pending:      dict[str, tuple[int, int, int]] = {}
        self._kda_pending_count: dict[str, int]                 = {}

    # Frames of consistent lower reading needed to force-correct a bad baseline
    CORRECTION_OVERRIDE = 20

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def parse_value(self, raw: str) -> float | None:
        """
        Convert a raw OCR string to a float.
        Handles suffixes: k/K -> x1000, m/M -> x1_000_000.
        Returns None if the string cannot be parsed as a number.
        """
        s = raw.strip().replace(",", "").replace(" ", "")
        if not s:
            return None
        try:
            if s[-1].lower() == "k":
                return float(s[:-1]) * 1_000
            if s[-1].lower() == "m":
                return float(s[:-1]) * 1_000_000
            return float(s)
        except (ValueError, IndexError):
            return None

    def _get_rule(self, field_name: str) -> dict | None:
        """Return the first matching rule for field_name, or None."""
        lower = field_name.lower()
        for key, rule in self.FIELD_RULES.items():
            if key in lower:
                return rule
        return None

    def _clear_pending(self, field_name: str) -> None:
        """Wipe any pending-confirmation state for a field."""
        self._pending_value.pop(field_name, None)
        self._pending_raw.pop(field_name,   None)
        self._pending_count.pop(field_name, None)
        self._correction_value.pop(field_name, None)
        self._correction_count.pop(field_name, None)

    # ------------------------------------------------------------------
    # KDA validation (shape check + slash-repair + never-decreases)
    # ------------------------------------------------------------------
    def _parse_kda_shape(self, raw: str):
        """
        Turn an OCR string into a (kills, deaths, assists) triple, or None
        if it can't be made to fit that shape.

        Handles the common failure where the SECOND slash is missed and two
        numbers run together, e.g. "6/519" (meant "6/5/19"). We try every
        way of splitting the merged digits and keep the first split where
        both halves look like plausible KDA numbers.
        """
        raw = raw.strip()

        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{1,2})", raw)
        if m:
            return tuple(int(g) for g in m.groups())

        # Exactly one slash found -> the second slash was likely lost.
        m = re.fullmatch(r"(\d{1,2})/(\d{2,4})", raw)
        if m:
            kills, blob = m.group(1), m.group(2)
            for i in range(1, len(blob)):
                d_str, a_str = blob[:i], blob[i:]
                if d_str.startswith("0") and d_str != "0":
                    continue
                if a_str.startswith("0") and a_str != "0":
                    continue
                deaths, assists = int(d_str), int(a_str)
                if 0 <= deaths <= 30 and 0 <= assists <= 30:
                    return (int(kills), deaths, assists)

        return None

    def validate_kda(self, field_name: str, new_raw: str, current_display: str) -> str:
        """
        Validate a KDA read: must fit (or be repairable to) kills/deaths/assists,
        and none of the three numbers may decrease during a match unless a
        lower reading repeats CONFIRM_NEEDED times in a row (handles a
        genuine baseline correction, same idea as the numeric fields above).
        """
        triple = self._parse_kda_shape(new_raw)
        if triple is None:
            print(f"[Validator] [{field_name}] KDA BAD SHAPE '{new_raw}' — keeping '{current_display}'")
            self._kda_pending.pop(field_name, None)
            self._kda_pending_count.pop(field_name, None)
            return current_display

        last = self._last_kda.get(field_name)
        if last is not None and any(new_v < last_v for new_v, last_v in zip(triple, last)):
            prev_pending = self._kda_pending.get(field_name)
            if prev_pending == triple:
                count = self._kda_pending_count.get(field_name, 1) + 1
            else:
                count = 1
            self._kda_pending[field_name] = triple
            self._kda_pending_count[field_name] = count

            if count >= self.CONFIRM_NEEDED:
                print(f"[Validator] [{field_name}] CONFIRMED KDA decrease {last} -> {triple} after {count} readings")
                self._last_kda[field_name] = triple
                self._kda_pending.pop(field_name, None)
                self._kda_pending_count.pop(field_name, None)
                return "/".join(str(v) for v in triple)

            print(f"[Validator] [{field_name}] BLOCKED KDA decrease {last} -> {triple} ({count}/{self.CONFIRM_NEEDED})")
            return current_display

        self._last_kda[field_name] = triple
        self._kda_pending.pop(field_name, None)
        self._kda_pending_count.pop(field_name, None)
        return "/".join(str(v) for v in triple)

    # ------------------------------------------------------------------
    # Main entry point
    # ------------------------------------------------------------------
    def validate(self, field_name: str, new_raw: str, current_display: str) -> str:
        """
        Validate new_raw (freshly OCR-read) for field_name.

        Returns:
            new_raw          - value passed all checks; accepted immediately
            current_display  - value failed a hard check (range/suffix/drop);
                               keep whatever is currently shown
            last valid raw   - value failed the jump check but is now
                               confirmed after CONFIRM_NEEDED repeats
        """
        rule = self._get_rule(field_name)
        if rule is None:
            return new_raw  # no rule — pass through

        # ---- 1. Parse ----
        new_num = self.parse_value(new_raw)
        if new_num is None:
            print(f"[Validator] [{field_name}] UNPARSEABLE '{new_raw}' — keeping '{current_display}'")
            self._clear_pending(field_name)
            return current_display

        # ---- 2. Hard block: million suffix ----
        if rule.get("block_m_suffix"):
            stripped = new_raw.strip().replace(",", "").replace(" ", "")
            if stripped and stripped[-1].lower() == "m":
                print(f"[Validator] [{field_name}] BLOCKED million-suffix '{new_raw}' (impossible in-match)")
                self._clear_pending(field_name)
                return current_display

        # ---- 3. Hard block: absolute range ----
        min_v, max_v = rule.get("min_val"), rule.get("max_val")
        if min_v is not None and new_num < min_v:
            print(f"[Validator] [{field_name}] BLOCKED {new_num} < min {min_v}")
            self._clear_pending(field_name)
            return current_display
        if max_v is not None and new_num > max_v:
            print(f"[Validator] [{field_name}] BLOCKED {new_num} > max {max_v}")
            self._clear_pending(field_name)
            return current_display

        # ---- 4. Drop checks ----
        if field_name in self._last_values:
            prev = self._last_values[field_name]
            drop = prev - new_num  # positive = value went down

            if drop > 0:
                # 4a. Floor percentage check — instant reject, no confirmation
                # e.g. 5084 → 114 (only 2.2% of previous — always noise)
                floor_pct = rule.get("min_floor_pct")
                if floor_pct is not None and prev > 0 and new_num < prev * floor_pct:
                    print(f"[Validator] [{field_name}] BLOCKED drop {prev}→{new_num} "
                          f"({new_num/prev*100:.1f}% of prev, floor={floor_pct*100:.0f}%)")
                    self._clear_pending(field_name)
                    return self._last_raw.get(field_name, current_display)

                # 4b. Max drop check
                max_drop = rule.get("max_drop")
                if max_drop is not None and drop > max_drop:
                    # Check if this field allows drop confirmation
                    # Fields with confirm_drop=True (e.g. gold) get a confirmation
                    # window before rejecting — large real drops can happen (e.g. new game)
                    # Fields with confirm_drop=False (e.g. towers, kills) reject instantly
                    if rule.get("confirm_drop"):
                        needed       = rule.get("confirm_needed", self.CONFIRM_NEEDED)
                        prev_pending = self._pending_value.get(field_name)
                        tolerance    = max_drop * 0.10

                        if prev_pending is not None and abs(new_num - prev_pending) <= tolerance:
                            self._pending_count[field_name] = self._pending_count.get(field_name, 1) + 1
                            self._pending_raw[field_name]   = new_raw
                        else:
                            self._pending_value[field_name] = new_num
                            self._pending_raw[field_name]   = new_raw
                            self._pending_count[field_name] = 1

                        count = self._pending_count[field_name]

                        if count >= needed:
                            confirmed_raw = self._pending_raw[field_name]
                            confirmed_num = self.parse_value(confirmed_raw) or new_num
                            print(f"[Validator] [{field_name}] CONFIRMED large drop "
                                  f"{prev}→{confirmed_num} after {count} readings — accepted")
                            self._last_values[field_name] = confirmed_num
                            self._last_raw[field_name]    = confirmed_raw
                            self._clear_pending(field_name)
                            return confirmed_raw
                        else:
                            print(f"[Validator] [{field_name}] PENDING drop {prev}→{new_num} "
                                  f"(drop={drop:.0f} > {max_drop}, {count}/{needed} confirmations)")
                            return self._last_raw.get(field_name, current_display)
                    else:
                        # Instant reject normally — but track correction counter.
                        # If OCR keeps reading the same lower value for
                        # CORRECTION_OVERRIDE frames, force-accept it as the
                        # baseline was likely wrong (e.g. 1 misread as 6).
                        prev_corr = self._correction_value.get(field_name)
                        tolerance = 0.5
                        if prev_corr is not None and abs(new_num - prev_corr) <= tolerance:
                            self._correction_count[field_name] = self._correction_count.get(field_name, 1) + 1
                        else:
                            self._correction_value[field_name] = new_num
                            self._correction_count[field_name] = 1

                        corr_count = self._correction_count[field_name]
                        if corr_count >= self.CORRECTION_OVERRIDE:
                            print(f"[Validator] [{field_name}] FORCE-CORRECTED bad baseline "
                                  f"{prev}→{new_num} after {corr_count} consistent readings")
                            self._last_values[field_name] = new_num
                            self._last_raw[field_name]    = new_raw
                            self._correction_value.pop(field_name, None)
                            self._correction_count.pop(field_name, None)
                            self._clear_pending(field_name)
                            return new_raw

                        print(f"[Validator] [{field_name}] BLOCKED drop {prev}→{new_num} "
                              f"(drop={drop}, max_drop={max_drop}, correction {corr_count}/{self.CORRECTION_OVERRIDE})")
                        self._clear_pending(field_name)
                        return self._last_raw.get(field_name, current_display)

        # ---- 5. First-value confirmation ----
        # If confirm_first is set and we have no baseline yet, require
        # CONFIRM_NEEDED readings before accepting the very first non-zero value.
        # This prevents a single misread from permanently locking the baseline.
        if rule.get("confirm_first") and field_name not in self._last_values and new_num != 0:
            needed       = rule.get("confirm_needed", self.CONFIRM_NEEDED)
            prev_pending = self._pending_value.get(field_name)
            tolerance    = 0.5  # allow ±0.5 for integer fields

            if prev_pending is not None and abs(new_num - prev_pending) <= tolerance:
                self._pending_count[field_name] = self._pending_count.get(field_name, 1) + 1
                self._pending_raw[field_name]   = new_raw
            else:
                self._pending_value[field_name] = new_num
                self._pending_raw[field_name]   = new_raw
                self._pending_count[field_name] = 1

            count = self._pending_count[field_name]

            if count >= needed:
                confirmed_raw = self._pending_raw[field_name]
                confirmed_num = self.parse_value(confirmed_raw) or new_num
                print(f"[Validator] [{field_name}] FIRST VALUE {confirmed_num} confirmed "
                      f"after {count} readings — accepted")
                self._last_values[field_name] = confirmed_num
                self._last_raw[field_name]    = confirmed_raw
                self._clear_pending(field_name)
                return confirmed_raw
            else:
                print(f"[Validator] [{field_name}] FIRST VALUE {new_num} pending "
                      f"({count}/{needed} confirmations)")
                return self._last_raw.get(field_name, current_display)

        # ---- 6. Jump (increase) check with confirmation window ----
        max_jump = rule.get("max_jump")
        if max_jump is not None and field_name in self._last_values:
            prev  = self._last_values[field_name]
            delta = new_num - prev  # only upward jumps reach here (drops handled above)

            if delta > max_jump:
                needed       = rule.get("confirm_needed", self.CONFIRM_NEEDED)
                prev_pending = self._pending_value.get(field_name)
                # Tolerance band: ±10% of max_jump to handle minor OCR jitter
                # on a genuinely large jump (e.g. 39265, 39270, 39280 all confirm)
                tolerance = max_jump * 0.10

                if prev_pending is not None and abs(new_num - prev_pending) <= tolerance:
                    # Same suspicious value seen again — increment counter
                    self._pending_count[field_name] = self._pending_count.get(field_name, 1) + 1
                    self._pending_raw[field_name]   = new_raw
                else:
                    # Different suspicious value — restart counter
                    self._pending_value[field_name] = new_num
                    self._pending_raw[field_name]   = new_raw
                    self._pending_count[field_name] = 1

                count = self._pending_count[field_name]

                if count >= needed:
                    # Confirmed — accept and clear pending state
                    confirmed_raw = self._pending_raw[field_name]
                    confirmed_num = self.parse_value(confirmed_raw) or new_num
                    print(f"[Validator] [{field_name}] CONFIRMED large jump +{delta:.0f} "
                          f"after {count} readings ({prev}→{confirmed_num}) — accepted")
                    self._last_values[field_name] = confirmed_num
                    self._last_raw[field_name]    = confirmed_raw
                    self._clear_pending(field_name)
                    return confirmed_raw
                else:
                    # Not yet confirmed — hold previous value
                    print(f"[Validator] [{field_name}] PENDING jump +{delta:.0f} > {max_jump} "
                          f"({count}/{needed} confirmations)")
                    return self._last_raw.get(field_name, current_display)

        # ---- 6. Passed all checks ----
        self._last_values[field_name] = new_num
        self._last_raw[field_name]    = new_raw
        self._clear_pending(field_name)
        return new_raw

    def reset(self, field_name: str | None = None) -> None:
        """Clear all stored state for one field, or all fields."""
        if field_name is None:
            self._last_values.clear()
            self._last_raw.clear()
            self._pending_value.clear()
            self._pending_raw.clear()
            self._pending_count.clear()
        else:
            self._last_values.pop(field_name, None)
            self._last_raw.pop(field_name,    None)
            self._clear_pending(field_name)

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
        self.copied_roi = None

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

    def copy_selected_field(self):
        """Remember the selected field so it can be pasted later."""
        if self.selected_id is None:
            return False
        for roi in self.rois:
            if roi['id'] == self.selected_id:
                self.copied_roi = dict(roi)
                self.copied_roi['rect'] = list(roi['rect'])
                return True
        return False

    def paste_field_centered(self):
        """Paste the copied field, centered in the part of the canvas you can currently see."""
        if not self.copied_roi:
            return None
        src = self.copied_roi

        # The visible working area = the scroll area's viewport, mapped into this widget.
        vp = self.scroll_area.viewport()
        visible = QRect(self.mapFrom(vp, QPoint(0, 0)), vp.size()).intersected(self.rect())
        center = visible.center() if not visible.isEmpty() else self.rect().center()

        w, h = max(1, self.width()), max(1, self.height())
        nw, nh = src['rect'][2], src['rect'][3]
        nx = max(0.0, min(center.x() / w - nw / 2, 1.0 - nw))
        ny = max(0.0, min(center.y() / h - nh / 2, 1.0 - nh))

        # Names are the keys in the JSON output, so make the pasted name unique.
        existing = {r['name'] for r in self.rois}
        base = f"{src['name']} copy" if src['name'] else "Area copy"
        new_name, n = base, 2
        while new_name in existing:
            new_name = f"{base} {n}"
            n += 1

        new_roi = dict(src)
        new_roi['id'] = self.area_counter
        new_roi['name'] = new_name
        new_roi['rect'] = [nx, ny, nw, nh]
        new_roi['is_on_scene'] = True
        self.rois.append(new_roi)
        self.selected_id = new_roi['id']
        self.area_counter += 1
        self.update()
        self.rois_changed.emit(self.rois)
        self.roi_selected.emit(self.selected_id)
        return new_roi['id']

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
        
        # --- FPS Tracker Variables ---
        self.fps_start_time = time.time()
        self.fps_counter = 0
        self.current_fps = 0

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

    def set_source_capture_card(self, device_index: int, device_name: str = ""):
        self.source_type = "capture_card"
        self.source_device_name = device_name
        self.source_path = device_index
        print(f"[CaptureCard] Selected '{device_name}' → index {device_index}")
        self._new_source_requested = True

    def update_rois(self, rois):
        self.active_rois = rois

    def _is_blank_frame(self, img):
        """
        Returns True if the captured frame is blank/black.
        Uses mean of the RGB channels (ignores alpha) for a robust check.
        A single bright pixel on a dark border can fool np.max(), so we use mean.
        Threshold of 8 means >96.9% of pixels must be near-black to be considered blank.
        """
        if img is None:
            return True
        return float(np.mean(img[:, :, :3])) < 8.0

    def capture_window_direct(self, hwnd):
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        w, h = rect.right - rect.left, rect.bottom - rect.top
        if w <= 0 or h <= 0: return None

        hwndDC = user32.GetWindowDC(hwnd)
        mfcDC = gdi32.CreateCompatibleDC(hwndDC)
        saveBitMap = gdi32.CreateCompatibleBitmap(hwndDC, w, h)
        gdi32.SelectObject(mfcDC, saveBitMap)

        # PW_RENDERFULLCONTENT (flag=3): captures GPU/DX/OpenGL content on Win8.1+
        # This is the best option for emulators, but many (MuMu, BlueStacks, LDPlayer)
        # still return a black bitmap because they composite entirely on the GPU.
        result = user32.PrintWindow(hwnd, mfcDC, 3)
        if result == 0:
            # Flag 2 = client area only (legacy fallback)
            result = user32.PrintWindow(hwnd, mfcDC, 2)

        img = None
        if result != 0:
            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biWidth = w
            bmi.bmiHeader.biHeight = -h
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0
            buffer = ctypes.create_string_buffer(w * h * 4)
            gdi32.GetDIBits(mfcDC, saveBitMap, 0, h, buffer, ctypes.byref(bmi), 0)
            img = np.frombuffer(buffer, dtype=np.uint8).reshape((h, w, 4)).copy()
            img[:, :, 3] = 255

        user32.ReleaseDC(hwnd, hwndDC)
        gdi32.DeleteDC(mfcDC)
        gdi32.DeleteObject(saveBitMap)

        # If PrintWindow gave us a blank frame (common for GPU-rendered emulators),
        # return None so the caller knows to fall back to MSS screen grab.
        if self._is_blank_frame(img):
            return None

        return img

    def preprocess_image(self, crop, threshold_val, thickness_val, roi_type=None):
        gray = cv2.cvtColor(crop, cv2.COLOR_BGRA2GRAY)

        # KDA text is tiny, so enlarge it before OCR (Tesseract reads small text badly).
        scale = 1
        if roi_type in ('KDA (K/D/A)', 'Gold Amount (K/M)'):
            scale = 4
            gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_LANCZOS4)

        if threshold_val <= 1:
            _, processed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        else:
            thresh_calc = int(255 * (threshold_val / 10.0))
            _, processed = cv2.threshold(gray, thresh_calc, 255, cv2.THRESH_BINARY)
            
        thick_calc = thickness_val - 5
        if thick_calc != 0:
            k_size = (abs(thick_calc) + 1) * scale    # keep the slider's meaning after enlarging
            kernel = np.ones((k_size, k_size), np.uint8)
            if thick_calc > 0:
                processed = cv2.erode(processed, kernel, iterations=1)
            else:
                processed = cv2.dilate(processed, kernel, iterations=1)

        if roi_type in ('KDA (K/D/A)', 'Gold Amount (K/M)'):
            # Make the text black on white(the minority colour is the text) ...
            if cv2.countNonZero(processed) < processed.size / 2:
                processed = cv2.bitwise_not(processed)
            # ... and add a white margin, which Tesseract needs for single-line reads.
            pad = 8 * scale
            processed = cv2.copyMakeBorder(processed, pad, pad, pad, pad,
                                           cv2.BORDER_CONSTANT, value=255)
        return processed

    def _process_single_roi(self, roi, processed_img):
        """OCR using image_to_data so we get a confidence score per read."""
        config = "--psm 7"
        if roi['type'] == 'Numbers Only':
            config = "-c tessedit_char_whitelist=0123456789 --psm 7"
        elif roi['type'] == 'Time Format':
            config = "-c tessedit_char_whitelist=0123456789:. --psm 7"
        elif roi['type'] == 'KDA (K/D/A)':
            config = "-c tessedit_char_whitelist=0123456789/ --psm 7"
        elif roi['type'] == 'Gold Amount (K/M)':
            config = "-c tessedit_char_whitelist=0123456789.kK --psm 7"

        data = pytesseract.image_to_data(
            processed_img, config=config,
            output_type=pytesseract.Output.DICT
        )

        # Join all recognized words in the ROI, and take the lowest
        # confidence among them as the ROI's overall confidence.
        # Tesseract returns -1 for confidence on words it couldn't score.
        words = []
        confidences = []
        for word, conf in zip(data['text'], data['conf']):
            word = word.strip()
            if word:
                words.append(word)
                conf = int(conf)
                if conf >= 0:
                    confidences.append(conf)

        text = " ".join(words)
        min_conf = min(confidences) if confidences else 0

        return roi, text, min_conf

    def _run_ocr_background(self, rois_copy, previews_copy, ocr_start_time, current_fps):
        """Processes all ROIs concurrently and emits the current video FPS."""
        extracted_data = {}
        areas_scanned = 0
        
        active_rois = [roi for roi in rois_copy if roi['id'] in previews_copy]

        # Minimum confidence (0-100) a reading must meet to be kept.
        # Falls back to 0 (accept everything) if a ROI has no 'confidence' key yet.
        if active_rois:
            with concurrent.futures.ThreadPoolExecutor(max_workers=len(active_rois)) as executor:
                futures =[]
                for roi in active_rois:
                    areas_scanned += 1
                    processed = previews_copy[roi['id']]
                    futures.append(executor.submit(self._process_single_roi, roi, processed))
                
                for future in concurrent.futures.as_completed(futures):
                    roi, text, min_conf = future.result()
                    if text:
                        # The "Conf. Th" slider is stored 1-10 in the UI;
                        # scale it to a rough 0-100 confidence floor.
                        conf_floor = roi.get('confidence', 1) * 10
                        if min_conf < conf_floor:
                            print(f"[OCR] [{roi.get('name')}] LOW CONFIDENCE {min_conf} < {conf_floor} — discarded '{text}'")
                            continue
                        safe_name = roi['name'] if roi['name'] else f"Area_{roi['id']}"
                        extracted_data[safe_name] = text

        # --- RAW OCR LOG (pre-validation) ---
        # Logs exactly what Tesseract returned before the OCRValidator
        # touches it, so we can measure raw OCR accuracy separately from
        # what the validator later accepts or rejects.
        try:
            if extracted_data:
                timestamp = time.strftime("%H:%M:%S")
                raw_line = f"[{timestamp}] " + " | ".join(f"{k}={v}" for k, v in extracted_data.items())
                raw_log_path = os.path.join(current_dir, "raw_ocr_log.txt")
                with open(raw_log_path, "a", encoding="utf-8") as rf:
                    rf.write(raw_line + "\n")
        except Exception:
            pass

        ocr_end_time = time.time()
        if extracted_data:
            process_ms = int((ocr_end_time - ocr_start_time) * 1000)
            
            # --- Pass current FPS in metadata ---
            metadata = {
                "process_time_ms": process_ms, 
                "areas_scanned": areas_scanned, 
                "fps": current_fps
            }
            self.ocr_signal.emit(metadata, extracted_data)

    def run(self):
        self.running = True
        with mss.mss() as sct:
            cap = None
            target_fps = 30.0 
            video_start_time = 0
            total_frames = 0
            
            # Reset FPS Trackers
            self.fps_start_time = time.time()
            self.fps_counter = 0
            self.current_fps = 0

            while self.running:
                loop_start = time.time() 

                if self._new_source_requested:
                    if cap is not None:
                        cap.release()
                        cap = None
                    # Wait for any in-progress open thread to finish
                    prev_ready = getattr(self, '_cap_ready', None)
                    if prev_ready and not prev_ready.is_set():
                        prev_ready.wait(timeout=6)
                    self._new_source_requested = False
                    self.ocr_counter = 0
                    self._cap_opening = False
                    self._cap_result = [None]
                    self._cap_ready = threading.Event()
                    self.msleep(500)

                frame = None

                # --- VIDEO SOURCE ---
                if self.source_type == "video" and self.source_path:
                    # ... (keep your video logic same as before)
                    if cap is None: 
                        cap = cv2.VideoCapture(self.source_path)
                        target_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
                        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                        video_start_time = time.time()
                    elapsed = time.time() - video_start_time
                    target_frame = int(elapsed * target_fps)
                    if target_frame >= total_frames and total_frames > 0:
                        video_start_time = time.time()
                        target_frame = 0
                    cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
                    ret, v_frame = cap.read()
                    if ret: frame = cv2.cvtColor(v_frame, cv2.COLOR_BGR2BGRA)

                # --- CAPTURE CARD SOURCE ---
                elif self.source_type == "capture_card" and self.source_path is not None:
                    if cap is None and not getattr(self, '_cap_opening', False):
                        self._cap_opening = True
                        self._cap_ready = threading.Event()
                        self._cap_result = [None]
                        src_index = self.source_path
                        def _open_cap(idx=src_index):
                            opened = None
                            time.sleep(1.0)  # wait for OS to fully release previous device
                            for backend in [cv2.CAP_MSMF, cv2.CAP_DSHOW, cv2.CAP_ANY]:
                                try:
                                    print(f"[CaptureCard] Trying index {idx} backend {backend}...")
                                    c = cv2.VideoCapture(idx, backend)
                                    is_open = c.isOpened()
                                    print(f"[CaptureCard] index {idx} backend {backend} → isOpened={is_open}")
                                    if is_open:
                                        c.set(cv2.CAP_PROP_FRAME_WIDTH,  1920)
                                        c.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
                                        c.set(cv2.CAP_PROP_FPS, 30)
                                        # warm up — read a few frames before handing off
                                        for _ in range(5):
                                            c.read()
                                        opened = c
                                        print(f"[CaptureCard] ✅ Opened index {idx} backend {backend}")
                                        break
                                    c.release()
                                    time.sleep(0.3)
                                except Exception as e:
                                    print(f"[CaptureCard] backend {backend} failed: {e}")
                            self._cap_result[0] = opened
                            self._cap_opening = False
                            self._cap_ready.set()
                        threading.Thread(target=_open_cap, daemon=True).start()

                    if cap is None and getattr(self, '_cap_ready', None) and self._cap_ready.is_set():
                        cap = self._cap_result[0]
                        if cap is None:
                            print(f"[CaptureCard] All backends failed — will not retry until source changes")
                            self.source_type = None  # stop retrying
                    if cap and cap.isOpened():
                        ret, v_frame = cap.read()
                        if ret:
                            frame = cv2.cvtColor(v_frame, cv2.COLOR_BGR2BGRA)
                        else:
                            print(f"[CaptureCard] read() returned False — device may still be warming up")
                            cap.release()
                            cap = None
                            self._cap_opening = False
                            self._cap_ready = threading.Event()
                            self._cap_result = [None]


                # --- SCREEN / EMULATOR SOURCE ---
                elif self.source_type == "screen" and self.source_path:
                    try:
                        wins = gw.getWindowsWithTitle(self.source_path)
                        if wins:
                            win = wins[0]

                            # 1. Try Direct Window Capture first (PrintWindow / GDI).
                            #    capture_window_direct() now returns None if the result
                            #    is a blank/black frame so we always fall through below
                            #    for GPU-rendered windows (emulators, DirectX apps, etc.)
                            frame = self.capture_window_direct(win._hWnd)

                            # 2. EMULATOR FIX: MSS reads directly from the composited
                            #    screen buffer — it captures what you *see* on screen,
                            #    including GPU/OpenGL/DirectX rendered content.
                            #    MuMuPlayer, BlueStacks, LDPlayer, NoxPlayer all need this.
                            #
                            #    Note: MSS requires the window to be VISIBLE and not
                            #    fully obscured.  If the window is minimised it will
                            #    return a blank grab — we skip in that case.
                            if frame is None:
                                # Restore the window if it is minimised so MSS can see it
                                try:
                                    if win.isMinimized:
                                        win.restore()
                                        self.msleep(150)   # brief pause for the OS to repaint
                                except Exception:
                                    pass

                                win_left   = win.left
                                win_top    = win.top
                                win_width  = win.width
                                win_height = win.height

                                if win_width > 0 and win_height > 0:
                                    mss_rect = {
                                        "top":    win_top,
                                        "left":   win_left,
                                        "width":  win_width,
                                        "height": win_height,
                                    }
                                    screenshot = sct.grab(mss_rect)
                                    # MSS returns BGRA — force alpha to fully opaque
                                    frame = np.array(screenshot)
                                    frame[:, :, 3] = 255

                                    # If MSS also returned blank (window hidden / off-screen)
                                    # discard so we don't feed garbage to OCR.
                                    if self._is_blank_frame(frame):
                                        frame = None
                    except Exception as e:
                        print(f"Capture Error: {e}")

                if frame is not None:
                    # FPS Calculation
                    self.fps_counter += 1
                    current_time = time.time()
                    if current_time - self.fps_start_time >= 1.0:
                        self.current_fps = self.fps_counter
                        self.fps_counter = 0
                        self.fps_start_time = current_time

                    self.frame_signal.emit(frame)
                    
                    # ROI Processing
                    previews = {}
                    fh, fw = frame.shape[:2]
                    scene_rois = [r for r in self.active_rois if r['is_on_scene']]
                    
                    for roi in scene_rois:
                        nx, ny, nw, nh = roi['rect']
                        x, y, w, h = int(nx * fw), int(ny * fh), int(nw * fw), int(nh * fh)
                        # Ensure crop is within frame boundaries
                        crop = frame[max(0, y):min(fh, y+h), max(0, x):min(fw, x+w)]
                        if crop.size > 0:
                            processed = self.preprocess_image(crop, roi['threshold'], roi['thickness'], roi['type'])
                            previews[roi['id']] = processed
                    
                    self.previews_signal.emit(previews)

                    # OCR Trigger
                    if self.ocr_enabled:
                        self.ocr_counter += 1
                        if self.ocr_counter >= 3: # Scan every ~3 frames for stability
                            self.active_futures = [f for f in self.active_futures if not f.done()]
                            if len(self.active_futures) < self.thread_count and scene_rois:
                                ocr_start_time = time.time()
                                rois_copy = [r.copy() for r in scene_rois]
                                prev_copy = {k: v.copy() for k, v in previews.items()}
                                future = self.executor.submit(self._run_ocr_background, rois_copy, prev_copy, ocr_start_time, self.current_fps)
                                self.active_futures.append(future)
                            self.ocr_counter = 0
                
                # Maintain Timing
                self.msleep(10)

            if cap: cap.release()

    def stop(self):
        self.running = False
        
class ReorderableTable(QTableWidget):
    """Field table whose rows can be dragged up/down to change their order."""
    row_moved = pyqtSignal(int, int)   # (old_row, new_row)

    def __init__(self, rows, cols):
        super().__init__(rows, cols)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QTableWidget.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragDropOverwriteMode(False)
        self._grab_ok = False
        self._dragging = False

    def mousePressEvent(self, event):
        # A row can only be grabbed by its Field (col 1) or Value (col 2) cell.
        idx = self.indexAt(event.position().toPoint())
        self._grab_ok = idx.isValid() and idx.column() in (1, 2)
        super().mousePressEvent(event)

    def startDrag(self, supportedActions):
        if not self._grab_ok:
            return
        self._dragging = True
        try:
            super().startDrag(supportedActions)
        finally:
            self._dragging = False

    def dropEvent(self, event):
        if not self._dragging:
            event.ignore()
            return
        src = self.currentRow()
        pos = event.position().toPoint()
        idx = self.indexAt(pos)
        if idx.isValid():
            dst = idx.row()
            if pos.y() > self.visualRect(idx).center().y():
                dst += 1                      # dropped on the lower half -> insert after
        else:
            dst = self.rowCount()             # dropped below the last row
        if dst > src:
            dst -= 1                          # account for the row being removed first
        # We move the data ourselves; tell Qt not to clear the dragged cells.
        event.setDropAction(Qt.DropAction.IgnoreAction)
        event.accept()
        self.viewport().update()
        if src >= 0 and dst != src:
            self.row_moved.emit(src, dst)


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
        
        self.ocr_validator = OCRValidator()

        self.engine = CaptureEngine()
        self.engine.frame_signal.connect(self.update_preview)
        self.engine.ocr_signal.connect(self.update_ocr_text)
        self.engine.previews_signal.connect(self.update_roi_preview)
        
        self.internal_update = False 
        self.init_ui()
        
        self.shortcut_reset = QShortcut(QKeySequence("Ctrl+Shift+R"), self)
        self.shortcut_reset.activated.connect(self.preview_overlay.reset_view)

        self.shortcut_rename = QShortcut(QKeySequence("Ctrl+R"), self)
        self.shortcut_rename.activated.connect(self.shortcut_rename_field)
        self.shortcut_copy = QShortcut(QKeySequence("Ctrl+C"), self)
        self.shortcut_copy.activated.connect(self.shortcut_copy_field)
        self.shortcut_paste = QShortcut(QKeySequence("Ctrl+V"), self)
        self.shortcut_paste.activated.connect(self.shortcut_paste_field)
        
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
        self.combo_source.addItems(["Select Source", "Open a Video File", "Screen Capture", "Capture Card (Camera)"])
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

        # --- Capture Card Widget ---
        self.capture_card_widget = QWidget()
        card_layout = QHBoxLayout(self.capture_card_widget)
        card_layout.setContentsMargins(0, 0, 0, 5)

        card_layout.addWidget(QLabel("Device Index:"))
        self.combo_card_index = QComboBox()
        self._populate_card_devices()
        if self.combo_card_index.count() == 0:
            self.combo_card_index.addItem("No devices found", None)

        self.combo_card_index.currentIndexChanged.connect(self.handle_card_pick)

        btn_refresh_card = QPushButton("Refresh")
        btn_refresh_card.setFixedWidth(70)
        btn_refresh_card.clicked.connect(self.refresh_card_list)

        card_layout.addWidget(self.combo_card_index, stretch=1)
        card_layout.addWidget(btn_refresh_card)

        self.capture_card_widget.hide()
        config_layout.addWidget(self.capture_card_widget)

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
        
        self.roi_table = ReorderableTable(0, 3)
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
        self.roi_table.row_moved.connect(self.on_table_row_moved)
        
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
        self.combo_type.addItems(["General Text", "Numbers Only", "Time Format", "KDA (K/D/A)", "Gold Amount (K/M)"])
        self.combo_type.currentIndexChanged.connect(self.sync_properties)
        grid.addWidget(self.combo_type, 0, 1)

        self.sl_thresh = QSlider(Qt.Orientation.Horizontal)
        self.sl_thick = QSlider(Qt.Orientation.Horizontal)
        self.sl_conf = QSlider(Qt.Orientation.Horizontal)

        # Live number readouts next to each slider.
        self.lbl_thresh_val = QLabel("1")
        self.lbl_thick_val = QLabel("5")
        self.lbl_conf_val = QLabel("60%")
        for lbl in [self.lbl_thresh_val, self.lbl_thick_val, self.lbl_conf_val]:
            lbl.setFixedWidth(40)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        for sl in [self.sl_thresh, self.sl_thick, self.sl_conf]:
            sl.setRange(1, 10)
            sl.setTickPosition(QSlider.TickPosition.TicksBelow)
            sl.setTickInterval(1)
            sl.valueChanged.connect(self.sync_properties)
            sl.setStyleSheet("QSlider::handle:horizontal { background: #888; width: 12px; border-radius: 6px; }")

        self.sl_thresh.valueChanged.connect(lambda v: self.lbl_thresh_val.setText(str(v)))
        self.sl_thick.valueChanged.connect(lambda v: self.lbl_thick_val.setText(str(v)))
        self.sl_conf.valueChanged.connect(lambda v: self.lbl_conf_val.setText(f"{v * 10}%"))

        grid.addWidget(QLabel("Binarize"), 1, 0)
        grid.addWidget(self.sl_thresh, 1, 1)
        grid.addWidget(self.lbl_thresh_val, 1, 2)

        grid.addWidget(QLabel("Cleanup/Dilate"), 2, 0)
        grid.addWidget(self.sl_thick, 2, 1)
        grid.addWidget(self.lbl_thick_val, 2, 2)

        grid.addWidget(QLabel("Conf. Th"), 3, 0)
        grid.addWidget(self.sl_conf, 3, 1)
        grid.addWidget(self.lbl_conf_val, 3, 2)

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

        self.btn_copy_crop = QPushButton("Copy Image")
        self.btn_copy_crop.clicked.connect(self.copy_crop_preview)
        props_main_layout.addWidget(self.btn_copy_crop)

        config_layout.addWidget(self.props_frame)
        config_layout.addStretch() 

        # ----- TAB 2: LIVE OUTPUT -----
        tab_output = QWidget()
        output_layout = QVBoxLayout(tab_output)
        output_layout.setContentsMargins(8, 8, 8, 8)
        
        self.meta_card = QFrame()
        self.meta_card.setStyleSheet("background-color: #222; border-radius: 4px; border: 1px solid #444; padding: 4px;")
        meta_layout = QHBoxLayout(self.meta_card)
        
        self.lbl_meta_frames = QLabel("Frames: 0")
        self.lbl_meta_ping = QLabel("Processing: -- ms")
        self.lbl_meta_areas = QLabel("Scanned: 0")
        
        for lbl in[self.lbl_meta_frames, self.lbl_meta_ping, self.lbl_meta_areas]:
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

    def _focused_text_widget(self):
        w = QApplication.focusWidget()
        return w if isinstance(w, (QLineEdit, QTextEdit)) else None

    def shortcut_copy_field(self):
        tw = self._focused_text_widget()
        if tw is not None:      # normal text copy while typing / selecting text
            tw.copy()
            return
        self.preview_overlay.copy_selected_field()

    def shortcut_paste_field(self):
        tw = self._focused_text_widget()
        if tw is not None:      # normal text paste while typing
            if not tw.isReadOnly():
                tw.paste()
            return
        self.preview_overlay.paste_field_centered()

    def shortcut_rename_field(self):
        roi_id = self.preview_overlay.selected_id
        if roi_id is None:
            return
        for row in range(self.roi_table.rowCount()):
            item = self.roi_table.item(row, 1)
            if item is not None and item.data(Qt.ItemDataRole.UserRole) == roi_id:
                self.roi_table.setCurrentItem(item)
                self.roi_table.setFocus()
                self.roi_table.editItem(item)
                editor = QApplication.focusWidget()
                if isinstance(editor, QLineEdit):
                    editor.selectAll()
                break

    def reset_to_defaults(self):
        if self.preview_overlay.selected_id is None: return
        self.combo_type.setCurrentText("General Text")
        self.sl_thresh.setValue(1)  
        self.sl_thick.setValue(5)   
        self.sl_conf.setValue(6)    
        self.lbl_thresh_val.setText("1")
        self.lbl_thick_val.setText("5")
        self.lbl_conf_val.setText("60%") 

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

    def on_table_row_moved(self, src, dst):
        rois = self.preview_overlay.rois
        if not (0 <= src < len(rois) and 0 <= dst < len(rois)):
            return

        # Remember each field's current value so it travels with its row.
        values = {}
        for i in range(self.roi_table.rowCount()):
            name_item = self.roi_table.item(i, 1)
            val_item = self.roi_table.item(i, 2)
            if name_item and val_item:
                values[name_item.data(Qt.ItemDataRole.UserRole)] = val_item.text()

        rois.insert(dst, rois.pop(src))       # reorder the real data
        self.sync_table_to_rois(rois)         # rebuild the table + update the OCR engine

        self.internal_update = True
        for i in range(self.roi_table.rowCount()):
            rid = self.roi_table.item(i, 1).data(Qt.ItemDataRole.UserRole)
            self.roi_table.item(i, 2).setText(values.get(rid, ""))
        self.internal_update = False
        self.preview_overlay.update()

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

    def handle_card_pick(self):
        idx = self.combo_card_index.currentData()
        name = self.combo_card_index.currentText()
        if idx is None:
            return
        self.preview_overlay.is_new_source = True
        import threading
        def _open():
            self.engine.set_source_capture_card(idx, name)
            if not self.engine.isRunning():
                self.engine.start()
        threading.Thread(target=_open, daemon=True).start()

    def refresh_card_list(self):
        self.combo_card_index.blockSignals(True)
        self.combo_card_index.clear()
        self._populate_card_devices()
        if self.combo_card_index.count() == 0:
            self.combo_card_index.addItem("No devices found", None)
        self.combo_card_index.blockSignals(False)

    def _populate_card_devices(self):
        """
        Enumerate video capture devices via pygrabber (DirectShow).
        Falls back to index probing if pygrabber is unavailable.
        """
        # --- Method 1: pygrabber (most reliable, gets exact DS names) ---
        try:
            from pygrabber.dshow_graph import FilterGraph
            graph = FilterGraph()
            device_names = graph.get_input_devices()
            print(f"[CardEnum] pygrabber devices: {device_names}")
            for i, name in enumerate(device_names):
                self.combo_card_index.addItem(name, i)
            print(f"[CardEnum] Final mapping: "
                  f"{[(self.combo_card_index.itemText(i), self.combo_card_index.itemData(i)) for i in range(self.combo_card_index.count())]}")
            return
        except Exception as e:
            print(f"[CardEnum] pygrabber unavailable ({e}), falling back to index probe")

        # --- Method 2: index probe fallback ---
        import winreg
        reg_names = []
        DSHOW_KEY = r"SOFTWARE\Classes\CLSID\{860BB310-5D01-11d0-BD3B-00A0C911CE86}\Instance"
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, DSHOW_KEY)
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        friendly, _ = winreg.QueryValueEx(subkey, "FriendlyName")
                        if friendly:
                            reg_names.append(friendly)
                    except FileNotFoundError:
                        pass
                    finally:
                        winreg.CloseKey(subkey)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception as e:
            print(f"[CardEnum] Registry read failed: {e}")

        print(f"[CardEnum] Registry fallback names: {reg_names}")
        probe_limit = max(len(reg_names) + 6, 12)
        found_count = 0
        for i in range(probe_limit):
            opened = False
            for backend in [cv2.CAP_MSMF, cv2.CAP_DSHOW, cv2.CAP_ANY]:
                try:
                    cap_test = cv2.VideoCapture(i, backend)
                    if cap_test.isOpened():
                        cap_test.release()
                        opened = True
                        break
                    cap_test.release()
                except Exception:
                    pass
            if opened:
                name = reg_names[found_count] if found_count < len(reg_names) else f"Device {i}"
                self.combo_card_index.addItem(name, i)
                print(f"[CardEnum] Index {i} → '{name}'")
                found_count += 1

        print(f"[CardEnum] Final mapping: "
              f"{[(self.combo_card_index.itemText(i), self.combo_card_index.itemData(i)) for i in range(self.combo_card_index.count())]}")

    def handle_source_change(self, index):
        source = self.combo_source.currentText()
        self.screen_widget.hide()
        self.capture_card_widget.hide()
        if source == "Open a Video File":
            path, _ = QFileDialog.getOpenFileName(self, "Select Video", "", "Video Files (*.mp4 *.avi *.mkv)")
            if path:
                self.preview_overlay.is_new_source = True
                self.engine.set_source_video(path)
                if not self.engine.isRunning(): self.engine.start()
        elif source == "Screen Capture":
            self.screen_widget.show()
            self.refresh_window_list()
        elif source == "Capture Card (Camera)":
            self.capture_card_widget.show()
            self.handle_card_pick()

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

            # setValue() above won't fire valueChanged if the number is
            # unchanged between ROIs, so refresh the readouts explicitly.
            self.lbl_thresh_val.setText(str(roi['threshold']))
            self.lbl_thick_val.setText(str(roi['thickness']))
            self.lbl_conf_val.setText(f"{roi['confidence'] * 10}%")

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
            # Reset validator state so previous game values don't block new reads
            self.ocr_validator.reset()
            # Also clear the displayed values in the table back to 0
            for i in range(self.roi_table.rowCount()):
                val_item = self.roi_table.item(i, 2)
                if val_item:
                    val_item.setText("0")
            self.tabs.setCurrentIndex(1)

    def update_preview(self, frame):
        h, w, c = frame.shape
        # Both capture_window_direct (GetDIBits) and MSS return BGRA.
        # Format_ARGB32 = BGRA on little-endian x86/x64 — correct colors without extra swap.
        q_img = QImage(frame.data, w, h, w * c, QImage.Format.Format_ARGB32)
        self.preview_overlay.set_frame(QPixmap.fromImage(q_img))

    def copy_crop_preview(self):
        img = getattr(self, 'last_crop_image', None)
        if img is None or img.size == 0:
            self.btn_copy_crop.setText("Select a field first")
        else:
            img = np.ascontiguousarray(img)
            # Enlarge small crops (sharp pixels) so the pasted picture is easy to see.
            grow = max(1, 400 // max(1, img.shape[1]))
            if grow > 1:
                img = cv2.resize(img, None, fx=grow, fy=grow, interpolation=cv2.INTER_NEAREST)
            h, w = img.shape
            qimg = QImage(img.data, w, h, w, QImage.Format.Format_Grayscale8).copy()
            QApplication.clipboard().setImage(qimg)
            self.btn_copy_crop.setText("Copied!")
        from PyQt6.QtCore import QTimer
        QTimer.singleShot(1200, lambda: self.btn_copy_crop.setText("Copy Image"))

    def update_roi_preview(self, previews_dict):
        roi_id = self.preview_overlay.selected_id
        if roi_id is not None and roi_id in previews_dict:
            processed = previews_dict[roi_id]
            self.last_crop_image = processed      # full-size image, used by the Copy Image button
            h, w = processed.shape
            q_img = QImage(processed.data, w, h, w, QImage.Format.Format_Grayscale8)
            pixmap = QPixmap.fromImage(q_img)
            scaled = pixmap.scaled(340, 60, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.lbl_crop_preview.setPixmap(scaled)

    def update_ocr_text(self, metadata, data_dict):
        # Update UI labels
        self.lbl_meta_frames.setText(f"FPS: {metadata.get('fps', 0)}")
        self.lbl_meta_ping.setText(f"Processing: {metadata.get('process_time_ms', '--')} ms")
        self.lbl_meta_areas.setText(f"Scanned: {metadata.get('areas_scanned', 0)}")
        
        # --- SELECTION PRESERVATION LOGIC ---
        cursor = self.ocr_output.textCursor()
        had_selection = cursor.hasSelection()
        
        if QApplication.mouseButtons() == Qt.MouseButton.LeftButton and self.ocr_output.underMouse():
            return

        scroll_pos = self.ocr_output.verticalScrollBar().value()
        sel_start = cursor.selectionStart()
        sel_end = cursor.selectionEnd()
        # ------------------------------------

        self.internal_update = True
        display_dict = {}
        roi_types = {roi['id']: roi.get('type') for roi in self.preview_overlay.rois}
        
        for i in range(self.roi_table.rowCount()):
            field_name = self.roi_table.item(i, 1).text()
            roi_id = self.roi_table.item(i, 1).data(Qt.ItemDataRole.UserRole)
            safe_name = field_name if field_name else f"Area_{roi_id}"
            
            current_value = self.roi_table.item(i, 2).text()
            
            if safe_name in data_dict:
                raw_val = data_dict[safe_name]
                # --- VALIDATION: reject bad OCR reads before storing ---
                if roi_types.get(roi_id) == 'KDA (K/D/A)':
                    validated_val = self.ocr_validator.validate_kda(safe_name, raw_val, current_value)
                else:
                    validated_val = self.ocr_validator.validate(safe_name, raw_val, current_value)
                self.roi_table.item(i, 2).setText(validated_val)
                display_dict[safe_name] = validated_val
            else:
                display_dict[safe_name] = current_value
                
        self.internal_update = False
        
        formatted_json = json.dumps(display_dict, indent=4)
        
        # Update the text in the UI
        self.ocr_output.setPlainText(formatted_json)

        # --- RESTORE SELECTION AND SCROLL ---
        if had_selection:
            new_cursor = self.ocr_output.textCursor()
            if sel_start == 0 and sel_end >= (len(self.ocr_output.toPlainText()) - 20):
                new_cursor.select(new_cursor.SelectionType.Document)
            else:
                text_len = len(formatted_json)
                new_cursor.setPosition(min(sel_start, text_len))
                new_cursor.setPosition(min(sel_end, text_len), new_cursor.MoveMode.KeepAnchor)
            self.ocr_output.setTextCursor(new_cursor)
        
        self.ocr_output.verticalScrollBar().setValue(scroll_pos)

        # ---------------------------------------------------------
        # NEW: EXPORT TO LOCAL JSON FILE FOR YOUR OVERLAY
        # ---------------------------------------------------------
        try:
            export_path = os.path.join(current_dir, "live_overlay_data.json")
            with open(export_path, "w", encoding="utf-8") as f:
                json.dump(display_dict, f, indent=4)
        except Exception as e:
            pass

        # ---------------------------------------------------------
        # GOLD LOG — appends one line per frame to gold-log.txt
        # Format: [HH:MM:SS] red gold: <val> | blue gold: <val>
        # ---------------------------------------------------------
        try:
            blue_gold = display_dict.get("blue-gold", "")
            red_gold  = display_dict.get("red-gold",  "")
            if blue_gold or red_gold:
                timestamp = time.strftime("%H:%M:%S")
                log_line  = f"[{timestamp}] red gold: {red_gold} | blue gold: {blue_gold}\n"
                log_path  = os.path.join(current_dir, "gold-log.txt")
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(log_line)
        except Exception:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = OCRApp()
    win.show()
    sys.exit(app.exec())