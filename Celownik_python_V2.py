import os
import sys
import time
import threading
import ctypes
from ctypes import wintypes

import win32api
import win32con
import win32gui
import keyboard
from PIL import Image
from pystray import Icon, MenuItem, Menu


# ---------- POMOCNICZE ŚCIEŻKI (PyInstaller) ----------

def resource_path(relative_path: str) -> str:
    """
    Zwraca ścieżkę do zasobów zarówno w trybie .py, jak i w .exe (PyInstaller).
    """
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.getcwd(), relative_path)


# ---------- USTAWIENIA ----------

SPLASH_FILE = resource_path("splash.png")
SPLASH_DURATION_MS = 1500

DOT_COLOR = (255, 0, 0)
DOT_SIZE = 6
FADE_DURATION_MS = 200
ICON_FILE = resource_path("celownik.ico")

MAGENTA = win32api.RGB(255, 0, 255)
stop_event = threading.Event()


# ---------- CTYPES STRUKTURY DLA SPLASH ----------

class POINT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
    ]


class SIZE(ctypes.Structure):
    _fields_ = [
        ("cx", wintypes.LONG),
        ("cy", wintypes.LONG),
    ]


class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_byte),
        ("BlendFlags", ctypes.c_byte),
        ("SourceConstantAlpha", ctypes.c_byte),
        ("AlphaFormat", ctypes.c_byte),
    ]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]

    def __init__(self, width, height):
        super().__init__()
        self.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        self.biWidth = width
        self.biHeight = -height  # top-down bitmapa
        self.biPlanes = 1
        self.biBitCount = 32
        self.biCompression = 0
        self.biSizeImage = 0
        self.biXPelsPerMeter = 0
        self.biYPelsPerMeter = 0
        self.biClrUsed = 0
        self.biClrImportant = 0


# ---------- FUNKCJA SPLASH (statyczny, stabilny) ----------

def show_splash():
    path = SPLASH_FILE  # już pełna ścieżka dzięki resource_path
    if not os.path.isfile(path):
        print(f"[SPLASH] Brak pliku splash: {path}")
        return

    try:
        img = Image.open(path).convert("RGBA")
    except Exception as e:
        print(f"[SPLASH] Nie udało się wczytać splash.png: {e}")
        return

    w, h = img.size
    data = img.tobytes("raw", "BGRA")

    hInstance = win32api.GetModuleHandle(None)
    className = "SplashLayered"

    wndClass = win32gui.WNDCLASS()
    wndClass.hInstance = hInstance
    wndClass.lpszClassName = className
    wndClass.lpfnWndProc = win32gui.DefWindowProc
    try:
        win32gui.RegisterClass(wndClass)
    except Exception:
        # klasa już zarejestrowana – ignorujemy
        pass

    screen_w = win32api.GetSystemMetrics(0)
    screen_h = win32api.GetSystemMetrics(1)
    x = screen_w // 2 - w // 2
    y = screen_h // 2 - h // 2

    hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
        className,
        None,
        win32con.WS_POPUP,
        x,
        y,
        w,
        h,
        None,
        None,
        hInstance,
        None,
    )

    hdc_screen = win32gui.GetDC(0)
    hdc_mem = win32gui.CreateCompatibleDC(hdc_screen)
    hbmp = win32gui.CreateCompatibleBitmap(hdc_screen, w, h)
    win32gui.SelectObject(hdc_mem, hbmp)

    bmi = BITMAPINFOHEADER(w, h)

    ctypes.windll.gdi32.SetDIBitsToDevice(
        hdc_mem,
        0,
        0,
        w,
        h,
        0,
        0,
        0,
        h,
        data,
        ctypes.byref(bmi),
        win32con.DIB_RGB_COLORS,
    )

    blend = BLENDFUNCTION()
    blend.BlendOp = win32con.AC_SRC_OVER
    blend.BlendFlags = 0
    blend.SourceConstantAlpha = 255
    blend.AlphaFormat = win32con.AC_SRC_ALPHA

    ctypes.windll.user32.UpdateLayeredWindow(
        hwnd,
        hdc_screen,
        ctypes.byref(POINT(x, y)),
        ctypes.byref(SIZE(w, h)),
        hdc_mem,
        ctypes.byref(POINT(0, 0)),
        0,
        ctypes.byref(blend),
        win32con.ULW_ALPHA,
    )

    win32gui.ReleaseDC(0, hdc_screen)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

    time.sleep(SPLASH_DURATION_MS / 1000)
    win32gui.DestroyWindow(hwnd)
    print("[SPLASH] OK")


# ---------------- OVERLAY ----------------

class DotOverlay:
    def __init__(self):
        self.dot_size = DOT_SIZE
        self.margin = DOT_SIZE
        self.window_size = self.dot_size + self.margin * 2
        self.visible = True
        self.class_name = "DotOverlay_Final"
        self.hwnd = None

        self._register()
        self._create()

    def _register(self):
        wc = win32gui.WNDCLASS()
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = self.class_name
        wc.lpfnWndProc = self._wnd_proc
        wc.hbrBackground = win32gui.CreateSolidBrush(MAGENTA)
        try:
            win32gui.RegisterClass(wc)
        except Exception as e:
            print(f"[REGISTER CLASS ERROR] {e}")

    def _create(self):
        screen_w = win32api.GetSystemMetrics(0)
        screen_h = win32api.GetSystemMetrics(1)
        x = screen_w // 2 - self.window_size // 2
        y = screen_h // 2 - self.window_size // 2

        style = win32con.WS_POPUP
        ex_style = (
            win32con.WS_EX_LAYERED
            | win32con.WS_EX_TOPMOST
            | win32con.WS_EX_TOOLWINDOW
            | win32con.WS_EX_TRANSPARENT
        )

        hinst = win32api.GetModuleHandle(None)
        self.hwnd = win32gui.CreateWindowEx(
            ex_style,
            self.class_name,
            None,
            style,
            x,
            y,
            self.window_size,
            self.window_size,
            None,
            None,
            hinst,
            None,
        )

        win32gui.SetLayeredWindowAttributes(
            self.hwnd,
            MAGENTA,
            255,
            win32con.LWA_COLORKEY | win32con.LWA_ALPHA,
        )

        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
        print("[FINAL] Overlay utworzony")

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_PAINT:
            hdc, ps = win32gui.BeginPaint(hwnd)

            bg = win32gui.CreateSolidBrush(MAGENTA)
            win32gui.FillRect(hdc, (0, 0, self.window_size, self.window_size), bg)

            pen = win32gui.CreatePen(
                win32con.PS_SOLID, 1, win32api.RGB(*DOT_COLOR)
            )
            brush = win32gui.CreateSolidBrush(win32api.RGB(*DOT_COLOR))

            prev_pen = win32gui.SelectObject(hdc, pen)
            prev_brush = win32gui.SelectObject(hdc, brush)

            left = self.margin
            top = self.margin
            right = self.margin + self.dot_size
            bottom = self.margin + self.dot_size

            win32gui.Ellipse(hdc, left, top, right, bottom)

            win32gui.SelectObject(hdc, prev_pen)
            win32gui.SelectObject(hdc, prev_brush)

            win32gui.DeleteObject(pen)
            win32gui.DeleteObject(brush)
            win32gui.DeleteObject(bg)

            win32gui.EndPaint(hwnd, ps)
            return 0

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def set_alpha(self, alpha):
        try:
            win32gui.SetLayeredWindowAttributes(
                self.hwnd,
                MAGENTA,
                int(alpha),
                win32con.LWA_COLORKEY | win32con.LWA_ALPHA,
            )
        except Exception as e:
            print(f"[ALPHA ERROR] {e}")

    def fade_in(self, duration_ms=FADE_DURATION_MS):
        steps = 15
        delay = duration_ms / steps / 1000
        for i in range(steps + 1):
            a = int(i * (255 / steps))
            self.set_alpha(a)
            time.sleep(delay)

    def fade_out(self, duration_ms=FADE_DURATION_MS):
        steps = 15
        delay = duration_ms / steps / 1000
        for i in range(steps + 1):
            a = int(255 - i * (255 / steps))
            self.set_alpha(a)
            time.sleep(delay)

    def toggle(self):
        print(f"[FINAL] toggle() visible={self.visible}")
        if self.visible:
            self.fade_out()
            win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
            self.visible = False
        else:
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
            self.fade_in()
            self.visible = True

    def close(self):
        print("[FINAL] Overlay close()")
        win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)


# ---------------- TRAY ----------------

class TrayManager:
    def __init__(self, overlay: DotOverlay):
        self.overlay = overlay
        self.icon = None
        self._create_icon()

    def _create_icon(self):
        icon_path = ICON_FILE  # już pełna ścieżka

        if not os.path.isfile(icon_path):
            print(f"[TRAY] Brak pliku ikony: {icon_path}, używam domyślnej.")
            img = Image.new("RGBA", (64, 64), (255, 0, 0, 255))
        else:
            print(f"[TRAY] Wczytuję ikonę: {icon_path}")
            img = Image.open(icon_path).convert("RGBA")

        menu = Menu(
            MenuItem("Pokaż/Ukryj", lambda icon, item: self.overlay.toggle()),
            MenuItem("Zakończ", lambda icon, item: self.exit()),
        )

        self.icon = Icon("Celownik", img, "Celownik", menu)

    def run(self):
        print("[FINAL] Tray start")
        self.icon.run()

    def exit(self):
        print("[FINAL] Tray exit")
        stop_event.set()
        self.icon.stop()


# ---------------- HOTKEYS ----------------

def listen_hotkeys(overlay: DotOverlay, tray: TrayManager):
    print("[FINAL] Rejestruję ALT+1 (toggle), ALT+Q (exit)")
    keyboard.add_hotkey("alt+1", overlay.toggle)

    def _exit():
        print("[FINAL] ALT+Q — wyjście")
        stop_event.set()
        tray.exit()

    keyboard.add_hotkey("alt+q", _exit)
    keyboard.wait()


# ---------------- MAIN ----------------

def run_all():
    show_splash()
    print("[FINAL] Start programu")

    overlay = DotOverlay()
    tray = TrayManager(overlay)

    t_tray = threading.Thread(target=tray.run, daemon=True)
    t_tray.start()

    t_hot = threading.Thread(
        target=listen_hotkeys, args=(overlay, tray), daemon=True
    )
    t_hot.start()

    print("[FINAL] Program działa — oczekiwanie na zdarzenia")

    while not stop_event.is_set():
        win32gui.PumpWaitingMessages()
        time.sleep(0.01)

    print("[FINAL] Koniec programu")
    overlay.close()


if __name__ == "__main__":
    run_all()

# Komenda do włączania celownika: python Celownik_python_V2.py