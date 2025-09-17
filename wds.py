import win32gui
import win32con
import win32clipboard
from mss import mss
from PIL import Image
import keyboard
import io
import ctypes
from ctypes import wintypes
import sys
import time
import ctypes

# ======================
# CONFIGURATION
# ======================
CAPTURE_HOTKEY = 'CTRL+ALT+S'
BACKGROUND_COLOR = (255, 255, 255)  # RGB: white
USE_OVERLAY = True  # Set to False to disable fullscreen overlay
OVERLAY_DISPLAY_TIME_SEC = 0.1  # Time to show overlay before capture
# ======================

if USE_OVERLAY:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
# Try to use DWM API for accurate window bounds (without shadows)
try:
    dwmapi = ctypes.WinDLL("dwmapi")
    DWMWA_EXTENDED_FRAME_BOUNDS = 9

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left",   wintypes.LONG),
            ("top",    wintypes.LONG),
            ("right",  wintypes.LONG),
            ("bottom", wintypes.LONG),
        ]

    def get_visible_window_rect(hwnd):
        """Get window rectangle excluding DWM shadow"""
        try:
            rect = RECT()
            size = ctypes.c_uint32(ctypes.sizeof(rect))
            result = dwmapi.DwmGetWindowAttribute(
                hwnd,
                DWMWA_EXTENDED_FRAME_BOUNDS,
                ctypes.byref(rect),
                size
            )
            if result == 0:
                return (rect.left, rect.top, rect.right, rect.bottom)
        except Exception as e:
            print(f"[ERR] DWM error: {e}")
        return win32gui.GetWindowRect(hwnd)

except Exception as e:
    print(f"[WARN] DWM API not available: {e}")

    def get_visible_window_rect(hwnd):
        return win32gui.GetWindowRect(hwnd)


# --- Optional: Fullscreen overlay to clean background ---
if USE_OVERLAY:
    import tkinter as tk

    def create_overlay():
        """Create a fullscreen solid-color overlay window"""
        root = tk.Tk()
        root.overrideredirect(True)  # No title bar
        
        # Get actual screen dimensions in PHYSICAL pixels
        root.withdraw()  # Hide temporarily to get correct DPI-scaling info
        root.update_idletasks()
        width = root.winfo_screenwidth()
        height = root.winfo_screenheight()
        root.deiconify()

        root.geometry(f"{width}x{height}+0+0")
        root.configure(bg="#%02x%02x%02x" % BACKGROUND_COLOR)
        # Do NOT set -topmost → so our app can stay on top naturally
        root.update()
        return root

# --- Helper functions ---
def get_parent_window(hwnd):
    """Get owner window"""
    parent = win32gui.GetParent(hwnd)
    if parent == 0:
        parent = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
    return parent

def get_window_text(hwnd):
    """Get window title"""
    try:
        return win32gui.GetWindowText(hwnd)
    except:
        return ""

def set_clipboard_bmp(image):
    """Copy PIL image to clipboard as BMP (CF_DIB)"""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()

        with io.BytesIO() as output:
            image.convert("RGB").save(output, format='BMP')
            data = output.getvalue()

        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data[14:])
        win32clipboard.CloseClipboard()
        print("[OK]  Image copied to clipboard (BMP)")

    except Exception as e:
        print(f"[ERR] Clipboard error: {e}")

# --- Main capture ---
def capture_combined_region():
    try:
        # === 1. Get the target window BEFORE showing overlay ===
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            print("[ERR] No active window")
            return

        title = get_window_text(hwnd)
        print(f"[INFO] Active window: '{title}' (HWND: {hwnd})")

        # Store dialog and parent rects
        dialog_rect = get_visible_window_rect(hwnd)
        print(f"       Dialog rect: {dialog_rect}")

        parent_hwnd = get_parent_window(hwnd)
        parent_rect = None

        if parent_hwnd and parent_hwnd != hwnd:
            parent_title = get_window_text(parent_hwnd)
            parent_rect = get_visible_window_rect(parent_hwnd)
            print(f"       Parent window: '{parent_title}' → {parent_rect}")

        # Define combined area
        if parent_rect:
            combined_rect = (
                min(parent_rect[0], dialog_rect[0]),
                min(parent_rect[1], dialog_rect[1]),
                max(parent_rect[2], dialog_rect[2]),
                max(parent_rect[3], dialog_rect[3]),
            )
            print(f"[OK]  Combined area: {combined_rect}")
        else:
            combined_rect = dialog_rect

        left, top, right, bottom = combined_rect
        width = right - left
        height = bottom - top

        if width <= 0 or height <= 0:
            print("[ERR] Invalid capture size")
            return

        # === 2. Show fullscreen overlay (our active app remains top-most) ===
        overlay = None
        if USE_OVERLAY:
            overlay = create_overlay()
            print("[INFO] Fullscreen overlay shown")
            time.sleep(OVERLAY_DISPLAY_TIME_SEC)

        # === 3. Capture the region ===
        with mss() as sct:
            monitor = {"top": top, "left": left, "width": width, "height": height}
            sct_data = sct.grab(monitor)
            screen_img = Image.frombytes("RGB", sct_data.size, sct_data.rgb)

        # === 4. Create final image with solid background ===
        final_img = Image.new("RGB", (width, height), BACKGROUND_COLOR)

        # Paste parent
        if parent_rect:
            # Coordinates relative to combined area
            parent_box = (
                parent_rect[0] - left,
                parent_rect[1] - top,
                parent_rect[2] - left,
                parent_rect[3] - top
            )
            parent_crop = screen_img.crop(parent_box)
            final_img.paste(parent_crop, (parent_box[0], parent_box[1]))

        # Paste dialog
        dialog_box = (
            dialog_rect[0] - left,
            dialog_rect[1] - top,
            dialog_rect[2] - left,
            dialog_rect[3] - top
        )
        dialog_crop = screen_img.crop(dialog_box)
        final_img.paste(dialog_crop, (dialog_box[0], dialog_box[1]))

        # === 5. Cleanup overlay ===
        if overlay:
            overlay.destroy()
            print("[INFO] Overlay removed")

        set_clipboard_bmp(final_img)
        print("[OK]  Capture complete: only parent + dialog pasted, background filled")

    except Exception as e:
        if 'overlay' in locals():
            try:
                overlay.destroy()
            except:
                pass
        print(f"[ERR] Capture failed: {type(e).__name__}: {e}")

# === Main ===
def main():
    print("=== win-dialog-shot v1.1 ===")
    print(f"  Capture: {CAPTURE_HOTKEY}")
    print("  Exit:    CTRL+C")
    print(f"  Background color: {BACKGROUND_COLOR} (RGB)")
    print(f"  Overlay mode: {'Enabled' if USE_OVERLAY else 'Disabled'}")
    print("  Only parent + dialog are pasted. Corners filled with color.")
    print("  Ready... Press CTRL+C to quit.")

    keyboard.add_hotkey(CAPTURE_HOTKEY, capture_combined_region)

    try:
        print("[RUN] Background listener is active. Waiting for hotkeys...")
        while True:
            time.sleep(1)  # Keep alive
    except KeyboardInterrupt:
        print("\n[INFO] Exit signal received (Ctrl+C). Shutting down...")
    except Exception as e:
        print(f"[ERR] Unexpected error: {e}")
    finally:
        print("[INFO] Goodbye!")
        sys.exit(0)

if __name__ == "__main__":
    main()