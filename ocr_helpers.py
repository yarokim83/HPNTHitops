"""
OCR-based text detection for PRMaker (Optimized Tesseract)
Enhanced preprocessing for maximum accuracy.
"""
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import os

try:
    import numpy as np
    _NUMPY_AVAILABLE = True
except Exception:
    _NUMPY_AVAILABLE = False

# Global Tesseract configuration
_tesseract_available = False

def init_tesseract():
    """
    Initialize Tesseract configuration.
    """
    global _tesseract_available
    
    try:
        import pytesseract
        
        # Check PATH first
        try:
            pytesseract.get_tesseract_version()
            _tesseract_available = True
            print("Tesseract found in PATH.")
            return True
        except:
            pass
        
        # Check common paths
        paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        
        for path in paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                _tesseract_available = True
                print(f"Tesseract found at: {path}")
                return True
                
    except Exception as e:
        print(f"Tesseract not available: {e}")
    
    print("WARNING: Tesseract OCR not available!")
    return False

# Alias for compatibility
init_ocr = init_tesseract

def _otsu_threshold(gray_image):
    """Compute Otsu's optimal binarization threshold for a grayscale PIL image.

    Falls back to 128 if numpy is unavailable.
    """
    if not _NUMPY_AVAILABLE:
        return 128
    try:
        arr = np.asarray(gray_image, dtype=np.uint8)
        hist, _ = np.histogram(arr, bins=256, range=(0, 256))
        total = arr.size
        if total == 0:
            return 128
        sum_total = np.dot(np.arange(256), hist)
        sum_b = 0.0
        w_b = 0.0
        max_var = 0.0
        threshold = 128
        for t in range(256):
            w_b += hist[t]
            if w_b == 0:
                continue
            w_f = total - w_b
            if w_f == 0:
                break
            sum_b += t * hist[t]
            m_b = sum_b / w_b
            m_f = (sum_total - sum_b) / w_f
            var_between = w_b * w_f * (m_b - m_f) ** 2
            if var_between > max_var:
                max_var = var_between
                threshold = t
        return int(threshold)
    except Exception:
        return 128

def _preprocess_extreme(image, target="text", invert=False):
    """
    Extreme preprocessing for maximum OCR accuracy.

    Args:
        image: PIL Image
        target: "text" for normal text, "header" for menu headers
        invert: If True, invert colors before binarization. Useful for
            light-on-dark UIs (e.g., dark themes), where the default pass
            would lose the text after thresholding.
    """
    # 1. Upscale for better small text
    w, h = image.size
    scale = 3 if target == "header" else 2
    new_w, new_h = w * scale, h * scale
    processed = image.resize((new_w, new_h), Image.Resampling.LANCZOS)

    # 2. Grayscale
    processed = processed.convert('L')

    # 3. Optional invert (dark-theme pass)
    if invert:
        processed = ImageOps.invert(processed)

    # 4. Contrast boost
    processed = ImageEnhance.Contrast(processed).enhance(2.0)

    # 5. Sharpening
    processed = processed.filter(ImageFilter.SHARPEN)

    # 6. Otsu binarization (adaptive global threshold)
    threshold = _otsu_threshold(processed)
    processed = processed.point(lambda p: 255 if p > threshold else 0)

    return processed, scale

class _OcrBox:
    __slots__ = ("left", "top", "width", "height", "confidence", "text")
    def __init__(self, l, t, w, h, conf=0, text=""):
        self.left = l
        self.top = t
        self.width = w
        self.height = h
        self.confidence = conf
        self.text = text

def _scan_for_matches(screenshot, target_text, region, invert):
    """Run a single OCR pass and return all matches as _OcrBox list."""
    import pytesseract

    target_type = "header" if region == "header" else "text"
    processed, scale = _preprocess_extreme(screenshot, target=target_type, invert=invert)

    # --psm 11: Sparse text (menus). --oem 1: LSTM mode.
    config = '--psm 11 --oem 1'
    data = pytesseract.image_to_data(processed, output_type=pytesseract.Output.DICT, config=config)

    matches = []
    target_lower = target_text.lower()
    n_boxes = len(data['text'])
    for i in range(n_boxes):
        try:
            conf_val = int(data['conf'][i])
        except (TypeError, ValueError):
            continue
        if conf_val < 30:
            continue
        text = data['text'][i].strip()
        if not text:
            continue
        if target_lower not in text.lower():
            continue
        left = int(data['left'][i] / scale)
        top = int(data['top'][i] / scale)
        width = int(data['width'][i] / scale)
        height = int(data['height'][i] / scale)
        matches.append(_OcrBox(left, top, width, height, conf_val, text))
    return matches

def find_text_in_image(screenshot, target_text, region="full"):
    """
    Find text using optimized Tesseract OCR.
    Runs a normal pass and (if needed) an inverted pass for dark-theme UIs,
    then returns the single highest-confidence match.

    Args:
        screenshot: PIL Image
        target_text: Text to search for
        region: "full", "top" (top 150px only), or "header" (top 100px, 3x scale)
    """
    if not _tesseract_available:
        if not init_tesseract():
            return None

    try:
        if region == "top":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(150, h)))
        elif region == "header":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(100, h)))

        all_matches = _scan_for_matches(screenshot, target_text, region, invert=False)
        # Inverted pass: helps dark themes / light-on-dark text.
        if not all_matches:
            all_matches = _scan_for_matches(screenshot, target_text, region, invert=True)

        if not all_matches:
            return None

        # Pick the highest-confidence match.
        best = max(all_matches, key=lambda b: b.confidence)
        print(f"OCR found '{best.text}' (conf: {best.confidence}%) at ({best.left}, {best.top})")
        return best

    except Exception:
        return None


def find_all_text_in_image(screenshot, target_text, region="full"):
    """
    Find ALL occurrences of text using Tesseract OCR.
    Returns a list of Box objects (sorted by confidence desc), or empty list.
    """
    if not _tesseract_available:
        if not init_tesseract():
            return []

    try:
        if region == "top":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(150, h)))
        elif region == "header":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(100, h)))

        results = _scan_for_matches(screenshot, target_text, region, invert=False)
        if not results:
            results = _scan_for_matches(screenshot, target_text, region, invert=True)

        results.sort(key=lambda b: b.confidence, reverse=True)
        for b in results:
            print(f"OCR found '{b.text}' (conf: {b.confidence}%) at ({b.left}, {b.top})")
        return results

    except Exception:
        return []

def find_inventory_menu(screenshot):
    """
    Finds the 'Inventory' menu item.
    Searches full screen as Inventory can appear in dropdown.
    """
    variations = ["Inventory", "Inven", "Stock"]
    
    for text in variations:
        result = find_text_in_image(screenshot, text, region="full")
        if result:
            return result
            
    return None

def find_purchase_request_menu(screenshot):
    """
    Finds 'Purchase Request' or 'Approve Purchase Request' menu.
    """
    variations = ["Purchase Request", "Purchase", "Request", "Approve"]
    
    for text in variations:
        result = find_text_in_image(screenshot, text, region="full")
        if result:
            return result
    
    return None


def find_mr_submenu(screenshot):
    """
    Finds M&R submenu in the dropdown/side area.
    """
    variations = ["Maintenance", "Protection", "Machinery", "M&R", "Hull", "Inventory"] 
    
    for text in variations:
        result = find_text_in_image(screenshot, text, region="full")
        if result:
            # Dropdown/Side menus are typically on the far left (X < 500)
            # and distinct from the main tile area.
            # Maintenance tile is at X~80 (Logical) or X~400 (Maximized).
            # But sidebar submenus are usually at X < 300.
            if result.left < 400:
                print(f"OCR: Found Submenu '{text}' at X={result.left}, Y={result.top}")
                return result
            
    return None

def find_maintenance_root_menu(screenshot):
    """
    Finds the 'Maintenance' top-level menu.
    CRITICAL: Only searches TOP 100px with maximum preprocessing.
    """
    targets = ["Maintenance", "Mainte", "Operation", "Admin", "Planning"]
    
    for text in targets:
        # Use "header" mode: top 100px + 3x scale
        box = find_text_in_image(screenshot, text, region="header")
        if box:
            # Double-check Y position (should be in top header)
            if box.top < 80:  # Very strict top-only check
                print(f"OCR: Found Root Menu '{text}' at Y={box.top}")
                return box
    
    return None

def find_maintenance_tile(screenshot):
    """
    Finds the 'Maintenance & Repair' tile on main screen.
    Searches FULL screen for text within the tile.
    """
    # Keywords that appear on the tile
    targets = ["Maintenance", "Repair", "M&R"]
    
    for text in targets:
        box = find_text_in_image(screenshot, text, region="full")
        if box:
            # Tile is usually in lower half of screen (Y > 200)
            if box.top > 150:
                print(f"OCR: Found Tile '{text}' at ({box.left}, {box.top})")
                return box
    
    return None
