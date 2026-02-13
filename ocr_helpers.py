"""
OCR-based text detection for PRMaker (Optimized Tesseract)
Enhanced preprocessing for maximum accuracy.
"""
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import os

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

def _preprocess_extreme(image, target="text"):
    """
    Extreme preprocessing for maximum OCR accuracy.
    
    Args:
        image: PIL Image
        target: "text" for normal text, "header" for menu headers
    """
    # 1. Upscale 3x for better small text
    w, h = image.size
    scale = 3 if target == "header" else 2
    new_w, new_h = w * scale, h * scale
    processed = image.resize((new_w, new_h), Image.Resampling.LANCZOS)
    
    # 2. Convert to grayscale
    processed = processed.convert('L')
    
    # 3. Increase contrast (2x)
    enhancer = ImageEnhance.Contrast(processed)
    processed = enhancer.enhance(2.0)
    
    # 4. Sharpening for clearer edges
    processed = processed.filter(ImageFilter.SHARPEN)
    
    # 5. Binarization (Otsu's method via PIL)
    # Convert to pure black and white for cleaner text
    threshold = 128  # Auto-threshold would be better but PIL doesn't have Otsu built-in
    processed = processed.point(lambda p: 255 if p > threshold else 0)
    
    return processed, scale

def find_text_in_image(screenshot, target_text, region="full"):
    """
    Find text using optimized Tesseract OCR.
    
    Args:
        screenshot: PIL Image
        target_text: Text to search for
        region: "full", "top" (top 150px only), or "header" (top 100px, 3x scale)
    """
    if not _tesseract_available:
        if not init_tesseract():
             return None
    
    try:
        import pytesseract
        
        # Region cropping for speed and accuracy
        if region == "top":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(150, h)))
        elif region == "header":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(100, h)))
        
        # Preprocessing
        target_type = "header" if region == "header" else "text"
        processed, scale = _preprocess_extreme(screenshot, target=target_type)
        
        # OCR with optimized config
        # --psm 11: Sparse text (good for menus)
        # --oem 1: LSTM neural net mode (more accurate)
        config = '--psm 11 --oem 1'
        data = pytesseract.image_to_data(processed, output_type=pytesseract.Output.DICT, config=config)
        
        n_boxes = len(data['text'])
        target_lower = target_text.lower()
        
        for i in range(n_boxes):
            conf_val = int(data['conf'][i])
            if conf_val < 30:  # Minimum confidence threshold
                continue
                
            text = data['text'][i].strip()
            if not text:
                continue
                
            # Match check (case insensitive, substring)
            if target_lower in text.lower():
                # Scale coordinates back to original size
                left = int(data['left'][i] / scale)
                top = int(data['top'][i] / scale)
                width = int(data['width'][i] / scale)
                height = int(data['height'][i] / scale)
                
                class Box:
                    def __init__(self, l, t, w, h):
                        self.left = l
                        self.top = t
                        self.width = w
                        self.height = h
                
                print(f"OCR found '{text}' (conf: {conf_val}%) at ({left}, {top})")
                return Box(left, top, width, height)
        
        return None
        
    except Exception as e:
        # Suppress frequent errors
        return None


def find_all_text_in_image(screenshot, target_text, region="full"):
    """
    Find ALL occurrences of text using Tesseract OCR.
    Returns a list of Box objects, or empty list if none found.
    """
    if not _tesseract_available:
        if not init_tesseract():
             return []
    
    try:
        import pytesseract
        
        if region == "top":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(150, h)))
        elif region == "header":
            w, h = screenshot.size
            screenshot = screenshot.crop((0, 0, w, min(100, h)))
        
        target_type = "header" if region == "header" else "text"
        processed, scale = _preprocess_extreme(screenshot, target=target_type)
        
        config = '--psm 11 --oem 1'
        data = pytesseract.image_to_data(processed, output_type=pytesseract.Output.DICT, config=config)
        
        class Box:
            def __init__(self, l, t, w, h):
                self.left = l
                self.top = t
                self.width = w
                self.height = h
        
        results = []
        n_boxes = len(data['text'])
        target_lower = target_text.lower()
        
        for i in range(n_boxes):
            conf_val = int(data['conf'][i])
            if conf_val < 30:
                continue
            text = data['text'][i].strip()
            if not text:
                continue
            if target_lower in text.lower():
                left = int(data['left'][i] / scale)
                top = int(data['top'][i] / scale)
                width = int(data['width'][i] / scale)
                height = int(data['height'][i] / scale)
                print(f"OCR found '{text}' (conf: {conf_val}%) at ({left}, {top})")
                results.append(Box(left, top, width, height))
        
        return results
        
    except Exception as e:
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
    Finds M&R submenu in the dropdown area.
    """
    # Search more keywords with lower confidence needs
    variations = ["Maintenance", "Protection", "Machinery", "M&R", "Hull"] 
    
    for text in variations:
        result = find_text_in_image(screenshot, text, region="full")
        if result:
            # Additional validation: should be in reasonable dropdown area
            if result.top > 50 and result.top < 400:  # Not in header, not in footer
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
