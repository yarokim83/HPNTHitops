import cv2
import os

def create_code_asset():
    # Source image (uploaded list screenshot)
    src_path = r"C:\Users\baewoong.kim\.gemini\antigravity\brain\71d054ab-f771-4289-b793-b10782a9ddb3\uploaded_media_1_1769583613930.png"
    
    if not os.path.exists(src_path):
        print(f"File not found: {src_path}")
        return

    img = cv2.imread(src_path)
    if img is None:
        print("Failed to load image.")
        return

    h, w, c = img.shape
    print(f"Image Dimensions: {w}x{h}")

    # Assumption: The target code '0501040106' is the LAST item in this list image.
    # Text height is typically 20-25px.
    # Let's crop the bottom 25 pixels.
    # Also we might need to ignore scrollbar on the right. Scrollbar width ~15px.
    
    # Crop logic: [y:y+h, x:x+w]
    item_height = 25
    y_start = h - item_height
    x_end = w - 15 # Exclude scrollbar
    
    crop = img[y_start:h, 0:x_end]
    
    out_dir = r"C:\Users\baewoong.kim\.gemini\PRMakerV1.0\assets"
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
        
    out_path = os.path.join(out_dir, "acc_code_0501040106.png")
    cv2.imwrite(out_path, crop)
    print(f"Saved asset to {out_path}")

if __name__ == "__main__":
    create_code_asset()
