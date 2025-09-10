import fitz  # PyMuPDF
from PIL import Image
import io
import cv2
import numpy as np
import os

def get_footer_image(pixmap, height_ratio=0.2):
    img = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
    width, height = img.size
    footer = img.crop((0, int(height * (1 - height_ratio)), width, height))
    return footer

def match_logo(logo_path, footer_pil, debug_path=None, min_matches=13):
    logo = cv2.imread(logo_path, cv2.IMREAD_COLOR)
    if logo is None:
        print("❌ Could not read logo image.")
        return False, 0

    logo = cv2.cvtColor(logo, cv2.COLOR_BGR2RGB)
    footer = np.array(footer_pil)

    akaze = cv2.AKAZE_create()
    kp1, des1 = akaze.detectAndCompute(logo, None)
    kp2, des2 = akaze.detectAndCompute(footer, None)

    if des1 is None or des2 is None:
        return False, 0

    bf = cv2.BFMatcher()
    matches = bf.knnMatch(des1, des2, k=2)

    good_matches = []
    for m, n in matches:
        if m.distance < 0.7 * n.distance:
            good_matches.append(m)

    match_count = len(good_matches)

    if match_count >= min_matches:
        src_pts = np.float32([kp1[m.queryIdx].pt for m in good_matches]).reshape(-1, 1, 2)
        dst_pts = np.float32([kp2[m.trainIdx].pt for m in good_matches]).reshape(-1, 1, 2)

        M, mask = cv2.findHomography(src_pts, dst_pts, cv2.RANSAC, 5.0)
        matchesMask = mask.ravel().tolist() if mask is not None else []
        inlier_count = sum(matchesMask)

        if debug_path:
            draw_params = dict(matchColor=(0, 255, 0), singlePointColor=None, matchesMask=matchesMask, flags=2)
            match_img = cv2.drawMatches(logo, kp1, footer, kp2, good_matches, None, **draw_params)
            cv2.imwrite(debug_path, cv2.cvtColor(match_img, cv2.COLOR_RGB2BGR))

        return inlier_count >= min_matches, inlier_count

    return False, match_count

def check_first_page_for_logo(pdf_path, logo_path, output_dir):
    try:
        doc = fitz.open(pdf_path)
        first_page = doc[0]
        pix = first_page.get_pixmap(dpi=150)
        footer_pil = get_footer_image(pix)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        footer_img_path = os.path.join(output_dir, f"{base_name}_logo_footer.png")
        debug_path = os.path.join(output_dir, f"{base_name}_matched_debug.png")

        has_logo, matches = match_logo(logo_path, footer_pil, debug_path=debug_path)

        if has_logo:
            footer_pil.save(footer_img_path)
            print(f"✅ Logo detected in: {pdf_path} | Matches: {matches}")
        else:
            print(f"❌ No logo in: {pdf_path} | Matches: {matches}")
    except Exception as e:
        print(f"⚠️ Error processing {pdf_path}: {e}")

def process_folder(folder_path, logo_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print("No PDF files found in folder.")
        return

    print(f"🔍 Processing {len(pdf_files)} PDF files...\n")
    for pdf_file in pdf_files:
        full_pdf_path = os.path.join(folder_path, pdf_file)
        check_first_page_for_logo(full_pdf_path, logo_path, output_dir)

# === Set Paths Here ===
FOLDER_PATH = input("Enter the folder path containing .pdf: ")
LOGO_PATH = r"C:\Users\SiddharthSrivastava\Downloads\Images\Logo"
OUTPUT_DIR = os.path.join(FOLDER_PATH, "output")

# === Run Batch Logo Detection ===
process_folder(FOLDER_PATH, LOGO_PATH, OUTPUT_DIR)