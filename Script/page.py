import fitz  # PyMuPDF
from PIL import Image, ImageDraw
from pyzbar.pyzbar import decode


def extract_text_positions(doc, search_text):
    text_positions = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(search_text)
        if text_instances:
            print(f"Found {len(text_instances)} occurrence(s) of text '{search_text}' on page {page_num + 1}.")
        for inst in text_instances:
            text_positions.append((page_num, inst))
    if not text_positions:
        print(f"No occurrences of the specified text '{search_text}' found in the document.")
    return text_positions


def detect_qr_codes_in_page(page, zoom=3):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    decoded_objects = decode(img)
    qr_rects = []
    for obj in decoded_objects:
        if obj.type == "QRCODE":
            rect = obj.rect
            qr_rects.append((rect.left, rect.top, rect.left + rect.width, rect.top + rect.height))
            print(f"Detected QR code at image coords: {qr_rects[-1]}")
    if not qr_rects:
        print(f"No QR codes detected on page {page.number + 1} at zoom {zoom}x.")
    return qr_rects, img


def convert_pdf_coords_to_image_coords(page, rect, zoom):
    x0, y0, x1, y1 = rect
    page_height = page.rect.height
    img_x0 = x0 * zoom
    img_x1 = x1 * zoom
    img_y0 = (page_height - y1) * zoom
    img_y1 = (page_height - y0) * zoom
    return (img_x0, img_y0, img_x1, img_y1)


def rects_are_next_to_each_other(text_rect, qr_rect, tolerance_x=100, tolerance_y=50):
    text_x0, text_y0, text_x1, text_y1 = text_rect
    qr_x0, qr_y0, qr_x1, qr_y1 = qr_rect

    horizontal_gap = qr_x0 - text_x1
    vertical_center_text = (text_y0 + text_y1) / 2
    vertical_center_qr = (qr_y0 + qr_y1) / 2
    vertical_diff = abs(vertical_center_text - vertical_center_qr)

    horizontal_close = 0 <= horizontal_gap <= tolerance_x
    vertical_close = vertical_diff <= tolerance_y

    debug_msg = (f"Checking adjacency - horizontal gap: {horizontal_gap}, vertical center diff: {vertical_diff} -> "
                 f"horizontal_close: {horizontal_close}, vertical_close: {vertical_close}")
    print(debug_msg)
    return horizontal_close and vertical_close


def draw_boxes(image, boxes, color, label=None):
    draw = ImageDraw.Draw(image)
    for i, box in enumerate(boxes):
        draw.rectangle(box, outline=color, width=3)
        if label:
            draw.text((box[0], box[1]-10), f"{label} {i+1}", fill=color)


def find_qr_code_next_to_text(pdf_path, search_text):
    doc = fitz.open(pdf_path)
    text_positions = extract_text_positions(doc, search_text)
    if not text_positions:
        return False

    for page_num, text_rect in text_positions:
        page = doc.load_page(page_num)
        zoom = 3
        qr_rects, img = detect_qr_codes_in_page(page, zoom)

        if not qr_rects:
            continue

        text_rect_img = convert_pdf_coords_to_image_coords(page, text_rect, zoom)
        print(f"Text box on image coords: {text_rect_img}")

        draw_boxes(img, [text_rect_img], color="blue", label="Text")
        draw_boxes(img, qr_rects, color="red", label="QR Code")
        debug_image_path = f"debug_page_{page_num + 1}.png"
        img.save(debug_image_path)
        print(f"Saved debug image with bounding boxes to: {debug_image_path}")

        for qr_rect in qr_rects:
            if rects_are_next_to_each_other(text_rect_img, qr_rect):
                print(f"QR code found next to specified text on page {page_num + 1}.")
                return True

    print("No QR code found next to the specified text.")
    return False


if __name__ == "__main__":
    pdf_path = r"C:\Users\SiddharthSrivastava\Downloads\T\B13412R1968575-ENG_ISI_04292025.pdf"
    search_text = "See claim details on following pages or go directly to MyUHCMedicare.com to view."

    found = find_qr_code_next_to_text(pdf_path, search_text)
    if found:
        print("Result: QR code exists next to the specified text.")
    else:
        print("Result: No QR code next to the specified text.")