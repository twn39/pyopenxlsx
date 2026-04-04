import pytest
import os
from pyopenxlsx import Workbook

def test_image_extraction():
    wb = Workbook()
    ws = wb.active

    # Need a small dummy image
    test_image_path = "test_image.png"
    with open(test_image_path, "wb") as f:
        # 1x1 transparent PNG
        f.write(b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\x0bIDATx\x9cc\xf8\xff\xff?\x00\x05\xfe\x02\xfe\xa7\x35\x81\x84\x00\x00\x00\x00IEND\xaeB`\x82')

    try:
        # Add image to worksheet
        ws.add_image(test_image_path, anchor="B2")
        
        # Save to flush drawing rels to the zip archive
        save_path = "test_image_save.xlsx"
        wb.save(save_path)
        
        # Test reading the image back from drawing
        drawing = ws._sheet.drawing()
        assert drawing.image_count() == 1
        
        img_item = drawing.image(0)
        assert img_item.row() == 1  # 0-indexed internally
        assert img_item.col() == 1  # 0-indexed internally

        binary_data = img_item.image_binary()
        assert len(binary_data) > 0
        assert binary_data.startswith(b'\x89PNG')
        
        # Explicitly close to release file handle on Windows
        wb.close()
    finally:
        if os.path.exists(test_image_path):
            os.remove(test_image_path)
        if os.path.exists("test_image_save.xlsx"):
            os.remove("test_image_save.xlsx")

