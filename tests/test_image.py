import pytest
import os
from pyopenxlsx import Workbook, ImageInfo
from pathlib import Path


def test_add_image(tmp_path):
    # Create a small 1x1 red PNG image
    img_path = tmp_path / "test.png"
    # PNG signature + IHDR + IDAT + IEND
    # This is a valid 1x1 red pixel PNG
    img_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xff\xff?0\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"
    img_path.write_bytes(img_data)

    xlsx_path = tmp_path / "test_img.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.add_image(str(img_path), anchor="B2", width=100, height=100)
    wb.save(str(xlsx_path))

    assert xlsx_path.exists()

    # Re-open and check (though we can't easily check image existence via API yet)
    wb2 = Workbook(str(xlsx_path))
    assert wb2.active.title == ws.title
    wb2.close()


def test_get_embedded_images(tmp_path):
    """Test extracting embedded images from an Excel file."""
    # Create a small 1x1 red PNG image
    img_path = tmp_path / "test.png"
    img_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xff\xff?0\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"
    img_path.write_bytes(img_data)

    xlsx_path = tmp_path / "test_img.xlsx"

    # Create workbook with image
    wb = Workbook()
    ws = wb.active
    ws.add_image(str(img_path), anchor="B2", width=100, height=100)
    wb.save(str(xlsx_path))
    wb.close()

    # Reopen and extract images
    wb2 = Workbook(str(xlsx_path))

    images = wb2.get_embedded_images()
    assert len(images) == 1

    img_info = images[0]
    assert isinstance(img_info, ImageInfo)
    assert img_info.extension == "png"
    assert img_info.name == "image1.png"
    assert img_info.path == "xl/media/image1.png"

    wb2.close()


def test_get_image_data(tmp_path):
    """Test getting image binary data."""
    # Create a small 1x1 red PNG image
    img_path = tmp_path / "test.png"
    img_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xff\xff?0\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"
    img_path.write_bytes(img_data)

    xlsx_path = tmp_path / "test_img.xlsx"

    # Create workbook with image
    wb = Workbook()
    ws = wb.active
    ws.add_image(str(img_path), anchor="B2", width=100, height=100)
    wb.save(str(xlsx_path))
    wb.close()

    # Reopen and get image data
    wb2 = Workbook(str(xlsx_path))

    # Get by full path
    data = wb2.get_image_data("xl/media/image1.png")
    assert isinstance(data, bytes)
    assert len(data) > 0
    # Check PNG signature
    assert data[:4] == b"\x89PNG"

    # Get by name only
    data2 = wb2.get_image_data("image1.png")
    assert data == data2

    wb2.close()


def test_extract_images(tmp_path):
    """Test extracting all images to a directory."""
    # Create a small 1x1 red PNG image
    img_path = tmp_path / "test.png"
    img_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xff\xff?0\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"
    img_path.write_bytes(img_data)

    xlsx_path = tmp_path / "test_img.xlsx"
    output_dir = tmp_path / "extracted_images"

    # Create workbook with image
    wb = Workbook()
    ws = wb.active
    ws.add_image(str(img_path), anchor="B2", width=100, height=100)
    wb.save(str(xlsx_path))
    wb.close()

    # Reopen and extract
    wb2 = Workbook(str(xlsx_path))
    extracted = wb2.extract_images(str(output_dir))

    assert len(extracted) == 1
    assert os.path.exists(extracted[0])

    # Verify the extracted file
    extracted_data = Path(extracted[0]).read_bytes()
    assert extracted_data[:4] == b"\x89PNG"

    wb2.close()


def test_get_image_not_found(tmp_path):
    """Test that getting non-existent image raises error."""
    xlsx_path = tmp_path / "test_empty.xlsx"

    wb = Workbook()
    wb.save(str(xlsx_path))
    wb.close()

    wb2 = Workbook(str(xlsx_path))

    with pytest.raises(RuntimeError, match="Image not found"):
        wb2.get_image_data("nonexistent.png")

    wb2.close()


def test_multiple_images(tmp_path):
    """Test handling multiple images."""
    # Create two different PNG images
    img_data1 = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xff\xff?0\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"
    img_data2 = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\x00\x00\x000\x00\x03\xff\x01\xfe\x8e\xfe\x1d\x00\x00\x00\x00IEND\xaeB`\x82"

    img_path1 = tmp_path / "test1.png"
    img_path2 = tmp_path / "test2.png"
    img_path1.write_bytes(img_data1)
    img_path2.write_bytes(img_data2)

    xlsx_path = tmp_path / "test_multi_img.xlsx"

    # Create workbook with two images
    wb = Workbook()
    ws = wb.active
    ws.add_image(str(img_path1), anchor="A1", width=50, height=50)
    ws.add_image(str(img_path2), anchor="C1", width=50, height=50)
    wb.save(str(xlsx_path))
    wb.close()

    # Reopen and check
    wb2 = Workbook(str(xlsx_path))
    images = wb2.get_embedded_images()

    assert len(images) == 2

    # Extract and verify
    output_dir = tmp_path / "multi_extracted"
    extracted = wb2.extract_images(str(output_dir))
    assert len(extracted) == 2

    wb2.close()
