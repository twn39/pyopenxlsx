# Images API

`pyopenxlsx` allows you to embed images directly into worksheets. The image anchoring and scaling logic relies on the Pillow (`PIL`) library, which must be installed for auto-sizing to work.

## Inserting an Image

Images are inserted using the `add_image()` method on a `Worksheet` object.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # 1. Simple Insertion (Auto-detected dimensions)
    # The top-left corner of the image will be at cell B2.
    # Note: Pillow (pip install pillow) must be installed to auto-detect width and height.
    ws.add_image("logo.png", anchor="B2")
    
    # 2. Resized Insertion (Maintain aspect ratio based on width)
    ws.add_image("photo.jpg", anchor="D5", width=300)
    
    # 3. Exact Dimensions Insertion (Overrides original aspect ratio)
    ws.add_image("banner.png", anchor="A10", width=600, height=100)
    
    wb.save("images.xlsx")
```

## Extracting Images

You can extract all embedded images from an existing workbook.

```python
from pyopenxlsx import load_workbook

wb = load_workbook("images.xlsx")

# Retrieve a list of ImageInfo objects describing the embedded images
image_infos = wb.get_embedded_images()
for img in image_infos:
    print(f"Found image: {img.name} (Type: {img.extension})")

# Extract all images to a specific directory on disk
extracted_paths = wb.extract_images("output_folder/")
print(f"Extracted to: {extracted_paths}")
```

## Advanced Binary Data Access

If you need to process images in memory (e.g., uploading to a cloud bucket or sending over an API) without saving to disk:

```python
for img in wb.get_embedded_images():
    # Fetch raw bytes of the image
    raw_bytes = wb.get_image_data(img.path)
    # Process raw_bytes...
```
