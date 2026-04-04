# Images and Shapes API

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

# Alternatively, extract raw bytes directly from the worksheet's drawing:
ws = wb.active
drawing = ws._sheet.drawing()
for i in range(drawing.image_count()):
    img_item = drawing.image(i)
    raw_bytes = img_item.image_binary() # Direct access to the image binary data (e.g. PNG/JPEG)
```

## Advanced Binary Data Access

If you need to process images in memory (e.g., uploading to a cloud bucket or sending over an API) without saving to disk:

```python
for img in wb.get_embedded_images():
    # Fetch raw bytes of the image
    raw_bytes = wb.get_image_data(img.path)
    # Process raw_bytes...
```

---

## Inserting Vector Shapes

`pyopenxlsx` provides support for adding native Excel vector shapes (like Rectangles, Arrows, Diamonds, etc.) using `add_shape()`. Unlike raster images, shapes are drawn natively by Excel and scale perfectly.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # 1. Add an Arrow pointing right
    ws.add_shape(
        row=2, col=2, 
        shape_type="Arrow", 
        name="MyArrow", text="Important!", 
        fill_color="FF0000", line_width=2.5,
        rotation=90
    )
    
    # 2. Add a styled Cloud
    ws.add_shape(
        row=5, col=5, 
        shape_type="Cloud",
        name="Cloudy", text="Data",
        width=200, height=150,
        flip_h=True
    )
    
    wb.save("shapes.xlsx")
```

### Legacy VML Shapes

For advanced manipulation of legacy form controls or specific comment-like shape behavior, `pyopenxlsx` exposes the underlying VML Drawing APIs:

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

vml = ws._sheet.vml_drawing()
shape = vml.create_shape()

# Set properties
shape.set_fill_color("#00FF00")
shape.set_type("#_x0000_t202")
shape.set_stroked(True)

# Anchor the shape
client_data = shape.client_data()
client_data.set_move_with_cells(True)
client_data.set_size_with_cells(True)
client_data.set_anchor("3, 15, 3, 10, 5, 15, 5, 10") # format: LeftCol, LeftOffset, TopRow, TopOffset, RightCol, RightOffset, BottomRow, BottomOffset
client_data.set_row(3)     # Logical link to Row 4
client_data.set_column(3)  # Logical link to Col D

# Adjust the style
style = shape.style()
style.show() # Remove default hidden attribute
style.set_position("absolute")
style.set_width(120)
style.set_height(40)
shape.set_style_obj(style) # Apply back to the shape

wb.save("vml_shapes.xlsx")
```

### Supported Shape Options (`**kwargs`)
You can extensively configure your shape using the following arguments:
- **Dimensions & Anchor**: `row`, `col`, `width`, `height`, `offset_x`, `offset_y`.
- **Two-Cell Anchor**: `end_row`, `end_col`, `end_offset_x`, `end_offset_y` (allow the shape to resize automatically as cells scale).
- **Appearance**: `fill_color` (ARGB without `#`), `line_color` (ARGB without `#`), `line_width` (float).
- **Transformations**: `rotation` (degrees), `flip_h` (bool), `flip_v` (bool).
- **Outline Styles**: `line_dash` ("dash", "sysDash", "dot", etc.), `arrow_start`, `arrow_end` ("triangle", "stealth", "diamond").
- **Text & Alignment**: `text` (str), `horz_align` ("l", "ctr", "r"), `vert_align` ("t", "ctr", "b").
