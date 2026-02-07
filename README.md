# pyopenxlsx

`pyopenxlsx` 是一个基于 [OpenXLSX](https://github.com/troldal/OpenXLSX) C++ 库的高性能 Python Excel 绑定库。它旨在提供比纯 Python 库（如 openpyxl）更快的读写速度，同时保持 Pythonic 的 API 设计。

## 核心特性

- **高性能**: 底层使用现代 C++17 编写的 OpenXLSX 库。
- **Pythonic API**: 提供了符合 Python 习惯的接口（如属性访问、迭代器、上下文管理器）。
- **异步支持**: 关键的 I/O 操作支持 `async/await`。
- **样式支持**: 支持字体、填充、边框、对齐方式和数字格式。
- **内存安全**: 结合了 C++ 的效率和 Python 的自动内存管理。

## 技术栈

| 组件 | 技术 |
|------|------|
| **C++ Core** | [OpenXLSX](https://github.com/troldal/OpenXLSX) |
| **Bindings** | [pybind11](https://github.com/pybind/pybind11) |
| **Build System** | [scikit-build-core](https://github.com/scikit-build/scikit-build-core) & [CMake](https://cmake.org/) |
| **Package Management** | [uv](https://github.com/astral-sh/uv) |
| **Testing** | [pytest](https://pytest.org/) & [pytest-cov](https://github.com/pytest-dev/pytest-cov) |

## 安装

### 从源码安装

```bash
# 使用 uv (推荐)
uv pip install .

# 或者使用 pip
pip install .
```

### 开发模式安装

```bash
uv pip install -e .
```

## 快速开始

### 创建并保存工作簿

```python
from pyopenxlsx import Workbook

# 创建新工作簿
with Workbook() as wb:
    ws = wb.active
    ws.title = "MySheet"
    
    # 写入数据
    ws["A1"].value = "Hello"
    ws["B1"].value = 42
    ws.cell(row=2, column=1).value = 3.14
    
    # 保存
    wb.save("example.xlsx")
```

### 读取工作簿

```python
from pyopenxlsx import load_workbook

wb = load_workbook("example.xlsx")
ws = wb["MySheet"]
print(ws["A1"].value)  # 输出: Hello
wb.close()
```

### 异步操作

```python
import asyncio
from pyopenxlsx import load_workbook_async

async def main():
    wb = await load_workbook_async("example.xlsx")
    ws = wb.active
    print(ws["A1"].value)
    await wb.save_async("example_saved.async.xlsx")
    await wb.close_async()

asyncio.run(main())
```

### 设置样式

```python
from pyopenxlsx import Workbook, Font, Fill, Border, Side, Alignment, XLColor

wb = Workbook()
ws = wb.active

# 创建样式组件
font = Font(name="Arial", size=14, bold=True, color=XLColor(255, 0, 0))
fill = Fill(pattern_type="solid", color=XLColor(255, 255, 0))
border = Border(
    left=Side(style="thin", color=XLColor(0, 0, 0)),
    right=Side(style="thin"),
    top=Side(style="thick"),
    bottom=Side(style="thin")
)
alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 应用样式
style_idx = wb.add_style(font=font, fill=fill, border=border, alignment=alignment)
ws["A1"].value = "Styled Cell"
ws["A1"].style_index = style_idx

wb.save("styles.xlsx")
```

### 插入图片

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 插入图片到 A1 单元格，自动保持宽高比
# 需要安装 Pillow: pip install pillow
ws.add_image("logo.png", anchor="A1", width=200)

# 或者指定精确的宽高
ws.add_image("banner.jpg", anchor="B5", width=400, height=100)

wb.save("images.xlsx")
```

---

## API 文档

### 模块导入

```python
from pyopenxlsx import (
    # 核心类
    Workbook, Worksheet, Cell, Range,
    load_workbook, load_workbook_async,
    
    # 样式类
    Font, Fill, Border, Side, Alignment, Style, Protection,
    
    # 枚举和常量
    XLColor, XLSheetState, XLLineStyle, XLPatternType, XLAlignmentStyle,
    XLProperty,
)
```

---

### `Workbook` 类

工作簿是 Excel 文件的顶层容器。

#### 构造函数

```python
Workbook(filename: str | None = None)
```

- `filename`: 可选。如果提供，则打开现有文件；否则创建新工作簿。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `active` | `Worksheet \| None` | 获取或设置当前活动的工作表。 |
| `sheetnames` | `list[str]` | 返回所有工作表名称的列表。 |
| `properties` | `DocumentProperties` | 访问文档属性（标题、作者等）。 |
| `styles` | `XLStyles` | 访问底层样式对象（高级用法）。 |
| `workbook` | `XLWorkbook` | 访问底层 C++ 工作簿对象（高级用法）。 |

#### 方法

| 方法 | 返回类型 | 描述 |
|------|----------|------|
| `save(filename=None)` | `None` | 保存工作簿。如果未指定 `filename`，则保存到原始路径。 |
| `save_async(filename=None)` | `Coroutine` | 异步保存工作簿。 |
| `close()` | `None` | 关闭工作簿并释放资源。 |
| `close_async()` | `Coroutine` | 异步关闭工作簿。 |
| `create_sheet(title=None, index=None)` | `Worksheet` | 创建新工作表。`title` 默认为 "Sheet1", "Sheet2" 等。 |
| `create_sheet_async(title=None, index=None)` | `Coroutine[Worksheet]` | 异步创建工作表。 |
| `remove(worksheet)` | `None` | 删除指定的工作表。 |
| `remove_async(worksheet)` | `Coroutine` | 异步删除工作表。 |
| `copy_worksheet(from_worksheet)` | `Worksheet` | 复制工作表并返回副本。 |
| `copy_worksheet_async(from_worksheet)` | `Coroutine[Worksheet]` | 异步复制工作表。 |
| `add_style(...)` | `int` | 创建新样式并返回样式索引。详见下方。 |
| `add_style_async(...)` | `Coroutine[int]` | 异步创建样式。 |

#### `add_style` 方法详解

```python
def add_style(
    font: Font | int | None = None,
    fill: Fill | int | None = None,
    border: Border | int | None = None,
    alignment: Alignment | None = None,
    number_format: str | int | None = None,
    protection: Protection | None = None,
) -> int:
```

**参数:**

- `font`: `Font` 对象、字体索引（int）或 `None`。
- `fill`: `Fill` 对象、填充索引（int）或 `None`。
- `border`: `Border` 对象、边框索引（int）或 `None`。
- `alignment`: `Alignment` 对象或 `None`。
- `number_format`: 格式字符串（如 `"0.00%"`）、内置格式 ID（如 `14` 表示日期）或 `None`。
- `protection`: `Protection` 对象或 `None`。

**返回:** 新创建的样式索引（`int`），可赋值给 `Cell.style_index`。

**示例:**

```python
# 使用 Style 对象一次性传递所有样式
from pyopenxlsx import Style, Font, Fill

style = Style(
    font=Font(bold=True),
    fill=Fill(color=XLColor(200, 200, 200)),
    number_format="0.00"
)
idx = wb.add_style(style)
```

#### 魔术方法

| 方法 | 描述 |
|------|------|
| `__getitem__(key)` | 通过名称获取工作表：`wb["Sheet1"]` |
| `__delitem__(key)` | 通过名称删除工作表：`del wb["Sheet1"]` |
| `__iter__()` | 迭代所有工作表：`for ws in wb: ...` |
| `__len__()` | 返回工作表数量：`len(wb)` |
| `__contains__(key)` | 检查工作表是否存在：`"Sheet1" in wb` |
| `__enter__() / __exit__()` | 支持上下文管理器：`with Workbook() as wb: ...` |

---

### `Worksheet` 类

工作表代表 Excel 文件中的一个 Sheet。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `title` | `str` | 获取或设置工作表名称。 |
| `index` | `int` | 获取或设置工作表在工作簿中的索引（从 0 开始）。 |
| `sheet_state` | `str` | 获取或设置可见性：`"visible"`, `"hidden"`, `"very_hidden"`。 |
| `max_row` | `int` | 返回已使用的最大行号。 |
| `max_column` | `int` | 返回已使用的最大列号。 |
| `rows` | `Iterator` | 按行迭代所有单元格。 |
| `merges` | `MergeCells` | 获取合并单元格信息。 |
| `protection` | `dict` | 获取工作表保护状态（只读）。 |

#### 方法

| 方法 | 返回类型 | 描述 |
|------|----------|------|
| `cell(row, column, value=None)` | `Cell` | 通过行列索引获取单元格（从 1 开始）。可选设置值。 |
| `range(address)` | `Range` | 获取单元格范围，如 `ws.range("A1:C3")`。 |
| `range(start, end)` | `Range` | 获取单元格范围，如 `ws.range("A1", "C3")`。 |
| `merge_cells(address)` | `None` | 合并单元格，如 `ws.merge_cells("A1:B2")`。 |
| `merge_cells_async(address)` | `Coroutine` | 异步合并单元格。 |
| `unmerge_cells(address)` | `None` | 取消合并单元格。 |
| `unmerge_cells_async(address)` | `Coroutine` | 异步取消合并。 |
| `append(iterable)` | `None` | 在最后一行之后追加一行数据。 |
| `append_async(iterable)` | `Coroutine` | 异步追加行。 |
| `set_column_format(column, style_index)` | `None` | 设置整列的默认样式。`column` 可以是 `int` 或 `str`（如 `"A"`）。 |
| `set_row_format(row, style_index)` | `None` | 设置整行的默认样式。 |
| `column(col)` | `Column` | 获取列对象，用于设置宽度等属性。 |
| `protect(...)` | `None` | 保护工作表。详见下方。 |
| `protect_async(...)` | `Coroutine` | 异步保护工作表。 |
| `unprotect()` | `None` | 取消工作表保护。 |
| `unprotect_async()` | `Coroutine` | 异步取消保护。 |
| `add_image(img_path, anchor, width, height)` | `None` | 插入图片。详见下方。 |
| `add_image_async(...)` | `Coroutine` | 异步插入图片。 |

#### `protect` 方法详解

```python
def protect(
    password: str | None = None,
    objects: bool = True,
    scenarios: bool = True,
    insert_columns: bool = False,
    insert_rows: bool = False,
    delete_columns: bool = False,
    delete_rows: bool = False,
    select_locked_cells: bool = True,
    select_unlocked_cells: bool = True,
) -> None:
```

#### `add_image` 方法详解

```python
def add_image(
    img_path: str,
    anchor: str = "A1",
    width: int | None = None,
    height: int | None = None,
) -> None:
```

- `img_path`: 图片文件路径（支持 PNG, JPG, GIF）。
- `anchor`: 图片左上角锚定的单元格地址。
- `width`, `height`: 图片尺寸（像素）。如果未提供，需要安装 Pillow 自动检测。

#### 魔术方法

| 方法 | 描述 |
|------|------|
| `__getitem__(key)` | 通过地址获取单元格：`ws["A1"]` |

---

### `Cell` 类

单元格是 Excel 中最基本的数据单元。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `value` | `Any` | 获取或设置单元格的值。支持 `str`, `int`, `float`, `bool`, `datetime`, `date`。 |
| `formula` | `Formula` | 获取或设置单元格公式（不含 `=` 前缀）。 |
| `style_index` | `int` | 获取或设置单元格的样式索引。 |
| `style` | `int` | `style_index` 的别名。 |
| `is_date` | `bool` | 判断单元格是否应用了日期格式。 |
| `comment` | `str \| None` | 获取或设置单元格批注。设置为 `None` 删除批注。 |
| `font` | `XLFont` | 获取单元格的字体对象（只读）。 |
| `fill` | `XLFill` | 获取单元格的填充对象（只读）。 |
| `border` | `XLBorder` | 获取单元格的边框对象（只读）。 |
| `alignment` | `XLAlignment` | 获取单元格的对齐对象（只读）。 |

#### 日期处理

当 `is_date` 为 `True` 时，`value` 属性会自动将 Excel 序列号转换为 Python `datetime` 对象：

```python
from datetime import datetime
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 设置日期值和格式
ws["A1"].value = datetime(2024, 1, 15)
style_idx = wb.add_style(number_format=14)  # 内置日期格式
ws["A1"].style_index = style_idx

# 读取时自动转换
print(ws["A1"].value)  # datetime.datetime(2024, 1, 15, 0, 0)
print(ws["A1"].is_date)  # True
```

#### 公式设置

**重要提示**：公式必须通过 `formula` 属性设置，而不是 `value` 属性。通过 `value` 设置以 `=` 开头的字符串会被当作普通文本处理。

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# ✅ 正确：使用 formula 属性（不需要 = 前缀）
ws["A1"].value = 10
ws["A2"].value = 20
ws["A3"].formula = "A1+A2"        # 设置公式
ws["A4"].formula = "SUM(A1:A2)"   # SUM 公式

# ❌ 错误：使用 value 会被当作字符串
ws["B1"].value = "=A1+A2"  # 这会显示为文本 "=A1+A2"

# 读取公式
print(ws["A3"].formula.text)  # 输出: A1+A2

# 清除公式
ws["A3"].formula.clear()

wb.save("formulas.xlsx")
```

---

### `Range` 类

范围表示一个矩形的单元格区域。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `address` | `str` | 返回范围地址，如 `"A1:C3"`。 |
| `num_rows` | `int` | 返回范围包含的行数。 |
| `num_columns` | `int` | 返回范围包含的列数。 |

#### 方法

| 方法 | 返回类型 | 描述 |
|------|----------|------|
| `clear()` | `None` | 清空范围内所有单元格的值。 |
| `clear_async()` | `Coroutine` | 异步清空范围。 |

#### 迭代

```python
# 迭代范围内的所有单元格
for cell in ws.range("A1:B2"):
    print(cell.value)
```

---

### 样式类

#### `Font` 类

```python
Font(
    name: str = "Arial",
    size: int = 11,
    bold: bool = False,
    italic: bool = False,
    color: XLColor | str | None = None,
)
```

#### `Fill` 类

```python
Fill(
    pattern_type: XLPatternType | str = XLPatternType.Solid,
    color: XLColor | str | None = None,
    background_color: XLColor | str | None = None,
)
```

#### `Border` 类

```python
Border(
    left: Side | str | None = None,
    right: Side | str | None = None,
    top: Side | str | None = None,
    bottom: Side | str | None = None,
    diagonal: Side | str | None = None,
)
```

#### `Side` 类

```python
Side(
    style: XLLineStyle | str = XLLineStyle.Thin,
    color: XLColor | str | None = None,
)
```

**可用的线条样式:** `"thin"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`, `"hair"`, `"medium"`, `"mediumDashed"`, `"mediumDashDot"`, `"mediumDashDotDot"`, `"slantDashDot"`

#### `Alignment` 类

```python
Alignment(
    horizontal: XLAlignmentStyle | str | None = None,
    vertical: XLAlignmentStyle | str | None = None,
    wrap_text: bool = False,
)
```

**可用的对齐方式:** `"left"`, `"center"`, `"right"`, `"general"`, `"top"`, `"bottom"`

#### `Protection` 类

```python
Protection(
    locked: bool = True,
    hidden: bool = False,
)
```

#### `Style` 类

组合样式容器，可一次性传递给 `add_style()`：

```python
Style(
    font: Font | None = None,
    fill: Fill | None = None,
    border: Border | None = None,
    alignment: Alignment | None = None,
    number_format: str | int | None = None,
    protection: Protection | None = None,
)
```

---

### `DocumentProperties` 类

通过 `Workbook.properties` 访问。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `title` | `str` | 文档标题。 |
| `subject` | `str` | 文档主题。 |
| `creator` | `str` | 创建者。 |
| `keywords` | `str` | 关键词。 |
| `description` | `str` | 描述/备注。 |
| `last_modified_by` | `str` | 最后修改者。 |
| `category` | `str` | 分类。 |
| `company` | `str` | 公司名称。 |

#### 字典式访问

```python
# 读取
print(wb.properties["title"])
print(wb.properties["creator"])

# 写入
wb.properties["title"] = "My Report"
wb.properties["creator"] = "Python Script"

# 删除
del wb.properties["keywords"]
```

---

### `Column` 类

通过 `Worksheet.column()` 获取。

#### 属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `width` | `float` | 获取或设置列宽（字符单位）。 |
| `hidden` | `bool` | 获取或设置列是否隐藏。 |
| `style_index` | `int` | 获取或设置列的默认样式索引。 |

---

### 辅助函数

#### `load_workbook`

```python
def load_workbook(filename: str) -> Workbook:
    """打开现有的 Excel 文件。"""
```

#### `load_workbook_async`

```python
async def load_workbook_async(filename: str) -> Workbook:
    """异步打开现有的 Excel 文件。"""
```

#### `is_date_format`

```python
def is_date_format(format_code: int | str) -> bool:
    """
    判断给定的格式代码是否表示日期/时间格式。
    
    - 对于 int：检查是否为内置日期格式 ID（14-22, 27-36, 45-47）。
    - 对于 str：启发式检查格式字符串中是否包含日期时间标记。
    """
```

---

### 枚举类型

#### `XLColor`

```python
XLColor(r: int, g: int, b: int)  # RGB 构造
XLColor(hex_string: str)         # 如 "#FF0000" 或 "FF0000"
```

#### `XLSheetState`

- `XLSheetState.Visible`
- `XLSheetState.Hidden`
- `XLSheetState.VeryHidden`

#### `XLLineStyle`

- `XLLineStyle.Thin`
- `XLLineStyle.Thick`
- `XLLineStyle.Dashed`
- `XLLineStyle.Dotted`
- `XLLineStyle.Double`
- `XLLineStyle.Hair`
- `XLLineStyle.Medium`
- `XLLineStyle.MediumDashed`
- `XLLineStyle.MediumDashDot`
- `XLLineStyle.MediumDashDotDot`
- `XLLineStyle.SlantDashDot`

#### `XLPatternType`

- `XLPatternType.None`
- `XLPatternType.Solid`
- `XLPatternType.Gray125`
- `XLPatternType.Gray0625`
- ... (更多图案类型)

#### `XLAlignmentStyle`

- `XLAlignmentStyle.General`
- `XLAlignmentStyle.Left`
- `XLAlignmentStyle.Center`
- `XLAlignmentStyle.Right`
- `XLAlignmentStyle.Top`
- `XLAlignmentStyle.Bottom`

---

## 运行测试

```bash
# 运行所有测试
uv run pytest

# 运行测试并查看覆盖率报告
uv run pytest --cov=src/pyopenxlsx --cov-report=term-missing

# 生成 HTML 覆盖率报告
uv run pytest --cov=src/pyopenxlsx --cov-report=html
# 报告位于 htmlcov/index.html
```

## 许可证

本项目采用 MIT 许可证。底层 OpenXLSX 库采用其自身的许可证。
