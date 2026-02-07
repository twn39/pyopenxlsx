import pytest
import openpyxl
import asyncio
import numpy as np
from pyopenxlsx import Workbook as PyWorkbook, load_workbook_async


# Fixture to provide a temporary file path
@pytest.fixture
def temp_xlsx_file(tmp_path):
    return str(tmp_path / "benchmark.xlsx")


# --- Write Benchmarks ---


def write_pyopenxlsx(rows, cols, filepath):
    wb = PyWorkbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c).value = f"R{r}C{c}"
    wb.save(filepath)


def write_openpyxl(rows, cols, filepath):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=f"R{r}C{c}")
    wb.save(filepath)


@pytest.mark.benchmark(group="write_small")
def test_write_small_pyopenxlsx(benchmark, temp_xlsx_file):
    benchmark(write_pyopenxlsx, 100, 10, temp_xlsx_file)


@pytest.mark.benchmark(group="write_small")
def test_write_small_openpyxl(benchmark, temp_xlsx_file):
    benchmark(write_openpyxl, 100, 10, temp_xlsx_file)


@pytest.mark.benchmark(group="write_large")
def test_write_large_pyopenxlsx(benchmark, temp_xlsx_file):
    benchmark(write_pyopenxlsx, 1000, 50, temp_xlsx_file)


@pytest.mark.benchmark(group="write_large")
def test_write_large_openpyxl(benchmark, temp_xlsx_file):
    benchmark(write_openpyxl, 1000, 50, temp_xlsx_file)


# --- Optimized Write Methods ---


def write_pyopenxlsx_set_cell_value(rows, cols, filepath):
    """Use set_cell_value() - bypasses Python Cell object creation"""
    wb = PyWorkbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.set_cell_value(r, c, f"R{r}C{c}")
    wb.save(filepath)


def write_pyopenxlsx_write_rows(rows, cols, filepath):
    """Use write_rows() - batch write Python lists"""
    wb = PyWorkbook()
    ws = wb.active
    data = [[f"R{r}C{c}" for c in range(1, cols + 1)] for r in range(1, rows + 1)]
    ws.write_rows(1, data)
    wb.save(filepath)


@pytest.mark.benchmark(group="write_large")
def test_write_large_set_cell_value(benchmark, temp_xlsx_file):
    """Test optimized set_cell_value() method"""
    benchmark(write_pyopenxlsx_set_cell_value, 1000, 50, temp_xlsx_file)


@pytest.mark.benchmark(group="write_large")
def test_write_large_write_rows(benchmark, temp_xlsx_file):
    """Test optimized write_rows() method with string data"""
    benchmark(write_pyopenxlsx_write_rows, 1000, 50, temp_xlsx_file)


def write_pyopenxlsx_bulk(rows, cols, filepath):
    wb = PyWorkbook()
    ws = wb.active
    # Create a numpy array of strings (or numbers, depending on support)
    # The original test used strings "R{r}C{c}".
    # Let's see if write_range supports string arrays or just numeric.
    # The docstring says "2D numpy array of doubles" for get_range_values (read),
    # but for write_range it says "2D numpy array or any object supporting the buffer protocol".
    # String arrays in numpy can be tricky for C++ buffer protocols unless handled specifically.
    # Let's try numeric first as it's the standard use case for "high performance",
    # or creates a numeric array.
    data = np.arange(rows * cols, dtype=np.float64).reshape(rows, cols)
    ws.write_range(1, 1, data)
    wb.save(filepath)


@pytest.mark.benchmark(group="write_large")
def test_write_large_bulk_pyopenxlsx(benchmark, temp_xlsx_file):
    benchmark(write_pyopenxlsx_bulk, 1000, 50, temp_xlsx_file)


# --- Read Benchmarks ---


@pytest.fixture
def large_file(tmp_path):
    filepath = str(tmp_path / "large_input.xlsx")
    wb = openpyxl.Workbook()  # Use openpyxl to generate the file to be fair/consistent
    ws = wb.active
    for r in range(1, 1001):
        for c in range(1, 21):
            ws.cell(row=r, column=c, value=f"Val_{r}_{c}")
    wb.save(filepath)
    return filepath


def read_pyopenxlsx(filepath):
    wb = PyWorkbook(filepath)
    ws = wb.active
    val = ws.cell(row=500, column=10).value
    return val


def read_openpyxl(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    val = ws.cell(row=500, column=10).value
    return val


@pytest.mark.benchmark(group="read")
def test_read_pyopenxlsx(benchmark, large_file):
    benchmark(read_pyopenxlsx, large_file)


@pytest.mark.benchmark(group="read")
def test_read_openpyxl(benchmark, large_file):
    benchmark(read_openpyxl, large_file)


# --- Iteration Benchmarks ---


def iterate_pyopenxlsx(filepath):
    wb = PyWorkbook(filepath)
    ws = wb.active
    count = 0
    # iterate over existing cells in the generated range
    # Note: access pattern might differ slightly, trying to match basic usage
    # Assuming we know the dimension or iterating available cells if API supports it
    # For now, let's iterate a known range like in the write test
    for r in range(1, 1001):
        for c in range(1, 21):
            ws.cell(row=r, column=c).value
            count += 1
    return count


def iterate_openpyxl(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    count = 0
    for row in ws.iter_rows(min_row=1, max_row=1000, min_col=1, max_col=20):
        for cell in row:
            count += 1
    return count


@pytest.mark.benchmark(group="iterate")
def test_iterate_pyopenxlsx(benchmark, large_file):
    benchmark(iterate_pyopenxlsx, large_file)


@pytest.mark.benchmark(group="iterate")
def test_iterate_openpyxl(benchmark, large_file):
    benchmark(iterate_openpyxl, large_file)


# --- Async/Concurrent Benchmarks ---


@pytest.fixture
def multiple_files(tmp_path):
    # Generate 10 medium sized files
    files = []
    data = np.arange(100 * 50, dtype=np.float64).reshape(100, 50)
    for i in range(10):
        fp = str(tmp_path / f"async_bench_{i}.xlsx")
        wb = PyWorkbook()
        ws = wb.active
        ws.write_range(1, 1, data)
        wb.save(fp)
        files.append(fp)
    return files


def read_files_sync(files):
    for fp in files:
        wb = PyWorkbook(fp)
        ws = wb.active
        _ = ws.get_range_data(1, 1, 100, 50)


async def read_files_async(files):
    async def read_one(fp):
        wb = await load_workbook_async(fp)
        ws = wb.active
        _ = await ws.get_range_data_async(1, 1, 100, 50)

    await asyncio.gather(*(read_one(fp) for fp in files))


def run_async_benchmark(benchmark, func, *args):
    # Helper to run async function in benchmark
    def wrapper():
        asyncio.run(func(*args))

    benchmark(wrapper)


@pytest.mark.benchmark(group="async_read")
def test_read_concurrent_sync(benchmark, multiple_files):
    benchmark(read_files_sync, multiple_files)


@pytest.mark.benchmark(group="async_read")
def test_read_concurrent_async(benchmark, multiple_files):
    run_async_benchmark(benchmark, read_files_async, multiple_files)


@pytest.fixture
def output_dir(tmp_path):
    d = tmp_path / "async_out"
    d.mkdir()
    return d


def write_files_sync(output_dir):
    data = np.arange(100 * 50, dtype=np.float64).reshape(100, 50)
    for i in range(10):
        fp = str(output_dir / f"sync_out_{i}.xlsx")
        wb = PyWorkbook()
        ws = wb.active
        ws.write_range(1, 1, data)
        wb.save(fp)


async def write_files_async(output_dir):
    data = np.arange(100 * 50, dtype=np.float64).reshape(100, 50)

    async def write_one(i):
        fp = str(output_dir / f"async_out_{i}.xlsx")
        wb = PyWorkbook()  # Constructor is sync/fast enough usually
        ws = wb.active
        await ws.write_range_async(1, 1, data)
        await wb.save_async(fp)

    await asyncio.gather(*(write_one(i) for i in range(10)))


@pytest.mark.benchmark(group="async_write")
def test_write_concurrent_sync(benchmark, output_dir):
    benchmark(write_files_sync, output_dir)


@pytest.mark.benchmark(group="async_write")
def test_write_concurrent_async(benchmark, output_dir):
    run_async_benchmark(benchmark, write_files_async, output_dir)


def write_files_loop_sync(output_dir):
    # Sequential loop write
    for i in range(5):  # Reduce count as loop write is slow
        fp = str(output_dir / f"loop_sync_{i}.xlsx")
        write_pyopenxlsx(100, 50, fp)


async def write_files_loop_async(output_dir):
    # Concurrent loop write using threads
    async def write_one(i):
        fp = str(output_dir / f"loop_async_{i}.xlsx")
        # Run the synchronous loop write in a thread
        await asyncio.to_thread(write_pyopenxlsx, 100, 50, fp)

    await asyncio.gather(*(write_one(i) for i in range(5)))


@pytest.mark.benchmark(group="async_loop_write")
def test_write_loop_concurrent_sync(benchmark, output_dir):
    benchmark(write_files_loop_sync, output_dir)


@pytest.mark.benchmark(group="async_loop_write")
def test_write_loop_concurrent_async(benchmark, output_dir):
    run_async_benchmark(benchmark, write_files_loop_async, output_dir)
