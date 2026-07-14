from typing import Any, Optional

def is_external_url(target: str) -> bool: ...
def link(
    worksheet: Any,
    cell_ref: str,
    target: str,
    *,
    text: Optional[str] = ...,
    tooltip: str = ...,
    internal: Optional[bool] = ...,
) -> None: ...
def unlink(worksheet: Any, cell_ref: str) -> None: ...
