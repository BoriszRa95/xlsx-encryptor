"""Hide any file inside the cell colours of an Excel sheet, and get it back byte for byte.

Usage:
    python3 weird.py        # menu: 1 to hide, 2 to extract
"""

import hashlib
import struct
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

MAGIC = b"WSTR"
MAX_CELLS = 1_048_575  # rows available in one Excel column


def pack(data: bytes, name: str) -> bytes:
    """Wrap a file in a header: MAGIC, name length (1 byte), data length (4 bytes), name."""
    name_b = name.encode()
    return MAGIC + struct.pack(">BI", len(name_b), len(data)) + name_b + data


def unpack(blob: bytes) -> tuple[str, bytes]:
    """Read the header back and return (filename, file bytes)."""
    if blob[:4] != MAGIC:
        raise ValueError("not one of ours: no WSTR marker")
    name_len, data_len = struct.unpack(">BI", blob[4:9])
    name = blob[9 : 9 + name_len].decode()
    data = blob[9 + name_len : 9 + name_len + data_len]
    if len(data) != data_len:
        raise ValueError("data is truncated")
    return name, data


def blob_to_xlsx(blob: bytes, out: Path) -> None:
    """Write every 3 bytes as one cell colour (red, green, blue) down column A."""
    blob += bytes((-len(blob)) % 3)
    cells = len(blob) // 3
    if cells > MAX_CELLS:
        raise ValueError(f"file too big: {cells} cells, limit is {MAX_CELLS}")

    wb = Workbook()
    ws = wb.active
    for i in range(0, len(blob), 3):
        r, g, b = blob[i], blob[i + 1], blob[i + 2]
        ws.cell(row=i // 3 + 1, column=1).fill = PatternFill(
            "solid", fgColor="FF%02X%02X%02X" % (r, g, b)
        )
    wb.save(out)


def xlsx_to_blob(xlsx: Path) -> bytes:
    """Read cell colours from column A back into one byte string."""
    wb = load_workbook(xlsx)
    ws = wb.active
    parts = []
    for row in range(1, ws.max_row + 1):
        rgb = ws.cell(row=row, column=1).fill.fgColor.rgb
        parts.append(bytes([int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)]))
    return b"".join(parts)


def ask(prompt: str) -> str:
    """Read a line from stdin, tolerating end of input (pipes, scripts)."""
    try:
        return input(prompt).strip().strip('"')
    except EOFError:
        print()
        return ""


def main() -> None:
    print("1 - hide a file in Excel")
    print("2 - extract a file from Excel")
    mode = ask("choose 1 or 2: ")

    if mode == "1":
        src = Path(ask("file to hide: "))
        if not src.is_file():
            sys.exit(f"not a file: {src}")
        out = Path(ask("save the .xlsx as: ") or "carrier.xlsx")
        if out.is_dir():
            out = out / "carrier.xlsx"
        out.parent.mkdir(parents=True, exist_ok=True)
        data = src.read_bytes()
        blob_to_xlsx(pack(data, src.name), out)
        if not out.exists():
            sys.exit("nothing was written, check blob_to_xlsx")
        print(f"saved: {out} ({len(data)} bytes hidden)")
        print("sha256 of original:", hashlib.sha256(data).hexdigest())

    elif mode == "2":
        src = Path(ask(".xlsx to read: "))
        if not src.is_file():
            sys.exit(f"not a file: {src}")
        name, data = unpack(xlsx_to_blob(src))
        out = Path(ask("save the file as: ") or name)
        if out.is_dir():
            out = out / name
        if out.exists():
            out = out.with_name(f"{out.stem}_restored{out.suffix}")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(data)
        print(f"saved: {out} ({len(data)} bytes)")
        print("sha256 of restored:", hashlib.sha256(data).hexdigest())

    else:
        sys.exit("enter 1 or 2")


if __name__ == "__main__":
    main()
