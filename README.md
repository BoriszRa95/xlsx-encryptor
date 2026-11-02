# Weird Storage

A Python tool that hides any file inside the cell colours of an Excel sheet, and gets it back byte for byte.

Each cell's background colour is three bytes: red, green and blue. So every 3 bytes of a file become one cell down column A. The picture opens as a normal `.xlsx`, and nothing about the file is visible from the outside.

## How it works

1. The file is read as bytes (numbers 0-255).
2. A small header is glued in front:
   - `WSTR` — 4-byte marker, so we know the sheet is ours
   - name length — 1 byte (`B`), filenames up to 255 characters
   - file length — 4 bytes (`I`), files up to 4 GB
   - the filename itself
3. Every 3 bytes become the RGB of one cell in column A.
4. Reading back is the same in reverse: cell colour -> 3 bytes -> cut the header off by its lengths.
5. The result is checked with SHA-256: equal hashes mean the bytes are identical.

The header is what makes extraction possible: without the filename and length you would not know where the file ends, because the last cell is padded with zeros.

## Usage
```
locate the file

python3 xslx_encoder.py
```

Then pick a mode:

- `1` — hide a file: enter the file path, then where to save the `.xlsx`
- `2` — extract a file: enter the `.xlsx` path, then where to save the file

You can also run it non-interactively by piping the answers:

```
printf "1\nphoto.jpg\ncarrier.xlsx\n" | python3 xslx_encoder.py
```

## Limits

- One Excel column holds 1 048 576 rows, so about 3 MB per file.
- Bigger files take longer and need a second column or a second sheet.
- The `.xlsx` grows with the file size; small files stay under a few hundred KB.

## Checking it works

Hashes from a real run (a 50 400-byte JPEG):

```
original:   d3ea0d54e265c8ff98ec17bcffe16a2a3a41018e31258a9e008baff67b7d6d13
restored:   d3ea0d54e265c8ff98ec17bcffe16a2a3a41018e31258a9e008baff67b7d6d13
```

`cmp` reports the two files are identical.

## Files

- `xlsx_encoder` — the tool
- `carrier.xlsx` — an example sheet with a file already hidden inside, so you can extract it yourself
