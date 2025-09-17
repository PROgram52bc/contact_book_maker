# Contact Book Maker

Generate a beautifully formatted, image-rich **Contact Book** PDF from an Excel spreadsheet.

This tool reads two sheets from a single workbook -- `info_current` and `info_previous` -- and lays out one contact per row with a photo, names, and optional details (address, phone, email, children). It can also create a table of contents and page numbers.

A quick demo output file can be found at [demo.pdf](demo.pdf).

---

## ✨ Features

- **Excel → PDF**: Reads `info_new.xlsx` and produces a ready-to-share PDF.
- **Photos auto-match by key**: Each row's `key` looks up a matching image file in `pictures/`.
- **Icons**: Optional icons for address, phone, email, and children.
- **Table of contents** (optional) and **page numbers** (optional).
- **Symmetric/reverse layout** controls.
- **Multi-language capable**: Fonts included for Latin and CJK text (see Fonts section).

## 🧱 Project Structure

    .
    ├── generate.py           # main script
    ├── requirements.txt      # Python dependencies
    ├── icons/
    │   ├── address.png
    │   ├── anonymous.jpg     # default photo if none found for a row
    │   ├── children.png
    │   ├── email.png
    │   └── phone.png
    ├── pictures/             # your contact photos go here (you create this)
    ├── HPSimplified_*.ttf    # bundled Latin font family
    ├── simkai.ttf            # bundled CJK font (KaiTi)
    └── msyh*.ttc             # bundled Microsoft YaHei TTCs (not used by default)

---

## 📦 Requirements

- **Python**: 3.9+ recommended (3.10+ ideal)
- **OS**: macOS, Windows, or Linux
- **Excel workbook**: `info_new.xlsx` with sheets:
  - `info_current`
  - `info_previous`

### Python Dependencies

Declared in `requirements.txt`:

    openpyxl
    pandas
    fpdf2
    imagesize
    pillow

---

## 🧰 Setup (macOS / Linux)

1) Install Python 3 (if not already)
   - macOS (Homebrew): `brew install python@3.11`
   - Debian/Ubuntu: `sudo apt-get update && sudo apt-get install -y python3 python3-venv python3-pip`

2) Create and activate a virtual environment

```bash
    cd /path/to/contact_book_maker
    python3 -m venv .venv
    source .venv/bin/activate
    python -m pip install --upgrade pip
    pip install -r requirements.txt
```

3) Prepare your data & photos
   - Place your Excel workbook at: `./info_new.xlsx`
   - Create a `pictures/` folder and add images named by each row's `key`
     (e.g., if `key = abc`, then `pictures/abc.jpg` or `pictures/abc.png`)

4) Run

```bash
    python generate.py                # outputs out_YYYYMMDDHHMMSS.pdf
    python generate.py mybook.pdf     # outputs to mybook.pdf
```

5) Deactivate venv (when done)

```bash
    deactivate
```

---

## 🧰 Setup (Windows PowerShell)

1) Install Python 3 from https://www.python.org/downloads/  
   Ensure "Add Python to PATH" is checked.

2) Create and activate a virtual environment

```
    cd C:\path\to\contact_book_maker
    py -3 -m venv .venv
    .venv\Scripts\Activate.ps1
    python -m pip install --upgrade pip
    pip install -r requirements.txt
```

3) Prepare your data & photos
   - Put `info_new.xlsx` in the project root.
   - Create `pictures\` and add photos named after each row's `key`
     (e.g., `pictures\abc.jpg`).

4) Run

```
    python generate.py              # outputs out_YYYYMMDDHHMMSS.pdf
    python generate.py mybook.pdf   # outputs mybook.pdf
```

5) Deactivate venv (when done)

```
    deactivate
```

If `Activate.ps1` is blocked, you may temporarily enable execution:
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

---

## 📄 Input Excel Format

`generate.py` expects two sheets in `info_new.xlsx`:

- `info_current`
- `info_previous`

Each sheet should have the following columns. Only `key` and `english_name` are strictly required; others are optional.

| Column             | Required | Purpose                                                             |
|--------------------|----------|---------------------------------------------------------------------|
| `key`              | ✅       | Unique ID for the row; used to find a matching photo in `pictures/` |
| `english_name`     | ✅       | Displayed prominently; used for section headings and TOC            |
| `chinese_name`     | ❌       | Optional secondary line under the English name                      |
| `children`         | ❌       | Text describing children (e.g., names)                              |
| `children_chinese` | ❌       | Chinese version of children text                                    |
| `address`          | ❌       | Postal address                                                      |
| `phone`            | ❌       | Phone number(s)                                                     |
| `email`            | ❌       | Email address(es)                                                   |

### Sheet Example (minimal)

| key     | english_name    |
| -----   | --------------- |
| alice_j | Alice Johnson   |
| bob_c   | Bob Chen        |

### Sheet Example (full)

| key     | english_name    | chinese_name   | children   | children_chinese   | address                  | phone            | email                  |
| -----   | --------------- | -------------- | ---------- | ------------------ | ------------------------ | ---------------- | ---------------------- |
| alice_j | Alice Johnson   |                | Evan, Mia  |                    | 123 Maple St, City, ST   | (555) 123-4567   | alice@example.com      |
| bob_c   | Bob Chen        | 陈博           | Ryan       | 瑞安               | 456 Oak Ave, City, ST    | (555) 987-6543   | bob.chen@example.com   |

---

## 🖼️ Photos

Place photos in `pictures/` with filenames based on the `key`.

- Allowed extensions: `.png`, `.jpg`, `.jpeg`
- If no matching file is found, the script uses `icons/anonymous.jpg`.

Examples:

    pictures/001.jpg
    pictures/002.png

---

## ▶️ How to Run

Basic:

```
    python generate.py
```

- Reads `info_new.xlsx` (sheets `info_current` and `info_previous`)
- Adds fonts and icons from the repo
- Generates a multi-page PDF with 3 entries per page
- Outputs `out_YYYYMMDDHHMMSS.pdf`

Custom output filename:

```
    python generate.py my_directory_2025.pdf

```

---

## ⚙️ Adjustable Parameters (edit `generate.py`)

Open `generate.py` and look for the configuration block near the top. Common parameters:

### Paths

```python
    image_dir     = "pictures"   # where your photos live
    icon_dir      = "icons"      # where the UI icons live
    default_image = os.path.join(icon_dir, "anonymous.jpg")
    icons = {
        "email":    os.path.join(icon_dir, "email.png"),
        "children": os.path.join(icon_dir, "children.png"),
        "address":  os.path.join(icon_dir, "address.png"),
        "phone":    os.path.join(icon_dir, "phone.png"),
    }
    fonts = {
        "kaiti": "./simkai.ttf",              # CJK font
        "hp":    "./HPSimplified_Rg.ttf",     # Latin font
    }
```

### Global switches

```python
    gen_toc          = True   # include a table of contents
    gen_page_num     = True   # include page numbers
    gen_header       = True   # include section headers
    symmetric_layout = True   # flip image/info on even pages
    reverse_layout   = True   # flip image/info on all pages (XOR with 'symmetric_layout')
```

### Layout and sizing

All units are in inches.

```python
    num_per_page   = 3     # entries per page (vertical stacking)
    item_height    = 2.0   # vertical space reserved per entry

    info_width     = 2.5   # width of the text area (including icon gutter)
    info_margin    = 0.0625
    icon_width     = 0.125

    img_width      = 1.5   # width reserved for the image column (including margins)
    img_margin     = 0.05

    header_height  = 0.2
    footer_height  = 0.2

    item_scale     = 1.0   # scale factor applied to the template and image sizing
```

### Field flow & icons

The script defines a top-down flow of fields; missing fields collapse vertically:

    english_name → chinese_name → children → children_chinese → address → phone → email

It also auto-positions icons next to the first line that applies:

```python
    icon_flow = [
        ("children_icon",  ["children", "children_chinese"], { 'type': 'I', ... }),
        ("address_icon",   ["address"],                      { 'type': 'I', ... }),
        ("phone_icon",     ["phone"],                        { 'type': 'I', ... }),
        ("email_icon",     ["email"],                        { 'type': 'I', ... }),
    ]
```

You can adjust order, add new icons, or remove any mapping. If you add new fields to your Excel sheets, add them to the `flow` list in the same section.

---

## 🧪 Examples

Generate with default filename:

```
    python generate.py
```

Generate with a custom filename:

```
    python generate.py ChurchDirectory_Fall2025.pdf
```

Change entries per page (e.g., 4 per page):

```
    # in generate.py
    num_per_page = 4
    item_height  = 1.7   # tweak to fit layout nicely
```

Disable symmetric layout and always keep image on the left:

```
    symmetric_layout = False
    reverse_layout   = False
```

Turn off table of contents:

```
    gen_toc = False
```

---

## 🖋️ Fonts

By default, the script loads:

- `hp` → `HPSimplified_Rg.ttf` (Latin; included)
- `kaiti` → `simkai.ttf` (CJK; included)

If you see missing glyphs, ensure the font you've assigned contains them, or add another TTF/OTF and register it in the `fonts` map, then reference that font name in the flow definitions.

**Note:** TTC files (`msyh.ttc`, etc.) are bundled but not referenced by default. `fpdf2` generally prefers TTF/OTF; stick to the provided TTFs unless you know you need TTC.

---

## 🪪 Output

- Portrait PDF sized to the calculated `page_width` × `page_height`, based on your layout values.
- Each entry includes:
  - Photo (from `pictures/` by `key`, otherwise `icons/anonymous.jpg`)
  - English name (required)
  - Optional lines in the configured order: Chinese name, children, address, phone, email
  - Icons auto-placed next to the first applicable line per icon type
- Optional TOC and page numbers if enabled.

---

## 🛠️ Troubleshooting

- "No module named …" → Activate your venv and `pip install -r requirements.txt`.
- "UnicodeEncodeError" / Missing characters → Use a font that supports your characters and ensure it is added via `fpdf.add_font` (already handled in the script for bundled fonts).
- Photos not showing → Check `pictures/<key>.jpg|png|jpeg` exists; otherwise the default `icons/anonymous.jpg` is used. Note that the keys should start with an alphabet.
- Wrong sheet names → Ensure your workbook contains sheets named exactly `info_current` and `info_previous`.
- Layout overflows or overlaps → Lower `num_per_page` or increase `item_height`; you can also reduce `item_scale`.

---

## 🧭 Roadmap / TODO

- Extract configs to a YAML file (so you won't need to edit `generate.py`).
- Provide a sample Excel sheet and example output PDF.
