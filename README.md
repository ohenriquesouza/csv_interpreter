# 📊 CSV Interpreter

Converts messy CSV files into clean, formatted Excel spreadsheets (.xlsx).

---

## 📸 Example

> Blur was applied to the images below to prevent data leakage.

**Before** — raw CSV, no formatting:

![Raw CSV](img/bluredcru.png)

**After** — processed and formatted spreadsheet:

![Formatted Excel](img/bluredarrumado.png)

---

## 🖥️ Desktop App (recommended)

Just run `CSVInterpreter.exe` — no Python or dependencies required.

![CSV Interpreter UI](img/UI.png)

1. Click **Selecionar arquivo CSV** and pick your file
2. Pick a **theme color** — all spreadsheet colors (header, footer, text) are derived automatically from your choice
3. Set how **empty cells** should appear in the output:
   - **N/A** — highlighted in red, italic grey text (default)
   - **Vazio** — blank cell, no formatting
   - **Zero** — fills with `0`, no formatting
4. Click **Processar e salvar** and choose the output folder
5. Toggle 🔒 before processing to generate a protected (read-only) spreadsheet

### Personalization

The theme system derives all spreadsheet colors from a single base color:

| Element | Derivation |
|---|---|
| Header fill | The chosen color |
| Header text | White or dark, based on luminance contrast |
| Footer fill | 72% lighter tint of the base color |
| Footer text | 45% darker shade of the base color |

Image columns (detected by name or file extension) are automatically moved to the rightmost position so data always comes first.

---

## 🐍 Running from source

### Requirements

Python 3.12+ is required. Install all dependencies with:

```bash
pip install -r requirements.txt
```

| Package | Version | Purpose |
|---|---|---|
| pandas | 2.2.3 | CSV reading and data manipulation |
| openpyxl | 3.1.5 | Excel file generation and formatting |
| pillow | 12.2.0 | Icon conversion (build only) |
| pyinstaller | 6.20.0 | .exe packaging (build only) |

> `tkinter` is used for the GUI and comes built-in with Python on Windows — no install needed.

### Run the GUI

```bash
python src/app.py
```

### Run via CLI (batch mode)

Place CSV files in the `data/` folder, then:

```bash
python src/main.py
```

Output files will be saved to the `output/` folder.
