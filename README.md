# SummaryPoints

![Tool Icon](asserts/iconPicture.png)

A desktop tool for merging PowerPoint files and exporting them as a single PDF.  
> Great for students and organized people

---

## Features

- **No setup required** — just run the `.exe`
- **Simple folder picker** (GUI-based)
- **Watermark removal**:
  Automatically removes  
  `"Evaluation Warning : The document was created with Spire.Presentation for Python"`
- **Supports multiple PowerPoint formats** (`.ppt`, `.pptx`, `.pps`, etc.)
- **Error handling** for corrupted or temporary files
- **Built using PyInstaller** – no Python or Office installation required

---

## Project Structure

**The `.exe` is under `Releases`**

```text
SummaryPoints/
├── concatPowerPoints.py # Main (and single) Python script
├── icon.ico # App icon
├── assets/
│ └── icon.png # Used in README
├── README.md # You're here 🙂
├── .gitignore # Ignore build artifacts
```

---

## How to Use

1. Run `SummaryPoints.exe`
2. Select a folder containing PowerPoint files
3. The tool will:
   - Merge all presentations
   - Remove Spire watermark
   - Save a clean `.pptx`
   - Convert it to `.pdf`
   - Delete the intermediate watermarked file

---

## Developer Notes

### Build the `.exe` Yourself

Create `dist/`, `build/`, and `pyinstaller.spec`, then run:

```bash
pyinstaller --onefile --noconsole --icon="icon.ico" ^
  --add-binary "path\to\Spire.Presentation.Base.dll;spire\presentation\lib" ^
  --add-binary "path\to\libSkiaSharp.dll;spire\presentation\lib" ^
  concatPowerPoints.py
```
Make sure you're in the right virtual environment and have all dependencies installed.

---

Powered by python-pptx, spire.presentation, pyinstaller, and tkinter