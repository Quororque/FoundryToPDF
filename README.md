# FoundryVTT Session Exporter

Built with ChatGPT-5

This utility script converts exported Foundry Virtual Tabletop (FoundryVTT) chat logs (`.json` format) into a well-formatted `.docx` document and optionally exports a bookmarked `.pdf` using Microsoft Word.

It uses .json files exported with Vauxs' Archives module. For more information, visit https://foundryvtt.com/packages/vauxs-archives

---

## Features

- Convert multiple Foundry session JSON files to a single `.docx`
- Add **Cast section** with character portraits
- Clean HTML formatting in chat messages
- Automatically **remove consecutive duplicate messages**
- Insert session titles as **Word headings** for automatic PDF bookmarks
- Add automatic page numbering starting at the session section
- Export to PDF via Word COM using PowerShell, with **bookmarks preserved**
- Compatible with bookmark-aware PDF viewers such as Edge, Acrobat, and SumatraPDF

---

## Requirements

- Python 3.8+
- Microsoft Word installed (for PDF export)
- PowerShell (included with Windows)
- Required Python packages:
  ```bash
  pip install python-docx beautifulsoup4 colorama
  ```

---

## Directory Structure

```
project/
├── foundry_to_docx.py
├── sessions/           # place your Foundry JSON logs here
├── config/
│   ├── config.txt      # optional general settings
│   └── actors.txt      # optional cast list
├── portraits/          # optional character portraits (JPG)
├── export/             # generated DOCX and PDF files
└── omitted/            # auto-generated deleted duplicates log
```

---

## Configuration

You can override default settings. Refer to example files!

## Usage

1. Place your exported Foundry JSON logs in the `sessions/` folder.  
2. Configure `config.txt` and `actors.txt`.
3. Run the script with the bundled .bat file.
4. The script will generate:
   - A DOCX transcript in `export/`
   - A PDF with bookmarks (if Word + PowerShell are available)
   - A separate DOCX file under `export/omitted/` listing deleted duplicate messages.

If you set:
```
PRINT2PDF = NO
```
in `config.txt`, the script will skip PDF generation and only output DOCX.

---

## License

Licensed under the GNU General Public License v3.0 (GPLv3) – see the LICENSE file for details.

---

## Acknowledgments

- ChatGPT-5
- python-docx for DOCX manipulation  
- Beautiful Soup for HTML cleaning  
- Microsoft Word COM for PDF generation  
- SumatraPDF for lightweight bookmark testing
