# Rows → Forms

Turn each row of a **CSV / XLSX** file into a clean, copyable form — entirely in your browser. No server, no upload, no build step. Just open the page and drop in a file.

---

## Features

- 📄 **CSV & Excel support** — `.csv`, `.xlsx`, `.xls` (parsed locally via [SheetJS](https://sheetjs.com/))
- 🧾 **One section per row** — every data row becomes a full-width, collapsible form
- 🏷️ **Configurable header row** — pick which row holds your column names (handy when there's preamble above the data)
- 🔖 **Custom section titles** — choose which column's value becomes each form's heading
- 📋 **Copy anything** — per-field copy buttons, plus a **Copy JSON** button that copies the whole row as a `{ "Column": "value" }` object
- 👁️ **Field visibility** — show/hide individual fields across all forms, with a searchable field list
- ✅ **Form selection** — tick the forms you care about and toggle **“show selected only”**
- 🔍 **Search / filter** across all forms
- 💾 **Saved sessions** — parsed files and your settings are stored in `localStorage`, so you can reload them without re-uploading
- 🌗 **Light & dark mode** (follows your OS)

---

## Usage

1. Open `index.html` in any modern browser (double-click it, or `open index.html`).
2. **Drop** a file onto the page or click **browse** (try the included `sample.csv`).
3. Use the config bar to:
   - set the **Header row**
   - pick the **Section title column**
   - choose **Visible fields**
   - select forms and toggle **Showing all / selected only**
4. Click any field's copy icon, or **Copy JSON** to grab a whole row.

Your work is auto-saved. Returning to the landing screen (**New file**) shows your **Saved sessions** — click one to pick up where you left off.

---

## Project structure

```
rows-to-form/
├── index.html    # markup + SheetJS CDN include
├── styles.css    # minimal, modern styling (light/dark)
├── app.js        # all app logic (parsing, rendering, storage)
├── sample.csv    # example data
└── README.md
```

---

## Notes & limitations

- **Internet is required for the library** — SheetJS is loaded from a CDN (`cdn.jsdelivr.net`). All parsing, rendering, and copying still happen **100% locally**; no data ever leaves the page. To run fully offline, download `xlsx.full.min.js` and swap the CDN `<script>` tag for the local file.
- **Sessions are stored in `localStorage`** (~5 MB per site). Very large spreadsheets may exceed this; if a save fails you'll see a toast and that session simply won't persist (the current view still works).
- Only the **first sheet** of an Excel workbook is read.

---

## Tech

Vanilla HTML, CSS, and JavaScript — no framework, no bundler. Single static page.

## License

MIT
