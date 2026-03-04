# 📄 → ⚡ DOCX to VBA Converter

> Turn any Word document into a self-replicating VBA macro — paste it on another PC and rebuild the document from scratch.

---

## What it does

Drop in a `.docx` file and this tool generates a complete VBA macro that, when pasted into Word's editor on **any other computer**, recreates the entire document — no file transfer needed.

Just copy the code, paste it into Word, hit F5, and your document rebuilds itself.

---

## Features

- **Preserves text formatting** — fonts, sizes, bold, italic, underline
- **Preserves colors** — RGB text colors converted to VBA-safe values
- **Handles tables** — recreates table structure and cell content
- **Paragraph styles** — Heading 1–5, Normal, Title, Subtitle, List Paragraph, and more
- **Text alignment** — left, center, right, justified
- **Page breaks** — inserted at the correct positions
- **Long document support** — auto-splits into multiple `DocPart001`, `DocPart002`... subs to stay within VBA's procedure size limit
- **Syntax highlighted output** — color-coded VBA right in the browser
- **One-click copy** — big green copy button so you can't miss it
- **Download as `.bas`** — save the macro file directly

---

## How to use

**1. Generate the macro**

Open `docx-to-vba.html` in any browser. No server, no install — it runs entirely client-side.

Drop your `.docx` file onto the upload zone (or click to browse), then hit **⚡ Generate VBA Macro**.

**2. Copy the code**

Click the green **⎘ COPY MACRO** button that appears above the output.

**3. Paste into Word on the target PC**

```
Alt + F11          → opens the VBA editor
Insert → Module    → creates a new module
Ctrl + V           → paste the macro
F5                 → run it
```

A new document will be created and populated automatically.

---

## Supported formatting

| Feature | Supported |
|---|---|
| Font name | ✅ |
| Font size | ✅ |
| Bold / Italic / Underline | ✅ |
| Text color (RGB) | ✅ |
| Paragraph alignment | ✅ |
| Paragraph styles (Heading, Normal, etc.) | ✅ |
| Tables (text content) | ✅ |
| Page breaks | ✅ |
| Images | ❌ (VBA limitation) |
| Headers / Footers | ❌ |
| Comments / Track changes | ❌ |

---

## Options

Toggle these in the sidebar before generating:

| Option | Default | Effect |
|---|---|---|
| Preserve font styles | On | Includes font name and size |
| Preserve text colors | On | Includes RGB color per run |
| Include tables | On | Generates table creation code |
| Page breaks | On | Inserts `wdPageBreak` at correct spots |
| Compact output | Off | Strips comments for smaller macro |

---

## Technical notes

**Why VBA and not just copy the file?**

Sometimes you can't transfer files — restricted machines, locked-down environments, air-gapped systems. A VBA macro is just text you can paste into any email, chat, or terminal and run immediately.

**How the parser works**

A `.docx` is a ZIP archive containing XML. The tool uses [JSZip](https://stuk.github.io/jszip/) to extract `word/document.xml` client-side, then walks the XML tree parsing paragraphs, runs, and tables without any server involved.

**String safety**

VBA strings have a ~1023 character per line limit and can't contain raw control characters or non-ANSI Unicode. Every string is tokenised — special characters are converted to `Chr(N)` calls, long literals are split at 150-char boundaries with `& _` line continuation.

**Why multiple subs?**

VBA procedures have an undocumented size limit (roughly 64KB of bytecode per sub). Large documents are automatically split into `DocPart001`, `DocPart002`, etc., all called from the main `RecreateDocument()` entry point.

---

## Running locally

No build step. Just open the HTML file:

```bash
# Clone
git clone https://github.com/your-username/docx-to-vba.git
cd docx-to-vba

# Open in browser
open docx-to-vba.html        # macOS
start docx-to-vba.html       # Windows
xdg-open docx-to-vba.html    # Linux
```

Everything runs in the browser. No Node, no Python, no server.

---

## Dependencies

| Library | Version | Purpose |
|---|---|---|
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | Unzips the .docx in-browser |
| [JetBrains Mono](https://fonts.google.com/specimen/JetBrains+Mono) | — | UI monospace font |
| [Syne](https://fonts.google.com/specimen/Syne) | — | UI display font |

Both fonts load from Google Fonts. JSZip loads from cdnjs. Everything else is vanilla JS — no frameworks.

---

## Limitations

- **Images are not supported.** VBA can insert images but requires a file path on disk, which can't be embedded in the macro itself.
- **Complex layouts** (text boxes, SmartArt, charts) are skipped.
- **Very large documents** will generate very large macros. Word's VBA editor can handle them but may be slow to paste.
- **Non-Latin characters** (Chinese, Arabic, emoji, etc.) are converted to `Chr(N)` which works for Basic Multilingual Plane characters but may not render correctly in all Word locales.

---

## License

MIT — do whatever you want with it.
