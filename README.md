# Dudely Editor PC

Edit any PDF on a computer, including scanned and photo PDFs, without changing its layout.

- Edit text: click a line already on the page and type. The old words are covered in the page's own
  background colour and the new words go in the same place, size, colour and weight.
- Add text, Pen / sign, White-out, Highlight, Tick, Cross. Undo (Ctrl+Z), Save (Ctrl+S), zoom.
- Save as Word: every page keeps its exact look and each line of text becomes editable text in Word.
- Word back to PDF: open a Word file made by Dudely Editor PC (edited in Word or not) and it becomes a PDF
  with the same layout. Other Word files open as pages and are saved through "Save as PDF (print dialog)".
- Photos: Auto improve (colour and light balance, sharper), brightness, contrast, colour, warmth, sharpness,
  clean-document mode, crop, turn, mirror. Download one (JPG or PNG), all (zip), or put them all in a PDF.
- Merge: join PDFs and photos into one PDF in the order you choose.
- Split: bulk (a new file every N pages) or by page ranges you type (e.g. 1-3, 5, 8-10, 12-).
  Several files come as one zip.
- Saving always makes a new copy; the original file is never changed.
- Files stay on the computer: nothing is uploaded. Works offline after the first visit.

## Put it online once (free, GitHub Pages)
1. github.com > + > New repository. Name: dudely-editor. Public. Create repository.
2. Click "uploading an existing file" (or Add file > Upload files), then "choose your files". Unzip this
   file on the PC, open the folder, press Ctrl+A to select ALL the files and click Open. Commit changes.
   (There are no folders: every file sits next to index.html.)
3. Check: the repository's main page lists index.html, app.js, tools.js, pdf.min.mjs,
   tesseract-core-lstm.wasm.js, eng.traineddata.gz and the other files (about 30).
4. Settings > Pages > Source: Deploy from a branch; Branch: main; folder: / (root) > Save.
5. After 1-2 minutes: https://YOUR-GITHUB-NAME.github.io/dudely-editor/

## Install on each PC (Windows, Chrome or Edge)
1. Open the link in Chrome or Microsoft Edge.
2. Click "Install on this PC" (top right) or the install icon at the end of the address bar > Install.
3. Dudely Editor PC gets its own window, a desktop icon and a Start menu entry, and appears in
   "Open with" for PDF files.

Libraries included: PDF.js (Apache-2.0), pdf-lib (MIT), Tesseract.js (Apache-2.0) with English text data,
JSZip (MIT), docx-preview (Apache-2.0).
