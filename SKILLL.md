---
name: word-to-pdf
description: >
  Runs a local web application that converts Word documents (.docx / .doc) to PDF.
  Use this skill whenever the user wants to convert Word files to PDF through a browser
  UI — i.e. browse/upload a Word file, click a Convert button, and download the PDF.
  Trigger on phrases like "word to pdf", "convert to pdf", "export as pdf",
  "open the converter", "run the word-to-pdf app", or "start the converter web app".
  Always use this skill even if the request sounds simple — it manages starting the
  backend + frontend, the conversion engine, and result reporting.
---

# Word → PDF Converter (Web App)

A local web application that lets the user **browse and upload** a Word file
(`.docx` / `.doc`), click **Convert to PDF**, and **download** the result. The PDF is
delivered through the browser, so it lands in the user's default **Downloads** folder.

- **Frontend:** Vue 3 + Vite — `frontend/` (dev server on **http://localhost:3000**)
- **Backend:** FastAPI — `backend/` (API on **http://localhost:8000**)
- **Conversion engine:** Microsoft Word via `docx2pdf` (COM automation).

---

## ⚙️ Configuration

```
FRONTEND_URL = http://localhost:3000
BACKEND_URL  = http://localhost:8000
FRONTEND_DIR = C:\Users\zarta\Documents\Learn AI\Rasool_AI_Library\frontend
BACKEND_DIR  = C:\Users\zarta\Documents\Learn AI\Rasool_AI_Library\backend
```

> **How to use:** Ask Claude to "start the word-to-pdf app" (or "open the converter").
> Claude starts both servers, then you open `FRONTEND_URL` in a browser, upload a Word
> file, click **Convert to PDF**, and the PDF downloads to your default Downloads folder.

---

## How Claude Should Execute This Skill

When triggered, follow these steps **in order**:

### Step 1 — Read Configuration
Parse the URLs and directories from the **Configuration** section above. Do not ask the
user for them.

### Step 2 — Ensure Dependencies
- Backend: `pip install -r backend/requirements.txt` (FastAPI, uvicorn, python-multipart,
  docx2pdf, pywin32). Requires **Microsoft Word installed** on this machine.
- Frontend: run `npm install` in `FRONTEND_DIR` if `node_modules` is missing.

### Step 3 — Start the Servers
```bash
# Backend (terminal 1)
cd "<BACKEND_DIR>"
python -m uvicorn main:app --port 8000

# Frontend (terminal 2)
cd "<FRONTEND_DIR>"
npm run dev
```
Start them in the background. Confirm liveness via `GET <BACKEND_URL>/healthz`.

### Step 4 — Point the User to the UI
Tell the user to open `FRONTEND_URL`, browse/upload a Word file, and click **Convert to PDF**.
The browser downloads the PDF to its default Downloads folder.

### Step 5 — Report Results
Report back:
- Whether both servers started (and on which URLs)
- Readiness of the conversion engine (`GET <BACKEND_URL>/readyz`)
- Any errors, with the reason

---

## API Reference (backend)

| Method | Path           | Purpose                                                        |
|--------|----------------|----------------------------------------------------------------|
| GET    | `/healthz`     | Liveness — process is up.                                       |
| GET    | `/readyz`      | Readiness — Microsoft Word automation is reachable.            |
| POST   | `/api/convert` | Upload one Word file (`multipart/form-data`, field `file`); returns the converted PDF as a download. |

---

## Notes & Edge Cases

- **Unsupported files**: Only `.docx` and `.doc` are accepted; others get **HTTP 400**
  with a clear message in the UI.
- **Empty / oversized files**: Empty uploads → 400; files over **25 MB** → 413.
- **Download location**: The PDF is returned by the browser, so it goes to the user's
  **default Downloads folder** (no server-side output folder involved).
- **Word required**: Conversion uses the installed Microsoft Word. If Word is missing,
  `/readyz` returns 503 and conversions fail with a descriptive error.
- **Large files**: Word may take a few seconds per file; conversion runs on a COM-initialized
  worker thread so the server stays responsive.

---

## Dependencies

- **Microsoft Word** — installed locally (drives `docx2pdf`).
- **Python 3.12+** — `backend/requirements.txt`.
- **Node 20+** — `frontend/` (Vue 3 + Vite).
