# FileForge — File Automation System
## Quick Start (3 steps)

### 1. Install dependencies
```bash
pip install flask openpyxl pandas werkzeug
```

### 2. Start the backend
```bash
python app.py
```
You should see: `🚀  File Automation Server running at http://localhost:5000`

### 3. Open the frontend
Just open `index.html` in any browser. No server needed for the frontend.

---

## How it works

```
index.html  ─── POST /upload ───►  app.py (Flask)
                                     │
                                     ├─ saves file to  uploads/
                                     ├─ spawns background thread
                                     └─ processes with pandas + openpyxl

index.html  ─── GET /status/:id ──► returns {"status": "done|processing|error"}
index.html  ─── GET /download/:id ► streams the .xlsx report
```

## Generated Report Sheets

| Sheet | Contents |
|-------|----------|
| 📊 Dashboard | File metadata, row/column counts, missing value summary |
| 📄 Raw Data | Original data with styled header + alternating rows |
| 🔍 Column Summary | Per-column type, null count, unique count, min/max/mean/std |
| ✅ Cleaned Data | Rows with nulls and duplicates removed |
| 📈 Numeric Stats | pandas `.describe()` output for all numeric columns |

## Customising the processing logic

Edit the `process_file()` function in `app.py`. The placeholder already:
- Loads CSV or Excel with pandas
- Builds summary stats
- Cleans data (drops nulls + duplicates)
- Writes a styled multi-sheet Excel report

Add your own business logic between steps 1 and 3.

## Project Structure

```
├── app.py          ← Flask backend
├── index.html      ← Frontend UI (open in browser)
├── uploads/        ← Temporary uploaded files (auto-cleaned)
└── outputs/        ← Generated reports (auto-cleaned 60s after download)
```

## Notes
- Max file size: 50 MB (configurable via `MAX_CONTENT_LENGTH` in app.py)
- Files are deleted automatically after download
- CORS is enabled for all origins (restrict in production)