# FileForge — File Automation System
## Quick Start (3 steps)

### 1. Install dependencies
```bash
pip install flask openpyxl pandas werkzeug psycopg2-binary
```

### 1b. Configure PostgreSQL
- Create a PostgreSQL database
- Set `DATABASE_URL` in your environment or `.env`
- The first app start will bootstrap the reference data from `db/*.xlsx`

### 2. Start the backend
```bash
python app.py
```
You should see: `🚀  File Automation Server running at http://localhost:5000`

For local development, keep `FLASK_DEBUG=True` so code changes auto-reload.

---

## How it works

```
client       ─── POST /upload ───►  app.py (Flask)
                                     │
                                     ├─ saves file to  uploads/
                                     ├─ spawns background thread
                                     └─ processes with pandas + openpyxl

client       ─── GET /status/:id ──► returns {"status": "done|processing|error"}
client       ─── GET /download/:id ► streams the .xlsx report
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
├── db/             ← Database/reference Excel files for the app
├── uploads/        ← Temporary uploaded files (auto-cleaned)
└── outputs/        ← Generated reports (auto-cleaned 60s after download)
```

## Notes
- Max file size: 50 MB (configurable via `MAX_CONTENT_LENGTH` in app.py)
- Files are deleted automatically after download
- CORS is enabled for all origins (restrict in production)
- Reference data is stored in PostgreSQL
- The Excel files in `db/` are only seed sources for the initial import
