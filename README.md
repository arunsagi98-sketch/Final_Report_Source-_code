# Ad Campaign Report Generator

Flask app for generating campaign reports from Excel or CSV inputs, with a browser UI and downloadable Excel/HTML output.

## Project Structure

- `backend/` - Flask backend and processing logic
- `frontend/` - single-page UI served by the backend
- `uploads/` - temporary uploaded files
- `outputs/` - generated report files
- `wsgi.py` - Gunicorn/WSGI entrypoint for VPS deployment
- `requirements.txt` - Python dependencies

## Local Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the backend:

```bash
python backend/app.py
```

4. Open the app in your browser:

```text
http://localhost:5000
```

The backend serves the frontend automatically.

## What the App Does

- Uploads campaign data files
- Supports Video and Banner report modes
- Loads App/URL and City database sheet names from Excel files
- Generates Excel and HTML reports
- Provides status, download, and view endpoints

## Hostinger VPS Deployment

This project is meant for a Hostinger VPS, not shared hosting, because it runs a Flask backend.

### First-time deploy

1. Upload the project to the VPS or clone it there.
2. SSH into the server.
3. Create a virtual environment.
4. Install dependencies:

```bash
pip install -r requirements.txt
```

5. Run the app with Gunicorn:

```bash
gunicorn --bind 127.0.0.1:8000 wsgi:app
```

6. Put Nginx in front of Gunicorn and proxy your domain to `127.0.0.1:8000`.

### Updating the Project on Hostinger

When you make changes locally:

1. Upload the changed files to the VPS or pull the latest code there.
2. Install any dependency changes again:

```bash
pip install -r requirements.txt
```

3. Restart the app service:

```bash
sudo systemctl restart <your-service-name>
```

If you are running Gunicorn manually, stop the old process and start it again with the same command.

## Important Notes

- The frontend uses the same site origin by default, so it works better on deployed domains.
- Do not open `frontend/index.html` directly in production. Use the Flask app URL.
- The app expects its reference Excel files in `backend/data/`.
- Temporary uploads and outputs are cleaned up by the app logic.

## Troubleshooting

- If City sheets or App/URL sheets do not load, confirm the backend is running and the Excel files exist in `backend/data/`.
- If the browser cannot reach the API, confirm the domain is proxying to the Flask process.
- If you change Python packages, reinstall from `requirements.txt`.

