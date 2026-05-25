"""
WSGI entry point for production deployment (Gunicorn, etc.)
Run with: gunicorn -w 4 -b 0.0.0.0:10000 wsgi:app
"""
import os
import sys
from pathlib import Path

# Add backend to path
backend_dir = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_dir))

from app import app

if __name__ == "__main__":
    app.run()
