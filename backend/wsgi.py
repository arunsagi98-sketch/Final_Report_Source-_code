"""
WSGI entry point for production deployment (Gunicorn, etc.)
Run with: gunicorn -w 4 -b 0.0.0.0:5000 wsgi:app
"""
import os
import sys
from pathlib import Path

# Add backend to path
sys.path.insert(0, str(Path(__file__).parent))

from app import app

if __name__ == "__main__":
    app.run()
