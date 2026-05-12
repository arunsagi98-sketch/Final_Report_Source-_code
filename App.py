import subprocess
import os
import sys

# This is a redirect script to the new backend folder
backend_dir = os.path.join(os.path.dirname(__file__), "backend")
os.chdir(backend_dir)

print(f"--- Redirecting to backend/app.py ---")
subprocess.run([sys.executable, "app.py"])
