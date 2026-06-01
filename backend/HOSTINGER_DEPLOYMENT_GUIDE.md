# Hostinger Deployment Guide for Flask Backend

## Prerequisites
- Hostinger account with cPanel/hosting
- Python 3.8+ support (check with Hostinger support)
- SSH access enabled
- FTP or File Manager access

---

## Step 1: Prepare Your Project for Production

### 1.1 Create Requirements File
Run this in your backend directory:
```bash
pip freeze > requirements.txt
```

This creates a list of all dependencies for installation on Hostinger.

### 1.2 Create Python WSGI Entry Point (wsgi.py)
Check if `wsgi.py` exists - if not, create it:

```python
# wsgi.py
import sys
import os
from pathlib import Path

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import Flask app
from app import app

if __name__ == "__main__":
    app.run()
```

### 1.3 Create .htaccess for Apache
Create this file in your project root:

```apache
# .htaccess
<IfModule mod_rewrite.c>
  RewriteEngine On
  RewriteCond %{REQUEST_FILENAME} !-f
  RewriteCond %{REQUEST_FILENAME} !-d
  RewriteRule ^(.*)$ /index.py/$1 [L]
</IfModule>

# Enable gzip compression
<IfModule mod_deflate.c>
  AddOutputFilterByType DEFLATE text/html text/plain text/xml application/json
</IfModule>

# Disable directory listing
Options -Indexes
```

### 1.4 Create app.conf Configuration File
```
# app.conf
PassEnv PATH
PassEnv PYTHONPATH
```

---

## Step 2: Upload to Hostinger

### Option A: Using FTP/File Manager (Easiest)

1. **Access File Manager in cPanel**
   - Login to Hostinger cPanel
   - Go to File Manager
   - Navigate to `public_html` folder

2. **Create Project Directory**
   ```
   public_html/
   └── backend/
       ├── app.py
       ├── reference_db.py
       ├── wsgi.py
       ├── requirements.txt
       ├── .htaccess
       ├── db/
       │   ├── App_Url Data base.xlsx
       │   └── City for Aoutomation.xlsx
       ├── uploads/
       ├── outputs/
       └── assets/
   ```

3. **Upload Files**
   - Upload all Python files, databases, and assets
   - Ensure `uploads/` and `outputs/` directories are writable (755 permissions)

### Option B: Using SSH (Advanced)

```bash
# SSH into Hostinger
ssh username@your-domain.com

# Navigate to public_html
cd public_html

# Clone your GitHub repo (if available)
git clone https://github.com/arunsagi98-sketch/Final_Report_Source-_code.git backend

# OR upload via SCP
scp -r ./backend username@your-domain.com:~/public_html/
```

---

## Step 3: Configure Python on Hostinger

### 3.1 Access cPanel > Setup Python App
1. **Login to cPanel**
2. **Find "Setup Python App" or "Python Selector"**
3. **Create New Application**
   - **Python version:** 3.9+ (recommended)
   - **Application root:** `/home/username/public_html/backend`
   - **Application URL:** `https://your-domain.com`
   - **Application startup file:** `wsgi.py`

### 3.2 Install Dependencies

In cPanel > Terminal (or via SSH):

```bash
cd ~/public_html/backend
pip install -r requirements.txt
```

Key packages to verify:
```
Flask==2.3.0
flask-cors==4.0.0
openpyxl==3.9.0
pandas==1.5.0
werkzeug==2.3.0
psycopg2-binary==2.9.0 (optional, for PostgreSQL)
```

---

## Step 4: Configure Environment & Permissions

### 4.1 Set Permissions

```bash
# Set correct permissions (via SSH or Terminal)
chmod 755 ~/public_html/backend
chmod 755 ~/public_html/backend/uploads
chmod 755 ~/public_html/backend/outputs
chmod 755 ~/public_html/backend/db
chmod 644 ~/public_html/backend/app.py
chmod 644 ~/public_html/backend/*.xlsx
```

### 4.2 Create .env File (Optional)

If you need environment variables:

```bash
# .env file
export FLASK_ENV=production
export FLASK_DEBUG=0
export MAX_CONTENT_LENGTH=52428800
export DB_DIR=/home/username/public_html/backend/db
```

### 4.3 Configure Domain

1. In cPanel > Domains
2. Point your domain to `public_html/backend` (or keep default)
3. Add SSL certificate (free Let's Encrypt available)

---

## Step 5: Restart Application

### In cPanel > Setup Python App:
- Find your application
- Click **Restart** button

Or via SSH:
```bash
# Kill existing processes
pkill -f "python.*app.py"

# Restart the application
~/public_html/backend/bin/python -m gunicorn wsgi:app
```

---

## Step 6: Test Endpoints

Once deployed, test your endpoints:

```bash
# Test db-sheets endpoint
curl https://your-domain.com/db-sheets

# Test city-sheets endpoint
curl https://your-domain.com/city-sheets

# Test upload endpoint
curl -X OPTIONS https://your-domain.com/upload
```

---

## Step 7: Configure Frontend to Use Hostinger Backend

Update your frontend `.env` file:

```env
REACT_APP_API_URL=https://your-domain.com
```

Or hardcode in your frontend:
```javascript
const API = 'https://your-domain.com';
```

---

## Troubleshooting

### Issue: 500 Internal Server Error
- Check cPanel error logs: `Home > Logs > Error Log`
- Verify Python version compatibility
- Check file permissions (should be 755 for dirs, 644 for files)

### Issue: Module Not Found (pandas, openpyxl, etc.)
```bash
# Reinstall all packages
pip install --upgrade -r requirements.txt

# Or install individually
pip install pandas openpyxl flask-cors
```

### Issue: Database File Not Found
```bash
# Verify paths in app.py
# Check if /db folder is uploaded with Excel files
# Permissions: chmod 644 *.xlsx
```

### Issue: CORS Errors in Frontend
- Confirm `CORS(app)` is in app.py
- Check cPanel logs for 403/401 errors
- Add `Access-Control-Allow-Origin: *` headers if needed

### Issue: Upload File Size Limited
- In app.py, verify: `app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024`
- In cPanel > PHP > Max Upload Size: set to 50MB+
- In .htaccess: `php_value upload_max_filesize 50M`

---

## Performance Optimization

### 1. Add Caching Headers (.htaccess)
```apache
<FilesMatch "\.(jpg|jpeg|png|gif|ico|svg|css|js|xlsx)$">
  Header set Cache-Control "max-age=2592000, public"
</FilesMatch>
```

### 2. Enable Gzip Compression
Already in `.htaccess` - auto-compresses JSON responses

### 3. Database Optimization
- Keep Excel files in `/db` directory
- Consider migrating to PostgreSQL for better performance
- Add indexes if using database

### 4. Clean Old Output Files
```bash
# Add to cron job (in cPanel > Cron Jobs)
# Run daily at 2 AM
0 2 * * * find /home/username/public_html/backend/outputs -mtime +1 -delete
```

---

## Estimated Costs & Hosting Options

### Hostinger Plans Suitable for This Project:
1. **Premium Plan** - $2.99-5.99/month
   - Good for small to medium projects
   - Includes Python support
   - 100GB SSD storage

2. **Business Plan** - $5.99-9.99/month
   - Better for high-traffic apps
   - More CPU/RAM
   - Email included

3. **Cloud Hosting** - $8.99-29.99/month
   - Best performance
   - Full server control
   - Recommended for production

---

## Post-Deployment Checklist

- [ ] Upload all project files to Hostinger
- [ ] Set correct file permissions (755/644)
- [ ] Install Python packages via pip
- [ ] Test all endpoints (db-sheets, city-sheets, upload)
- [ ] Configure domain/SSL
- [ ] Update frontend API URL
- [ ] Enable error logging in cPanel
- [ ] Set up automated backups
- [ ] Add cron job for cleanup
- [ ] Monitor performance in cPanel > Metrics

---

## Support

**Hostinger Support:** support.hostinger.com
**Python Docs:** python.org/docs
**Flask Docs:** flask.palletsprojects.com
**cPanel Docs:** docs.cpanel.net

