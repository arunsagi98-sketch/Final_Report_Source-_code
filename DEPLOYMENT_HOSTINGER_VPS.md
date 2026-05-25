# Deployment Guide: Flask Backend on Hostinger VPS

## Prerequisites
- Hostinger VPS account with SSH access
- Python 3.8+ installed on server
- A domain or IP address pointing to your VPS

## Step 1: Connect to Your VPS via SSH

```bash
ssh root@your_vps_ip_address
# Or: ssh your_username@your_vps_ip_address
```

## Step 2: Install System Dependencies

```bash
# Update system
sudo apt update && sudo apt upgrade -y

# Install Python and pip
sudo apt install python3 python3-pip python3-venv -y

# Install other dependencies
sudo apt install git curl wget -y
```

## Step 3: Create Application Directory

```bash
# Create app directory
mkdir -p /home/your_username/apps/final-report-backend
cd /home/your_username/apps/final-report-backend

# Clone your repository (or upload files)
git clone https://github.com/arunsagi98-sketch/Final_Report_Source-_code.git
cd Final_Report_Source-_code/backend
```

## Step 4: Set Up Python Virtual Environment

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Upgrade pip
pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt

# Install production server
pip install gunicorn
```

## Step 5: Configure Environment Variables

```bash
# Edit .env file
nano .env
```

Update with your settings:
```
FLASK_ENV=production
FLASK_DEBUG=False
SECRET_KEY=your-random-secret-key-here
FRONTEND_URL=https://your-frontend-domain.com
DATABASE_DIR=/home/your_username/databases/
HOST=0.0.0.0
PORT=5000
```

## Step 6: Create Systemd Service File

```bash
sudo nano /etc/systemd/system/final-report-api.service
```

Paste this:
```ini
[Unit]
Description=Final Report Generator Backend
After=network.target

[Service]
User=your_username
WorkingDirectory=/home/your_username/apps/final-report-backend/Final_Report_Source-_code/backend
Environment="PATH=/home/your_username/apps/final-report-backend/venv/bin"
ExecStart=/home/your_username/apps/final-report-backend/venv/bin/gunicorn -w 4 -b 0.0.0.0:5000 wsgi:app
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

## Step 7: Enable and Start Service

```bash
# Reload systemd
sudo systemctl daemon-reload

# Enable service to start on boot
sudo systemctl enable final-report-api.service

# Start the service
sudo systemctl start final-report-api.service

# Check status
sudo systemctl status final-report-api.service

# View logs
sudo journalctl -u final-report-api.service -f
```

## Step 8: Set Up Nginx as Reverse Proxy

```bash
sudo apt install nginx -y

# Create nginx config
sudo nano /etc/nginx/sites-available/final-report-api
```

Paste this:
```nginx
server {
    listen 80;
    server_name your_domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_connect_timeout 60s;
        proxy_send_timeout 60s;
        proxy_read_timeout 60s;
    }
}
```

Enable and start Nginx:
```bash
sudo ln -s /etc/nginx/sites-available/final-report-api /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl start nginx
sudo systemctl enable nginx
```

## Step 9: Set Up SSL Certificate (HTTPS)

```bash
sudo apt install certbot python3-certbot-nginx -y

sudo certbot --nginx -d your_domain.com
```

## Step 10: Configure Firewall

```bash
sudo ufw allow 22/tcp  # SSH
sudo ufw allow 80/tcp  # HTTP
sudo ufw allow 443/tcp # HTTPS
sudo ufw enable
```

## Step 11: Update Frontend to Use New Backend URL

In your frontend `index.html`, update:
```javascript
const API = window.BACKEND_URL || "https://your_domain.com";
```

Or set in HTML:
```html
<script>
  window.BACKEND_URL = "https://your_domain.com";
</script>
```

## Verification

Test your API:
```bash
# From your local machine
curl https://your_domain.com/db-sheets
```

Should return JSON with sheet names.

## Useful Commands

```bash
# View service logs
sudo journalctl -u final-report-api.service -f

# Restart service
sudo systemctl restart final-report-api.service

# Stop service
sudo systemctl stop final-report-api.service

# Check if running on port 5000
sudo netstat -tulpn | grep 5000

# Check Nginx logs
sudo tail -f /var/log/nginx/error.log
sudo tail -f /var/log/nginx/access.log
```

## Troubleshooting

**Service won't start:**
```bash
sudo systemctl status final-report-api.service
sudo journalctl -u final-report-api.service -n 50
```

**Nginx returns 502 Bad Gateway:**
- Make sure Flask is running: `sudo systemctl status final-report-api.service`
- Check firewall: `sudo ufw status`

**Can't connect to database files:**
- Ensure paths in `.env` are correct
- Check file permissions: `ls -la /home/your_username/databases/`

## Next Steps

1. Keep your application updated: `git pull` in the app directory
2. Monitor logs regularly
3. Set up automated backups for generated reports
4. Consider setting up Cron jobs for cleanup of old uploads
