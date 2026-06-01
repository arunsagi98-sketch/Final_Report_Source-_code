#!/bin/bash
# Hostinger Deployment Setup Script
# Run this after uploading to Hostinger via SSH

set -e  # Exit on error

echo "========================================="
echo "Flask App Deployment Setup for Hostinger"
echo "========================================="

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Get current directory
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo -e "${GREEN}Project directory: $PROJECT_DIR${NC}"

# 1. Set file permissions
echo -e "\n${YELLOW}Step 1: Setting file permissions...${NC}"
chmod 755 "$PROJECT_DIR"
chmod 755 "$PROJECT_DIR/uploads"
chmod 755 "$PROJECT_DIR/outputs"
chmod 755 "$PROJECT_DIR/db"
chmod 644 "$PROJECT_DIR"/*.py
chmod 644 "$PROJECT_DIR"/db/*.xlsx
chmod 644 "$PROJECT_DIR/.htaccess"
chmod 644 "$PROJECT_DIR/.env.production"
echo -e "${GREEN}✓ Permissions set correctly${NC}"

# 2. Create logs directory
echo -e "\n${YELLOW}Step 2: Creating logs directory...${NC}"
mkdir -p "$PROJECT_DIR/logs"
chmod 755 "$PROJECT_DIR/logs"
echo -e "${GREEN}✓ Logs directory created${NC}"

# 3. Install Python dependencies
echo -e "\n${YELLOW}Step 3: Installing Python dependencies...${NC}"
if command -v pip3 &> /dev/null; then
    pip3 install --upgrade pip setuptools wheel
    pip3 install -r "$PROJECT_DIR/requirements.txt"
    echo -e "${GREEN}✓ Dependencies installed${NC}"
else
    echo -e "${RED}✗ pip3 not found. Please install dependencies manually:${NC}"
    echo "   python3 -m pip install -r $PROJECT_DIR/requirements.txt"
fi

# 4. Verify database files exist
echo -e "\n${YELLOW}Step 4: Verifying database files...${NC}"
if [ -f "$PROJECT_DIR/db/App_Url Data base.xlsx" ]; then
    echo -e "${GREEN}✓ App database file found${NC}"
else
    echo -e "${RED}✗ App database file NOT found at $PROJECT_DIR/db/App_Url Data base.xlsx${NC}"
fi

if [ -f "$PROJECT_DIR/db/City for Aoutomation.xlsx" ]; then
    echo -e "${GREEN}✓ City database file found${NC}"
else
    echo -e "${RED}✗ City database file NOT found at $PROJECT_DIR/db/City for Aoutomation.xlsx${NC}"
fi

# 5. Test if app can be imported
echo -e "\n${YELLOW}Step 5: Testing Python app import...${NC}"
if python3 -c "from app import app; print('✓ Flask app imported successfully')" 2>/dev/null; then
    echo -e "${GREEN}✓ App imports successfully${NC}"
else
    echo -e "${RED}✗ Failed to import app. Check dependencies and app.py syntax${NC}"
fi

# 6. Create .env file if needed
echo -e "\n${YELLOW}Step 6: Creating .env file (if needed)...${NC}"
if [ ! -f "$PROJECT_DIR/.env" ]; then
    cp "$PROJECT_DIR/.env.production" "$PROJECT_DIR/.env"
    echo -e "${YELLOW}Note: Please edit .env file with your configuration${NC}"
    echo "    nano $PROJECT_DIR/.env"
else
    echo -e "${GREEN}✓ .env file already exists${NC}"
fi

# 7. Summary
echo -e "\n${GREEN}=========================================${NC}"
echo -e "${GREEN}✓ Deployment setup completed!${NC}"
echo -e "${GREEN}=========================================${NC}"

echo -e "\n${YELLOW}Next steps:${NC}"
echo "1. Edit .env file with your configuration"
echo "   nano $PROJECT_DIR/.env"
echo ""
echo "2. Restart your Python app in cPanel > Setup Python App"
echo ""
echo "3. Test your endpoints:"
echo "   curl https://your-domain.com/db-sheets"
echo "   curl https://your-domain.com/city-sheets"
echo ""
echo "4. Update frontend API URL in .env file"
echo ""
echo -e "${YELLOW}Troubleshooting:${NC}"
echo "- Check logs: tail -f $PROJECT_DIR/logs/app.log"
echo "- Check cPanel error logs"
echo "- Verify file permissions: ls -la $PROJECT_DIR"
echo ""
