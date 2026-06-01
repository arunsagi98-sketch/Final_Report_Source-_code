@echo off
REM Hostinger Deployment Setup Script for Windows
REM Run this after uploading to Hostinger Windows hosting

setlocal enabledelayedexpansion

echo =========================================
echo Flask App Deployment Setup for Hostinger
echo =========================================
echo.

REM Get project directory
set PROJECT_DIR=%~dp0

echo Project directory: %PROJECT_DIR%
echo.

REM 1. Create directories if they don't exist
echo Step 1: Creating directories...
if not exist "%PROJECT_DIR%logs" mkdir "%PROJECT_DIR%logs"
if not exist "%PROJECT_DIR%uploads" mkdir "%PROJECT_DIR%uploads"
if not exist "%PROJECT_DIR%outputs" mkdir "%PROJECT_DIR%outputs"
echo ✓ Directories verified
echo.

REM 2. Install Python dependencies
echo Step 2: Installing Python dependencies...
python -m pip install --upgrade pip
python -m pip install -r "%PROJECT_DIR%requirements.txt"
if %errorlevel% equ 0 (
    echo ✓ Dependencies installed successfully
) else (
    echo ✗ Failed to install dependencies
    pause
    exit /b 1
)
echo.

REM 3. Verify database files
echo Step 3: Verifying database files...
if exist "%PROJECT_DIR%db\App_Url Data base.xlsx" (
    echo ✓ App database file found
) else (
    echo ✗ App database file NOT found
)

if exist "%PROJECT_DIR%db\City for Aoutomation.xlsx" (
    echo ✓ City database file found
) else (
    echo ✗ City database file NOT found
)
echo.

REM 4. Test Python app import
echo Step 4: Testing Python app import...
python -c "from app import app; print('✓ Flask app imported successfully')"
if %errorlevel% neq 0 (
    echo ✗ Failed to import app
    pause
    exit /b 1
)
echo.

REM 5. Create .env file if needed
echo Step 5: Creating .env file...
if not exist "%PROJECT_DIR%.env" (
    copy "%PROJECT_DIR%.env.production" "%PROJECT_DIR%.env"
    echo Note: Please edit .env file with your configuration
    echo        Edit: %PROJECT_DIR%.env
) else (
    echo ✓ .env file already exists
)
echo.

REM 6. Summary
echo =========================================
echo ✓ Deployment setup completed!
echo =========================================
echo.

echo Next steps:
echo 1. Edit .env file with your configuration
echo    File: %PROJECT_DIR%.env
echo.
echo 2. In cPanel, go to "Setup Python App" and:
echo    - Create new application
echo    - Set Application Root to your domain path
echo    - Set Startup File to wsgi.py
echo    - Click "Create"
echo.
echo 3. Restart your app in cPanel
echo.
echo 4. Test your endpoints:
echo    https://your-domain.com/db-sheets
echo    https://your-domain.com/city-sheets
echo.
echo 5. Update your frontend API URL
echo.

echo Troubleshooting:
echo - Check cPanel error logs
echo - Verify file permissions
echo - Check that all database files are uploaded
echo.

pause
