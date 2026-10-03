@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating a private Python environment...
    python -m venv .venv
    if errorlevel 1 goto error
)

".venv\Scripts\python.exe" -c "import django, openpyxl, defusedxml" >nul 2>&1
if errorlevel 1 (
    echo Installing the free Django and Excel dependencies...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto error
)

echo Preparing the local database...
".venv\Scripts\python.exe" manage.py migrate --noinput
if errorlevel 1 goto error

".venv\Scripts\python.exe" manage.py setup_workspace_owner --if-needed
if errorlevel 1 goto error

echo Starting the website at http://127.0.0.1:8001/
start "My Web App" http://127.0.0.1:8001/
".venv\Scripts\python.exe" manage.py runserver --insecure 127.0.0.1:8001
goto end

:error
echo.
echo Could not start the website. Check that Python is installed and internet is available for the first install.
pause

:end
