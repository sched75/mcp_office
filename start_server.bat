@echo off
REM ========================================
REM MCP Office Automation Server Launcher
REM ========================================

echo.
echo ========================================
echo   MCP Office Automation Server
echo ========================================
echo.

REM Check if venv exists
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment not found!
    echo.
    echo Please create it first:
    echo   python -m venv venv
    echo   venv\Scripts\activate.bat
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM Activate virtual environment
echo [1/3] Activating virtual environment...
call venv\Scripts\activate.bat

REM Check if dependencies are installed
echo [2/3] Checking dependencies...
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo [WARNING] pywin32 not found. Installing dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies!
        pause
        exit /b 1
    )
)

REM Start the server
echo [3/3] Starting MCP Office Automation server...
echo.
echo ========================================
echo Server is running...
echo Press Ctrl+C to stop
echo ========================================
echo.

python -m src.server

REM On exit
echo.
echo ========================================
echo Server stopped
echo ========================================
echo.
pause
