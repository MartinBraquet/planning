@echo off
setlocal
echo Génération des planning...
echo.

pushd ".runner"

REM Create venv if missing
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
    call .venv\Scripts\activate.bat
    echo Upgrading pip...
    python -m pip install --upgrade pip
    echo Installing required package
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

REM Show Python version
python --version

REM Run your script
python run.py

REM Show completion message
echo.
echo Fini. Voir fichiers dans le dossier "planning".
endlocal
pause