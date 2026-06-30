@echo off
REM ============================================================
REM  Fab Sol Compliance Build - v5 (pyproject.toml based)
REM ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "LOG_FILE=%~dp0build_log_fab_sol.txt"
if exist "%LOG_FILE%" del /q "%LOG_FILE%"

echo Starting build... output will be saved to:
echo   %LOG_FILE%
echo.
echo Please wait 5-10 minutes (don't close this window)...
echo.

call :main > "%LOG_FILE%" 2>&1
set "BUILD_RESULT=%errorlevel%"

echo.
echo ============================================================
if !BUILD_RESULT! equ 0 (
    echo  BUILD SUCCESS!
    if exist "%~dp0dist\Fab_Sol_Compliance.exe" (
        echo  Output: %~dp0dist\Fab_Sol_Compliance.exe
        for %%A in ("%~dp0dist\Fab_Sol_Compliance.exe") do (
            set /a "size_mb=%%~zA / 1048576"
            echo  Size:   !size_mb! MB
        )
    )
) else (
    echo  BUILD FAILED!
    echo  Check build_log_fab_sol.txt for details
)
echo ============================================================
echo.
echo Opening build_log_fab_sol.txt in Notepad...
timeout /t 2 >nul
start "" notepad "%LOG_FILE%"
exit /b !BUILD_RESULT!


:main
echo ============================================================
echo  Fab Sol Compliance Build - v5 (pyproject.toml based)
echo  Started: %DATE% %TIME%
echo ============================================================
echo.

REM ----- Verify project files -----
echo [PREFLIGHT] Checking project files...
if not exist "%~dp0pyproject.toml" (
    if not exist "%~dp0requirements-build.txt" (
        echo [ERROR] No pyproject.toml OR requirements-build.txt found!
        echo         At least one must exist in: %~dp0
        exit /b 1
    )
    echo   Will use requirements-build.txt ^(pyproject.toml not found^)
    set "USE_PYPROJECT=0"
) else (
    echo   Will use pyproject.toml
    set "USE_PYPROJECT=1"
)
if not exist "%~dp0tpm_compliance_fabsol.py" (
    echo [ERROR] tpm_compliance_fabsol.py not found!
    exit /b 1
)
echo   [OK] All required files present
echo.

REM ----- Step 1: Find Python -----
echo [STEP 1] Finding Python...
set "PY_CMD="
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist "%%~P" (
        set "PY_CMD=%%~P"
        echo   Found Python.org Python at: %%~P
        goto :found_python
    )
)
where py >nul 2>&1 && (set "PY_CMD=py -3" & echo   Found py launcher & goto :found_python)
where python >nul 2>&1 && (set "PY_CMD=python" & echo   Found python in PATH & goto :found_python)
echo [ERROR] No Python found
exit /b 1

:found_python
!PY_CMD! --version
!PY_CMD! -c "import sys; print('  Executable:', sys.executable)"
echo.

REM ----- Step 2: Clean old artifacts -----
echo [STEP 2] Cleaning old artifacts...
if exist "%~dp0build_venv" rd /s /q "%~dp0build_venv"
if exist "%~dp0dist" rd /s /q "%~dp0dist"
if exist "%~dp0build" rd /s /q "%~dp0build"
if exist "%~dp0Fab_Sol_Compliance.spec" del /q "%~dp0Fab_Sol_Compliance.spec"
echo   Done
echo.

REM ----- Step 3: Create venv -----
echo [STEP 3] Creating venv...
!PY_CMD! -m venv "%~dp0build_venv" --copies --clear
if !errorlevel! neq 0 (
    echo [ERROR] venv creation failed
    exit /b 1
)
set "VENV_PY=%~dp0build_venv\Scripts\python.exe"
echo   Venv Python: !VENV_PY!
"!VENV_PY!" --version
echo.

REM ----- Step 4: Upgrade pip -----
echo [STEP 4] Upgrading pip...
set "PIP_FLAGS=--trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org"
"!VENV_PY!" -m pip install --upgrade pip !PIP_FLAGS!
echo.

REM ----- Step 5: Install from pyproject.toml or requirements-build.txt -----
echo [STEP 5] Installing all packages...

if "!USE_PYPROJECT!"=="1" (
    echo   Using pyproject.toml [build] dependencies...
    "!VENV_PY!" -m pip install !PIP_FLAGS! ".[build]"
) else (
    echo   Using requirements-build.txt...
    "!VENV_PY!" -m pip install !PIP_FLAGS! -r requirements-build.txt
)

if !errorlevel! neq 0 (
    echo [ERROR] Package install failed
    exit /b 1
)
echo   [OK] All packages installed
echo.

REM ----- Verify PyInstaller bootloader -----
set "BOOTSTRAP_FILE=%~dp0build_venv\Lib\site-packages\PyInstaller\loader\pyiboot01_bootstrap.py"
if not exist "!BOOTSTRAP_FILE!" (
    echo [WARN] PyInstaller bootstrap missing - force reinstalling...
    "!VENV_PY!" -m pip install !PIP_FLAGS! --force-reinstall --no-cache-dir pyinstaller
    if not exist "!BOOTSTRAP_FILE!" (
        echo [ERROR] PyInstaller still broken
        exit /b 1
    )
)
echo   [OK] PyInstaller ready
echo.

echo [VERIFY] Installed packages:
"!VENV_PY!" -m pip list
echo.

REM ----- Step 6: Run PyInstaller -----
echo [STEP 6] Building exe...

"!VENV_PY!" -m PyInstaller ^
    --onefile ^
    --console ^
    --name Fab_Sol_Compliance ^
    --noconfirm --clean ^
    --collect-all polars ^
    --collect-all polars_runtime_32 ^
    --collect-all fastexcel ^
    --collect-submodules openpyxl ^
    --collect-submodules xlsxwriter ^
    --collect-submodules rich ^
    --collect-submodules questionary ^
    --collect-submodules prompt_toolkit ^
    --hidden-import pyxlsb ^
    --hidden-import pandas ^
    --hidden-import pandas._libs.tslibs.base ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\ffi-8.dll;." ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\libcrypto-3-x64.dll;." ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\libssl-3-x64.dll;." ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\liblzma.dll;." ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\libbz2.dll;." ^
    --add-binary "C:\ProgramData\miniforge3\Library\bin\sqlite3.dll;." ^
    --exclude-module matplotlib ^
    --exclude-module tkinter ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    --exclude-module IPython ^
    --exclude-module jupyter ^
    --exclude-module pytest ^
    --exclude-module hypothesis ^
    --exclude-module pandas.tests ^
    --exclude-module pandas.io.tests ^
    --exclude-module numpy.tests ^
    --exclude-module asyncssh ^
    "%~dp0tpm_compliance_fabsol.py"

if !errorlevel! neq 0 (
    echo [ERROR] PyInstaller build failed
    exit /b 1
)

REM ----- Verify output -----
echo.
if exist "%~dp0dist\Fab_Sol_Compliance.exe" (
    for %%A in ("%~dp0dist\Fab_Sol_Compliance.exe") do (
        echo [SUCCESS] Built: %%~fA
        echo           Size: %%~zA bytes
    )
    exit /b 0
) else (
    echo [ERROR] Output file missing
    exit /b 1
)