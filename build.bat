@echo off
REM Builds Topshot.exe using PyInstaller. The entry point is
REM launcher.py, which runs a downloaded update if there is one (see
REM make_update.py) and the app built into the .exe otherwise.
REM Run this from the ImageSearchBuild folder after `pip install -r requirements.txt`.

REM Pre-download the CLIP model weights (once) so they can be bundled into
REM the .exe - end users then never wait on the first-run download.
if not exist "model_cache" (
    echo Model cache not found - downloading CLIP weights once before building...
    py download_model.py
)

py -m PyInstaller --noconfirm --onedir --windowed --name Topshot ^
    --collect-all open_clip ^
    --collect-all torch ^
    --collect-all torchvision ^
    --collect-all msal ^
    --collect-all customtkinter ^
    --collect-all transformers ^
    --collect-all tokenizers ^
    --hidden-import unicodedata ^
    --hidden-import tkinter_app ^
    --hidden-import app_version ^
    --hidden-import sharepoint_client ^
    --hidden-import config ^
    --add-data "model_cache;model_cache" ^
    --add-data "assets;assets" ^
    --version-file "version_info.txt" ^
    --icon "assets\app_icon.ico" ^
    launcher.py

REM Strip bundled third-party license text (torch's kineto/dynolog trees in
REM particular nest ~150 folders deep). Pure attribution files, unused at
REM runtime, but their long paths risk exceeding Windows' 260-char MAX_PATH
REM on extraction elsewhere and silently dropping files/folders.
for /d /r "dist\Topshot\_internal" %%D in (*.dist-info) do (
    if exist "%%D\licenses" rd /s /q "%%D\licenses"
)

echo.
echo Build complete. Find Topshot.exe in dist\Topshot\
