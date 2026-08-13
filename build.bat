@echo off
REM Builds ImageSearch.exe from tkinter_app.py using PyInstaller.
REM Run this from the ImageSearchBuild folder after `pip install -r requirements.txt`.

REM Pre-download the CLIP model weights (once) so they can be bundled into
REM the .exe - end users then never wait on the first-run download.
if not exist "model_cache" (
    echo Model cache not found - downloading CLIP weights once before building...
    py download_model.py
)

py -m PyInstaller --noconfirm --onedir --windowed --name ImageSearch ^
    --collect-all open_clip ^
    --collect-all torch ^
    --collect-all torchvision ^
    --collect-all msal ^
    --add-data "model_cache;model_cache" ^
    --add-data "assets;assets" ^
    tkinter_app.py

REM Strip bundled third-party license text (torch's kineto/dynolog trees in
REM particular nest ~150 folders deep). Pure attribution files, unused at
REM runtime, but their long paths risk exceeding Windows' 260-char MAX_PATH
REM on extraction elsewhere and silently dropping files/folders.
for /d /r "dist\ImageSearch\_internal" %%D in (*.dist-info) do (
    if exist "%%D\licenses" rd /s /q "%%D\licenses"
)

echo.
echo Build complete. Find ImageSearch.exe in dist\ImageSearch\
