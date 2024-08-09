@echo off
setlocal enabledelayedexpansion

REM Definir la carpeta de descargas y las carpetas de destino
set "source_folder=%USERPROFILE%\Downloads"
set "destination_folder_images=%USERPROFILE%\Documents\Images"
set "destination_folder_documents=%USERPROFILE%\Documents\Documents"
set "destination_folder_music=%USERPROFILE%\Documents\Music"
set "destination_folder_videos=%USERPROFILE%\Documents\Videos"
set "destination_folder_others=%USERPROFILE%\Documents\Others"

REM Crear las carpetas de destino si no existen
if not exist "%destination_folder_images%" mkdir "%destination_folder_images%"
if not exist "%destination_folder_documents%" mkdir "%destination_folder_documents%"
if not exist "%destination_folder_music%" mkdir "%destination_folder_music%"
if not exist "%destination_folder_videos%" mkdir "%destination_folder_videos%"
if not exist "%destination_folder_others%" mkdir "%destination_folder_others%"

REM Mover archivos según su extensión
for %%f in ("%source_folder%\*") do (
    set "filename=%%~nxf"
    set "ext=%%~xf"
    
    if /I "!ext!"==".jpg" (
        move "%%f" "%destination_folder_images%"
    ) else if /I "!ext!"==".png" (
        move "%%f" "%destination_folder_images%"
    ) else if /I "!ext!"==".doc" (
        move "%%f" "%destination_folder_documents%"
    ) else if /I "!ext!"==".docx" (
        move "%%f" "%destination_folder_documents%"
    ) else if /I "!ext!"==".pdf" (
        move "%%f" "%destination_folder_documents%"
    ) else if /I "!ext!"==".mp3" (
        move "%%f" "%destination_folder_music%"
    ) else if /I "!ext!"==".mp4" (
        move "%%f" "%destination_folder_videos%"
    ) else (
        move "%%f" "%destination_folder_others%"
    )
)

echo Carpeta de descargas ordenada exitosamente.
pause
