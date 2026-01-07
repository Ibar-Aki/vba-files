@echo off
setlocal enabledelayedexpansion

rem Change to the directory where this script resides
cd /d "%~dp0"

set "REFERENCE_DOC=reference.docx"

for %%F in (*.md) do (
    echo Converting "%%F" to DOCX...
    if exist "!REFERENCE_DOC!" (
        pandoc "%%F" -f gfm -t docx --reference-doc="!REFERENCE_DOC!" --toc --toc-depth=3 --extract-media="media_%%~nF" -o "%%~nF.docx"
    ) else (
        pandoc "%%F" -f gfm -t docx --toc --toc-depth=3 --extract-media="media_%%~nF" -o "%%~nF.docx"
    )
)

echo Done.
endlocal
