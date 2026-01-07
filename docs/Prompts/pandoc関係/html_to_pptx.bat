@echo off
setlocal enabledelayedexpansion

rem Change to the directory where this script resides
cd /d "%~dp0"

rem Use reference.pptx if it exists for consistent styling
set "REFERENCE=reference.pptx"

for %%F in (*.html) do (
    echo Converting "%%F" to PPTX...
    if exist "!REFERENCE!" (
        pandoc "%%F" -f html -t pptx -s --reference-doc="!REFERENCE!" -o "%%~nF.pptx"
    ) else (
        pandoc "%%F" -f html -t pptx -s -o "%%~nF.pptx"
    )
)

echo Done.
endlocal
