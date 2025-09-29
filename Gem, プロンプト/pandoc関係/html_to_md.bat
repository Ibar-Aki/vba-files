@echo off
setlocal enabledelayedexpansion

rem Change to the directory where this script resides
cd /d "%~dp0"

rem Convert all HTML files in this folder to Markdown using Pandoc
for %%F in (*.html) do (
    echo Converting "%%F" to Markdown...
    pandoc "%%F" -f html -t gfm --wrap=preserve -o "%%~nF.md"
)

echo Done.
endlocal
