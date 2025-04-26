@echo off
echo Building Word Document Merger executable...
echo Developed by Aaqib Jeelani - instagram.com/aaqibjeelani

REM Create the sample directory if it doesn't exist
if not exist "sample" mkdir sample

REM Check if icon exists, otherwise run the icon creation script
if not exist "app_icon.ico" (
    echo Creating application icon...
    python create_icon.py
)

REM Run PyInstaller to create the executable
echo Building executable with PyInstaller...
pyinstaller --name=WordDocumentMerger ^
            --onefile ^
            --icon=app_icon.ico ^
            --noconsole ^
            --add-data="sample;sample" ^
            --add-data="app_icon.ico;." ^
            merge_word_docs.py

REM Copy sample directory to dist for convenience
if exist "dist" (
    if exist "sample" (
        if not exist "dist\sample" (
            echo Copying sample directory to dist...
            xcopy "sample" "dist\sample\" /E /I /Y
        )
    )
)

echo.
echo Build completed!
echo The executable can be found in the 'dist' directory
echo.
echo Developed by Aaqib Jeelani - instagram.com/aaqibjeelani
pause 