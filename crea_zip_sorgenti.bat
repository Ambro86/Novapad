@echo off
set ZIP_NAME=rustnotepad_source.zip
echo Creazione dell'archivio %ZIP_NAME% in corso...

if exist "%ZIP_NAME%" del "%ZIP_NAME%"

:: Usa tar (incluso in Windows 10/11) per creare lo zip solo con il codice da modificare
tar -a -c -f "%ZIP_NAME%" assets src i18n Cargo.toml Cargo.lock build.rs bridge scripts packager vcpkg.json app.manifest index.html donations*.txt CHANGELOG* README*.md

echo.
echo Archivio %ZIP_NAME% creato con successo!
pause
