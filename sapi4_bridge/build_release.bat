@echo off
cargo build --release --target i686-pc-windows-msvc
if %ERRORLEVEL% EQU 0 (
    copy /Y "target\i686-pc-windows-msvc\release\sapi4_bridge.exe" "..\dll\sapi4_bridge_32.exe"
    echo Copied to ..\dll\sapi4_bridge_32.exe
)
