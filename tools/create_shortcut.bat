@echo off
rem Creates a desktop shortcut "Reactivity Simulator" that starts app.py without a console window
rem and uses assets\app.ico as its icon. Run once: double-click this file or "tools\create_shortcut.bat".
setlocal
set "ROOT=%~dp0.."
for %%I in ("%ROOT%") do set "ROOT=%%~fI"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$py = (Get-Command pythonw -ErrorAction SilentlyContinue).Source;" ^
  "if (-not $py) { Write-Host 'pythonw.exe not found in PATH'; exit 1 };" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Reactivity Simulator.lnk');" ^
  "$s.TargetPath = $py; $s.Arguments = '\"%ROOT%\app.py\"'; $s.WorkingDirectory = '%ROOT%';" ^
  "$s.IconLocation = '%ROOT%\assets\app.ico'; $s.Description = 'Reactivity simulator (KhAI)'; $s.Save();" ^
  "Write-Host 'Shortcut created on the Desktop: Reactivity Simulator'"
endlocal
pause
