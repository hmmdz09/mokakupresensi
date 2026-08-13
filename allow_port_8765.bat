@echo off
netsh advfirewall firewall add rule name="MOKAKU Presensi Port 8765" dir=in action=allow protocol=TCP localport=8765 profile=private
if errorlevel 1 (
  echo GAGAL: jalankan file ini sebagai Administrator.
  pause
  exit /b 1
)
echo BERHASIL: port TCP 8765 diizinkan pada jaringan Private.
pause
