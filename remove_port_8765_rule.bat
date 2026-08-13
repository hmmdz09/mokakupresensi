@echo off
netsh advfirewall firewall delete rule name="MOKAKU Presensi Port 8765"
if errorlevel 1 (
  echo GAGAL: jalankan file ini sebagai Administrator.
  pause
  exit /b 1
)
echo BERHASIL: aturan firewall port 8765 dihapus.
pause
