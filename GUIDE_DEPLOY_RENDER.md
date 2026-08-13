# 🚀 Panduan Deploy Web Presensi MOKAKU FPMIPA 2026 ke Render.com (100% GRATIS)

Panduan ini akan membantu Anda meng-upload dan menjalankan aplikasi **Presensi MOKAKU FPMIPA 2026** di **Render.com** dengan **Database PostgreSQL Cloud Gratis** agar data 900+ mahasiswa tersimpan permanen dan aman.

---

## 🛠️ Langkah 1: Upload Project ke GitHub

1. Buka [GitHub.com](https://github.com) dan buat repository baru (misal: `presensi-mokaku-fpmipa-2026`).
2. Di laptop Anda (folder `C:\Users\hamdi\presensi-mokaku-fpmipa-2026`), jalankan perintah git berikut:

```bash
git init
git add .
git commit -m "Siap deploy ke Render dengan PostgreSQL"
git branch -M main
git remote add origin https://github.com/USERNAME_ANDA/presensi-mokaku-fpmipa-2026.git
git push -u origin main
```

---

## ☁️ Langkah 2: Deploy ke Render.com via Blueprint (Otomatis)

1. Buka [Render.com](https://render.com) dan login/daftar menggunakan akun GitHub Anda.
2. Di dashboard Render, klik tombol **"New +"** di pojok kanan atas ➔ pilih **"Blueprint"**.
3. Sambungkan (connect) repository `presensi-mokaku-fpmipa-2026` dari GitHub Anda.
4. Beri nama Service, lalu klik **"Apply"**.
5. Render akan secara otomatis membaca file `render.yaml` dan membuat:
   * **Web Service FastAPI (Python)**
   * **Database PostgreSQL Cloud Gratis**

---

## 🔐 Kredensial Default Login Admin Render

Setelah deployment selesai (berstatus *Live*):
* **URL Web:** (Diberikan oleh Render, misal `https://presensi-mokaku-fpmipa-2026.onrender.com`)
* **Username Admin:** `admin`
* **Password Admin:** `Mokaku2026!`

*(Anda bisa mengubah password admin di bagian Environment Variables di dashboard Render).*

---

## ✅ Mengapa Pilihan Ini Terbaik & Aman untuk 900 Mahasiswa?

1. **100% Gratis:** Paket Free Tier Render Web Service + Render PostgreSQL tidak dipungut biaya.
2. **Data Permanen (Persistent):** Menggunakan Database PostgreSQL resmi, data presensi 900 mahasiswa tersimpan aman dan tidak akan hilang meski server restart/redeploy.
3. **Export Excel Tetap Lancar:** Fitur rekap Excel (`export.xlsx`) tetap bisa dipakai kapan saja oleh admin.
