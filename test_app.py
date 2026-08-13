import io
import sqlite3
from pathlib import Path

import openpyxl
from fastapi.testclient import TestClient

import app as appmod


def reset_db(path):
    appmod.DB_PATH = path
    if appmod.DB_PATH.exists():
        appmod.DB_PATH.unlink()
    appmod.init_db()


def test_complete_attendance_flow(tmp_path):
    # Seluruh data uji wajib berada di database sementara agar tidak
    # mencemari presensi.db yang dipakai saat pengujian manual.
    reset_db(tmp_path / 'presensi-test.db')
    client = TestClient(appmod.app)

    response = client.post('/register', data={
        'nim': '2406050', 'name': 'M. Hamdi',
        'email': 'hamdi@example.test',
        'study_program': 'Pendidikan Fisika'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert 'QR Presensi' in response.text
    assert '2406050' in response.text

    client.get('/logout')
    response = client.post('/login', data={'nim': '2406050', 'email': 'hamdi@example.test'}, follow_redirects=True)
    assert response.status_code == 200
    assert 'Riwayat Kehadiran' in response.text

    # Registrasi dan login hanya membuat sesi autentikasi. Kehadiran baru
    # boleh tercatat setelah QR dipindai admin atau NIM dicatat manual.
    with appmod.db() as con:
        assert con.execute('SELECT COUNT(*) FROM attendance').fetchone()[0] == 0

    with appmod.db() as con:
        user = con.execute('SELECT * FROM users WHERE nim=?', ('2406050',)).fetchone()
        session = con.execute('SELECT * FROM sessions LIMIT 1').fetchone()
    token = appmod.sign_qr(user)

    client.get('/logout')
    response = client.post('/admin/login', data={'username': 'admin', 'password': 'Mokaku2026!'}, follow_redirects=True)
    assert response.status_code == 200
    assert 'Scanner Kehadiran' in response.text

    scan = client.post(f'/admin/api/scan/{session["id"]}', json={'token': token})
    assert scan.status_code == 200
    assert scan.json()['ok'] is True
    assert scan.json()['nim'] == '2406050'

    duplicate = client.post(f'/admin/api/scan/{session["id"]}', json={'token': token})
    assert duplicate.json()['ok'] is False
    assert duplicate.json()['duplicate'] is True

    invalid = client.post(f'/admin/api/scan/{session["id"]}', json={'token': 'MOKAKU26:2406050:fake:bad'})
    assert invalid.json()['ok'] is False

    export = client.get('/admin/export.xlsx')
    assert export.status_code == 200
    assert export.headers['content-type'].startswith('application/vnd.openxmlformats')
    workbook = openpyxl.load_workbook(io.BytesIO(export.content))
    sheet = workbook['Rekap Presensi']
    values = list(sheet.values)
    assert values[0][0:4] == ('NIM', 'Nama', 'Email', 'Jenis Akun')
    assert values[1][0] == '2406050'
    assert values[1][3] == 'Peserta'
    assert 'Hadir' in values[1]
    assert values[1][-3] == 1
    assert values[1][-2] == 100
    log = workbook['Log Kehadiran Real-Time']
    log_values = list(log.values)
    assert log_values[0][0:8] == ('No', 'Timestamp Scan', 'Tanggal', 'Waktu', 'NIM', 'Nama', 'Email', 'Jenis Akun')
    assert log_values[1][4] == '2406050'
    assert log_values[1][7] == 'Peserta'
    assert log_values[1][8] == 'Pendidikan Fisika'
    assert log_values[1][9] == '-'
    assert log_values[1][13] == 'QR Scanner'

    client.get('/logout')
    history = client.post('/login', data={'nim': '2406050', 'email': 'hamdi@example.test'}, follow_redirects=True)
    assert 'Hadir' in history.text
