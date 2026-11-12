---
name: skripsi-writing
description: Write and repair Indonesian thesis/skripsi documents (BAB I–V) with academic structure, EYD/PUEBI language, campus template formatting, realtime Word editing, table/TOC fixes, and render verification. Use when the user mentions skripsi, thesis, BAB, abstrak, latar belakang, pedoman format, or academic final-project documents.
license: MIT
compatibility: opencode
metadata:
  audience: students
  format: docx
---

# SKILL: Penulisan & Perbaikan Skripsi (Indonesia)

Dua pekerjaan berbeda, dua alat berbeda:

| Pekerjaan | Alat | Word |
|-----------|------|------|
| Menulis/mengubah isi | `python-docx` + `open ... revert true` | **terbuka, sekali** |
| Menyeragamkan format, tabel, section, field | pipeline OOXML | **tertutup** |

Aturan pokok: **Word dibuka sekali per sesi dan dibiarkan terbuka.** Jangan
buka–tutup untuk tiap perubahan; setiap `save` Word menulis ulang XML dan bisa
membuang atribut hasil pipeline.

---

## 0. Sebelum menyentuh apa pun

1. Kerjakan **salinan** (`skripsi_fix/`), jangan file asli (`skripsi/`).
2. Baca dokumen dulu, jangan menebak: ekstrak teks + tabel + gaya.
3. Cari template resmi kampus (mis. `Format PA 2026.docx`). Kalau tidak ada,
   catat asumsi format di `docs/`.
4. Tulis temuan di `docs/`, bukan di chat.

---

## 1. Tanya struktur, jangan mengarang

Minimal sebelum menulis: judul, nama, NIM, prodi, tahun, penomoran bab,
judul sub-bab, dan **data/angka yang harus muncul**.

Aturan data:

- Angka penelitian (jumlah responden, skor SUS, hasil uji) **hanya** dari
  berkas bukti (laporan pengujian, log, screenshot). Tidak ada bukti → jangan
  tulis.
- Fitur yang diuji harus sama dengan yang ada di aplikasi dan laporan uji.
- Klaim "layak/praktis/efektif" tanpa instrumen → hapus atau turunkan menjadi
  deskripsi fungsional.

---

## 2. Selaraskan format ke template (sekali, di awal)

Baca template, catat selisihnya, lalu terapkan lewat skrip — bukan manual.

| Aspek | Aturan tipikal template PA | Cara cek |
|-------|---------------------------|----------|
| Paragraf isi | first line 1,00 cm; spasi 6/6 pt; 1,5 baris; justify | `Normal` |
| Heading 1 (BAB) | indent 0; spasi 0/18 pt; tengah; bernomor | `Heading 1` |
| Heading 2 | indent 0; spasi 0/0 pt; kiri | `Heading 2` |
| Heading 3 | indent 0; **miring saja** (bukan tebal) | `Heading 3` |
| Caption | spasi sesudah 0; 1,0 baris; tengah | `Caption` |
| Nomor subbab | `4.1`, bukan `4.1.` | `w:lvlText` |
| Sel tabel | padding atas/bawah 0, kiri/kanan 0,19 cm; baris judul diulang | `tblCellMar`, `tblHeader` |
| Lebar tabel | ≤ lebar teks section (portrait 14 cm) | `tblGrid` |

Pola skrip wajib: **backup bertanggal → dry run → `--apply` → laporan**.
Contoh nama: `align_template.py`, `finalize.py`, `verify_format.py`.

Hal yang sering terlewat:

- Override langsung (`w:spacing`, `w:ind`, `w:jc`) pada paragraf heading/caption
  harus dibersihkan, kalau tidak gaya template tidak berpengaruh.
- Nomor subbab diambil dari `numbering.xml`; ubah `lvlText` pada abstractNum
  yang benar-benar dipakai heading, jangan semua.
- `DAFTAR PUSTAKA` harus ber-style Heading 1 + `numId 0` + page break, kalau
  tidak ia hilang dari DAFTAR ISI.

---

## 3. Tulis isi secara realtime

Word sudah terbuka; jangan menutupnya. Satu langkah kerja = satu blok:

```bash
VENV=/path/ke/repo/.venv-mcp/bin/python

$VENV work/pipeline/realtime.py status
$VENV work/pipeline/realtime.py open                 # buka sekali
$VENV work/pipeline/realtime.py focus "Landasan Teori"
$VENV work/pipeline/realtime.py tail                 # paragraf terakhir
$VENV work/pipeline/realtime.py pdf work/render/cek.pdf
```

Di Python:

```python
from realtime import apply

def tulis_bab4(doc):
    p = doc.add_paragraph("4.1 Hasil Implementasi", style="Heading 2")
    doc.add_paragraph("Isi paragraf ...", style="Normal")
    return ["BAB IV: heading + 1 paragraf"]

apply(tulis_bab4, anchor="4.1 Hasil Implementasi", label="bab4")
```

`apply` melakukan: backup → tulis file → `open revert true` → `execute find`
ke jangkar (jendela Word bergulir ke situ) → aktifkan Word. Hasilnya terlihat
seperti mengetik, tanpa tutup-buka.

Jangan pakai AppleScript loop `paragraph i` untuk mencari teks: 2000+ paragraf
membuat Word menggantung. Pakai `execute find` (native, ~1 detik).

---

## 4. Tabel lebar → section landscape

Tabel 8 kolom tidak muat di 14 cm: judul kolom pecah jadi "N o", "Tahu n".
Jangan perkecil font atau paksa 14 cm. Buat section landscape khusus tabel itu:

1. `sectPr` portrait ditempel pada paragraf sebelum caption (menutup section
   portrait).
2. `sectPr` landscape ditempel pada paragraf sesudah tabel (menutup section
   landscape).
3. `pgSz w=16838 h=11906 orient=landscape`; tabel dilebarkan ke lebar teks
   landscape (A4 − margin kiri − margin kanan = 22,7 cm).
4. Hapus `pgNumType` pada section lanjutan supaya nomor halaman tidak
   mengulang dari 1.

Lebar tabel harus dihitung **per section**, bukan angka mati.

---

## 5. DAFTAR ISI: bake sekali, tanpa prompt

Prompt "Update Table of Contents" muncul karena `w:updateFields=true`.
Sekali per dokumen:

```bash
$VENV work/pipeline/bake_fields.py --mark      # tandai field dirty
# buka Word sekali → tekan Return pada prompt → Word membangun ulang TOC
# simpan dari Word
$VENV work/pipeline/bake_fields.py --clear     # hapus updateFields + dirty
$VENV work/pipeline/bake_fields.py --status    # pastikan bersih
```

Hasil akhir: entri DAFTAR ISI/GAMBAR/TABEL sudah terisi dan prompt tidak
pernah muncul lagi. `do Visual Basic` dari AppleScript **tidak** berfungsi
(error -2741) — jangan diandalkan.

---

## 6. Verifikasi: PASS/FAIL + render

Dua lapis, keduanya wajib:

```bash
$VENV work/pipeline/verify_format.py --write    # 18 pemeriksaan, laporan .md
$VENV work/pipeline/realtime.py pdf work/render/cek.pdf
pdfinfo work/render/cek.pdf | grep Pages
pdftotext -layout work/render/cek.pdf work/render/cek.txt
pdftoppm -f 21 -l 21 -r 100 -png work/render/cek.pdf work/render/hal21
```

Baca gambarnya untuk halaman kritis: tabel lebar (landscape), caption,
halaman transisi antar bab. Angka di laporan tidak menggantikan melihat hasil.

Verifikasi dijalankan **setelah** Word menyimpan, karena Word membuang
`tblCellMar` yang sama dengan default dan menghapus flag field.

---

## 7. Anti-pattern

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Mengarang data, angka, sitasi, screenshot | Minta bukti; tandai `[DATA]` dan laporkan |
| Buka–tutup Word tiap perubahan | `open ... revert true`, satu sesi |
| Kirim 40 paragraf satu per satu via AppleScript | Satu blok per `apply` |
| Loop 2000 paragraf untuk cari teks | `execute find` |
| Paksa tabel 8 kolom ke 14 cm | Section landscape |
| Biarkan `updateFields` menyala | Bake lalu `--clear` |
| Klaim "sudah sesuai pedoman" tanpa cek | `verify_format.py` + render |
| Ubah file asli | Kerjakan salinan, backup bertanggal |
