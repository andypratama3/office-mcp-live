---
name: office-word-mcp
description: Create, read, and edit Microsoft Word (.docx) files via the Office-Word MCP server, an OOXML pipeline, and live AppleScript editing of the open document. Use when the user mentions Word documents, DOCX read/write, live Word editing, tables, headings, TOC/field updates, or document automation.
license: MIT
compatibility: opencode
metadata:
  audience: developers
  format: docx
---

# SKILL: Office-Word MCP — Kerja Word Tanpa Tutup-Buka

Skill ini mengatur **cara** mengedit `.docx`: kapan lewat file (OOXML),
kapan lewat dokumen yang sedang terbuka (live), dan kapan Word boleh dibuka.

Repo: fork `gongrzhe/office-word-mcp-server` + helper live macOS.

---

## 0. Tiga jalur, satu aturan

| Jalur | Mekanisme | Pakai untuk | Word harus |
|-------|-----------|-------------|------------|
| **OOXML pipeline** (python-docx) | tulis file di disk | gaya, section, numbering, tabel, field | **tertutup** |
| **Live** (AppleScript ke Word) | tulis ke memori Word | isi paragraf/judul, lihat langsung | **terbuka** |
| **MCP tools** (`word-document-server`) | python-docx dibungkus tool | baca isi/outline, tabel, cari teks | bebas |

**Aturan utama: Word dibuka SEKALI per sesi, dokumen dibiarkan terbuka sampai
pekerjaan selesai.** Jangan pola buka → simpan → tutup → buka lagi untuk tiap
perubahan. Setiap `save` Word menulis ulang XML dan bisa membuang atribut yang
kita set (lihat §4).

Urutan kerja yang benar:

```
1. Fase struktur (Word TERTUTUP)  → pipeline OOXML, verifikasi skrip
2. Buka Word SEKALI                → dokumen tampil, TOC/field di-bake
3. Fase isi (Word TERBUKA)         → live write, terlihat realtime
4. Selesai                         → simpan sekali, ekspor PDF, bersihkan flag
```

Kalau harus kembali ke fase struktur: simpan + tutup Word dulu, jalankan
pipeline, lalu buka lagi **satu kali**. Bukan berkali-kali.

---

## 1. Fase struktur — pipeline OOXML

Jangan pakai Word untuk mengubah format massal. Tulis skrip kecil, pola:
`dry run` → `--apply` → `verify` → `render`.

```bash
VENV=/path/ke/repo-office/.venv-mcp/bin/python
$VENV pipeline/align_template.py            # laporan saja
$VENV pipeline/align_template.py --apply    # + backup otomatis
$VENV pipeline/verify_format.py --write     # PASS/FAIL + laporan markdown
```

Yang wajib ada di setiap skrip: backup bertanggal, mode `--apply`, dan
laporan perubahan (bukan diam-diam menulis).

Contoh isi pipeline (lihat repo skripsi):

| Skrip | Tugas |
|-------|-------|
| `align_template.py` | gaya, spasi, numbering, padding tabel, lebar tabel per section |
| `finalize.py` | heading daftar pustaka, field TOC `dirty`, `w:updateFields` |
| `fix_wide_table.py` | section landscape untuk tabel lebar (mis. 22,7 cm) |
| `verify_format.py` | pemeriksaan PASS/FAIL, tanpa mengubah dokumen |

### Lebar tabel mengikuti section, bukan angka mati

Portrait: lebar teks 14,00 cm. Landscape: lebar kertas − margin kiri − margin
kanan (A4 = 22,70 cm dengan margin 4/3). Tabel 8 kolom **tidak** dibiarkan
dipaksa 14 cm — judul kolom pecah jadi "N o", "Tahu n". Solusinya section
landscape khusus tabel itu, lalu kembali portrait.

---

## 2. Fase isi — live, terlihat seperti mengetik

Dokumen sudah terbuka; semua tulisan dikirim ke Word, bukan ke file.

```bash
$VENV pipeline/live.py status     # Word jalan? dokumen terbuka? ada dialog?
$VENV pipeline/live.py open       # buka sekali, tidak pernah menutup
$VENV pipeline/live.py tail       # 5 paragraf terakhir + gayanya
```

Pola aman menulis blok:

1. Baca semua teks paragraf **sekali** panggilan (`every paragraph`), hitung
   indeks target di Python. Jangan loop `paragraph i` 2000× dari AppleScript.
2. Tulis blok dengan satu panggilan: tempel teks berisi `return` pada paragraf
   jangkar, sehingga paragraf baru mewarisi gaya jangkar.
3. Terapkan gaya per paragraf baru dalam **satu** panggilan berisi `repeat`.
4. Verifikasi: baca ulang teks + gaya paragraf yang baru.

Jangan memakai `paste_copy_after` atau clipboard: Word for Mac menggantung
saat ada dialog modal dan tidak ada timeout yang menyelamatkan.

### Nomor paragraf Word ≠ nomor python-docx

Word menghitung paragraf **di dalam sel tabel** juga (contoh: 2.042 vs 817).
Jangan pernah memakai indeks dari satu sisi di sisi lain. Untuk live, selalu
baca indeks dari Word; untuk OOXML, selalu dari `doc.paragraphs`.

---

## 3. Field TOC: bake sekali, jangan beri prompt

Prompt **"Update Table of Contents"** muncul karena `w:updateFields=true`.
Dua langkah, sekali saja per dokumen:

```
A. Nyalakan  : semua fldChar diberi dirty="true", settings += w:updateFields
B. Buka Word : prompt muncul → tekan Return (Yes) → Word membangun ulang TOC
C. Simpan    : save (tanpa Save As), lalu tutup
D. Matikan   : hapus w:updateFields dan seluruh atribut dirty
```

Setelah D, dokumen tidak pernah bertanya lagi. Jangan biarkan flag menyala di
file yang diserahkan ke penulis.

```bash
$VENV pipeline/bake_fields.py --mark     # A
# ... buka Word sekali, tekan Return, simpan, tutup ...
$VENV pipeline/bake_fields.py --clear    # D
```

`do Visual Basic` dari AppleScript **tidak bisa** dipakai (error -2741);
update TOC lewat Word sendiri, bukan lewat skrip.

---

## 4. Yang Word buang saat menyimpan

Word menulis ulang XML. Yang hilang dan harus diperiksa **setelah** Word
menyimpan:

| Atribut | Perilaku Word | Sikap |
|---------|---------------|-------|
| `w:tblCellMar` 0/108 | dihapus karena sama dengan default | anggap setara default, jangan FAIL |
| `w:updateFields` | hilang setelah field di-bake | memang diinginkan |
| `dirty="true"` | hilang setelah field di-update | memang diinginkan |
| padding/gaya tabel lain | dipertahankan | tetap diverifikasi |

Karena itu: **verifikasi dijalankan setelah Word menyimpan**, bukan sebelum.
Angka PASS sebelum save bisa berubah jadi FAIL sesudahnya.

---

## 5. Render untuk verifikasi visual

Word yang sedang terbuka bisa mengekspor PDF tanpa menyimpan `.docx`:

```applescript
tell application "Microsoft Word"
  save as active document file name "/path/out.pdf" file format format PDF
end tell
```

Lalu periksa objeknya, bukan hanya teksnya:

```bash
pdfinfo out.pdf | grep Pages
pdftotext -layout out.pdf out.txt
pdftoppm -f 21 -l 21 -r 100 -png out.pdf hal21   # halaman landscape tabel lebar
```

`soffice` tidak selalu ada di macOS; jangan diandalkan sebagai satu-satunya
jalur render.

---

## 6. MCP server langsung

Untuk baca isi/outline atau operasi batch saat Word tidak terlibat:

```bash
PYTHONPATH=/path/ke/repo-office $VENV -c "
import asyncio
from word_document_server.tools import document_tools, content_tools
async def main():
    print(await content_tools.add_paragraph('lap.docx', 'Isi...'))
    print(await document_tools.get_document_text('lap.docx'))
asyncio.run(main())
"
```

Pendaftaran ke OpenCode:

```json
{
  "mcp": {
    "word-document-server": {
      "type": "local",
      "command": "/path/ke/repo-office/.venv-mcp/bin/python",
      "args": ["/path/ke/repo-office/word_mcp_server.py"],
      "environment": {
        "PYTHONPATH": "/path/ke/repo-office",
        "MCP_TRANSPORT": "stdio"
      },
      "disabled": false
    }
  }
}
```

Uji sebelum percaya: kirim `initialize` + `tools/list` via stdio dan pastikan
server membalas (versi ini 54 tool, FastMCP 4.1.8). Kalau balasan kosong,
tunggu lebih lama — server butuh ~2–3 detik untuk siap.

---

## 7. Anti-pattern

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Buka–tutup Word tiap perubahan | Satu sesi, dokumen tetap terbuka |
| Ubah file `.docx` saat Word membukanya | Live write, atau tutup Word dulu |
| `paste_copy_after` / clipboard | Sisip teks lewat object model |
| `do Visual Basic` | Update field lewat Word + `updateFields` |
| Loop `paragraph i` 2000× | Ambil semua teks sekali, hitung di Python |
| Anggap `python-docx` melihat TOC | TOC ada di dalam `w:sdt`; baca lewat XML |
| Verifikasi sebelum Word menyimpan | Verifikasi sesudah, karena Word menulis ulang |
| Klaim "sudah rapi" tanpa render | `pdftoppm` halaman kritis, lihat gambarnya |
| Biarkan `updateFields` menyala | Bake, lalu hapus flag |
