---
name: jurnal-writing
description: Write academic journal articles (IMRAD structure, bilingual abstract, citations) in Indonesian or English as Word documents. Use when the user mentions jurnal, journal article, paper, artikel ilmiah, IMRAD, abstrak/abstract, or submitting to a journal.
license: MIT
compatibility: opencode
metadata:
  audience: researchers
  format: docx
---

# SKILL: Penulisan Artikel Jurnal

Skill menulis artikel jurnal (Indonesia/Inggris) berstruktur **IMRAD** sebagai
dokumen Word, memakai `smart_write.py` (skill `office-word-mcp`) + hardening
`italic-foreign.py` dari repo office-skripsi.

---

## 0. Alur standar

```
1. Tanya target  →  2. Tulis live IMRAD  →  3. Hardening  →  4. Cek kelayakan submit
```

### 1. Tanya target dulu

- Jurnal tujuan (nama, template tersedia? Sinta/Scopus?) — template jurnal
  **selalu menang** atas struktur default di bawah
- Bahasa: Indonesia / Inggris / keduanya (abstrak dua bahasa?)
- Gaya sitasi (APA / IEEE / Vancouver) + minimal jumlah referensi
- Batas halaman/kata + data & hasil yang wajib dimuat (jangan mengarang!)

### 2. Struktur default (kalau tanpa template)

```bash
V=/path/ke/repo-office/.venv-mcp/bin/python
# Artikel sudah TERBUKA di Word; dia satu-satunya target (tidak ada buka/tutup)

$V smart_write.py heading "JUDUL ARTIKEL MAKSIMAL 14 KATA" --level 1 --after "<baris terakhir template>"
$V smart_write.py append "Nama Penulis1, Nama Penulis2" "Afiliasi, email korespondensi" --after "JUDUL ARTIKEL MAKSIMAL 14 KATA"
$V smart_write.py heading "ABSTRAK" --level 1 --after "Afiliasi, email korespondensi"
$V smart_write.py append "Satu paragraf 150-250 kata: masalah, tujuan, metode, hasil utama + angka, simpulan. Diakhiri Kata kunci: 3-5 frasa." --after "ABSTRAK"
$V smart_write.py heading "ABSTRACT" --level 1 --after "Kata kunci:"
$V smart_write.py append "English version of the abstract, same structure." --after "ABSTRACT"
$V smart_write.py heading "1. PENDAHULUAN" --level 1 --after "English version of the abstract"
$V smart_write.py append "Gap penelitian -> rumusan -> tujuan -> kontribusi. Setiap klaim penting bersitasi." --after "1. PENDAHULUAN"
$V smart_write.py heading "2. METODE" --level 1 --after "Setiap klaim penting bersitasi"
$V smart_write.py append "Desain, data/sampel, instrumen, analisis - cukup detail untuk direplikasi." --after "2. METODE"
$V smart_write.py heading "3. HASIL DAN PEMBAHASAN" --level 1 --after "cukup detail untuk direplikasi"
$V smart_write.py append "Sajikan temuan objektif, lalu bahas makna + bandingkan dengan studi sebelumnya." --after "3. HASIL DAN PEMBAHASAN"
$V smart_write.py heading "4. KESIMPULAN" --level 1 --after "bandingkan dengan studi sebelumnya"
$V smart_write.py append "Jawab tujuan, sebut keterbatasan, arah riset lanjut. Tanpa referensi baru." --after "4. KESIMPULAN"
$V smart_write.py heading "DAFTAR PUSTAKA" --level 1 --after "Tanpa referensi baru"
$V smart_write.py append "Penulis. (Tahun). Judul. Jurnal, volume(halaman)." --after "DAFTAR PUSTAKA" --style "Daftar Paragraf1"
$V smart_write.py toc
$V smart_write.py save
```

Aturan CLI ini:

- **Setiap sisip butuh `--after`** (teks jangkar = paragraf yang sudah ada).
  Jangan mengandalkan "akhir dokumen": `count characters` Word berhenti di
  akhir section, bukan akhir berkas (lihat `docs/REALTIME-WORD.md`).
- Jangan pakai dokumen uji terpisah. Kalau artikel belum ada, buat file kosong
  dulu (MCP `create_document` / python-docx), buka di Word, baru jalankan
  perintah di atas dengan jangkar baris pertama.
- Tabel dan gambar belum bisa live: sisipkan lewat pipeline OOXML saat Word
  ditutup, atau MCP tool yang menulis file; sesudah itu lanjut live lagi.
- `delete` hanya menghapus paragraf yang **diawali** teks yang dicari.

Kaedah isi:

| Bagian | Aturan praktis |
|--------|----------------|
| Judul | Spesifik, ada variabel + konteks, tanpa singkatan tak umum |
| Abstrak | Tanpa sitasi, tanpa singkatan tanpa kepanjangan, ada angka hasil |
| Pendahuluan | Pola corong: umum → gap → tujuan; sitasi primer >80% |
| Metode | Replikabel: data, alat, parameter, uji statistik |
| Hasil | Fakta dulu, interpretasi di pembahasan; tiap tabel/gambar dirujuk di teks |
| Kesimpulan | Klaim ≤ bukti; sebut limitasi |

### 3. Hardening

```bash
# Italic istilah asing saja (terutama artikel Indonesia campur istilah Inggris)
python3 /path/ke/office-skripsi/scripts/italic-foreign.py audit Artikel.docx --strip-other
python3 /path/ke/office-skripsi/scripts/italic-foreign.py apply Artikel.docx --strip-other
# Verifikasi render
soffice --headless --convert-to pdf Artikel.italic.docx && pdftotext -layout Artikel.italic.pdf cek.txt
```

### 4. Cek kelayakan submit

- [ ] Semua sitasi in-text ada di Daftar Pustaka (dan sebaliknya)
- [ ] Tabel/gambar bernomor urut + dirujuk di teks
- [ ] Abstrak ≤ batas kata jurnal, kata kunci 3–5
- [ ] Tidak ada data/angka yang dikarang — tiap angka tertelusur ke sumber
- [ ] Similarity siap cek (parafrasa definisi umum dengan kalimat sendiri)

---

## Anti-patterns

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Sitasi dari sumber sekunder/abstrak saja | Baca & sitasi sumber primer |
| Klaim melebihi data | Batasi simpulan pada yang terbukti |
| Satu kalimat penuh miring karena 1 istilah Inggris | `italic-foreign.py --strip-other` |
| Mengarang angka/tabel | Minta data, tandai `[DATA]` bila kosong |
