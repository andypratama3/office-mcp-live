# Format Skripsi PA — Checklist Terukur

Checklist ini disusun dari template resmi kampus (`Format PA 2026.docx`) dan
diuji pada dokumen 120 halaman (817 paragraf, 30 tabel, 40 gambar, 7 section).
Semua nilai di bawah bisa diperiksa ulang dengan skrip, bukan dikira-kira.

---

## 1. Halaman

| Item | Nilai |
|------|-------|
| Ukuran kertas | A4 (21,00 × 29,70 cm) |
| Margin | kiri 4 cm, atas/kanan/bawah 3 cm |
| Jarak header/footer | 1,27 cm |
| Lebar teks portrait | 14,00 cm |
| Lebar teks landscape | 22,70 cm (21 − ... lihat §5) |

---

## 2. Gaya paragraf

| Gaya | First line | Spasi sebelum/sesudah | Baris | Perataan |
|------|-----------|----------------------|-------|----------|
| Normal (isi) | 1,00 cm | 6 / 6 pt | 1,5 | justify |
| Heading 1 (BAB) | 0 | 0 / 18 pt | 1,5 | tengah |
| Heading 2 (subbab) | 0 | 0 / 0 pt | 1,5 | kiri |
| Heading 3 | 0 | 0 / 0 pt | 1,5 | justify |
| Caption | 0 | 0 / 0 pt | 1,0 | tengah |

Catatan penting:

- **Heading 3 hanya miring**, bukan tebal. Banyak dokumen salah di sini.
- Override langsung pada paragraf heading (`w:spacing`, `w:ind`, `w:jc`)
  membatalkan gaya template. Bersihkan sebelum menyimpulkan "sudah sesuai".
- `Normal` mewarisi indentasi ke heading; heading wajib `ind firstLine=0`.

---

## 3. Penomoran

- Bab: `BAB I`, `BAB II`, … (angka Romawi, kapital).
- Subbab: `1.1`, `1.2`, **tanpa titik di akhir** (bukan `1.1.`).
- Sub-subbab: `1.1.1`, `1.1.2`.
- Nomor berasal dari `numbering.xml` (`w:lvlText` = `%1.%2`), bukan diketik.
- Ubah hanya abstractNum yang dipakai gaya heading (`numId` yang menempel di
  `Heading 1/2/3`), jangan semua abstractNum.

---

## 4. Tabel

| Item | Nilai |
|------|-------|
| Padding sel | atas/bawah 0, kiri/kanan 0,19 cm (108 twips) |
| Baris judul | diulang di halaman berikutnya (`w:tblHeader`) |
| Layout | `fixed` agar lebar kolom dihormati |
| Lebar | ≤ lebar teks section |

Word menghapus `tblCellMar` yang nilainya sama dengan default (0/108). Itu
normal — jangan dianggap gagal. Verifikasi harus memperlakukan "tidak ada
`tblCellMar`" sebagai sama dengan 0/0,19/0/0,19 cm.

---

## 5. Tabel lebar → section landscape

Tabel 8 kolom pada lebar 14 cm memecah judul kolom menjadi "N o", "Tahu n".
Solusi yang dipakai:

1. Paragraf **sebelum caption** diberi `sectPr` portrait (menutup section
   portrait).
2. Paragraf **sesudah tabel** diberi `sectPr` landscape (menutup section
   landscape, kembali portrait untuk isi berikutnya).
3. `w:pgSz w="16838" h="11906" w:orient="landscape"`.
4. Lebar tabel = 16838 − margin kiri (2268) − margin kanan (1701) = 12.869
   twips ≈ 22,7 cm.
5. `w:pgNumType` dihapus dari section lanjutan supaya nomor halaman tidak
   mengulang dari 1.

Periksa hasilnya dengan render halaman landscape, bukan hanya teks PDF.

---

## 6. DAFTAR ISI, GAMBAR, TABEL

- DAFTAR ISI/GAMBAR/TABEL adalah **field** (`TOC`), bukan teks manual.
- Field berada di dalam blok `w:sdt`, sehingga tidak terlihat oleh
  `doc.paragraphs`. Hitung entri lewat XML (`w:pStyle` = `TOC1`/`TOC2`/…).
- Prompt "Update Table of Contents" muncul bila `w:updateFields=true`.
  Alur sekali pakai: `--mark` → buka Word → Return → simpan → `--clear`.
- `DAFTAR PUSTAKA` wajib Heading 1 + `numId 0` + page break, kalau tidak ia
  hilang dari DAFTAR ISI.

---

## 7. Pemeriksaan otomatis (contoh hasil)

```
18 PASS / 0 FAIL
Ukuran halaman A4            PASS   21.00 x 29.70 cm
Margin 4/3/3/3 cm            PASS   L4.00 R3.00 T3.00 B3.00
Header/footer 1,27 cm        PASS   1.27 / 1.27
Normal: first line 1,00 cm   PASS   1.00 cm
Normal: spasi 6/6 pt         PASS   6/6 pt
Normal: jarak 1,5 baris      PASS   1.5
Heading 1: indent 0, 0/18 pt PASS   first=0.0 after=18
Heading 2: indent 0, 0/0 pt  PASS   after=0 pt
Heading 3: miring saja       PASS   bold=None italic=True
Caption: spasi sesudah 0     PASS   0 pt
Semua tabel ≤ lebar section  PASS   terlebar 22.70 cm di landscape
Padding sel                  PASS   0 menyimpang
Baris judul tabel berulang   PASS   0 tanpa tblHeader
Nomor subbab tanpa titik     PASS   -
DAFTAR ISI terisi            PASS   87 entri isi, 62 entri gambar/tabel
DAFTAR PUSTAKA ber-heading   PASS   Heading 1
```

Setelah PASS, tetap render ke PDF dan lihat halaman kritis: sampul, daftar
isi, tabel landscape, transisi bab, daftar pustaka.
