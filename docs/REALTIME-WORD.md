# Realtime Word Editing (macOS)

Mengedit dokumen Word yang **sedang terbuka** tanpa menutup/membukanya, dengan
jendela Word ikut bergulir ke titik kerja. Dipakai untuk dokumen besar
(contoh: 118 halaman, 40 gambar, 30 tabel, 7 section).

Inti: `office-mcp-live/live_engine.py` (AppleScript murni), diuji oleh
`tests/test_live_engine.py`.

---

## Aturan utama

1. **Jangan pernah menulis file `.docx` di belakang Word.** Word memegang
   salinan di memori; tulisan dari python-docx akan tertimpa saat Word
   menyimpan. Semua perubahan lewat AppleScript, lalu `save active document`.
2. **`open ... revert true` tidak dipakai.** Pada Word for Mac, perintah itu
   diabaikan (terbukti dari uji pada dokumen uji), jadi tidak ada "reload".
3. **Satu dokumen kerja saja.** Jangan membuka dokumen uji/lain saat mengedit;
   `active document` harus tetap dokumen yang dikerjakan.
4. **Ekspor PDF hanya ke folder dokumen.** `save as ... format PDF` gagal
   (`-1708`) bila target di folder lain.
5. **Verifikasi setelah save.** Word menulis ulang XML (mis. menghapus
   `tblCellMar` yang sama dengan default), jadi periksa hasil akhir dari file,
   bukan dari keadaan sebelum simpan.

---

## Operasi yang terbukti

```applescript
tell application "Microsoft Word"
  with timeout of 240 seconds
    -- pindah ke awal dokumen (find selalu mulai dari story selection)
    tell active document
      set r0 to create range start 0 end 0
      select r0
    end tell

    -- cari kemunculan ke-N
    repeat with i from 1 to 2
      set f to execute find (find object of selection) find text "DAFTAR PUSTAKA" ¬
        match forward true match case true
    end repeat

    -- baca posisi: HARUS di luar `tell active document`
    set s0 to start of content of (text object of selection)
    set e0 to end of content of (text object of selection)
  end timeout
end tell
```

| Kebutuhan | Cara yang benar |
|---|---|
| Pindah view | `select (create range start X end Y)` |
| Baca posisi | `start/end of content of (text object of selection)` **di luar** `tell active document` |
| Tambah paragraf | `insert text (¶ & teks) at (create range start P end P)`; `P` = `end of content` paragraf jangkar |
| Ganti teks | `execute find ... replace replace all` pada range `0..count characters` |
| Hapus paragraf | `delete text object of (paragraph 1 of text object of selection)` |
| Hapus page break manual | `content of (create range start P end P+1)` → kosongkan bila `chr 12` |
| Gaya paragraf | `set style of (paragraph 1 of text object of selection)` |
| Gaya Heading 3 | `set bold of font object of (Word style "Heading 3") to true` |
| Perbarui TOC | `update (table of contents i)` untuk semua TOC |
| Simpan | `save active document` |
| Ekspor PDF | `save as active document file name "/folder/dokumen/x.pdf" file format format PDF` |

Tambahan penting:

- **Paragraf, bukan selection.** `insert text` pada selection yang di-collapse
  bisa menaruh teks di ujung dokumen. Selalu pakai offset absolut.
- **Jangan gabungkan find dan baca offset dalam satu `tell active document`.**
  `text object of selection` menjadi `missing value`; lakukan find dulu, lalu
  baca offset di blok `tell application` yang sama tetapi di luar `tell active
  document`.
- **`execute find` selalu mulai dari story selection.** Reset ke `range 0..0`
  sebelum mencari, dan ulangi `execute find` untuk kemunculan ke-N.
- **Hasil find berupa boolean tidak bisa di-coerce langsung** di dalam blok
  `tell`; simpan ke flag string (`"yes"`/`"no"`) lalu periksa.
- **Jangkar jangan memakai teks yang ada di DAFTAR ISI.** Cari kemunculan ke-2
  (isi dokumen), bukan ke-1 (field TOC).
- **`paragraph N of text object` bisa memberi posisi salah** untuk teks di
  dalam field. Hitung batas paragraf dengan memindai `¶` (ASCII 13) di sekitar
  posisi hasil find.
- **Sebagian tabel menolak `row`/`cell`.** Untuk tabel seperti itu, ambil
  `text object of table`, lalu `execute find` di dalam selection tabel untuk
  memformat baris judul. Perhatikan juga bahwa nomor `table i` Word bisa
  bergeser satu setelah tabel tertentu; cocokkan dulu isi tabel.

---

## Jebakan lain yang sudah teruji

1. **`page number of selection`** → `missing value`. Pakai
   `get selection information (selection) information type active end page
   number`.
2. **`wrap find` bukan enum valid.** Cukup `match forward true`.
3. **`activate` tidak selalu memindahkan Space desktop.** Bila perlu:
   `open -a "Microsoft Word"`.
4. **`content of paragraph N`** mengembalikan objek gaya. Yang benar:
   `content of text object of (paragraph N ...) as string`.
5. **`list format` bisa dipakai untuk melepas nomor**:
   `apply number default (list format of (text object of selection))`
   bersifat toggle; panggil sekali lagi bila nomor justru muncul.
6. **Padding sel 0,19 cm mengurangi lebar teks tiap kolom.** Untuk tabel 8
   kolom, hitung ulang lebar kolom (`tblW` total tetap 14 cm) atau turunkan
   ukuran font; jangan mengandalkan lebar lama.
7. **Judul tabel/gambar tidak otomatis tebal.** Baris judul tabel perlu
   di-`bold`; gaya `Caption` memang tidak tebal menurut template.
8. **Gaya `Heading 3` dokumen ini = tebal + miring.** Template hanya miring;
   target dokumen yang dipakai adalah tebal + miring.
9. **`count characters of text object of active document` BUKAN panjang
   dokumen.** Nilainya berhenti di akhir **section tempat selection berada**
   (mis. 135.523 dari total ±137.000), dan `count paragraphs of active
   document` = 0. Kalau dipakai sebagai posisi "akhir dokumen", teks akan
   masuk ke tengah paragraf terakhir section itu. Karena itu penambahan
   paragraf **selalu** memakai jangkar (`insert_paragraphs`), tidak pernah
   offset "akhir dokumen".
10. **`delete text object of (paragraph ...)` menghapus SELURUH paragraf.**
   Bila teks yang dicari ada di tengah paragraf, kalimat sebelum dan sesudahnya
   ikut hilang. `delete_paragraph()` sekarang menolak bila paragraf tidak
   diawali teks yang dicari (`force=True` untuk memaksa).

---

## Alur kerja

```
pilih jangkar di isi dokumen (bukan TOC)
        ↓
hitung posisi absolut (paragraph_range)
        ↓
ubah lewat AppleScript (insert / style / format)
        ↓
selection dipindah ke titik kerja → view ikut bergulir
        ↓
update TOC bila judul/halaman berubah
        ↓
save active document
        ↓
verifikasi dari file + ekspor PDF ke folder dokumen
```

---

## CLI `smart_write.py` (live murni)

```bash
V=.venv-mcp/bin/python

$V smart_write.py status                      # nama, path, halaman, tersimpan
$V smart_write.py info                        # + jumlah paragraf/kata/gambar
$V smart_write.py read --from 0 --chars 2000
$V smart_write.py goto "5.2 Saran" --n 2      # pindahkan view ke titik kerja
$V smart_write.py show "Adi. (2015)"          # baca satu paragraf
$V smart_write.py style "Judul" "Heading 2"
$V smart_write.py replace "lama" "baru"
$V smart_write.py append "paragraf baru" --after "Adi. (2015)" --style "Daftar Paragraf1"
$V smart_write.py heading "Metode" --after "Pendahuluan" --level 2
$V smart_write.py pagebreak --after "LAMPIRAN A"
$V smart_write.py delete "Penanda uji"        # hanya paragraf yang diawali teks ini
$V smart_write.py toc                         # perbarui daftar isi/gambar/tabel
$V smart_write.py save
$V smart_write.py pdf "/folder/dokumen/keluaran.pdf"
```

`--file PATH` (sebelum atau sesudah perintah) memastikan dokumen aktif memang
berkas itu; kalau bukan, CLI berhenti dengan kode 2 tanpa menyentuh apa pun.
Tidak ada mode file, tidak ada reload, tidak ada tutup-buka: dokumen yang
terbuka adalah satu-satunya target.

---

## API `live_engine.py`

```python
import live_engine as le

le.status()                    # name, path, saved, pages, page
le.goto("Landasan Teori", 2)   # pindah selection/view
le.paragraph_range("Adi. (2015)", 1)   # (start, end-of-content)
le.insert_paragraphs("Adi. (2015)", [("Daftar Paragraf1", "Teks baru")])
le.set_style("Judul", "Heading 2")
le.delete_paragraph("Teks uji")        # menolak bila paragraf tidak diawali teks itu
le.delete_paragraph("Teks uji", force=True)
le.replace_all("lama", "baru")
le.set_page_break_before("LAMPIRAN A", True)
le.delete_blank_before("BAB V PENUTUP")
le.update_toc()
le.save()
le.export_pdf("/folder/dokumen/keluaran.pdf")
```

Uji:

```bash
.venv-mcp/bin/python tests/test_live_engine.py            # kompilasi semua skrip
WORD_LIVE_TEST=1 .venv-mcp/bin/python tests/test_live_engine.py   # baca dokumen nyata
```

Uji tulis yang aman: sisipkan penanda unik setelah satu paragraf, baca kembali,
hapus penanda, lalu pastikan paragraf tetangga tidak berubah.

---

## Kapan memakai pipeline OOXML

Pipeline OOXML (`work/pipeline/`) dipakai untuk **audit dan verifikasi**:
`verify_format.py`, `realtime.py`, `fix_wide_table.py`. Perubahan struktural
besar (section, numbering, tabel) sekarang bisa dilakukan live, tetapi lebih
aman disiapkan sebagai skrip idempoten dan diverifikasi setelah `save`.

Yang tidak bisa/tidak boleh lewat jalur file saat Word terbuka:

- menulis `.docx` dengan python-docx lalu berharap Word memuat ulang;
- memakai `open ... revert true`;
- membuka dokumen kedua sebagai "dokumen uji".
