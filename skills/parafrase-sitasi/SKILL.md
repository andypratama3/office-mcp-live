---
name: parafrase-sitasi
description: Turunkan kemiripan teks (Turnitin/similarity) secara sah lewat parafrase tataran ide, sintesis, dan sitasi yang benar, lalu ukur tumpang-tindih n-gram terhadap sumber yang kamu berikan. Use when the user mentions Turnitin, similarity, plagiarisme, parafrase, kutipan, sitasi, atau patchwriting pada skripsi/artikel.
license: MIT
compatibility: opencode
metadata:
  audience: researchers
  format: docx
---

# SKILL: Parafrase, Sintesis, dan Sitasi (Turunkan Kemiripan Secara Sah)

Skill ini menurunkan **kemiripan teks** dengan cara yang diakui akademik:
memahami sumber, menulis ulang pada tataran ide, menyintesis beberapa sumber,
dan memberi atribusi. Alat ukurnya `scripts/overlap_check.py` — menghitung
tumpang-tindih n-gram draf terhadap sumber yang **kamu** sediakan.

Batas jujur: skor Turnitin hanya bisa diukur oleh Turnitin (korpusnya tidak
publik). Yang bisa dijamin alat ini: draf tidak lagi menyalin rangkaian kata
dari sumber yang diberikan, dan setiap ide beratribusi. Jangan pernah menjanjikan
"pasti lolos Turnitin".

---

## 0. Tiga hal yang sering dicampur

| Istilah | Arti | Cara turun yang sah |
|---|---|---|
| **Similarity** | % teks yang cocok dengan korpus Turnitin | parafrase ide + sintesis + sitasi |
| **Plagiarisme** | memakai ide/kata orang tanpa atribusi | sitasi + kutipan benar |
| **AI writing** | deteksi tulisan mesin | tulis ulang dengan suara sendiri |

Similarity tinggi belum tentu plagiat (daftar pustaka, istilah baku, kutipan),
dan similarity rendah belum tentu bebas plagiat (parafrase tanpa sitasi tetap
plagiat). Jangan kejar angka; kejar atribusi dan penulisan ulang yang benar.

Fakta teknis yang perlu diketahui:

- Turnitin mencocokkan **rangkaian kata** dengan indeks (web, jurnal, tugas
  mahasiswa sebelumnya, repositori). Metadata file, author, thumbnail, dan
  properti DOCX **tidak** ikut dihitung.
- Persentase = proporsi teks yang cocok, bukan vonis. Pengecualian kutipan,
  daftar pustaka, dan ambang "small match" **diatur dosen**, bukan bawaan.
- Detektor AI Turnitin adalah modul terpisah dari similarity.
- Mengganti sinonim, mengacak urutan kata, menyisipkan karakter tak terlihat,
  atau menaruh teks sebagai gambar **bukan** solusi: tetap terdeteksi sebagai
  patchwriting, merusak dokumen, dan melanggar aturan akademik.

---

## 1. Diagnosa sebelum menulis ulang

Jalankan dulu pengukur pada draf yang ada (langkah 4). Lalu klasifikasikan
temuan:

| Gejala | Diagnosis | Tindakan |
|---|---|---|
| Potongan panjang (>15 kata) identik | copy-paste | tulis ulang dari catatan ide, bukan dari teks |
| Kalimat sama dengan 1–2 sinonim ditukar | patchwriting | ubah struktur klausa, bukan kata |
| Kutipan benar tetapi tanpa tanda kutip | kutipan tak beratribusi | beri tanda kutip + halaman + sitasi |
| Banyak kalimat pendek mirip tersebar | mozaik sumber | sintesis 2–3 sumber per paragraf |
| Bab metode mirip tugas sendiri | self-plagiarism | sitasi karya sendiri + tulis ulang konteks |

---

## 2. SOP menulis dari sumber (inti penurunan)

```
1. Pilih sumber relevan (prioritaskan jurnal & buku, bukan blog)
2. Baca sampai paham argumennya, bukan kalimatnya
3. Catat ide di catatan: 1 kalimat sendiri + nomor halaman
4. TUTUP sumbernya
5. Tulis paragraf dari catatan (tanpa melihat teks sumber)
6. Bandingkan dengan sumber: ada frasa yang sama?
7. Beri sitasi untuk setiap ide yang bukan milikmu
```

Kunci langkah 4: selama teks sumber terlihat, otak menyalin struktur permukaan.
Menulis dari catatan memaksa rekode ide.

Catatan yang baik: `Sumber A (hlm. 12): SI = orang + teknologi + prosedur;
penekanan pada aliran informasi untuk keputusan.` Bukan: menyalin kalimat sumber.

---

## 3. Teknik parafrase yang benar (tataran ide)

Yang **boleh dan efektif**:

1. **Ubah struktur klausa** — dua klausa koordinat → satu klausa sebab-akibat;
   kalimat pasif → aktif; nomina aksi ("pengumpulan data") → verba
   ("mengumpulkan data").
2. **Abstraksi/konkretisasi** — sumber menyebut contoh spesifik, kamu angkat
   prinsipnya; sumber menyebut prinsip, kamu beri contoh sendiri.
3. **Sudut pandang** — sumber menjelaskan definisi, kamu menjelaskan
   konsekuensinya bagi kasusmu.
4. **Sintesis** — gabungkan 2–3 sumber: "A menekankan X (A, 2021), B
   menambahkan Y (B, 2022), namun keduanya belum membahas Z."
5. **Kontribusi** — tutup paragraf dengan posisi/implikasi kamu sendiri.

Yang **dilarang** (dan sia-sia):

- tukar sinonim kata per kata;
- acak urutan frasa;
- tambah/hapus kata sambung agar tidak persis;
- terjemahan mesin bolak-balik;
- karakter mirip/homoglyph, spasi tak terlihat, teks putih;
- menaruh teks sebagai gambar;
- menyembunyikan teks di properti dokumen.

Semua itu meninggalkan jejak stilistika atau merusak karya. Bukan strategi.

### Contoh

Sumber: "Sistem informasi adalah kombinasi dari teknologi informasi dan aktivitas
orang yang menggunakan teknologi tersebut untuk mendukung operasi dan manajemen."

Patchwriting (buruk, masih ~45% n-gram sama):
> Sistem informasi merupakan kombinasi teknologi informasi dan aktivitas orang
> yang memakai teknologi tersebut guna mendukung operasi serta manajemen.

Parafrase ide (benar):
> Organisasi membutuhkan aliran data yang tertata agar keputusan harian tidak
> bergantung pada tebakan. Teknologi hanya satu unsur; manusia dan prosedur
> menentukan apakah data itu benar-benar dipakai (Sumber, tahun).

Perhatikan: tidak ada frasa sumber yang dipertahankan, maknanya sama, atribusi
tetap ada.

---

## 4. Ukur dengan alat

```bash
python3 scripts/overlap_check.py \
  --draft bab2.docx \
  --source "jurnal-a.txt" --source "buku-b.md" \
  --ngram 6 --max-coverage 25
```

Keluaran: coverage %, match terpanjang, rincian per sumber, potongan yang terlalu
dekat, dan kalimat berisiko patchwriting. Kode keluar `1` bila melewati ambang.

Pilihan penting:

| Opsi | Guna |
|---|---|
| `--ignore-quotes` | cek parafrase saja; kutipan langsung tidak dihitung |
| `--ngram 8` | lebih ketat untuk skripsi (frasa 8 kata) |
| `--max-coverage 15` | target internal yang lebih aman |
| `--json` | untuk otomasi/CI atau laporan |

Panduan internal (bukan angka Turnitin): coverage terhadap satu sumber
**< 15%** dan match terpanjang **< 12 kata**. Di atas itu, parafrase ulang
bagian yang ditandai, jangan seluruh dokumen.

`.docx` dibaca langsung (butuh `python-docx`); `.txt`, `.md`, `.html` juga.
PDF: konversi dulu (`pdftotext -layout file.pdf file.txt`).

---

## 5. Kutipan langsung

- Pakai hanya bila kata aslinya penting: definisi hukum, istilah baku, kalimat
  yang akan dianalisis.
- ≤ 40 kata: tanda kutip di dalam paragraf + `(Nama, tahun, hlm. X)`.
- > 40 kata: blok terpisah, indentasi, tanpa tanda kutip, tetap + halaman.
- Anggaran: **< 10%** total teks (idealnya < 5%). Kutipan yang menumpuk tetap
  menaikkan similarity meski benar.
- Jangan mengutip untuk menghindari menulis. Kutipan adalah bukti, bukan isi.

---

## 6. Sitasi (APA 7 ringkas)

| Situasi | Dalam teks | Daftar pustaka |
|---|---|---|
| Parafrase 1 penulis | (Sutrisno, 2021) | Sutrisno. (2021). *Judul*. Penerbit. |
| Dua penulis | (Sutrisno & Rahayu, 2021) | keduanya dicantumkan |
| Tiga+ penulis | (Sutrisno dkk., 2021) | semua sampai 20 penulis |
| Kutipan langsung | (Sutrisno, 2021, hlm. 45) | sama |
| Tanpa tahun | (Sutrisno, t.t.) | (t.t.) |
| Institusi | (BPS, 2023) | Badan Pusat Statistik. (2023). |
| Sitasi sekunder | (dalam Sutrisno, 2021) | hanya sumber yang dibaca |

Aturan praktis:

- Satu klaim → satu sitasi. Kalimat berisi tiga klaim dari tiga sumber →
  tiga sitasi.
- Parafrase tetap disitasi. Mengubah kata tidak menghapus kewajiban atribusi.
- Nama penulis bisa jadi subjek kalimat: "Sutrisno (2021) menunjukkan bahwa…"
- Konsisten satu gaya di seluruh dokumen; urut abjad di daftar pustaka.
- Data diri/responden penelitianmu sendiri tidak disitasi, tetapi prosedurnya
  dijelaskan.

Jika kampus memakai gaya lain (IEEE, MLA, Chicago), ganti formatnya tetapi
pertahankan prinsip: penulis, tahun, halaman untuk kutipan.

---

## 7. Sintesis: cara menurunkan kemiripan paling tahan lama

Satu paragraf = beberapa sumber + posisi kamu:

```
[Sumber A] menegaskan … (A, 2021).
[Sumber B] menambahkan … (B, 2022).
Namun keduanya belum menjelaskan … .
Oleh karena itu, penelitian ini … .
```

Efeknya: tidak ada satu sumber yang mendominasi, kalimat menjadi milikmu, dan
similarity turun karena struktur serta diksi bukan jiplakan satu teks.

---

## 8. Checklist sebelum submit

- [ ] Setiap ide orang lain punya sitasi (termasuk yang sudah diparafrase).
- [ ] Kutipan langsung bertanda kutip + halaman; total < 10%.
- [ ] Tidak ada potongan > 12 kata yang identik dengan sumber.
- [ ] `overlap_check.py` < 15% per sumber; kalimat bertanda sudah ditulis ulang.
- [ ] Daftar pustaka lengkap, konsisten, semua disitasi, semua sitasi ada di daftar.
- [ ] Tidak ada trik karakter/gambar/teks tersembunyi.
- [ ] Bab metode menjelaskan prosedur sendiri, bukan menyalin prosedur paper lain.
- [ ] Karya sendiri sebelumnya disitasi bila dipakai ulang.
- [ ] Baca ulang dengan suara — kalimat yang janggal adalah sisa parafrase mekanis.
- [ ] Simpan versi draf + sumber untuk bukti proses.

---

## 9. Batas dan pelaporan

- Alat ukur membandingkan **hanya** file yang kamu berikan. Turnitin memakai
  korpus jauh lebih besar; skor akhir bisa berbeda (naik atau turun).
- Tidak ada cara sah untuk "menjamin" skor tertentu. Yang bisa dilaporkan:
  coverage internal, jumlah kutipan, kelengkapan sitasi.
- Jika dosen meminta revisi similarity, minta **similarity report** mereka:
  lihat sumber mana yang cocok, lalu parafrase/sitasi bagian itu saja.
- Jangan pernah menyerahkan hasil yang mengklaim "bebas Turnitin".

---

## 10. Rujukan

- Howard, R. M. (1993). A plagiarism pentimento. *Journal of Teaching Writing*.
- Keck, C. (2006). The use of paraphrase in summary writing. *Journal of Second
  Language Writing*.
- American Psychological Association. (2020). *Publication Manual* (7th ed.).
- Turnitin. (n.d.). *Interpreting the similarity score* — panduan resmi bahwa
  skor adalah proporsi teks cocok, bukan vonis plagiarisme.
- Bretag, T., & Mahmud, S. (2009). Self-plagiarism or appropriate textual
  re-use? *Journal of Academic Ethics*.
