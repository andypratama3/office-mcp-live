#!/usr/bin/env python3
"""smart_write.py - satu pintu tulis Word, LANGSUNG di dokumen yang terbuka.

Tidak ada mode file, tidak ada reload, tidak ada tutup-buka. Setiap perintah
dikirim lewat AppleScript ke jendela Word yang tampil, lalu selection dipindah
ke titik kerja sehingga layar ikut bergulir. Simpan dilakukan eksplisit dengan
perintah `save` (atau otomatis pada `append`/`heading`/`pagebreak` kecuali
diberi `--no-save`).

Pakai:
  smart_write.py status
  smart_write.py info
  smart_write.py read [--from 0] [--chars 3000]
  smart_write.py show "teks" [--n 1]
  smart_write.py goto "teks" [--n 2]
  smart_write.py style "teks" "Heading 2" [--n 1]
  smart_write.py delete "teks" [--n 1]
  smart_write.py replace "lama" "baru"
  smart_write.py append "paragraf 1" "paragraf 2" --after "jangkar" [--style Gaya] [--n 1]
  smart_write.py heading "Judul" --after "jangkar" [--level 2] [--n 1]
  smart_write.py pagebreak --after "jangkar" [--n 1]
  smart_write.py toc
  smart_write.py save
  smart_write.py pdf "/folder/dokumen/keluaran.pdf"

Sisipkan paragraf baru selalu butuh jangkar (`--after`): posisi jangkar dihitung
dari ¶ paragraf itu, bukan dari `count characters` dokumen, karena properti itu
ternyata hanya menghitung sampai akhir paragraf/section tempat selection berada.
`delete` menolak menghapus paragraf yang tidak dimulai dengan teks yang dicari,
kecuali diberi `--force`.

Opsi global:
  --file PATH   pastikan dokumen aktif memang PATH ini (kalau bukan: berhenti)
"""

from __future__ import annotations

import argparse
import os
import sys

import live_engine as le


def die(msg: str, code: int = 1) -> "NoReturn":  # type: ignore[name-defined]
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def check_file(path: str | None) -> None:
    if not path:
        return
    info = le.status()
    want = os.path.abspath(path)
    if os.path.abspath(info["path"]) != want:
        die(
            "Dokumen aktif bukan target.\n"
            f"  target : {want}\n"
            f"  aktif  : {info['path']}\n"
            "Tutup/aktifkan dokumen yang benar dulu; skrip ini tidak membuka "
            "dokumen lain.",
            2,
        )


def main() -> int:
    ap = argparse.ArgumentParser(add_help=True, description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", default=None, help="pastikan dokumen aktif adalah berkas ini")
    ap.add_argument("--n", type=int, default=1, help="kemunculan ke-N (default 1)")

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--file", default=argparse.SUPPRESS)
    common.add_argument("--n", type=int, default=argparse.SUPPRESS)

    sub = ap.add_subparsers(dest="cmd", required=True)

    def cmd(name: str):
        return sub.add_parser(name, parents=[common])

    cmd("status")
    cmd("info")
    p = cmd("read")
    p.add_argument("--from", dest="start", type=int, default=0)
    p.add_argument("--chars", type=int, default=3000)
    p = cmd("show")
    p.add_argument("text")
    p = cmd("goto")
    p.add_argument("text")
    p = cmd("style")
    p.add_argument("text")
    p.add_argument("name")
    p = cmd("delete")
    p.add_argument("text")
    p.add_argument("--force", action="store_true",
                   help="hapus walau paragraf tidak dimulai dengan teks itu")
    p = cmd("replace")
    p.add_argument("old")
    p.add_argument("new")
    p = cmd("append")
    p.add_argument("lines", nargs="+")
    p.add_argument("--after", required=True, help="teks jangkar paragraf")
    p.add_argument("--style", default="Normal")
    p.add_argument("--no-save", action="store_true")
    p = cmd("heading")
    p.add_argument("text")
    p.add_argument("--after", required=True, help="teks jangkar paragraf")
    p.add_argument("--level", type=int, default=1)
    p.add_argument("--no-save", action="store_true")
    p = cmd("pagebreak")
    p.add_argument("--after", required=True, help="teks jangkar paragraf")
    p.add_argument("--no-save", action="store_true")
    cmd("toc")
    cmd("save")
    p = cmd("pdf")
    p.add_argument("target")

    a = ap.parse_args()
    check_file(a.file)

    try:
        if a.cmd == "status":
            info = le.status()
            for k in ("name", "path", "saved", "pages", "page"):
                print(f"{k}: {info[k]}")
        elif a.cmd == "info":
            info = le.status()
            print(f"dokumen : {info['name']}")
            print(f"halaman : {info['pages']} (selection di {info['page']})")
            print(f"paragraf: {le.count_paragraphs()}")
            print(f"kata    : {le.count_words()}")
            print(f"gambar  : {le.count_inline_shapes()}")
            print(f"tersimpan: {info['saved']}")
        elif a.cmd == "read":
            out = le.run(
                "tell active document\n"
                f"  set r to create range start {a.start} end {a.start + a.chars}\n"
                "  return (content of r) as string\n"
                "end tell",
                timeout=180,
            )
            sys.stdout.write(out.replace("\r", "\n"))
        elif a.cmd == "show":
            print(le.paragraph_text(a.text, a.n))
        elif a.cmd == "goto":
            if not le.goto(a.text, a.n, select=True):
                die(f"tidak ditemukan: {a.text!r} (kemunculan {a.n})")
            print(f"halaman {le.page()}: {a.text}")
        elif a.cmd == "style":
            le.set_style(a.text, a.name, a.n)
            print(f"[gaya] {a.text!r} -> {a.name}")
        elif a.cmd == "delete":
            le.delete_paragraph(a.text, a.n, force=a.force)
            print(f"[hapus] paragraf {a.text!r}")
        elif a.cmd == "replace":
            if not le.replace_all(a.old, a.new):
                die(f"tidak ditemukan: {a.old!r}")
            print(f"[ganti] {a.old!r} -> {a.new!r}")
        elif a.cmd == "append":
            le.insert_paragraphs(a.after, [(a.style, t) for t in a.lines], a.n)
            if not a.no_save:
                le.save()
            print(f"[tambah] {len(a.lines)} paragraf ({a.style}) setelah {a.after!r}")
        elif a.cmd == "heading":
            style = f"Heading {a.level}"
            le.insert_paragraphs(a.after, [(style, a.text)], a.n)
            if not a.no_save:
                le.save()
            print(f"[tambah] {style}: {a.text}")
        elif a.cmd == "pagebreak":
            le.insert_page_break(a.after, a.n)
            if not a.no_save:
                le.save()
            print(f"[tambah] page break setelah {a.after!r}")
        elif a.cmd == "toc":
            print(f"[toc] {le.update_toc()} daftar diperbarui")
        elif a.cmd == "save":
            le.save()
            print("[simpan] ok")
        elif a.cmd == "pdf":
            print(le.export_pdf(a.target))
    except le.WordError as exc:
        die(f"Word: {exc}")
    except LookupError as exc:
        die(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
