#!/usr/bin/env python3
"""live_engine.py - edit dokumen Word yang SEDANG TERBUKA, tanpa reload.

Semua perubahan dikirim langsung ke objek Word yang tampil, jadi terlihat
seketika seperti mengetik sendiri. Tidak ada `save as`, tidak ada tutup-buka,
tidak ada penulisan file di belakang Word.

Fakta yang sudah diuji di Word for Mac (Office 365, macOS 26):

1. `open ... revert true` DIABAIKAN Word. Menulis file lewat python-docx lalu
   berharap Word memuat ulang TIDAK bekerja. Satu-satunya cara agar Word
   menampilkan versi baru adalah menutup lalu membuka dokumen - dan itu yang
   justru dihindari di sini. Karena itu seluruh perubahan harus lewat AppleScript.
2. `execute find` hanya mencari di "story" tempat selection berada (isi utama,
   header, sel tabel, text box). Kalau selection tertinggal di dalam tabel,
   pencarian teks di body akan gagal. Karena itu setiap pencarian dimulai dengan
   menyeleksi range 0..0 (awal dokumen).
3. `paragraph of selection` -> missing value. Yang benar:
   `paragraph 1 of text object of selection`.
4. `content of paragraph N` mengembalikan objek; pakai
   `content of text object of paragraph N as string`.
5. `insert at end of text object of selection` menyisipkan tepat di titik
   selection (bisa memecah paragraf). Untuk menambah paragraf baru, sisipkan
   `ASCII character 13 & teks` pada `end of content` paragraf jangkar.
6. `save as ... file format format PDF` hanya berhasil bila tujuannya berada di
   folder yang sama dengan dokumen (sandbox Word). Folder lain -> error -1708.
7. Loop `paragraph i` ribuan kali membuat Word 100% CPU. Selalu pakai find.

API tingkat tinggi dipakai oleh `smart_write.py`, CLI, dan MCP (`live_tools.py`).
"""

from __future__ import annotations

import os
import subprocess
import time

CR = "ASCII character 13"   # paragraph mark
FF = "ASCII character 12"   # manual page break


class WordError(RuntimeError):
    pass


# --------------------------------------------------------------------------- #
# AppleScript helpers
# --------------------------------------------------------------------------- #
def _q(text: str) -> str:
    """String literal AppleScript yang aman untuk tanda kutip/backslash."""
    out = []
    for ch in text:
        if ch == "\\":
            out.append("\\\\")
        elif ch == '"':
            out.append('\\"')
        elif ch == "\n":
            out.append('" & ' + CR + ' & "')
        elif ch == "\r":
            out.append('" & ' + CR + ' & "')
        elif ch == "\t":
            out.append('" & tab & "')
        else:
            out.append(ch)
    return '"' + "".join(out) + '"'


def _osa(script: str, timeout: int = 120) -> str:
    try:
        proc = subprocess.run(
            ["osascript", "-e", script], capture_output=True, text=True, timeout=timeout
        )
    except subprocess.TimeoutExpired as exc:
        raise WordError(
            f"Word tidak merespons dalam {timeout} detik (dialog modal?). "
            "Tutup dialog di layar lalu ulangi."
        ) from exc
    if proc.returncode != 0:
        raise WordError(proc.stderr.strip() or "osascript gagal")
    return proc.stdout.strip()


def run(body: str, timeout: int = 120) -> str:
    """Jalankan blok AppleScript di dalam `tell application "Microsoft Word"`."""
    return _osa(
        'tell application "Microsoft Word"\n'
        f"with timeout of {max(10, timeout - 5)} seconds\n{body}\nend timeout\nend tell",
        timeout=timeout,
    )


def _escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


# --------------------------------------------------------------------------- #
# Status & navigasi
# --------------------------------------------------------------------------- #
def status() -> dict:
    out = run(
        'set d to active document\n'
        'set out to (name of d) & tab & (posix full name of d) & tab '
        '& ((saved of d) as string) & tab '
        '& ((get selection information (selection) information type number of pages in document) as string) & tab '
        '& ((get selection information (selection) information type active end page number) as string)',
        timeout=60,
    )
    parts = out.split("\t")
    return {
        "name": parts[0],
        "path": parts[1],
        "saved": parts[2] == "true",
        "pages": int(parts[3]) if parts[3].isdigit() else None,
        "page": int(parts[4]) if parts[4].isdigit() else None,
    }


def goto(text: str, occurrence: int = 1, select: bool = False) -> bool:
    """Pindahkan selection (dan view) ke kemunculan ke-`occurrence` dari `text`.

    Pencarian selalu dimulai dari awal dokumen supaya tidak terjebak di dalam
    tabel/header tempat selection terakhir berada. `select=True` memilih teks
    yang ditemukan; default hanya menaruh titik sisip di awalnya.

    Fungsi ini sengaja hanya memindahkan selection. Pembacaan posisi ada di
    `paragraph_range()`; mencampur keduanya membuat satu panggilan AppleEvent
    membaca `missing value` untuk `text object of selection`.
    """
    if occurrence < 1:
        raise ValueError("occurrence mulai dari 1")
    out = run(
        "tell active document\n"
        "  set r0 to create range start 0 end 0\n"
        "  select r0\n"
        "end tell\n"
        'set flag to "no"\n'
        f"repeat with i from 1 to {occurrence}\n"
        "  try\n"
        '    set f to execute find (find object of selection) '
        f'find text {_q(text)} match forward true match case true\n'
        '    if f then set flag to "yes"\n'
        "  on error\n"
        "  end try\n"
        '  if flag is "no" then exit repeat\n'
        "end repeat\n"
        "if flag is \"no\" then return \"no\"\n"
        + (
            'if (count of characters of (text object of selection)) > 0 then\n'
            "  set s to start of content of (text object of selection)\n"
            "  set e to end of content of (text object of selection)\n"
            "  tell active document\n"
            "    set r to create range start s end e\n"
            "    select r\n"
            "  end tell\n"
            "end if\n"
            if select
            else ""
        )
        + 'return "yes"',
        timeout=240,
    )
    return out.strip() == "yes"


def page() -> int:
    """Nomor halaman tempat selection berada."""
    out = run(
        "return (get selection information (selection) "
        "information type active end page number) as string",
        timeout=60,
    )
    return int(out) if out.isdigit() else 0


def count_words() -> int:
    return int(run("return (count of words of active document) as string", timeout=180))


def count_paragraphs() -> int:
    return int(run("return (count of paragraphs of active document) as string", timeout=180))


def paragraph_text(text: str, occurrence: int = 1) -> str:
    """Isi paragraf yang memuat `text` (tanpa paragraph mark)."""
    start, end = paragraph_range(text, occurrence)
    return run(
        "tell active document\n"
        f"  set r to create range start {start} end {end}\n"
        "  return (content of r) as string\n"
        "end tell",
        timeout=120,
    )


def find_all(text: str, limit: int = 500) -> list[int]:
    """Nomor halaman semua kemunculan `text` (berhenti di `limit`)."""
    out = run(
        "tell active document\n"
        "  set r0 to create range start 0 end 0\n"
        "  select r0\n"
        "end tell\n"
        'set pages to ""\n'
        f"repeat {limit} times\n"
        "  try\n"
        '    set f to execute find (find object of selection) '
        f'find text {_q(text)} match forward true match case true\n'
        "  on error\n"
        '    set f to false\n'
        "  end try\n"
        "  if not f then exit repeat\n"
        '  set pages to pages & ((get selection information (selection) '
        'information type active end page number) as string) & ","\n'
        "end repeat\n"
        "return pages",
        timeout=600,
    )
    return [int(x) for x in out.split(",") if x.strip().isdigit()]


def paragraph_range(text: str, occurrence: int = 1) -> tuple[int, int]:
    """Posisi absolut (start, end-of-content) paragraf yang memuat `text`.

    `end-of-content` = posisi paragraph mark (¶) paragraf itu: titik sisip yang
    tepat untuk menambah paragraf baru sesudahnya.

    Dua jebakan Word yang ditangani di sini:
    - `start/end of content of (text object of selection)` hanya bisa dibaca
      DI LUAR blok `tell active document`, padahal `create range` harus DI DALAM
      blok itu. Karena itu pembacaan offset dan pemindaian ¶ dipisah.
    - posisi dihitung dengan memindai ¶ sebelum/sesudah teks, bukan lewat
      `paragraph ... of text object` yang memberi posisi salah untuk field.
    """
    out = run(
        "tell active document\n"
        "  set r0 to create range start 0 end 0\n"
        "  select r0\n"
        "end tell\n"
        'set flag to "no"\n'
        f"repeat with i from 1 to {occurrence}\n"
        "  try\n"
        '    set f to execute find (find object of selection) '
        f'find text {_q(text)} match forward true match case true\n'
        '    if f then set flag to "yes"\n'
        "  on error\n"
        "  end try\n"
        '  if flag is "no" then exit repeat\n'
        "end repeat\n"
        'if flag is "no" then return "NONE"\n'
        "try\n"
        "  set s0 to start of content of (text object of selection)\n"
        "  set e0 to end of content of (text object of selection)\n"
        "on error\n"
        "  return \"FIELD\"\n"
        "end try\n"
        "if s0 is missing value then return \"FIELD\"\n"
        "set el to e0\n"
        "repeat 60 times\n"
        "  tell active document\n"
        "    set r to create range start el end (el + 120)\n"
        "  end tell\n"
        "  set t to (content of r) as string\n"
        "  set AppleScript's text item delimiters to (ASCII character 13)\n"
        "  set parts to text items of t\n"
        "  set AppleScript's text item delimiters to \"\"\n"
        "  if (count of parts) > 1 then\n"
        "    set el to el + (length of item 1 of parts)\n"
        "    exit repeat\n"
        "  end if\n"
        "  set el to el + 120\n"
        "end repeat\n"
        "set sl to s0\n"
        "repeat 60 times\n"
        "  if sl <= 0 then exit repeat\n"
        "  set startAt to sl - 120\n"
        "  if startAt < 0 then set startAt to 0\n"
        "  tell active document\n"
        "    set r to create range start startAt end sl\n"
        "  end tell\n"
        "  set t to (content of r) as string\n"
        "  set AppleScript's text item delimiters to (ASCII character 13)\n"
        "  set parts to text items of t\n"
        "  set AppleScript's text item delimiters to \"\"\n"
        "  if (count of parts) > 1 then\n"
        "    set sl to sl - (length of last item of parts)\n"
        "    exit repeat\n"
        "  end if\n"
        "  set sl to startAt\n"
        "end repeat\n"
        "return (sl as string) & \",\" & (el as string)",
        timeout=240,
    )
    if out.strip() == "NONE":
        raise LookupError(f"teks tidak ditemukan: {text!r}")
    if out.strip() == "FIELD":
        raise WordError(
            "jangkar berada di dalam field (mis. daftar isi); "
            "pakai teks di isi dokumen sebagai jangkar"
        )
    start, end = out.split(",")
    return int(start), int(end)


def insert_paragraphs(anchor: str, items: list[tuple[str, str]], occurrence: int = 1) -> list[str]:
    """Sisipkan paragraf baru tepat setelah paragraf yang memuat `anchor`.

    `items` = [(nama gaya Word, teks), ...]. Gaya "" berarti mewarisi gaya
    paragraf jangkar.

    Metodenya murni offset: `end of content` paragraf jangkar adalah posisi ¶.
    Menyisipkan `¶ + teks` di posisi itu menghasilkan paragraf baru tepat
    sesudah jangkar, tanpa memecah kalimat dan tanpa menyentuh paragraf lain.
    Selection lalu dipindah ke paragraf baru supaya gaya bisa diterapkan dan
    view Word ikut bergulir ke titik kerja.
    """
    _, pos = paragraph_range(anchor, occurrence)
    if pos <= 0:
        raise WordError(f"posisi paragraf jangkar tidak valid: {pos}")

    done = []
    for style, text in items:
        run(
            "tell active document\n"
            f"  set r to create range start {pos} end {pos}\n"
            f"  insert text ({CR} & {_q(text)}) at r\n"
            "end tell"
        )
        new_start = pos + 1
        if style:
            run(
                "tell active document\n"
                f"  set r to create range start {new_start} end {new_start}\n"
                "  select r\n"
                f"  set style of (paragraph 1 of text object of selection) to {_q(style)}\n"
                "end tell"
            )
        pos = pos + 1 + len(text)
        done.append(text)
    return done


def set_style(text: str, style: str, occurrence: int = 1) -> None:
    if not goto(text, occurrence):
        raise LookupError(f"teks tidak ditemukan: {text!r}")
    run(f"set style of (paragraph 1 of text object of selection) to {_q(style)}")


def delete_paragraph(text: str, occurrence: int = 1, force: bool = False) -> None:
    """Hapus SATU paragraf utuh yang memuat `text`.

    Pengaman: bila paragraf itu tidak dimulai dengan `text`, penghapusan
    ditolak. Tanpa pengaman ini, teks yang berada di tengah paragraf akan
    menghapus seluruh paragraf beserta isinya.
    """
    content = paragraph_text(text, occurrence)
    if not force and not content.strip().startswith(text.strip()):
        raise WordError(
            "paragraf tidak dimulai dengan teks itu, jadi tidak dihapus.\n"
            f"  dicari : {text[:60]!r}\n"
            f"  isi    : {content.strip()[:120]!r}\n"
            "Pakai force=True (CLI: --force) hanya bila yakin satu paragraf "
            "itu memang harus hilang."
        )
    if not goto(text, occurrence):
        raise LookupError(f"paragraf tidak ditemukan: {text!r}")
    run("delete text object of (paragraph 1 of text object of selection)")


def set_page_break_before(text: str, value: bool = True, occurrence: int = 1) -> None:
    if not goto(text, occurrence):
        raise LookupError(f"teks tidak ditemukan: {text!r}")
    run(
        "set page break before of (paragraph 1 of text object of selection) to "
        + ("true" if value else "false")
    )


def strip_leading_page_break(text: str, occurrence: int = 1) -> bool:
    """Hapus manual page break (chr 12) di awal paragraf, kalau ada."""
    if not goto(text, occurrence):
        raise LookupError(f"teks tidak ditemukan: {text!r}")
    out = run(
        "tell active document\n"
        "  set para to paragraph 1 of text object of selection\n"
        "  set p0 to start of content of (text object of para)\n"
        "  set r to create range start p0 end (p0 + 1)\n"
        "  set c to (content of r) as string\n"
        f"  if c is {FF} then\n"
        "    set content of r to \"\"\n"
        "    return \"dihapus\"\n"
        "  end if\n"
        "end tell\n"
        'return "tidak ada"',
    )
    return out.strip() == "dihapus"


def delete_blank_before(text: str, occurrence: int = 1, keep: int = 0) -> int:
    """Hapus paragraf kosong tepat sebelum paragraf `text` (sisakan `keep`).

    Mengembalikan jumlah paragraph mark yang dihapus. Paragraf kosong = tidak
    berisi teks dan tidak memuat gambar/objek.
    """
    start, _ = paragraph_range(text, occurrence)
    if start <= 1:
        raise WordError("paragraf terlalu dekat dengan awal dokumen")
    out = run(
        "tell active document\n"
        f"  set r to create range start ({start} - 60) end {start}\n"
        "  set t to (content of r) as string\n"
        "  set n to 0\n"
        "  repeat with i from (length of t) to 1 by -1\n"
        "    if character i of t is (ASCII character 13) then\n"
        "      set n to n + 1\n"
        "    else\n"
        "      exit repeat\n"
        "    end if\n"
        "  end repeat\n"
        "  return n as string\n"
        "end tell"
    )
    trailing = int(out) if out.isdigit() else 0
    blanks = max(0, trailing - 1)          # satu ¶ adalah akhir paragraf teks
    to_delete = max(0, blanks - keep)
    if to_delete:
        run(
            "tell active document\n"
            f"  set r to create range start ({start} - {to_delete}) end {start}\n"
            "  set content of r to \"\"\n"
            "end tell"
        )
    return to_delete


def replace_all(find_text: str, replace_text: str, whole_document: bool = True) -> bool:
    """Ganti semua kemunculan secara live (Find & Replace milik Word)."""
    scope = (
        "count characters of text object"
        if whole_document
        else "count characters of text object of selection"
    )
    out = run(
        "tell active document\n"
        f"  set r to create range start 0 end ({scope})\n"
        "  tell find object of r\n"
        f"    set f to execute find find text {_q(find_text)} "
        f"replace with {_q(replace_text)} replace replace all match case true\n"
        "  end tell\n"
        "end tell\n"
        "return (f as string)",
        timeout=300,
    )
    return out.strip() == "true"


def insert_page_break(text: str, occurrence: int = 1, before: bool = False) -> None:
    """Sisipkan manual page break (chr 12) sebelum/sesudah paragraf `text`."""
    start, end = paragraph_range(text, occurrence)
    pos = start if before else end
    run(
        "tell active document\n"
        f"  set r to create range start {pos} end {pos}\n"
        f"  insert text {FF} at r\n"
        "end tell"
    )


def count_inline_shapes() -> int:
    return int(run("return (count of inline shapes of active document) as string", timeout=60))


def insert_image(image_path: str, anchor: str | None = None,
                 width_cm: float | None = None, own_paragraph: bool = True) -> None:
    """Sisipkan gambar pada paragraf jangkar (default: paragraf baru sesudahnya)."""
    if not os.path.exists(image_path):
        raise FileNotFoundError(image_path)
    pos = None
    if anchor:
        _, pos = paragraph_range(anchor)
        if own_paragraph:
            run(
                "tell active document\n"
                f"  set r to create range start {pos} end {pos}\n"
                f"  insert text {CR} at r\n"
                "end tell"
            )
            pos = pos + 1          # awal paragraf kosong yang baru
        run(
            "tell active document\n"
            f"  set r to create range start {pos} end {pos}\n"
            "  select r\n"
            "end tell"
        )
    abs_path = os.path.abspath(image_path)
    before = count_inline_shapes()
    _osa(
        'set the clipboard to (read (POSIX file "' + _escape(abs_path) + '") as «class PNGf»)',
        timeout=60,
    )
    run(
        "tell active document\n"
        "  paste object selection\n"
        "end tell",
        timeout=180,
    )
    after = count_inline_shapes()
    if after <= before:
        raise WordError(
            "gambar tidak tersisip (clipboard bukan PNG / posisi tidak valid). "
            f"inline shape {before} -> {after}"
        )
    if width_cm:
        run(
            "tell active document\n"
            "  set s to last inline shape\n"
            f"  set width of s to {width_cm} * 28.3464567\n"
            "end tell"
        )


# --------------------------------------------------------------------------- #
# Daftar isi, penyimpanan, ekspor
# --------------------------------------------------------------------------- #
def update_toc() -> int:
    """Perbarui seluruh field daftar isi/gambar/tabel secara live."""
    out = run(
        "tell active document\n"
        "  set n to count of tables of contents\n"
        "  repeat with i from 1 to n\n"
        "    update (table of contents i)\n"
        "  end repeat\n"
        "  return n as string\n"
        "end tell",
        timeout=600,
    )
    return int(out) if out.isdigit() else 0


def save() -> None:
    run("save active document", timeout=300)


def export_pdf(target: str) -> str:
    """Ekspor PDF ke folder dokumen (batas sandbox Word)."""
    info = status()
    doc_dir = os.path.dirname(info["path"])
    target = os.path.abspath(target)
    if os.path.dirname(target) != doc_dir:
        raise WordError(
            "Word hanya boleh menulis PDF di folder dokumennya sendiri "
            f"({doc_dir}). Tujuan sekarang: {os.path.dirname(target)}"
        )
    if os.path.exists(target):
        os.remove(target)
    run(
        f'save as active document file name {_q(target)} file format format PDF',
        timeout=900,
    )
    if not os.path.exists(target):
        raise WordError(f"PDF tidak terbentuk: {target}")
    return target


def open_document(path: str) -> None:
    """Buka dokumen hanya bila memang belum terbuka (tidak pernah menutup)."""
    info = run('return name of every document', timeout=30)
    names = [n.strip() for n in info.split(",") if n.strip()]
    if os.path.basename(path) not in names:
        run(f'open POSIX file {_q(os.path.abspath(path))}\nactivate', timeout=180)


if __name__ == "__main__":
    print(status())
