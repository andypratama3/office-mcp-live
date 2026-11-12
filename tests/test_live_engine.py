"""Uji live_engine TANPA menyentuh Word: semua skrip AppleScript harus valid.

Kesalahan paling mahal di jalur live adalah skrip AppleScript yang gagal
dikompilasi atau salah mengutip teks (tanda kutip, backslash, baris baru).
Uji ini menangkap script dari setiap fungsi publik lalu mengompilasinya dengan
`osacompile` (kompilasi saja, tidak dijalankan), jadi dokumen yang sedang
dibuka tidak tersentuh.

Jalankan:

    .venv-mcp/bin/python -m pytest tests/test_live_engine.py -q
    .venv-mcp/bin/python tests/test_live_engine.py
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import live_engine as le

SCRIPTS: list[str] = []
SHAPES = [2]


def fake_run(body: str, timeout: int = 120) -> str:
    script = (
        'tell application "Microsoft Word"\n'
        f"with timeout of {max(10, timeout - 5)} seconds\n{body}\nend timeout\nend tell"
    )
    SCRIPTS.append(script)
    if "save as" in body:
        target = body.split('file name "')[1].split('"')[0]
        Path(target).write_bytes(b"%PDF-1.4")
    if "posix full name" in body:
        return "uji.docx\t/tmp/uji.docx\ttrue\t10\t1"
    if "start of content of (text object of selection)" in body:
        return "100,115"
    if "count of inline shapes" in body:
        SHAPES[0] += 1
        return str(SHAPES[0])
    if "count of words" in body or "count of paragraphs" in body:
        return "42"
    if "active end page number" in body:
        return "12" if "repeat" not in body else "3,4,"
    if "set pages to" in body:
        return "3,4,"
    if "count of tables of contents" in body or "tables of contents" in body:
        return "3"
    if "create range start 100 end 115" in body:
        return "ZZUJI paragraf uji"
    if "set t to (content of r)" in body:
        return "\r\r"
    if 'return "yes"' in body:
        return "yes"
    return "ok"


def fake_osa(script: str, timeout: int = 120) -> str:
    SCRIPTS.append(script)
    if "save as" in script:
        target = script.split('file name "')[1].split('"')[0]
        Path(target).write_bytes(b"%PDF-1.4")
    return "ok"


def compile_script(script: str) -> None:
    with tempfile.NamedTemporaryFile(suffix=".scpt", delete=False) as fh:
        target = fh.name
    try:
        proc = subprocess.run(
            ["osacompile", "-e", script, "-o", target],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if proc.returncode != 0:
            raise AssertionError(
                "AppleScript tidak valid:\n" + proc.stderr.strip() + "\n\n" + script
            )
    finally:
        Path(target).unlink(missing_ok=True)


def exercise() -> None:
    asli_run, asli_osa = le.run, le._osa
    le.run = fake_run          # type: ignore[assignment]
    le._osa = fake_osa         # type: ignore[assignment]

    import base64

    png = Path(tempfile.mkdtemp()) / "uji.png"
    png.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
    )

    le.goto('Judul "kutip" \\ backslash', occurrence=2)
    le.find_all("Gambar 3.11")
    le.paragraph_range("DAFTAR PUSTAKA", 2)
    le.insert_paragraphs(
        "DAFTAR PUSTAKA",
        [("Heading 2", 'Teks "kutip" & garis\\miring'), ("Normal", "Baris\nbaru\ttab")],
        occurrence=2,
    )
    le.set_style("DAFTAR PUSTAKA", "Heading 1", 2)
    le.delete_paragraph("ZZUJI", 1)
    le.set_page_break_before("HASIL DAN PEMBAHASAN", True, 3)
    le.strip_leading_page_break("DAFTAR PUSTAKA", 2)
    le.delete_blank_before("HASIL DAN PEMBAHASAN", 3)
    le.replace_all('a "b" c', 'd \\ e')
    le.insert_page_break("BAB V PENUTUP", 1, before=True)
    le.insert_image(str(png), anchor="Gambar 2.1", width_cm=10)
    le.update_toc()
    le.save()
    le.export_pdf("/tmp/keluaran.pdf")
    le.status()
    le.count_inline_shapes()
    le.count_words()
    le.page()
    le.run, le._osa = asli_run, asli_osa


def test_semua_skrip_valid() -> None:
    SCRIPTS.clear()
    exercise()
    assert SCRIPTS, "tidak ada skrip yang dihasilkan"
    for script in SCRIPTS:
        compile_script(script)


def test_kutip_dan_baris_baru_di_escape() -> None:
    assert le._q('a"b') == '"a\\"b"'
    assert le._q("a\\b") == '"a\\\\b"'
    assert "ASCII character 13" in le._q("baris\nbaru")
    assert "tab" in le._q("kolom\ttab")


def test_hapus_menolak_paragraf_tengah() -> None:
    """delete_paragraph tidak boleh menghapus paragraf yang hanya *memuat* teks."""
    asli_text, asli_goto = le.paragraph_text, le.goto
    le.paragraph_text = lambda *a, **k: "Kalimat panjang yang memuat penanda di tengah."  # type: ignore[assignment]
    le.goto = lambda *a, **k: True  # type: ignore[assignment]
    try:
        try:
            le.delete_paragraph("penanda")
        except le.WordError as exc:
            assert "tidak dihapus" in str(exc)
        else:
            raise AssertionError("delete_paragraph seharusnya menolak")
    finally:
        le.paragraph_text, le.goto = asli_text, asli_goto  # type: ignore[assignment]


def test_word_read_only() -> None:
    """Uji nyata pada Word (hanya baca). Aktifkan dengan WORD_LIVE_TEST=1."""
    import os

    if os.environ.get("WORD_LIVE_TEST") != "1":
        print("lewati uji Word (set WORD_LIVE_TEST=1 untuk menjalankan)")
        return
    info = le.status()
    assert info["name"], "tidak ada dokumen aktif"
    assert le.goto("DAFTAR PUSTAKA", 1), "judul tidak ditemukan"
    assert le.page() and le.page() > 0
    start, end = le.paragraph_range("DAFTAR PUSTAKA", 2)
    assert 0 < start < end
    assert le.find_all("Gambar") != []
    print(f"Word OK: {info['name']} ({info['pages']} halaman)")


if __name__ == "__main__":
    test_kutip_dan_baris_baru_di_escape()
    test_semua_skrip_valid()
    test_hapus_menolak_paragraf_tengah()
    test_word_read_only()
    print(f"OK: {len(SCRIPTS)} skrip AppleScript valid")
