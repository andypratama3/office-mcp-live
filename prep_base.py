#!/usr/bin/env python3
"""Siapkan dokumen kerja: salin file asli, pertahankan sampul + styles/header/footer,
hapus seluruh isi lama, lalu tambahkan style khusus skripsi."""
import copy
import os
import re
import shutil
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SRC = "/Users/andypratama3/Development/Key/skripsi_key/skripsi/PROPOSAL_PROYEK_AKHIR_SYARIFAH_KHAILA_NAJWA _HUMAIRA.docx"
DST = "/Users/andypratama3/Development/Key/skripsi_key/skripsi_fix/PROPOSAL_PROYEK_AKHIR_SYARIFAH_KHAILA_NAJWA_HUMAIRA.docx"


def add_style(doc, name, *, base="Normal", size=12, bold=False, italic=False,
              align=None, first_line=None, left=None, hanging=None,
              before=0, after=0, line=1.5, keep_next=False, keep_lines=False,
              outline=None, tab_right=None, color=None):
    styles = doc.styles
    if name in [s.name for s in styles]:
        st = styles[name]
    else:
        st = styles.add_style(name, 1)  # WD_STYLE_TYPE.PARAGRAPH
    st.base_style = styles[base]
    f = st.font
    f.name = "Times New Roman"
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    if color:
        f.color.rgb = RGBColor.from_string(color)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Times New Roman")
    pf = st.paragraph_format
    if align is not None:
        pf.alignment = align
    if first_line is not None:
        pf.first_line_indent = Cm(first_line)
    if left is not None:
        pf.left_indent = Cm(left)
    if hanging is not None:
        pf.first_line_indent = Cm(-hanging)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    if tab_right is not None:
        pf.tab_stops.add_tab_stop(Cm(tab_right), 2, 1)  # RIGHT, DOTS
    ppr = st.element.get_or_add_pPr()
    if outline is not None:
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), str(outline))
        ppr.append(ol)
    # hilangkan numbering warisan
    for num in ppr.findall(qn("w:numPr")):
        ppr.remove(num)
    return st


def main():
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    shutil.copy2(SRC, DST)
    doc = Document(DST)
    body = doc.element.body

    # --- temukan paragraf terakhir sampul (mengandung sectPr) ---
    kids = list(body)
    cut = None
    for i, el in enumerate(kids):
        if el.tag == qn("w:p"):
            if el.find(".//" + qn("w:sectPr")) is not None:
                txt = "".join(t.text or "" for t in el.iter(qn("w:t")))
                if "2025" in txt:
                    cut = i
                    break
    if cut is None:
        raise SystemExit("Paragraf penutup sampul tidak ditemukan")
    print("Sampul berakhir pada elemen ke-", cut)

    # simpan sectPr untuk dipakai ulang nanti
    cover_sect = kids[cut].find(".//" + qn("w:sectPr"))
    sect_xml = {}
    all_sect = body.findall(".//" + qn("w:sectPr"))
    print("jumlah sectPr:", len(all_sect))
    for i, s in enumerate(all_sect):
        sect_xml[i] = s

    # hapus semua setelah sampul, kecuali sectPr level-body
    body_sect = body.find(qn("w:sectPr"))
    for el in kids[cut + 1:]:
        if el is body_sect:
            continue
        body.remove(el)

    # daftar style khusus
    add_style(doc, "Judul Bab", size=14, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=18,
              keep_next=True, outline=0)
    add_style(doc, "Judul Subbab", size=12, bold=True,
              align=WD_ALIGN_PARAGRAPH.LEFT, before=12, after=6,
              keep_next=True, outline=1)
    add_style(doc, "Judul Subsubbab", size=12, bold=True,
              align=WD_ALIGN_PARAGRAPH.LEFT, before=10, after=4,
              keep_next=True, outline=2)
    add_style(doc, "Isi Tanpa Inden", align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=0)
    add_style(doc, "Butir", align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=0,
              left=1.27, hanging=0.75, before=0, after=0)
    add_style(doc, "Isi Tabel", size=11, align=WD_ALIGN_PARAGRAPH.LEFT,
              first_line=0, line=1.0, before=0, after=0)
    add_style(doc, "Judul Tabel", size=11, align=WD_ALIGN_PARAGRAPH.CENTER,
              first_line=0, line=1.0, before=12, after=6, keep_next=True)
    add_style(doc, "Judul Gambar", size=11, align=WD_ALIGN_PARAGRAPH.CENTER,
              first_line=0, line=1.0, before=6, after=12)
    add_style(doc, "Sumber", size=10, align=WD_ALIGN_PARAGRAPH.LEFT,
              first_line=0, line=1.0, before=2, after=12)
    add_style(doc, "Entri Daftar", align=WD_ALIGN_PARAGRAPH.LEFT, first_line=0,
              left=0, hanging=0, line=1.5, before=0, after=0, tab_right=15.5)
    add_style(doc, "Tengah", align=WD_ALIGN_PARAGRAPH.CENTER, first_line=0)
    add_style(doc, "Referensi", align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=0,
              left=1.27, hanging=1.27, before=0, after=6)
    add_style(doc, "Keterangan Gambar", size=10, align=WD_ALIGN_PARAGRAPH.CENTER,
              first_line=0, line=1.0, before=0, after=10)

    # paragraf kosong penutup agar penulisan live punya jangkar
    p = doc.add_paragraph("")
    p.style = doc.styles["Isi Tanpa Inden"]

    doc.save(DST)

    # simpan sectPr asli untuk dipakai tahap akhir
    import json
    data = {}
    for i, s in sect_xml.items():
        data[str(i)] = s.xml
    with open("/Users/andypratama3/Development/Key/skripsi_key/work/sectpr.json", "w") as fh:
        json.dump(data, fh)
    print("OK ->", DST)
    print("jumlah paragraf sekarang:", len(doc.paragraphs))


if __name__ == "__main__":
    main()
