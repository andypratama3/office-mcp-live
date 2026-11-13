#!/usr/bin/env python3
"""Ukur tumpang-tindih n-gram draf terhadap sumber yang diberikan.

Alat bantu menulis: memeriksa seberapa dekat draf dengan teks sumber (indikasi
patchwriting), lalu menunjukkan bagian mana yang perlu diparafrase atau dikutip.
Ini BUKAN pengganti Turnitin: korpusnya hanya file yang kamu berikan sendiri.

Contoh:
    python3 overlap_check.py --draft bab2.docx --source jurnal1.pdf.txt \
        --source buku.md --ngram 6 --max-coverage 25

Kode keluar: 0 = di bawah ambang, 1 = melewati ambang, 2 = kesalahan pemakaian.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

TOKEN_RE = re.compile(r"[^\W_]+", re.UNICODE)
QUOTE_RE = re.compile(r"[\"“«']([^\"”»']{1,2000})[\"”»']")
SENTENCE_RE = re.compile(r"[^.!?\n]+[.!?]?", re.UNICODE)


def read_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        try:
            from docx import Document
        except ImportError as exc:  # pragma: no cover
            raise SystemExit(
                f"butuh python-docx untuk membaca {path} (pip install python-docx)"
            ) from exc
        doc = Document(str(path))
        chunks = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    chunks.append(cell.text)
        return "\n".join(chunks)
    raw = path.read_text(encoding="utf-8", errors="replace")
    if suffix in {".html", ".htm", ".xml", ".xhtml"}:
        raw = re.sub(r"<[^>]+>", " ", raw)
    return raw


def tokenize(text: str):
    return [(m.group(0).lower(), m.start(), m.end()) for m in TOKEN_RE.finditer(text)]


def quoted_spans(text: str):
    return [(m.start(1), m.end(1)) for m in QUOTE_RE.finditer(text)]


def in_spans(pos: int, spans) -> bool:
    return any(start <= pos < end for start, end in spans)


def ngrams(tokens, n: int):
    for i in range(len(tokens) - n + 1):
        yield tuple(tokens[i : i + n])


def analyze(draft: str, sources: dict[str, str], n: int, ignore_quotes: bool,
            flag_ratio: float, min_flag_words: int):
    draft_tokens = tokenize(draft)
    quote_boxes = quoted_spans(draft)

    eligible = []
    quoted_words = 0
    for token, start, _end in draft_tokens:
        is_quote = in_spans(start, quote_boxes)
        if is_quote:
            quoted_words += 1
        if ignore_quotes and is_quote:
            continue
        eligible.append((token, start))

    words = [t for t, _ in eligible]
    owner: dict[tuple, str] = {}
    per_source_index: dict[str, set] = {}
    for name, text in sources.items():
        src_tokens = [t for t, _, _ in tokenize(text)]
        index = set(ngrams(src_tokens, n))
        per_source_index[name] = index
        for gram in index:
            owner.setdefault(gram, name)

    covered_by = [None] * len(words)
    for i, gram in enumerate(ngrams(words, n)):
        name = owner.get(gram)
        if name is None:
            continue
        for j in range(i, i + n):
            if covered_by[j] is None:
                covered_by[j] = name

    covered_words = sum(1 for c in covered_by if c)
    total = len(words) or 1

    runs = []
    start = None
    for i, c in enumerate(covered_by + [None]):
        if c and start is None:
            start = i
        elif not c and start is not None:
            runs.append((start, i))
            start = None
    runs.sort(key=lambda r: r[1] - r[0], reverse=True)

    top_runs = []
    for a, b in runs[:10]:
        if b - a < min_flag_words:
            continue
        top_runs.append({
            "words": b - a,
            "source": covered_by[a],
            "text": " ".join(words[a:b]),
            "draft_token_range": [a, b],
        })

    per_source = {}
    for name, index in per_source_index.items():
        mark = [False] * len(words)
        for i, gram in enumerate(ngrams(words, n)):
            if gram in index:
                for j in range(i, i + n):
                    mark[j] = True
        per_source[name] = round(100.0 * sum(mark) / total, 1)

    sentences = []
    for m in SENTENCE_RE.finditer(draft):
        text = m.group(0).strip()
        if not text:
            continue
        spans = [(t, s) for t, s, _ in tokenize(text)]
        if len(spans) < min_flag_words:
            continue
        local = [t for t, _ in spans]
        hit = 0
        for i in range(max(1, len(local) - n + 1)):
            if tuple(local[i : i + n]) in owner:
                hit += n
        ratio = min(hit, len(local)) / len(local)
        if ratio >= flag_ratio:
            sentences.append({
                "ratio": round(100 * ratio, 1),
                "words": len(local),
                "text": text[:400],
            })
    sentences.sort(key=lambda s: s["ratio"], reverse=True)

    return {
        "draft_words_checked": len(words),
        "quoted_words_skipped": quoted_words if ignore_quotes else 0,
        "ngram": n,
        "coverage_pct": round(100.0 * covered_words / total, 1),
        "longest_match_words": (runs[0][1] - runs[0][0]) if runs else 0,
        "per_source_pct": per_source,
        "top_matches": top_runs,
        "patchwriting_sentences": sentences[:8],
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--draft", required=True, help="draf (.docx/.txt/.md/.html)")
    ap.add_argument("--source", action="append", default=[], required=True,
                    help="sumber pembanding; boleh diulang")
    ap.add_argument("--ngram", type=int, default=6, help="panjang n-gram (default 6)")
    ap.add_argument("--ignore-quotes", action="store_true",
                    help="kutipan langsung tidak dihitung (cek parafrase saja)")
    ap.add_argument("--flag-ratio", type=float, default=0.4,
                    help="ambang rasio kalimat mirip (default 0.4)")
    ap.add_argument("--min-words", type=int, default=8,
                    help="panjang minimum potongan yang dilaporkan (default 8)")
    ap.add_argument("--max-coverage", type=float, default=25.0,
                    help="ambang coverage persen untuk kode keluar 1 (default 25)")
    ap.add_argument("--json", action="store_true", help="keluaran JSON")
    args = ap.parse_args(argv)

    if args.ngram < 3:
        ap.error("--ngram minimal 3")

    draft_path = Path(args.draft)
    if not draft_path.exists():
        ap.error(f"draf tidak ditemukan: {draft_path}")
    sources = {}
    for item in args.source:
        p = Path(item)
        if not p.exists():
            ap.error(f"sumber tidak ditemukan: {p}")
        sources[p.name] = read_text(p)
    if not sources:
        ap.error("minimal satu --source")

    result = analyze(
        read_text(draft_path), sources, args.ngram, args.ignore_quotes,
        args.flag_ratio, args.min_words,
    )
    result["draft"] = draft_path.name
    result["sources"] = list(sources)
    result["max_coverage"] = args.max_coverage
    result["status"] = (
        "LULUS" if result["coverage_pct"] <= args.max_coverage else "PERIKSA ULANG"
    )

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Draf            : {result['draft']}")
        print(f"Sumber          : {', '.join(result['sources'])}")
        print(f"N-gram          : {args.ngram}")
        print(f"Coverage        : {result['coverage_pct']}% "
              f"(ambang {args.max_coverage}%) -> {result['status']}")
        print(f"Match terpanjang: {result['longest_match_words']} kata")
        if result["quoted_words_skipped"]:
            print(f"Kutipan dilewati: {result['quoted_words_skipped']} kata")
        print("Per sumber      : " + ", ".join(
            f"{k}={v}%" for k, v in result["per_source_pct"].items()))
        if result["top_matches"]:
            print("\nPotongan yang terlalu dekat:")
            for item in result["top_matches"]:
                print(f"  - {item['words']} kata [{item['source']}]: {item['text'][:180]}")
        if result["patchwriting_sentences"]:
            print("\nKalimat berisiko patchwriting:")
            for item in result["patchwriting_sentences"]:
                print(f"  - {item['ratio']}% ({item['words']} kata): {item['text'][:180]}")
        if result["status"] != "LULUS":
            print("\nTindakan: parafrase pada tataran ide, bukan tukar sinonim; "
                  "kutip langsung hanya bila perlu; tambahkan sitasi.")

    return 0 if result["status"] == "LULUS" else 1


if __name__ == "__main__":
    sys.exit(main())
