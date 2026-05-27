"""
Parst die drei Idem-Dalje-docx-Dateien zu data/corpus.json
Format: Liste von (SR-Satz, DE-Wort-für-Wort-Übersetzung)-Paaren pro Kurs.

Dient ausschliesslich als interne Referenz für die Beispielsatz-Generierung,
NICHT für die Karten in der App. Hilft, authentische Sätze und belegte Formen
aus dem Kursmaterial zu erkennen.

Birkenbihl-Format in den docx: SR-Satz steht als normaler Paragraph,
unmittelbar gefolgt vom DE-Satz in **fett** mit Wort-für-Wort-Übersetzung
und grammatischen Klammern wie (du-)gehst, (es-)bedeutet.
"""
import json
import re
import subprocess
import sys
from pathlib import Path


def extract_pairs(md_text: str) -> list[dict]:
    """Erwartet pandoc-md-Output: deutsche Übersetzungen sind in **fett**.
    Heuristik: serbischer Satz endet bevor ein **-block beginnt.
    """
    pairs = []
    # In pandoc-extract sind absätze durch leerzeilen getrennt
    blocks = [b.strip() for b in md_text.split("\n\n") if b.strip()]
    i = 0
    while i < len(blocks) - 1:
        current = blocks[i]
        next_b = blocks[i + 1]
        # Heuristik: serbischer Satz hat KEINE fett-marker am Start,
        # deutscher Birkenbihl-Satz beginnt mit **
        if not current.startswith("**") and next_b.startswith("**"):
            sr = clean_whitespace(current)
            de = clean_whitespace(next_b)
            # Filter Listenkopf-Überschriften aus (kurze Zeilen, häufig **xxx**)
            if len(sr) > 5 and len(de) > 5:
                pairs.append({"sr": sr, "de": de})
        i += 1
    return pairs


def clean_whitespace(s: str) -> str:
    """Mehrfach-Spaces zu einem reduzieren, ** entfernen, trim."""
    s = re.sub(r"\*+", "", s)
    s = re.sub(r"\s+", " ", s)
    s = s.strip(" -•")
    return s.strip()


def main(uploads_dir: Path, out_path: Path):
    corpus = {"courses": {}}
    for level in (1, 2, 3):
        docx_path = uploads_dir / f"Serbisch_-_Idemo_Dalje_{level}_-_Texte.docx"
        if not docx_path.exists():
            print(f"SKIP: {docx_path} not found")
            continue
        # Markdown via pandoc (kommt mit extract-text)
        md = subprocess.run(
            ["extract-text", str(docx_path)],
            capture_output=True, text=True, check=True,
        ).stdout
        pairs = extract_pairs(md)
        corpus["courses"][f"idemo_{level}"] = pairs
        print(f"Course {level}: {len(pairs)} Birkenbihl pairs")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(corpus, f, ensure_ascii=False, separators=(",", ":"))
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    uploads = root / "source"
    out = root / "data" / "corpus.json"
    main(uploads, out)
