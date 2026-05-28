"""
Generiert m/f/n-Triplets für alle Adjektive in der Wortliste.

Strategie:
1. Override-Map: manuell kuratierte Sonderfälle (Konsonant-Cluster,
   -ao/-eo, Komparative, unveränderliche Lehnwörter).
2. Heuristik für regelmäßige Fälle (mit/ohne beweglichem Vokal).
3. Output nach data/forms.json (merged mit existierenden Daten).

Bekannte Limitationen:
- Heuristik trifft etwa 90-95% der regelmäßigen Adjektive richtig.
- Konsonant-Wechsel (sladak→slatka) und Sonder-Stämme (zao→zla)
  sind in Override gepflegt, soweit erkannt.
- Korrekturen via Robin's "Fehler melden"-Knopf in der App.
"""
import json
import re
from pathlib import Path

from openpyxl import load_workbook


VOWELS = set("aeiouAEIOU")
CONSONANTS = set("bcčćdđfghjklmnpqrsštvwxzžBCČĆDĐFGHJKLMNPQRSŠTVWXZŽ")

# Manuell kuratierte Spezialfälle.
# Key: erstes Wort der SL-Spalte (vor "/").
# Value: [m, f, n]
OVERRIDE = {
    # ---- Unveränderlich (Lehnwörter, Farben, Sonderformen) ----
    "roze": ["roze", "roze", "roze"],

    # ---- -ao/-eo Adjektive (Stamm endet auf -al/-el, l→o im Maskulin) ----
    "topao": ["topao", "topla", "toplo"],
    "okrugao": ["okrugao", "okrugla", "okruglo"],
    "svetao": ["svetao", "svetla", "svetlo"],
    "kiseo": ["kiseo", "kisela", "kiselo"],
    "zreo": ["zreo", "zrela", "zrelo"],
    "beo": ["beo", "bela", "belo"],
    # -ao Partizipien als Adjektive (Stamm endet auf -l-, regulär dekliniert)
    "radoznao": ["radoznao", "radoznala", "radoznalo"],
    "preostao": ["preostao", "preostala", "preostalo"],
    "odrastao": ["odrastao", "odrasla", "odraslo"],

    # ---- Konsonant-Cluster mit -stan (t entfällt) ----
    "bolestan": ["bolestan", "bolesna", "bolesno"],
    "koristan": ["koristan", "korisna", "korisno"],
    "radostan": ["radostan", "radosna", "radosno"],
    "svestan": ["svestan", "svesna", "svesno"],

    # ---- -ao/-eo Sondergruppe: Stamm endet auf -l, l fällt zu -o im m ----
    "zao": ["zao", "zla", "zlo"],
    "ceo": ["ceo", "cela", "celo"],
    "veseo": ["veseo", "vesela", "veselo"],
    "debeo": ["debeo", "debela", "debelo"],
    "smeo": ["smeo", "smela", "smelo"],

    # ---- Konsonant-Wechsel bei beweglichem Vokal vor -k ----
    "sladak": ["sladak", "slatka", "slatko"],   # d → t
    "gladak": ["gladak", "glatka", "glatko"],   # d → t
    "nizak": ["nizak", "niska", "nisko"],       # z → s
    "redak": ["redak", "retka", "retko"],       # d → t
    "kratak": ["kratak", "kratka", "kratko"],   # regulär
    "težak": ["težak", "teška", "teško"],
    "blizak": ["blizak", "bliska", "blisko"],   # z → s
    "uzak": ["uzak", "uska", "usko"],           # z → s

    # ---- "bestimmte" Adjektive auf -i (deklinieren regulär a/o) ----
    "električni": ["električni", "električna", "električno"],
    "kućni": ["kućni", "kućna", "kućno"],
    "lični": ["lični", "lična", "lično"],
    "mali": ["mali", "mala", "malo"],
    "dupli": ["dupli", "dupla", "duplo"],
    "isti": ["isti", "ista", "isto"],
    "opšti": ["opšti", "opšta", "opšte"],   # -e statt -o im Neutrum
    "poslovni": ["poslovni", "poslovna", "poslovno"],
    "međunarodni": ["međunarodni", "međunarodna", "međunarodno"],
    "kulturni": ["kulturni", "kulturna", "kulturno"],
    "naučni": ["naučni", "naučna", "naučno"],

    # ---- Komparative auf -ji (deklinieren auf -a/-e, nicht -o) ----
    "raniji": ["raniji", "ranija", "ranije"],
    "bolji": ["bolji", "bolja", "bolje"],

    # ---- Aktive Partizipien auf -ći (deklinieren auf -a/-e) ----
    "važeći": ["važeći", "važeća", "važeće"],
    "iznenađujući": ["iznenađujući", "iznenađujuća", "iznenađujuće"],
    "odgovarajući": ["odgovarajući", "odgovarajuća", "odgovarajuće"],
    "budući": ["budući", "buduća", "buduće"],

    # ---- Weitere häufige Sonderformen ----
    "lep": ["lep", "lepa", "lepo"],
    "jak": ["jak", "jaka", "jako"],
    "suv": ["suv", "suva", "suvo"],
    "živ": ["živ", "živa", "živo"],

    # Possessivpronomen-artige Adjektive
    "njihov": ["njihov", "njihova", "njihovo"],
}


def has_movable_a(word: str) -> bool:
    """Heuristik: Wort hat beweglichen Vokal 'a' wenn es auf 'Konsonant + a + Endkonsonant' endet,
    UND vor der Konsonant-Gruppe noch ein anderer Vokal/Konsonant kommt (also nicht jednosilbig wie 'jak')."""
    if len(word) < 4:
        return False
    last = word[-1]
    if last not in "knrlmcčćš":  # häufige Endkonsonanten nach beweglichem a
        return False
    if word[-2] != "a":
        return False
    # Konsonant vor dem a?
    if word[-3] in VOWELS:
        return False
    return True


def generate_mfn(word: str) -> list[str]:
    """Generiere [m, f, n] heuristisch. Bei Override greift Override."""
    if word in OVERRIDE:
        return OVERRIDE[word]

    # -ao/-eo: zugrundeliegender Stamm endet auf -l, im Maskulin wurde l→o
    if word.endswith("ao") and len(word) >= 4:
        # Bekannte Spezialfälle stehen im Override. Default-Heuristik:
        # topao→topla pattern (l-Stamm, kein bw. Vokal)
        stem = word[:-2] + "l"
        return [word, stem + "a", stem + "o"]
    if word.endswith("eo") and len(word) >= 3:
        stem = word[:-2] + "el"
        return [word, stem + "a", stem + "o"]

    # Bestimmte Form auf -i (Partizipien, "definite" Adjektive)
    if word.endswith("i") and len(word) >= 3:
        stem = word[:-1]
        # Auf -ji oder -ći: f endet auf -a, n auf -e
        if word.endswith("ji") or word.endswith("ći"):
            return [word, stem + "a", stem + "e"]
        # Sonst: f endet auf -a, n auf -o
        return [word, stem + "a", stem + "o"]

    # Unveränderlich wenn Endung auf -e (vermutlich Lehnwort)
    if word.endswith("e") and len(word) <= 5:
        return [word, word, word]

    # Beweglicher Vokal
    if has_movable_a(word):
        # težak → teška / teško: das "a" vor dem Endkonsonant fällt weg
        stem = word[:-2] + word[-1]  # entferne "a"
        return [word, stem + "a", stem + "o"]

    # Standardfall: regelmäßig
    return [word, word + "a", word + "o"]


def main():
    repo = Path(__file__).resolve().parent.parent
    xlsx = repo / "source" / "wortliste.xlsx"
    forms_path = repo / "data" / "forms.json"

    wb = load_workbook(xlsx, read_only=True)
    ws = wb["Wortliste_V2"]
    rows = list(ws.iter_rows(values_only=True))

    # Existierende forms.json laden (preserve manual entries)
    existing = {}
    if forms_path.exists():
        with forms_path.open() as f:
            existing = json.load(f)

    generated = 0
    for r in rows[1:]:
        if len(r) < 11:
            continue
        nr, de, sl, sc, fil, wa, *_ = r
        if not (wa and "Adjektiv" in str(wa) and de and sl):
            continue
        key = f"{de}|{sl}"
        # Skip if already manually set
        if key in existing:
            continue
        first = str(sl).split(" / ")[0].strip()
        if not first:
            continue
        triplet = generate_mfn(first)
        existing[key] = {"mfn": triplet}
        generated += 1

    forms_path.parent.mkdir(parents=True, exist_ok=True)
    with forms_path.open("w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"Generated {generated} new Adjektiv-Triplets. Total entries in forms.json: {len(existing)}")


if __name__ == "__main__":
    main()
