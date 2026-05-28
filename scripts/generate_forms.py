"""
Generiert m/f/n-Triplets für Adjektive und Genus/Plural/Gen.Sg für Nomen.

Strategie:
- Heuristik für regelmäßige Fälle (Endungs-basiert).
- Override-Maps für Sonderfälle: bewegliche Vokale, Stammwechsel,
  feminine i-Deklination, maskuline Personennamen auf -a.
- Output nach data/forms.json (merged mit existierenden Daten).

Bekannte Limitationen:
- Heuristik trifft ca. 90-95% richtig. Restliche Fehler via Fehler-Knopf
  in der App melden und in der Override-Liste hier pflegen.
- Genitiv-Singular nur bei kuratierten beweglichen-Vokal-Fällen,
  weil heuristisch nicht zuverlässig erkennbar.
- Plural-Hinweis nur bei wirklich unregelmäßigen Stämmen, regulär
  Plurale werden weggelassen um Karten-Clutter zu vermeiden.
"""
import json
import re
from pathlib import Path

from openpyxl import load_workbook


# =====================================================================
# ADJEKTIV-OVERRIDES
# =====================================================================
ADJ_OVERRIDE = {
    "roze": ["roze", "roze", "roze"],
    "topao": ["topao", "topla", "toplo"],
    "okrugao": ["okrugao", "okrugla", "okruglo"],
    "svetao": ["svetao", "svetla", "svetlo"],
    "kiseo": ["kiseo", "kisela", "kiselo"],
    "zreo": ["zreo", "zrela", "zrelo"],
    "beo": ["beo", "bela", "belo"],
    "radoznao": ["radoznao", "radoznala", "radoznalo"],
    "preostao": ["preostao", "preostala", "preostalo"],
    "odrastao": ["odrastao", "odrasla", "odraslo"],
    "bolestan": ["bolestan", "bolesna", "bolesno"],
    "koristan": ["koristan", "korisna", "korisno"],
    "radostan": ["radostan", "radosna", "radosno"],
    "svestan": ["svestan", "svesna", "svesno"],
    "zao": ["zao", "zla", "zlo"],
    "ceo": ["ceo", "cela", "celo"],
    "veseo": ["veseo", "vesela", "veselo"],
    "debeo": ["debeo", "debela", "debelo"],
    "smeo": ["smeo", "smela", "smelo"],
    "sladak": ["sladak", "slatka", "slatko"],
    "gladak": ["gladak", "glatka", "glatko"],
    "nizak": ["nizak", "niska", "nisko"],
    "redak": ["redak", "retka", "retko"],
    "kratak": ["kratak", "kratka", "kratko"],
    "težak": ["težak", "teška", "teško"],
    "blizak": ["blizak", "bliska", "blisko"],
    "uzak": ["uzak", "uska", "usko"],
    "električni": ["električni", "električna", "električno"],
    "kućni": ["kućni", "kućna", "kućno"],
    "lični": ["lični", "lična", "lično"],
    "mali": ["mali", "mala", "malo"],
    "dupli": ["dupli", "dupla", "duplo"],
    "isti": ["isti", "ista", "isto"],
    "opšti": ["opšti", "opšta", "opšte"],
    "poslovni": ["poslovni", "poslovna", "poslovno"],
    "međunarodni": ["međunarodni", "međunarodna", "međunarodno"],
    "kulturni": ["kulturni", "kulturna", "kulturno"],
    "naučni": ["naučni", "naučna", "naučno"],
    "raniji": ["raniji", "ranija", "ranije"],
    "bolji": ["bolji", "bolja", "bolje"],
    "važeći": ["važeći", "važeća", "važeće"],
    "iznenađujući": ["iznenađujući", "iznenađujuća", "iznenađujuće"],
    "odgovarajući": ["odgovarajući", "odgovarajuća", "odgovarajuće"],
    "budući": ["budući", "buduća", "buduće"],
    "lep": ["lep", "lepa", "lepo"],
    "jak": ["jak", "jaka", "jako"],
    "suv": ["suv", "suva", "suvo"],
    "živ": ["živ", "živa", "živo"],
    "njihov": ["njihov", "njihova", "njihovo"],
}

VOWELS = set("aeiouAEIOU")


def adj_has_movable_a(word):
    if len(word) < 4:
        return False
    if word[-1] not in "knrlmcčćš":
        return False
    if word[-2] != "a":
        return False
    if word[-3] in VOWELS:
        return False
    return True


def gen_adj_mfn(word):
    if word in ADJ_OVERRIDE:
        return ADJ_OVERRIDE[word]
    if word.endswith("ao") and len(word) >= 4:
        stem = word[:-2] + "l"
        return [word, stem + "a", stem + "o"]
    if word.endswith("eo") and len(word) >= 3:
        stem = word[:-2] + "el"
        return [word, stem + "a", stem + "o"]
    if word.endswith("i") and len(word) >= 3:
        stem = word[:-1]
        if word.endswith("ji") or word.endswith("ći"):
            return [word, stem + "a", stem + "e"]
        return [word, stem + "a", stem + "o"]
    if word.endswith("e") and len(word) <= 5:
        return [word, word, word]
    if adj_has_movable_a(word):
        stem = word[:-2] + word[-1]
        return [word, stem + "a", stem + "o"]
    return [word, word + "a", word + "o"]


# =====================================================================
# NOMEN-OVERRIDES
# =====================================================================

# Maskulin endend auf -a (Verwandtschaft, Anrede, slawische Personenwörter)
MASC_ON_A = {
    "tata", "deda", "kolega", "sluga", "vojvoda", "starešina",
    "vladika", "gazda", "čika", "papa", "Sava",
}

# Maskulin endend auf -o (Lehnwörter und l-Stamm-Wechsel)
# l-Stamm: posao → posla, udeo → udela (im Maskulin l→o)
MASC_ON_O = {
    "auto": None,
    "metro": None,
    "posao": "posla",   # Gen.Sg
    "udeo": "udela",
}

# Pluralia tantum: nur Plural existent, Genus ist klassifizierend
NEUT_PL_TANTUM = {"vrata", "leđa", "kola"}
FEM_PL_TANTUM = {
    "naočare", "naočale",
    "pantalone", "farmerke",
    "makaze", "makazice",
    "merdevine", "stepenice",
    "novine", "studije",
    # Plural-Formen femininer -a Wörter (Robin's xlsx listet sie im Pl.)
    "rukavice", "papuče", "čarape", "patike", "cipele", "čizme",
}

# Feminin endend auf Konsonant (i-Deklination, Abstrakta auf -ost/-nost und Klassiker)
FEM_ON_CONS = {
    "noć", "stvar", "smrt", "reč", "ljubav", "vlast", "čast", "krv",
    "peć", "sol", "jesen", "moć", "pomoć", "vest", "mast", "kost",
    "narav", "savest", "starost", "mladost", "smelost", "vrednost",
    "stvarnost", "stranica",
    # Klassiker abseits -ost
    "stvar", "nit", "stvar", "želja", "smrt",
}

# Neutrum-Ausnahmen auf Konsonant (selten: -e/-ena Stämme)
# In Serbisch eher selten, lass ich aus.

# Pluralform-Wechsel (nur unregelmäßige eintragen)
NOUN_PLURAL = {
    "čovek": "ljudi",
    "dete": "deca",
    "brat": "braća",
    "gospodin": "gospoda",
    "oko": "oči",
    "uho": "uši",
    "uvo": "uši",
    "vlastelin": "vlastela",
    "knez": "kneževi",
    "vlas": "vlasi",
    "Srbin": "Srbi",
    "građanin": "građani",
    "seljanin": "seljani",
    "ostrvljanin": "ostrvljani",
    "katolik": "katolici",
    "neprijatelj": "neprijatelji",
    "prijatelj": "prijatelji",
}

# Genitiv-Singular bei beweglichem Vokal (kuratiert)
NOUN_GEN = {
    "otac": "oca",
    "pas": "psa",
    "san": "sna",
    "vetar": "vetra",
    "novac": "novca",
    "starac": "starca",
    "udarac": "udarca",
    "trenutak": "trenutka",
    "sastanak": "sastanka",
    "početak": "početka",
    "doručak": "doručka",
    "ručak": "ručka",
    "ulazak": "ulaska",
    "izlazak": "izlaska",
    "dolazak": "dolaska",
    "polazak": "polaska",
    "boraca": "borca",
    "gubitak": "gubitka",
    "dobitak": "dobitka",
    "spavanje": None,  # placeholder, hat keinen Gen-Hinweis nötig
    "lonac": "lonca",
    "lakat": "lakta",
    "nokat": "nokta",
    "palac": "palca",
    "venac": "venca",
    "kupac": "kupca",
    "stranac": "stranca",
    "borac": "borca",
    "vrabac": "vrapca",
    "vepar": "vepra",
    "popa": None,
    "đak": "đaka",  # nicht beweglich
    "izbor": "izbora",  # nicht beweglich
    "trgovac": "trgovca",
    "ljubavnik": "ljubavnika",  # nicht beweglich
    "muzika": None,
    "tovar": "tovara",  # nicht beweglich
    "razlog": "razloga",  # nicht beweglich
    "prozor": "prozora",  # nicht beweglich
}

# Bereinige NOUN_GEN: nur Einträge mit beweglichem Vokal behalten
NOUN_GEN = {k: v for k, v in NOUN_GEN.items() if v is not None and v != k + "a"}


def gen_noun(word):
    """Returns dict with {genus, plural?, gen?}"""
    out = {}
    # Genus
    if word in NEUT_PL_TANTUM:
        out["genus"] = "n"
    elif word in FEM_PL_TANTUM:
        out["genus"] = "f"
    elif word in MASC_ON_A:
        out["genus"] = "m"
    elif word in MASC_ON_O:
        out["genus"] = "m"
        gen_form = MASC_ON_O[word]
        if gen_form:
            out["gen"] = gen_form
    elif word in FEM_ON_CONS:
        out["genus"] = "f"
    elif word.endswith("a"):
        out["genus"] = "f"
    elif word.endswith("o") or word.endswith("e"):
        out["genus"] = "n"
    else:
        # Konsonant: m default, plus -ost/-nost ausnahme
        if word.endswith("ost") or word.endswith("nost"):
            out["genus"] = "f"
        else:
            out["genus"] = "m"

    # Plural unregelmäßig
    if word in NOUN_PLURAL:
        out["plural"] = NOUN_PLURAL[word]

    # Genitiv-Singular bei beweglichem Vokal
    if word in NOUN_GEN and "gen" not in out:
        out["gen"] = NOUN_GEN[word]

    return out


# =====================================================================
# MAIN
# =====================================================================
def main():
    repo = Path(__file__).resolve().parent.parent
    xlsx = repo / "source" / "wortliste.xlsx"
    forms_path = repo / "data" / "forms.json"

    wb = load_workbook(xlsx, read_only=True)
    ws = wb["Wortliste_V2"]
    rows = list(ws.iter_rows(values_only=True))

    existing = {}
    if forms_path.exists():
        with forms_path.open() as f:
            existing = json.load(f)

    adj_generated = 0
    noun_generated = 0
    for r in rows[1:]:
        if len(r) < 11:
            continue
        nr, de, sl, sc, fil, wa, *_ = r
        if not (wa and de and sl):
            continue
        key = f"{de}|{sl}"
        if key in existing:
            continue
        first = str(sl).split(" / ")[0].strip()
        if not first:
            continue

        if "Adjektiv" in str(wa):
            existing[key] = {"mfn": gen_adj_mfn(first)}
            adj_generated += 1
        elif "Nomen" in str(wa):
            existing[key] = gen_noun(first)
            noun_generated += 1

    forms_path.parent.mkdir(parents=True, exist_ok=True)
    with forms_path.open("w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"Adjektive: {adj_generated} neu generiert.")
    print(f"Nomen:     {noun_generated} neu generiert.")
    print(f"Total in forms.json: {len(existing)}")


if __name__ == "__main__":
    main()
