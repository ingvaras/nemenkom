#!/usr/bin/env python3
"""
Atliekų išvežimo grafiko generatorius (adresas paduodamas per Secrets).

Nuskaito UAB "Nemenčinės komunalininkas" puslapį, randa naujausius
pakuočių / stiklo / buitinių atliekų grafikus, ištraukia datas jūsų
adresui ir sugeneruoja .ics failą kalendoriaus prenumeratai.

Kiekvienas iš trijų dokumentų yra KITOKIO formato, todėl kiekvienam
yra atskira ištraukimo funkcija:

  * parse_pakuotes_pdf  - PDF, kur adreso bloke yra jūsų gatvė ir eilutė
                          "Pakuotė 10 d. 11 d. 9 d." (po vieną dieną 3 mėn.).
  * parse_stiklas_pdf   - PDF, kur gyvenvietės grupė "(išskyrus kai kurias gatves)"
                          turi VIENĄ dieną per ketvirtį; kuriam mėnesiui ji
                          priklauso, nustatoma pagal skaičiaus x-poziciją
                          lentelėje (Liepa / Rugpjūtis / Rugsėjis stulpeliai).
  * parse_buitines_xlsx - XLSX, kur datos surašytos tekstu ("8d., 22d.,")
                          po kelias per mėnesį, mėnesiai - stulpeliuose.

SVARBU: jei scriptas negali rasti jūsų gatvės duomenų patikimai, jis
NEPARAŠO klaidingų datų į kalendorių, tik įrašo įspėjimą į warnings.log.

Paleidimas: python3 generate_calendar.py
Reikalavimai: pip install requests pdfplumber openpyxl icalendar
"""

import re
import io
import os
import sys
import datetime
from collections import defaultdict
from urllib.parse import unquote

import requests
import pdfplumber
import openpyxl
from icalendar import Calendar, Event, Alarm

BASE_URL = "https://www.nemenkom.lt/buitiniu-ir-pakuociu-atlieku-surinkimo-grafikas/"

# Adresas NELAIKOMAS kode (repo viešas). Jis paduodamas per aplinkos kintamuosius,
# kurie GitHub Actions'e ateina iš užšifruotų Secrets (ADDR_VILLAGE, ADDR_STREET).
# Lokaliai testuojant: ADDR_VILLAGE=... ADDR_STREET=... python3 generate_calendar.py
VILLAGE = os.environ.get("ADDR_VILLAGE", "").strip()  # pvz. gyvenvietės fragmentas
STREET = os.environ.get("ADDR_STREET", "").strip()    # pvz. gatvės pavadinimas

# Atskiri raktai PDF failams (pakuotės/stiklas). Jei nenustatyti - naudojami
# tie patys kaip XLSX. Reikalingi ten, kur tas pats kaimas kartojasi kitoje
# seniūnijoje ir PDF blokai grupuojami kitaip nei XLSX eilutės.
PDF_VILLAGE = os.environ.get("ADDR_PDF_VILLAGE", "").strip() or VILLAGE
PDF_STREET = os.environ.get("ADDR_PDF_STREET", "").strip() or STREET

MONTH_NAMES = {
    "sausio": 1, "sausis": 1, "vasario": 2, "vasaris": 2, "kovo": 3, "kovas": 3,
    "balandžio": 4, "balandis": 4, "gegužės": 5, "gegužis": 5, "birželio": 6, "birželis": 6,
    "liepos": 7, "liepa": 7, "rugpjūčio": 8, "rugpjūtis": 8, "rugsėjo": 9, "rugsėjis": 9,
    "spalio": 10, "spalis": 10, "lapkričio": 11, "lapkritis": 11, "gruodžio": 12, "gruodis": 12,
}

# Mėnesių kilmininkas (genitive) pavadinimui "... rytoj vasario 18"
MONTH_GENITIVE = {
    1: "sausio", 2: "vasario", 3: "kovo", 4: "balandžio", 5: "gegužės", 6: "birželio",
    7: "liepos", 8: "rugpjūčio", 9: "rugsėjo", 10: "spalio", 11: "lapkričio", 12: "gruodžio",
}

# Kada dieną prieš išvežimą turi suveikti priminimas (vakare, kai reikia
# išnešti konteinerį). 18 = 18:00 vietos laiku priminimo dieną.
REMINDER_HOUR = 18

# WARNINGS - detalūs pranešimai (viskas rašoma į warnings.log).
# PROBLEMS - trumpi realių problemų žymekliai; jei netuščias, skriptas baigiasi
#            klaidos kodu (exit 1) -> GitHub Actions parausta -> ateina el. laiškas.
# NOTES    - informacinės pastabos (pvz. stiklo datos parinkimas pagal poziciją),
#            NElaikoma klaida.
WARNINGS = []
PROBLEMS = []
NOTES = []


def log_warning(msg):
    WARNINGS.append(msg)
    print(f"[ĮSPĖJIMAS] {msg}", file=sys.stderr)


def log_problem(category, msg):
    PROBLEMS.append(category)
    WARNINGS.append(f"PROBLEMA [{category}]: {msg}")
    print(f"[PROBLEMA] {category}: {msg}", file=sys.stderr)


def log_note(msg):
    NOTES.append(msg)
    print(f"[INFO] {msg}", file=sys.stderr)


# --------------------------------------------------------------------------
# Puslapio nuskaitymas
# --------------------------------------------------------------------------

def _newness(href):
    """Naujumo įvertis iš failo pavadinimo: (metai, vėliausias mėnuo).

    Pvz. „2026 m. liepos, rugpjūčio, rugsėjo mėn." -> (2026, 9).
    Jei puslapyje kabo keli to paties tipo failai (senas + naujas), imamas
    didžiausias įvertis. Nepavykus nustatyti - (0, 0), tad tinkamai pavadintas
    failas visada laimi prieš bevardį.
    """
    name = unquote(href).lower()
    ym = re.search(r'(20\d{2})', name)
    year = int(ym.group(1)) if ym else 0
    months = [MONTH_NAMES[w] for w in re.findall(r'[a-ząčęėįšųūž]+', name) if w in MONTH_NAMES]
    return (year, max(months) if months else 0)


def find_current_links():
    """Scrape the schedule page for the 3 current document links.

    Iš visų .pdf/.xlsx nuorodų kiekvienai rūšiai parenkama NAUJAUSIA (pagal
    pavadinime esančius metus/mėnesį), o ne pirma pasitaikiusi HTML'e.
    """
    resp = requests.get(BASE_URL, timeout=30)
    resp.raise_for_status()
    html = resp.text

    all_hrefs = re.findall(r'href="([^"]+\.(?:pdf|xlsx))"', html, re.IGNORECASE)

    def pick(keyword, exclude=None):
        cands = [
            h for h in all_hrefs
            if keyword.lower() in unquote(h).lower()
            and (exclude is None or exclude.lower() not in unquote(h).lower())
        ]
        # naujausias; jei įverčio nustatyti nepavyko, max lieka ties pirmu kandidatu
        return max(cands, key=_newness) if cands else None

    links = {}
    # (label, keyword, exclude): stiklo faile taip pat yra žodis „pakuočių",
    # todėl pakuotėms atmetame nuorodas, kuriose yra „Stiklo".
    for label, keyword, exclude in [
        ("pakuotes", "Pakuo", "Stiklo"),
        ("stiklas", "Stiklo", None),
        ("buitines", "Buitini", None),
    ]:
        href = pick(keyword, exclude)
        if href:
            links[label] = href
        else:
            log_warning(f"Nepavyko rasti '{keyword}' nuorodos puslapyje - patikrinkite rankiniu būdu: {BASE_URL}")
    return links


def download(url):
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    return resp.content


def extract_pdf_text(content):
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


# --------------------------------------------------------------------------
# 1) PAKUOTĖS (PDF): adreso bloke yra jūsų gatvė, eilutė "Pakuotė N d. N d. N d."
# --------------------------------------------------------------------------

def parse_table_pdf(content, kind_label="Atliekos"):
    """Bendras lentelinis PDF skaitytuvas (pakuotės ir stiklas).

    Kiekviena lentelės eilutė = vienas adresų blokas, stulpeliai = mėnesiai.
    Imame eilutę, kurios tekste yra ir PDF_VILLAGE (seniūnija), ir PDF_STREET
    (kaimas). Taip neįmanoma pagauti gretimo bloko, kaip būdavo ieškant
    artimiausio skaičiaus tekste.
    """
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                month_cols = {}
                for row in table:
                    found = {}
                    for ci, cell in enumerate(row or []):
                        if cell:
                            num = MONTH_NAMES.get(cell.replace("\n", " ").strip().lower())
                            if num:
                                found[ci] = num
                    if len(found) >= 2:
                        month_cols = found
                        break
                if not month_cols:
                    continue

                for row in table:
                    text = " ".join((c or "").replace("\n", " ") for c in row)
                    if PDF_VILLAGE not in text or PDF_STREET not in text:
                        continue
                    pairs = []
                    for ci, cell in enumerate(row):
                        if ci in month_cols and cell:
                            for d in re.findall(r"(\d{1,2})\s*d", cell):
                                pairs.append((month_cols[ci], int(d)))
                    if pairs:
                        return pairs

    log_warning(f"{kind_label}: nerasta lentelės eilutė su jūsų seniūnija ir kaimu.")
    return None


def parse_pakuotes_pdf(content, kind_label="Pakuotės"):
    """Grąžina [(month_num, day), ...] arba None."""
    text = extract_pdf_text(content)

    header = re.search(r'Atliekos\s+(\w+)\s+(\w+)\s+(\w+)', text)
    if not header:
        log_warning(f"{kind_label}: nerasta antraštė su mėnesiais - PDF formatas pasikeitė.")
        return None
    months = []
    for w in header.groups():
        num = MONTH_NAMES.get(w.lower())
        if not num:
            log_warning(f"{kind_label}: nežinomas mėnuo antraštėje: '{w}'.")
            return None
        months.append(num)

    # "Pakuotė N d. N d. N d." eilutė yra tame pačiame adreso bloke kaip
    # jūsų gatve. Grupės eina viena po kitos, todėl teisinga eilutė yra
    # ARČIAUSIA prie jūsų gatvės (skaičiuojant simbolių atstumą) - ją ir imame.
    # Taip veikia nepriklausomai nuo to, ar skaičiai stovi prieš ar po gatvės,
    # ir nesugriūva pasikeitus mėnesiams (mėnesiai imami iš antraštės).
    day_pat = re.compile(
        r'Pakuot[ėe]\s+(\d{1,2})\s*d\.\s*(\d{1,2})\s*d\.\s*(\d{1,2})\s*d\.'
    )
    all_matches = [(mm.start(), mm.groups()) for mm in day_pat.finditer(text)]
    if not all_matches:
        log_warning(f"{kind_label}: nerasta nė vienos 'Pakuotė N d. N d. N d.' eilutės - PDF formatas pasikeitė.")
        return None

    best = None
    for m in re.finditer(re.escape(PDF_STREET), text):
        # patikra, kad tai tikrai jūsų gyvenvietės blokas (o ne kita gyvenvietė)
        if PDF_VILLAGE not in text[max(0, m.start() - 1200): m.end() + 200]:
            continue
        street_pos = m.start()
        nearest = min(all_matches, key=lambda pm: abs(pm[0] - street_pos))
        best = nearest[1]
        break

    if best is None:
        log_warning(f"{kind_label}: nerastas jūsų adreso blokas su datomis - PDF formatas pasikeitė arba adresas kitoje grupėje.")
        return None
    return list(zip(months, (int(g) for g in best)))


# --------------------------------------------------------------------------
# 2) STIKLAS (PDF): gyvenvietės grupė "(išskyrus ...)" - VIENA diena per
#    ketvirtį; kuriam mėnesiui priklauso, nustatoma pagal skaičiaus x-poziciją.
# --------------------------------------------------------------------------

def _group_words_into_lines(words, tol=3.0):
    """Sugrupuoja pdfplumber žodžius į eilutes pagal 'top' koordinatę."""
    rows = defaultdict(list)
    for w in words:
        rows[round(w["top"] / tol)].append(w)
    lines = []
    for key in sorted(rows):
        ws = sorted(rows[key], key=lambda x: x["x0"])
        lines.append({
            "top": min(x["top"] for x in ws),
            "words": ws,
            "text": " ".join(x["text"] for x in ws),
        })
    return lines


def parse_stiklas_pdf(content, kind_label="Stiklas"):
    """Grąžina [(month_num, day), ...] arba None.

    Stiklo PDF yra lentelė: kairėje - seniūnijos/kaimų blokas, dešinėje - po
    stulpelį kiekvienam mėnesiui. Ieškome eilutės su savo kaimu (PDF_STREET),
    patikriname, kad artimiausia "... sen." antraštė virš jos yra mūsų
    (PDF_VILLAGE), ir skaičius paverčiame mėnesiais pagal x-poziciją.
    """
    pairs = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            month_x = {}
            for w in words:
                num = MONTH_NAMES.get(w["text"].strip().lower())
                if num:
                    month_x.setdefault(num, w["x0"])
            if len(month_x) < 2:
                continue

            lines = _group_words_into_lines(words)
            for li, line in enumerate(lines):
                text = " ".join(w["text"] for w in line["words"])
                if not re.search(r"(^|\s)" + re.escape(PDF_STREET), text):
                    continue

                # Artimiausia "sen." antraštė virš (arba pačioje) eilutėje.
                sen_line = None
                for back in range(li, max(-1, li - 8), -1):
                    t = " ".join(w["text"] for w in lines[back]["words"])
                    if " sen." in t:
                        sen_line = t
                        break
                if sen_line is None or PDF_VILLAGE not in sen_line:
                    continue

                for w in line["words"]:
                    if not re.fullmatch(r"\d{1,2}", w["text"].strip()):
                        continue
                    best = min(month_x.items(), key=lambda kv: abs(kv[1] - w["x0"]))
                    if abs(best[1] - w["x0"]) <= 40:
                        pairs.append((best[0], int(w["text"])))
                if pairs:
                    return pairs

    log_warning(f"{kind_label}: nerasta jūsų eilutė su kaimu - PDF formatas pasikeitė?")
    return None

def parse_buitines_xlsx(content, kind_label="Buitinės atliekos"):
    """Grąžina [(month_num, day), ...] arba None."""
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows())

        # Antraštės eilutė = ta, kurioje >=3 langeliai yra mėnesių pavadinimai.
        month_cols = {}      # column_index -> month_num
        header_idx = None
        for ri, row in enumerate(rows):
            found = {}
            for cell in row:
                if isinstance(cell.value, str):
                    num = MONTH_NAMES.get(cell.value.strip().lower())
                    if num:
                        found[cell.column] = num
            if len(found) >= 3:
                month_cols, header_idx = found, ri
                break
        if not month_cols:
            continue

        # Duomenų eilutė su mūsų adresu.
        for row in rows[header_idx + 1:]:
            row_text = " ".join(str(c.value) for c in row if c.value is not None)
            if VILLAGE in row_text and STREET in row_text:
                pairs = []
                for cell in row:
                    if cell.column in month_cols and cell.value is not None:
                        for d in re.findall(r'(\d{1,2})\s*d', str(cell.value)):
                            pairs.append((month_cols[cell.column], int(d)))
                if pairs:
                    return pairs
                log_warning(f"{kind_label}: rasta jūsų eilutė, bet nepavyko ištraukti dienų.")
                return None

        log_warning(f"{kind_label}: XLSX lape nerasta jūsų adreso eilutės.")
        return None

    log_warning(f"{kind_label}: XLSX faile nerasta antraštės su mėnesiais.")
    return None


# --------------------------------------------------------------------------
# Datos -> .ics
# --------------------------------------------------------------------------

def to_dates(month_day_pairs, year):
    """[(month_num, day), ...] -> [datetime.date, ...]."""
    dates = []
    for month_num, day in month_day_pairs:
        try:
            dates.append(datetime.date(year, month_num, int(day)))
        except ValueError as e:
            log_warning(f"Neteisinga data {year}-{month_num}-{day}: {e}")
    return dates


def make_ics(events):
    """
    events: list of (pickup_date, type_label)
      pickup_date - diena, kada realiai atvažiuoja šiukšliavežis.
      type_label  - pvz. "Buitinės atliekos".

    Kiekvienam įvykiui sukuriamas visos dienos įvykis DIENĄ PRIEŠ išvežimą
    (nes konteinerį reikia išnešti vakare), pavadinimu
    "Buitinės atliekos rytoj vasario 18", su priminimu (VALARM) tos pačios
    vakaro dienos 18:00 - jį telefone galima "snūzinti".

    DTSTAMP yra deterministinis (ne dabartinis laikas), kad tas pats grafikas
    generuotų BAITAIS IDENTIŠKĄ failą - tada GitHub Actions nedaro naujo
    commit'o, o telefonas neimportuoja iš naujo, kol grafikas realiai
    nepasikeitė.
    """
    cal = Calendar()
    cal.add("prodid", "-//atlieku grafikas//nemenkom.lt//")
    cal.add("version", "2.0")
    cal.add("x-wr-calname", "Atliekų grafikas")

    for pickup_date, type_label in sorted(events, key=lambda e: (e[0], e[1])):
        remind_date = pickup_date - datetime.timedelta(days=1)
        month = MONTH_GENITIVE[pickup_date.month]
        summary = f"{type_label} rytoj {month} {pickup_date.day}"

        ev = Event()
        ev.add("summary", summary)
        # Visos dienos įvykis priminimo dieną (diena prieš išvežimą).
        ev.add("dtstart", remind_date)
        ev.add("dtend", remind_date + datetime.timedelta(days=1))
        # Deterministinis DTSTAMP (priminimo dienos vidurnaktis).
        ev.add("dtstamp", datetime.datetime.combine(remind_date, datetime.time(0, 0)))
        ev["uid"] = f"{pickup_date.isoformat()}-{type_label.replace(' ', '')}@atliekos"

        # Priminimas priminimo dienos 18:00 (REMINDER_HOUR val. po vidurnakčio).
        alarm = Alarm()
        alarm.add("action", "DISPLAY")
        alarm.add("description", summary)
        alarm.add("trigger", datetime.timedelta(hours=REMINDER_HOUR))
        ev.add_component(alarm)

        cal.add_component(ev)
    return cal.to_ical()


def handle_category(links, key, type_label, parser, is_xlsx_ok=False):
    """Apdoroja vieną atliekų rūšį; grąžina įvykių sąrašą, žymi problemas."""
    if key not in links:
        log_problem(type_label, f"nuorodos į grafiką nerasta puslapyje.")
        return []
    url = requests.compat.urljoin(BASE_URL, links[key])
    content = download(url)
    if is_xlsx_ok and url.lower().endswith(".xlsx"):
        result = parser(content, kind_label=type_label)
    else:
        result = parser(content, kind_label=type_label)
    if not result:
        # detalią priežastį jau įrašė parseris; pažymime kaip realią problemą
        log_problem(type_label, "nepavyko ištraukti nė vienos datos (formatas pasikeitė?).")
        return []
    return [(d, type_label) for d in to_dates(result, datetime.date.today().year)]


def main():
    if not VILLAGE or not STREET:
        print(
            "KLAIDA: nenustatyti ADDR_VILLAGE ir/ar ADDR_STREET aplinkos kintamieji.\n"
            "GitHub Actions: pridėkite juos kaip Secrets (Settings -> Secrets -> Actions).\n"
            "Lokaliai: ADDR_VILLAGE=... ADDR_STREET=... python3 generate_calendar.py",
            file=sys.stderr,
        )
        sys.exit(2)

    links = find_current_links()
    all_events = []

    all_events += handle_category(links, "pakuotes", "Pakuočių atliekos", parse_table_pdf)
    all_events += handle_category(links, "stiklas", "Stiklo atliekos", parse_table_pdf)

    # Buitinės - XLSX (arba PDF, jei kada nors pasikeistų)
    if "buitines" not in links:
        log_problem("Buitinės atliekos", "nuorodos į grafiką nerasta puslapyje.")
    else:
        url = requests.compat.urljoin(BASE_URL, links["buitines"])
        content = download(url)
        if url.lower().endswith(".xlsx"):
            result = parse_buitines_xlsx(content)
        else:
            result = parse_pakuotes_pdf(content, kind_label="Buitinės atliekos")
        if result:
            all_events += [(d, "Buitinės atliekos") for d in to_dates(result, datetime.date.today().year)]
        else:
            log_problem("Buitinės atliekos", "nepavyko ištraukti nė vienos datos (formatas pasikeitė?).")

    if all_events:
        ics_bytes = make_ics(all_events)
        with open("atliekos.ics", "wb") as f:
            f.write(ics_bytes)
        print(f"OK: sugeneruota {len(all_events)} įvykių -> atliekos.ics")
    else:
        print("NIEKO nesugeneruota - žr. warnings.log", file=sys.stderr)

    # warnings.log perrašomas (ne pridedamas), kad be pakeitimų failas nesikeistų.
    with open("warnings.log", "w", encoding="utf-8") as f:
        if PROBLEMS:
            f.write(f"BŪSENA: PROBLEMA - nepavyko atnaujinti: {', '.join(sorted(set(PROBLEMS)))}\n\n")
        else:
            f.write("BŪSENA: OK - visi grafikai perskaityti sėkmingai.\n\n")
        if WARNINGS:
            f.write("Detalės:\n")
            for w in WARNINGS:
                f.write("  " + w + "\n")
        if NOTES:
            f.write("\nInformacija:\n")
            for n in NOTES:
                f.write("  " + n + "\n")

    # Jei buvo realių problemų - baigiame klaidos kodu, kad GitHub Actions
    # paraustų ir atsiųstų el. laišką. Failai jau įrašyti/commit'inami.
    if PROBLEMS:
        print(f"::error::Nepavyko atnaujinti: {', '.join(sorted(set(PROBLEMS)))}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
