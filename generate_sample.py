"""
Generate a correctly formatted sample BFW DTA ML Satz file.
All records use exact fixed-width field positions matching the parser.

Run: python generate_sample.py
Output: static/sample.ml
"""


def pad(s: str, width: int, char: str = " ") -> str:
    """Pad / truncate string to exact width."""
    s = str(s) if s else ""
    return (s + char * width)[:width]


def build_record(fields: list[tuple[int, str]], total_length: int, satztyp: str) -> str:
    """
    Build a fixed-width record.
    fields: list of (start_pos, value) tuples
    The end marker is the satztyp character at the last position.
    """
    buf = [" "] * total_length
    for start, value in fields:
        for i, ch in enumerate(value):
            if start + i < total_length:
                buf[start + i] = ch
    # End marker
    buf[total_length - 1] = satztyp
    return "".join(buf)


def a_satz(version="03.10", kundennr="BFWMUS001", kode="37",
           ordnung="0000000010000", kordnung="TESTDAT001") -> str:
    fields = [
        (0,  "A"),
        (1,  pad(version,  5)),
        (6,  pad(kundennr, 10)),
        (16, pad(kode,     2)),
        (18, pad(ordnung,  13)),
        (31, pad(kordnung, 20)),
    ]
    return build_record(fields, 128, "A")


def l_satz(version="03.10", kundennr="BFWMUS001", kode="37",
           ordnung="000000001    ", kordnung="PROP-EXT-001",
           brutto_netto="1",
           anfang="01012024", ende="31122024",
           strasse="Musterstraße 12", land="DEU",
           plz="10117", ort="Berlin",
           liegnr="0000000001", bezeichnung="Musteranlage Berlin Mitte") -> str:
    fields = [
        (0,   "L"),
        (1,   pad(version,     5)),
        (6,   pad(kundennr,   10)),
        (16,  pad(kode,        2)),
        (18,  pad(ordnung,    13)),   # BFW Ordnungsbegriff (9 prop + 4 unit blank)
        (31,  pad(kordnung,   20)),   # Kundlicher Ordnungsbegriff → external_id
        (51,  pad(brutto_netto, 1)),
        (52,  pad(anfang,      8)),   # Abrechnungszeitraum start (DDMMYYYY)
        (60,  pad(ende,        8)),   # Abrechnungszeitraum end
        (67,  pad(strasse,    35)),   # Street [67:102]
        (102, pad(land,        3)),   # Country code [102:105]
        (105, pad(plz,        10)),   # Postal code [105:115]
        (115, pad(ort,        35)),   # City [115:150]
        (150, pad(liegnr,     10)),   # Liegenschaftsnummer [150:160]
        (160, pad(bezeichnung, 40)),  # Grundstücksbezeichnung [160:200]
    ]
    return build_record(fields, 2048, "L")


def m_satz(version="03.10", kundennr="BFWMUS001", kode="37",
           prop_nr="000000001", unit_nr="0001",
           kordnung="UNIT-EXT-001",
           name1="Müller Hans", name2="", name3="", name4="",
           strasse="Musterstraße 12", land="DEU", plz="10117", ort="Berlin",
           einzug="01012022", auszug="", leerstand="0") -> str:
    ordnung = pad(prop_nr, 9) + pad(unit_nr, 4)  # 13 chars total
    fields = [
        (0,   "M"),
        (1,   pad(version,  5)),
        (6,   pad(kundennr, 10)),
        (16,  pad(kode,     2)),
        (18,  pad(ordnung,  13)),  # BFW Ordnungsbegriff
        (31,  pad(kordnung, 20)),  # Kundlicher Ordnungsbegriff → estate unit external id
        (51,  pad(name1,    35)),  # Nutzername1 [51:86]
        (86,  pad(name2,    35)),  # Nutzername2 [86:121]
        (121, pad(name3,    35)),  # Nutzername3 [121:156]
        (156, pad(name4,    35)),  # Nutzername4 [156:191]
        (191, pad(strasse,  35)),  # Nutzer Straße [191:226]
        (226, pad(land,      3)),  # Nutzer Land [226:229]
        (229, pad(plz,      10)),  # Nutzer PLZ [229:239]
        (239, pad(ort,      35)),  # Nutzer Stadt [239:274]
        (797, pad(einzug,    8)),  # Einzugsdatum [797:805]
        (805, pad(auszug,    8)),  # Auszugsdatum [805:813]
        (813, pad(leerstand, 1)),  # Leerstand [813:814]
    ]
    return build_record(fields, 2048, "M")


def b_satz(version="03.10", kundennr="BFWMUS001", kode="37",
           prop_nr="000000001",
           currency="EUR",
           start="01012023", ende="31122023",
           brennstoff="0001",  # 0001=GAS
           geraet_nr="12345678") -> str:
    ordnung = pad(prop_nr, 13)
    fields = [
        (0,   "B"),
        (1,   pad(version,   5)),
        (6,   pad(kundennr,  10)),
        (16,  pad(kode,      2)),
        (18,  pad(ordnung,   13)),   # BFW Ordnungsbegriff [18:31]
        (31,  pad(currency,  5)),    # Währungscode [31:36]
        (36,  pad(start,     8)),    # Abrechnungszeitraum Start [36:44]
        (44,  pad(ende,      8)),    # Abrechnungszeitraum Ende [44:52]
        (52,  pad(brennstoff, 4)),   # Brennstoffart [52:56]
        (281, pad(geraet_nr, 20)),   # Gerätenummer [281:301]
    ]
    return build_record(fields, 1024, "B")


def main():
    lines = []

    # A-Satz (header)
    lines.append(a_satz(
        ordnung="0000000010000",
        kordnung="IMMO-TESTDAT001"
    ))

    # ── Property 1: Berlin ───────────────────────────────────────────
    PROP1 = "000000001"
    lines.append(l_satz(
        ordnung=PROP1 + "    ",
        kordnung="PROP-BERLIN-001",
        strasse="Musterstraße 12",
        land="DEU",
        plz="10117",
        ort="Berlin",
        liegnr="P001",
        bezeichnung="Wohnanlage Musterstraße Berlin",
        anfang="01012024",
        ende="31122024",
    ))

    # Tenants for property 1
    lines.append(m_satz(
        prop_nr=PROP1, unit_nr="0001",
        kordnung="BERLIN-WE-0001",
        name1="Müller Hans", name2="und Sabine",
        einzug="01032020",
        auszug="",
        strasse="Musterstraße 12", land="DEU", plz="10117", ort="Berlin",
    ))
    lines.append(m_satz(
        prop_nr=PROP1, unit_nr="0002",
        kordnung="BERLIN-WE-0002",
        name1="Schmidt Anna",
        einzug="15012015",
        auszug="31122024",
        strasse="Musterstraße 12", land="DEU", plz="10117", ort="Berlin",
    ))

    # Device for property 1 (Gas heating system)
    lines.append(b_satz(
        prop_nr=PROP1,
        brennstoff="0001",  # GAS
        start="01012023",
        ende="31122023",
        geraet_nr="GASMETER-BLN-001",
    ))

    # ── Property 2: München ──────────────────────────────────────────
    PROP2 = "000000002"
    lines.append(l_satz(
        ordnung=PROP2 + "    ",
        kordnung="PROP-MUNICH-001",
        strasse="Kaiserplatz 7",
        land="DEU",
        plz="80331",
        ort="München",
        liegnr="P002",
        bezeichnung="Residenz am Kaiserplatz München",
        anfang="01012024",
        ende="31122024",
    ))

    # Tenants for property 2
    lines.append(m_satz(
        prop_nr=PROP2, unit_nr="0001",
        kordnung="MUC-WE-0001",
        name1="Weber Thomas",
        einzug="01032019",
        auszug="",
        strasse="Kaiserplatz 7", land="DEU", plz="80331", ort="München",
    ))
    lines.append(m_satz(
        prop_nr=PROP2, unit_nr="0002",
        kordnung="MUC-WE-0002",
        name1="Leerstand",
        einzug="01012024",
        auszug="31012024",
        leerstand="1",
        strasse="Kaiserplatz 7", land="DEU", plz="80331", ort="München",
    ))

    # Device for property 2 (Oil heating system)
    lines.append(b_satz(
        prop_nr=PROP2,
        brennstoff="0000",  # OEL
        start="01012023",
        ende="31122023",
        geraet_nr="OILMETER-MUC-001",
    ))

    output = "\r\n".join(lines) + "\r\n"

    import os
    out_path = os.path.join(os.path.dirname(__file__), "static", "sample.ml")
    with open(out_path, "w", encoding="latin-1") as f:
        f.write(output)

    print("Generated %d records -> %s" % (len(lines), out_path))
    print("Record lengths: A=%d, L=%d, M=%d, B=%d" % (
        len(lines[0]), len(lines[1]), len(lines[2]), len(lines[4])))


if __name__ == "__main__":
    main()
