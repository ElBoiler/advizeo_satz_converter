"""
BFW DTA ML Satz Parser
======================
Parses the German Heizkosten data exchange format (BFW DTA v03.10).
Fixed-width ASCII records, one per line.

Record types:
  A-Satz (128 B)  – File header / Ordnungsbegriffe
  L-Satz (2048 B) – Liegenschaft (property)
  M-Satz (2048 B) – Mieter (tenant)
  B-Satz (1024 B) – Bewegung / Brennstoff (fuel/device data)
  K-Satz (1024 B) – Kosten (costs) – parsed but not converted
  D-Satz (1024 B) – Abrechnungsergebnisse – parsed but not converted
"""

from __future__ import annotations
import re
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import List, Optional


# ---------------------------------------------------------------------------
# Fuel-type code → advizeo device type
# ---------------------------------------------------------------------------
BRENNSTOFF_TO_DEVICE_TYPE: dict[str, str] = {
    "0":    "OEL",   # Heizöl
    "0000": "OEL",
    "1":    "GAS",   # Erdgas
    "0001": "GAS",
    "2":    "WMZ",   # Fernwärme / District heating
    "0002": "WMZ",
    "3":    "PEL",   # Pellets
    "0003": "PEL",
    "4":    "WMZ",   # Holzhackschnitzel
    "0004": "WMZ",
    "5":    "WMZ",   # Kohle / Koks
    "0005": "WMZ",
    "6":    "WMZ",   # Strom für Wärmepumpe
    "0006": "WMZ",
    "7":    "WMZ",   # Biogas
    "0007": "WMZ",
    "9":    "WMZ",   # Sonstige
    "0009": "WMZ",
}


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def _parse_date(raw: str) -> str:
    """Convert BFW date string → DD.MM.YYYY, or '' if blank/zero."""
    s = raw.strip()
    digits = re.sub(r"\D", "", s)
    if not digits or set(digits) == {"0"}:
        return ""
    if len(digits) >= 8:
        dd, mm, yyyy = digits[:2], digits[2:4], digits[4:8]
        return f"{dd}.{mm}.{yyyy}"
    if len(digits) == 6:
        dd, mm, yy = digits[:2], digits[2:4], digits[4:6]
        year = 2000 + int(yy) if int(yy) < 50 else 1900 + int(yy)
        return f"{dd}.{mm}.{year}"
    return ""


def _str(raw: str, start: int, end: int) -> str:
    """Safe slice + strip from a padded record string."""
    return raw[start:end].strip() if len(raw) > start else ""


# ---------------------------------------------------------------------------
# Record dataclasses
# ---------------------------------------------------------------------------

@dataclass
class ASatz:
    """A-Satz – Ordnungsbegriffe / File header (128 bytes)."""
    raw: str
    version: str = field(init=False)
    bfw_kundennummer: str = field(init=False)
    betriebskode: str = field(init=False)
    bfw_ordnungsbegriff: str = field(init=False)
    kundlicher_ordnungsbegriff: str = field(init=False)

    def __post_init__(self):
        r = self.raw.ljust(128)
        self.version = _str(r, 1, 6)
        self.bfw_kundennummer = _str(r, 6, 16)
        self.betriebskode = _str(r, 16, 18)
        self.bfw_ordnungsbegriff = _str(r, 18, 31)
        self.kundlicher_ordnungsbegriff = _str(r, 31, 51)


@dataclass
class LSatz:
    """L-Satz – Liegenschaftssatz / Property record (2048 bytes)."""
    raw: str
    version: str = field(init=False)
    bfw_kundennummer: str = field(init=False)
    betriebskode: str = field(init=False)
    bfw_ordnungsbegriff: str = field(init=False)
    kundlicher_ordnungsbegriff: str = field(init=False)
    brutto_netto: str = field(init=False)
    abrechnungszeitraum_anfang: str = field(init=False)
    abrechnungszeitraum_ende: str = field(init=False)
    strasse: str = field(init=False)
    laendercode: str = field(init=False)
    postleitzahl: str = field(init=False)
    stadt: str = field(init=False)
    liegenschaftsnummer: str = field(init=False)
    grundstuecksbezeichnung: str = field(init=False)

    def __post_init__(self):
        r = self.raw.ljust(2048)
        self.version = _str(r, 1, 6)
        self.bfw_kundennummer = _str(r, 6, 16)
        self.betriebskode = _str(r, 16, 18)
        self.bfw_ordnungsbegriff = _str(r, 18, 31)
        self.kundlicher_ordnungsbegriff = _str(r, 31, 51)
        self.brutto_netto = _str(r, 51, 52)
        # Date fields at positions 53-57 (5 chars) per spec; treat as 8-char window
        self.abrechnungszeitraum_anfang = _parse_date(r[52:60])
        self.abrechnungszeitraum_ende = _parse_date(r[60:68])
        self.strasse = _str(r, 67, 102)
        self.laendercode = _str(r, 102, 105)
        self.postleitzahl = _str(r, 105, 115)
        self.stadt = _str(r, 115, 150)
        self.liegenschaftsnummer = _str(r, 150, 160)
        self.grundstuecksbezeichnung = _str(r, 160, 200)

    @property
    def property_number(self) -> str:
        """First 9 chars of BFW-Ordnungsbegriff = Liegenschaftsnummer."""
        return self.bfw_ordnungsbegriff[:9].strip() if self.bfw_ordnungsbegriff else ""

    @property
    def external_id(self) -> str:
        """Best available external identifier for advizeo."""
        return (
            self.kundlicher_ordnungsbegriff
            or self.liegenschaftsnummer
            or self.property_number
            or self.bfw_kundennummer
        )

    @property
    def real_estate_name(self) -> str:
        if self.grundstuecksbezeichnung:
            return self.grundstuecksbezeichnung
        parts = [self.strasse, self.postleitzahl, self.stadt]
        return " ".join(p for p in parts if p) or self.external_id


@dataclass
class MSatz:
    """M-Satz – Mietersatz / Tenant record (2048 bytes)."""
    raw: str
    version: str = field(init=False)
    bfw_kundennummer: str = field(init=False)
    betriebskode: str = field(init=False)
    bfw_ordnungsbegriff: str = field(init=False)
    kundlicher_ordnungsbegriff: str = field(init=False)
    nutzername1: str = field(init=False)
    nutzername2: str = field(init=False)
    nutzername3: str = field(init=False)
    nutzername4: str = field(init=False)
    nutzer_strasse: str = field(init=False)
    nutzer_laendercode: str = field(init=False)
    nutzer_plz: str = field(init=False)
    nutzer_stadt: str = field(init=False)
    einzugsdatum: str = field(init=False)
    auszugsdatum: str = field(init=False)
    leerstand: str = field(init=False)
    # Financial fields (advance payments) – extracted but not mapped to advizeo
    heizung_vorauszahlung_brutto: str = field(init=False)
    warmwasser_vorauszahlung_brutto: str = field(init=False)

    def __post_init__(self):
        r = self.raw.ljust(2048)
        self.version = _str(r, 1, 6)
        self.bfw_kundennummer = _str(r, 6, 16)
        self.betriebskode = _str(r, 16, 18)
        self.bfw_ordnungsbegriff = _str(r, 18, 31)
        self.kundlicher_ordnungsbegriff = _str(r, 31, 51)
        self.nutzername1 = _str(r, 51, 86)
        self.nutzername2 = _str(r, 86, 121)
        self.nutzername3 = _str(r, 121, 156)
        self.nutzername4 = _str(r, 156, 191)
        self.nutzer_strasse = _str(r, 191, 226)
        self.nutzer_laendercode = _str(r, 226, 229)
        self.nutzer_plz = _str(r, 229, 239)
        self.nutzer_stadt = _str(r, 239, 274)
        # Occupancy dates at positions 798-813 (0-indexed 797-813)
        self.einzugsdatum = _parse_date(r[797:805])
        self.auszugsdatum = _parse_date(r[805:813])
        self.leerstand = _str(r, 813, 814)
        # Financial (informational)
        self.heizung_vorauszahlung_brutto = _str(r, 443, 451)
        self.warmwasser_vorauszahlung_brutto = _str(r, 475, 483)

    @property
    def property_number(self) -> str:
        """First 9 chars of BFW Ordnungsbegriff."""
        return self.bfw_ordnungsbegriff[:9].strip() if self.bfw_ordnungsbegriff else ""

    @property
    def apartment_number(self) -> str:
        """Chars 9-12 of BFW Ordnungsbegriff."""
        if self.bfw_ordnungsbegriff and len(self.bfw_ordnungsbegriff) > 9:
            return self.bfw_ordnungsbegriff[9:13].strip()
        return ""

    @property
    def estate_unit_external_id(self) -> str:
        return self.kundlicher_ordnungsbegriff or self.apartment_number

    @property
    def tenant_name(self) -> str:
        parts = [
            self.nutzername1,
            self.nutzername2,
            self.nutzername3,
            self.nutzername4,
        ]
        return " ".join(p for p in parts if p)

    @property
    def is_vacant(self) -> bool:
        return bool(self.leerstand and self.leerstand != "0")


@dataclass
class BSatz:
    """B-Satz – Bewegung / Property fuel & device data (1024 bytes)."""
    raw: str
    version: str = field(init=False)
    bfw_kundennummer: str = field(init=False)
    betriebskode: str = field(init=False)
    bfw_ordnungsbegriff: str = field(init=False)
    waehrungscode: str = field(init=False)
    abrechnungszeitraum_start: str = field(init=False)
    abrechnungszeitraum_ende: str = field(init=False)
    brennstoffart: str = field(init=False)
    heizwert: str = field(init=False)
    warmwasser_verbrauch: str = field(init=False)
    warmwasser_prozentsatz: str = field(init=False)
    geraete_nr: str = field(init=False)
    masseinheit: str = field(init=False)

    def __post_init__(self):
        r = self.raw.ljust(1024)
        self.version = _str(r, 1, 6)
        self.bfw_kundennummer = _str(r, 6, 16)
        self.betriebskode = _str(r, 16, 18)
        self.bfw_ordnungsbegriff = _str(r, 18, 31)
        self.waehrungscode = _str(r, 31, 36)
        self.abrechnungszeitraum_start = _parse_date(r[36:44])
        self.abrechnungszeitraum_ende = _parse_date(r[44:52])
        self.brennstoffart = _str(r, 52, 56)
        self.heizwert = _str(r, 56, 66)
        self.warmwasser_verbrauch = _str(r, 174, 184)
        self.warmwasser_prozentsatz = _str(r, 184, 192)
        self.masseinheit = _str(r, 277, 281)
        self.geraete_nr = _str(r, 281, 301)

    @property
    def property_number(self) -> str:
        return self.bfw_ordnungsbegriff[:9].strip() if self.bfw_ordnungsbegriff else ""

    @property
    def device_type_advizeo(self) -> str:
        """Map BFW Brennstoffart code to advizeo device type."""
        code = self.brennstoffart.lstrip("0") or "0"
        return (
            BRENNSTOFF_TO_DEVICE_TYPE.get(self.brennstoffart)
            or BRENNSTOFF_TO_DEVICE_TYPE.get(code)
            or "WMZ"
        )


@dataclass
class DSatz:
    """D-Satz – Abrechnungsergebnisse / Settlement results (1024 bytes)."""
    raw: str
    bfw_ordnungsbegriff: str = field(init=False)
    kundlicher_ordnungsbegriff: str = field(init=False)
    abrechnungszeitraum_ende: str = field(init=False)
    abrechnungsbetrag_brutto: str = field(init=False)
    nutzer_name: str = field(init=False)

    def __post_init__(self):
        r = self.raw.ljust(1024)
        self.bfw_ordnungsbegriff = _str(r, 18, 31)
        self.kundlicher_ordnungsbegriff = _str(r, 31, 51)
        self.abrechnungszeitraum_ende = _parse_date(r[51:57])
        self.abrechnungsbetrag_brutto = _str(r, 57, 65)
        self.nutzer_name = _str(r, 165, 200)

    @property
    def property_number(self) -> str:
        return self.bfw_ordnungsbegriff[:9].strip() if self.bfw_ordnungsbegriff else ""


# ---------------------------------------------------------------------------
# Comgy 3×128-byte format detection + parser
# ---------------------------------------------------------------------------

# Pattern for Comgy M-header line:
#   M{digits}/{digits}/{PPPP}-{UUU}{DDMMYY}{DDMMYY}
# Example: M000000000010069100010887/341217/0001-001010125311225
_COMGY_RE = re.compile(r"^M\d+/\d+/(\d{4})-(\d{3})(\d{6})(\d{6})")


def _is_comgy_format(lines: list) -> bool:
    """Return True if file looks like Comgy compact 3-line record format."""
    if not lines:
        return False
    # BFW DTA records are ≥ 1024 chars; Comgy lines are ≤ 130 chars
    if any(len(l) > 300 for l in lines[:10] if l.strip()):
        return False
    # At least the first M-type line must match the Comgy header pattern
    for line in lines[:6]:
        if line.startswith("M") and _COMGY_RE.match(line):
            return True
    return False


def _parse_comgy(lines: list) -> dict:
    """
    Parse Comgy 3-line compact tenant format.

    Each tenant record = 3 × 128-byte lines:
      Line 1 (M-header): M{kundennr}/{code}/{PPPP}-{UUU}{DDMMYY}{DDMMYY}
      Line 2 (address):  {name:27}{PLZ:5} {city:21}{street:26}
      Line 3 (readings): {area_cm²:6}{meter readings...}

    Returns the same dict structure as parse_ml_satz(), using SimpleNamespace
    objects so that excel_generator.py and _build_preview work without changes.
    """
    non_empty = [l for l in lines if l.strip()]
    errors: List[str] = []
    props: dict = {}    # prop_num → SimpleNamespace (LSatz-compatible)
    tenants: List    = []  # SimpleNamespace objects (MSatz-compatible)

    i = 0
    while i < len(non_empty):
        line1 = non_empty[i]
        m = _COMGY_RE.match(line1)

        if not m:
            # L-Satz summary record or stray line – skip
            i += 1
            continue

        if i + 2 >= len(non_empty):
            i += 1
            continue

        line2 = non_empty[i + 1].ljust(128)
        line3 = non_empty[i + 2].ljust(128)

        prop_num  = m.group(1)   # "0001"
        unit_num  = m.group(2)   # "001"
        start_raw = m.group(3)   # DDMMYY e.g. "010125"
        end_raw   = m.group(4)   # DDMMYY e.g. "311225"

        # Line 2: name + address (fixed-width positions)
        tenant_name = line2[0:27].strip()
        plz         = line2[27:32].strip()
        city        = line2[33:54].strip()
        street      = line2[54:80].strip()

        # Line 3: area in m² × 100 (first 6 digits)
        area_raw = line3[:6].strip()
        try:
            area_m2 = round(int(area_raw) / 100, 2) if area_raw.isdigit() else 0.0
        except Exception:
            area_m2 = 0.0

        # Synthetic property record (first tenant's address wins per property)
        if prop_num not in props:
            name = f"{street}, {plz} {city}".strip(", ")
            props[prop_num] = SimpleNamespace(
                property_number  = prop_num,
                external_id      = prop_num,
                real_estate_name = name or prop_num,
                strasse          = street,
                postleitzahl     = plz,
                stadt            = city,
                laendercode      = "DE",
                grundstuecksbezeichnung = "",
                bfw_ordnungsbegriff     = prop_num,
                kundlicher_ordnungsbegriff = prop_num,
            )

        # Synthetic tenant record
        tenants.append(SimpleNamespace(
            property_number          = prop_num,
            apartment_number         = unit_num,
            estate_unit_external_id  = unit_num,
            bfw_ordnungsbegriff      = f"{prop_num}{unit_num}",
            kundlicher_ordnungsbegriff = unit_num,
            nutzername1              = tenant_name,
            tenant_name              = tenant_name,
            einzugsdatum             = "",   # not in Comgy format
            auszugsdatum             = "",   # not in Comgy format
            is_vacant                = not bool(tenant_name),
            area_m2                  = area_m2,
            billing_start            = _parse_date(start_raw),
            billing_end              = _parse_date(end_raw),
        ))

        i += 3  # consume all 3 lines of this group

    if not tenants:
        errors.append(
            "Keine Mieterdaten in der Comgy-Datei gefunden. "
            "/ No tenant data found in Comgy file."
        )

    n_props = len(props)
    errors.insert(0,
        f"[Comgy] Format erkannt: {n_props} Liegenschaft(en), {len(tenants)} Mieter. "
        f"Keine Geraetedaten (B-Satz) im Comgy-Format verfuegbar. "
        f"/ Comgy format detected: {n_props} propert(y/ies), {len(tenants)} tenants. "
        f"No device records in Comgy format."
    )

    l_saetze = sorted(props.values(), key=lambda x: x.property_number)
    return {
        "format":   "comgy",
        "a_saetze": [],
        "l_saetze": l_saetze,
        "m_saetze": tenants,
        "b_saetze": [],
        "d_saetze": [],
        "errors":   errors,
    }


# ---------------------------------------------------------------------------
# Main parse function
# ---------------------------------------------------------------------------

def parse_ml_satz(content: str) -> dict:
    """
    Parse BFW DTA ML Satz file content.

    Returns a dict with keys:
      a_saetze, l_saetze, m_saetze, b_saetze, d_saetze  – typed record lists
      errors   – list of warning/error strings
    """
    a_saetze: List[ASatz] = []
    l_saetze: List[LSatz] = []
    m_saetze: List[MSatz] = []
    b_saetze: List[BSatz] = []
    d_saetze: List[DSatz] = []
    errors: List[str] = []
    unknown_types: set = set()

    lines = content.splitlines()
    if not lines:
        errors.append("Datei ist leer / File is empty.")
        return {
            "a_saetze": [], "l_saetze": [], "m_saetze": [],
            "b_saetze": [], "d_saetze": [], "errors": errors,
        }

    # Auto-detect Comgy compact format (3×128-byte lines) vs BFW DTA (long lines)
    if _is_comgy_format(lines):
        return _parse_comgy(lines)

    for i, line in enumerate(lines, 1):
        if not line.strip():
            continue
        satztyp = line[0]
        try:
            if satztyp == "A":
                a_saetze.append(ASatz(raw=line))
            elif satztyp == "L":
                l_saetze.append(LSatz(raw=line))
            elif satztyp == "M":
                m_saetze.append(MSatz(raw=line))
            elif satztyp == "B":
                b_saetze.append(BSatz(raw=line))
            elif satztyp == "D":
                d_saetze.append(DSatz(raw=line))
            elif satztyp in ("K", "E", "W"):
                pass  # Known but not converted
            else:
                if satztyp not in unknown_types:
                    errors.append(
                        f"Zeile {i}: Unbekannter Satztyp '{satztyp}' – wird übersprungen."
                    )
                    unknown_types.add(satztyp)
        except Exception as exc:
            errors.append(f"Zeile {i}: Fehler beim Parsen – {exc}")

    if not l_saetze and not m_saetze and not b_saetze:
        errors.append(
            "Keine L-, M- oder B-Sätze gefunden. "
            "Bitte prüfen Sie, ob die Datei im BFW DTA Format vorliegt."
        )

    return {
        "a_saetze": a_saetze,
        "l_saetze": l_saetze,
        "m_saetze": m_saetze,
        "b_saetze": b_saetze,
        "d_saetze": d_saetze,
        "errors": errors,
    }
