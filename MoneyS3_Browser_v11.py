# -*- coding: utf-8 -*-
"""
Money S3 Browser v0.11
Read-only prohlížeč databází Money S3 přes COM objekt mon2kdbe.BFTable.

Nově:
- vydané i přijaté faktury
- automatická detekce VFaktPol.DAT / PFaktPol.DAT
- automatické hledání vazby hlavička faktury -> položky
- záložka Položky + raw pole vybrané položky
- CSV export faktur
- experimentální dohledání účetních zápisů / MD-Dal
- zobrazení předkontací, středisek, zakázek a činností
- bezpečnostní upozornění, nápověda a postup zálohy

Požadavky:
- Windows
- 32bit Python 3.x
- pywin32
- registrovaný COM objekt mon2kdbe.BFTable

Program nikdy nevolá Append/Update/Delete.
"""

import csv
import os
import struct
import sys
import platform
import tempfile
import subprocess
import webbrowser
import html
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, date
from pathlib import Path

try:
    import win32com.client
    import win32timezone  # required for COM DATE values in frozen PyInstaller builds
except ImportError:
    win32com = None


APP_TITLE = "Money S3 Browser v0.11"
DEFAULT_ROOTS = [
    r"C:\Users\Public\Documents\Solitea\Money S3",
    r"C:\Users\Public\Documents\CIGLER SOFTWARE\Money S3",
]


def is_32bit_python():
    return struct.calcsize("P") * 8 == 32


def safe_str(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    return str(value)


def norm(value):
    """Normalizace hodnoty pro hledání vazby mezi tabulkami."""
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        try:
            if float(value).is_integer():
                return str(int(value))
        except Exception:
            pass
        return str(value)
    s = str(value).strip()
    if not s:
        return None
    return s.lower()


class MoneyTable:
    """Read-only wrapper nad mon2kdbe.BFTable."""

    def __init__(self, path):
        self.path = str(path)
        self.obj = None
        self.columns = []

    def open(self):
        self.obj = win32com.client.Dispatch("mon2kdbe.BFTable")
        rc = self.obj.Open(self.path)
        if rc != 0 or self.obj.IsamError != 0:
            err = self.obj.IsamError
            try:
                self.obj.Close()
            except Exception:
                pass
            self.obj = None
            raise RuntimeError(
                f"Nelze otevřít {self.path}\nOpen={rc}, IsamError={err}"
            )
        self.columns = [self.obj.ColName(i) for i in range(self.obj.ColCount)]
        return self

    def close(self):
        if self.obj is not None:
            try:
                self.obj.Close()
            finally:
                self.obj = None

    def __enter__(self):
        return self.open()

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def read_all(self):
        rows = []
        if self.obj.RecCount <= 0:
            return rows

        self.obj.Top()
        while not self.obj.EOF:
            row = {}
            for col in self.columns:
                try:
                    row[col] = self.obj.Value(col)
                except Exception as exc:
                    row[col] = f"<chyba čtení: {exc}>"
            rows.append(row)
            self.obj.Next()
        return rows


def first_present(row, names, default=""):
    for name in names:
        if name in row:
            val = row.get(name)
            if val not in (None, ""):
                return val
    return default


def first_existing_file(directory, wanted_name):
    try:
        for p in Path(directory).iterdir():
            if p.is_file() and p.name.lower() == wanted_name.lower():
                return p
    except Exception:
        pass
    return None


def detect_year(rows, kind):
    candidates = (
        ["Vystaveno", "DatVyst", "DatUcPr", "PlnenoDPH"]
        if kind == "V"
        else ["Prijato", "DatUcPr", "PlnenoDPH", "Vystaveno"]
    )
    for row in rows[:20]:
        val = first_present(row, candidates, None)
        if isinstance(val, (datetime, date)):
            return val.year
        if val:
            text = str(val)
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(text[:10], fmt).year
                except Exception:
                    pass
    return None


def find_invoice_tables(root):
    root = Path(root)
    found = []
    if not root.exists():
        return found

    for base, dirs, files in os.walk(root):
        file_map = {f.lower(): f for f in files}
        for low_name, kind, title, item_name in (
            ("vfaktury.dat", "V", "Vydané faktury", "VFaktPol.DAT"),
            ("pfaktury.dat", "P", "Přijaté faktury", "PFaktPol.DAT"),
        ):
            if low_name in file_map:
                p = Path(base) / file_map[low_name]
                parts = [x.name for x in p.parents]
                agenda = next((x for x in parts if x.upper().startswith("AGENDA.")), "")
                rokdir = next((x for x in parts if x.upper().startswith("ROK.")), "")
                found.append({
                    "path": p,
                    "items_path": first_existing_file(base, item_name),
                    "kind": kind,
                    "title": title,
                    "agenda": agenda,
                    "rokdir": rokdir,
                })

    found.sort(key=lambda x: (x["agenda"], x["rokdir"], x["kind"]))
    return found


def relation_name_score(name):
    n = name.lower()
    exact = {
        "faktura", "vfaktura", "pfaktura", "cfaktura",
        "cisfakt", "cislofak", "cislofakt", "fakturacis",
        "faktcislo", "fakt_cislo", "fa_cislo", "parent",
    }
    if n in exact:
        return 120
    score = 0
    if "fakt" in n:
        score += 80
    if "guid" in n:
        score += 60
    if "dokl" in n:
        score += 45
    if "parent" in n or "master" in n:
        score += 45
    if "cis" in n:
        score += 20
    if n == "cislo":
        score -= 25  # typicky vlastní ID řádku, ne FK
    return score


def find_related_items(header, item_rows):
    """
    Najde položky patřící k faktuře.

    U VFaktPol/PFaktPol je v ověřených datech Money S3 skutečná vazba:
        VFaktury/PFaktury.Cislo == VFaktPol/PFaktPol.Cislo

    CisloPoloz je pořadové číslo položky, nikoliv číslo faktury.
    Původní heuristika mohla při shodě náhodou vybrat právě CisloPoloz.
    """
    if not item_rows:
        return [], "tabulka položek je prázdná"

    # 1) Primární a ověřená vazba Money S3.
    h_cislo = norm(header.get("Cislo"))
    if h_cislo is not None and "Cislo" in item_rows[0]:
        matches = [row for row in item_rows if norm(row.get("Cislo")) == h_cislo]
        if matches:
            def item_order(row):
                for key in ("Poradi", "CisloPoloz"):
                    v = row.get(key)
                    try:
                        return (0, float(v))
                    except (TypeError, ValueError):
                        pass
                return (1, 0)

            matches.sort(key=item_order)
            return matches, "Cislo = Cislo"

    # 2) Fallback pro případ jiné/starší struktury databáze.
    # Záměrně ignorujeme CisloPoloz, Poradi a podobná pořadová pole.
    excluded = {
        "cislopoloz", "poradi", "cislokat", "cisloskpol",
        "cislozafak", "cisloskpol", "cislozafak",
    }

    header_keys = [
        ("GUID", 120),
        ("Doklad", 90),
        ("EvCisDokl", 85),
        ("UDoklad", 80),
        ("VarSymbol", 50),
    ]

    hvals = []
    for hname, hscore in header_keys:
        if hname in header:
            nv = norm(header.get(hname))
            if nv is not None:
                hvals.append((hname, nv, hscore))

    if not hvals:
        return [], "v hlavičce není použitelný identifikátor"

    best = None
    for col in item_rows[0].keys():
        if col.lower() in excluded:
            continue

        nscore = relation_name_score(col)
        if nscore <= 0:
            continue

        for hname, hval, hscore in hvals:
            matches = [row for row in item_rows if norm(row.get(col)) == hval]
            if not matches:
                continue

            score = nscore + hscore + min(len(matches), 20) * 3
            if hname == "GUID" and "guid" in col.lower():
                score += 100

            candidate = (score, matches, f"{col} = {hname}")
            if best is None or candidate[0] > best[0]:
                best = candidate

    if best is None:
        return [], "vazba nebyla automaticky nalezena"

    matches = best[1]

    def item_order(row):
        for key in ("Poradi", "CisloPoloz"):
            v = row.get(key)
            try:
                return (0, float(v))
            except (TypeError, ValueError):
                pass
        return (1, 0)

    matches.sort(key=item_order)
    return matches, best[2]



def item_summary(row):
    qty = first_present(
        row, ["PocetMJ", "Mnozstvi", "Mnoz", "Pocet", "PocetJedn", "MJ_Pocet"]
    )
    price = first_present(
        row, ["Cena", "JednCena", "CenaJedn", "CenaMJ", "CenaBezDPH", "CenaZakl"]
    )
    stored_total = first_present(
        row,
        ["CelkemSDPH", "Celkem", "CenaCelkem", "CenaSDPH", "Castka",
         "CenaPoSleve", "Zaklad"],
        None,
    )
    total = stored_total
    total_calculated = False
    if total in (None, ""):
        try:
            if qty not in (None, "") and price not in (None, ""):
                total = float(qty) * float(price)
                total_calculated = True
        except (TypeError, ValueError):
            total = ""

    return {
        "popis": first_present(
            row, ["Popis", "Nazev", "Text", "Polozka", "Oznaceni", "Zkratka", "SkladNazev"]
        ),
        "poznamka": first_present(
            row, ["Poznamka", "Pozn", "TextPozn", "Komentar", "PoznamkaPol"]
        ),
        "mnozstvi": qty,
        "mj": first_present(row, ["Jednotka", "MJ", "MerJedn", "Jedn", "MJ_Zkratka"]),
        "cena": price,
        "dph": first_present(row, ["SazbaDPH", "DPH", "ProcentoDPH", "DPHSazba", "KodDPH"]),
        "celkem": total,
        "celkem_vypocteno": total_calculated,
    }




def format_amount(value):
    if value in (None, ""):
        return ""
    try:
        n = float(value)
        if n.is_integer():
            return f"{n:,.0f}".replace(",", " ")
        return f"{n:,.2f}".replace(",", " ")
    except (TypeError, ValueError):
        return safe_str(value)


def format_quantity(value):
    if value in (None, ""):
        return ""
    try:
        n = float(value)
        if n.is_integer():
            return str(int(n))
        return f"{n:g}"
    except (TypeError, ValueError):
        return safe_str(value)


def first_by_alias(row, aliases, default=""):
    """Case-insensitive varianta first_present pro neznámé/starší struktury."""
    lower_map = {str(k).lower(): k for k in row.keys()}
    for alias in aliases:
        real = lower_map.get(alias.lower())
        if real is not None:
            value = row.get(real)
            if value not in (None, ""):
                return value
    return default


def first_real_value(row, aliases, default=""):
    """
    Vrátí první smysluplnou hodnotu.
    Money může mít v některých D_/O_ polích text 'anonymizovano',
    zatímco skutečná obchodní adresa zůstává v *_Ob polích.
    """
    placeholders = {
        "anonymizovano", "anonymizováno", "anonymized", "anonymised",
        "<anonymizovano>", "<anonymizováno>"
    }
    lower_map = {str(k).lower(): k for k in row.keys()}
    fallback = None

    for alias in aliases:
        real = lower_map.get(alias.lower())
        if real is None:
            continue
        value = row.get(real)
        if value in (None, ""):
            continue

        text = safe_str(value).strip()
        if not text:
            continue

        if text.lower() in placeholders:
            if fallback is None:
                fallback = value
            continue

        return value

    return default if fallback is None else fallback


def accounting_table_score(path, columns):
    """Heuristické skóre tabulky, která může obsahovat účetní deník/zápisy."""
    name = Path(path).stem.lower()
    cols = [str(c).lower() for c in columns]
    score = 0

    if "denik" in name:
        score += 8
    if "ucet" in name or "uct" in name:
        score += 3

    for c in cols:
        if c in {"md", "d", "dal", "madati", "strana"}:
            score += 5
        if "ucet" in c:
            score += 3
        if "predkont" in c:
            score += 3
        if "castk" in c or c in {"cena", "celkem"}:
            score += 2
        if "doklad" in c:
            score += 2
        if c in {"varsymbol", "parsymbol", "parovsymbol"}:
            score += 2
        if "datum" in c or c.startswith("dat"):
            score += 1
        if "popis" in c:
            score += 1
        if "zdroj" in c:
            score += 1
    return score


def invoice_row_match_score(invoice, row):
    """
    Heuristicky zjistí, zda účetní řádek souvisí s aktuální fakturou.
    Používá jen identifikátory, ne částku, aby se nespojovaly cizí doklady.
    """
    values = []

    for key, weight in (
        ("GUID", 12),
        ("Doklad", 10),
        ("EvCisDokl", 9),
        ("UDoklad", 8),
        ("VarSymbol", 8),
        ("iDokladID", 8),
        ("ParSymbol", 7),
    ):
        if key in invoice:
            v = norm(invoice.get(key))
            if v is not None:
                values.append((key.lower(), v, weight))

    if not values:
        return 0

    score = 0
    for col, raw in row.items():
        c = str(col).lower()
        rv = norm(raw)
        if rv is None:
            continue

        relevant_name = (
            "doklad" in c or "guid" in c or "symbol" in c or
            "zdroj" in c or c in {"idokladid", "parsymbol", "varsymbol"}
        )
        if not relevant_name:
            continue

        for source_name, wanted, weight in values:
            if rv == wanted:
                score += weight
                if source_name in c or ("doklad" in c and "doklad" in source_name):
                    score += 3

    return score


def accounting_row_summary(row):
    """
    Sjednotí běžné názvy polí účetního zápisu.
    Umí i jednosloupcový model Ucet + ProtiUcet + Strana.
    """
    datum = first_by_alias(
        row, ["Datum", "DatUcPr", "DatVyst", "Vystaveno", "DatDokl", "DatUct"]
    )
    doklad = first_by_alias(
        row, ["Doklad", "CisDokl", "CisloDokladu", "EvCisDokl", "UDoklad"]
    )
    popis = first_by_alias(row, ["Popis", "Text", "Poznamka", "Pozn"])

    md = first_by_alias(
        row,
        ["UcMD", "MD", "MaDati", "MáDáti", "UcetMD", "MDUcet", "Ucet_MD",
         "UcetMaDati", "UcetMDKod", "KodUctuMD"],
    )
    dal = first_by_alias(
        row,
        ["UcD", "D", "Dal", "UcetD", "UcetDal", "DalUcet", "Ucet_Dal",
         "UcetDalKod", "KodUctuDal"],
    )

    ucet = first_by_alias(
        row, ["Ucet", "KodUctu", "UcetKod", "CisloUctu", "UcetCis"]
    )
    protiucet = first_by_alias(
        row, ["ProtiUcet", "Protiucet", "ProtUcet", "UcetProt", "UcetProti"]
    )
    strana = first_by_alias(row, ["Strana", "MD_D", "StranaUctu"])
    strana_norm = safe_str(strana).strip().lower()

    if not md and not dal and ucet:
        if strana_norm in {"md", "m", "1", "má dáti", "ma dati", "madati"}:
            md, dal = ucet, protiucet
        elif strana_norm in {"d", "dal", "2"}:
            md, dal = protiucet, ucet

    castka = first_by_alias(
        row,
        ["Castka", "CastkaMD", "CastkaDal", "Celkem", "Cena",
         "CastkaCM", "Hodnota", "Obrat"],
    )

    return {
        "datum": datum,
        "doklad": doklad,
        "popis": popis,
        "md": md,
        "dal": dal,
        "ucet": ucet,
        "protiucet": protiucet,
        "strana": strana,
        "castka": castka,
        "predkont": first_by_alias(
            row, ["PredKontac", "Predkontac", "PredKont", "Predkontace"]
        ),
        "zdroj": first_by_alias(row, ["Zdroj", "Source", "TypZdroj"]),
    }

class MoneyBrowserApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1520x920")
        self.minsize(1150, 680)

        self.datasets = {}
        self.current_dataset = None
        self.current_rows = []
        self.filtered_rows = []
        self.item_cache = {}
        self.current_item_rows = []
        self.accounting_table_cache = {}
        self.current_accounting_rows = []
        self.current_accounting_raw = []

        self._build_ui()
        self.after(100, self._startup_checks)

    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("InvoiceTitle.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("InvoiceTotal.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Warning.TLabel", font=("Segoe UI", 10, "bold"), foreground="red")
        style.configure(
            "Warning.TLabel",
            font=("Segoe UI", 10, "bold"),
            foreground="red",
        )

        top = ttk.Frame(self, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Money S3 Browser v0.11", style="InvoiceTitle.TLabel").pack(
            side="left", padx=(0, 12)
        )
        ttk.Label(top, text="Kořen dat:").pack(side="left")
        self.root_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.root_var, width=90).pack(
            side="left", padx=6, fill="x", expand=True
        )
        ttk.Button(top, text="Vybrat…", command=self.choose_root).pack(side="left")
        ttk.Button(top, text="Detekovat databáze", command=self.scan).pack(
            side="left", padx=(6, 0)
        )

        search_bar = ttk.Frame(self, padding=(6, 0, 6, 6))
        search_bar.pack(fill="x")
        ttk.Label(search_bar, text="Hledat ve fakturách:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        ttk.Entry(search_bar, textvariable=self.search_var).pack(
            side="left", padx=6, fill="x", expand=True
        )
        ttk.Button(search_bar, text="Export CSV", command=self.export_current_dataset).pack(
            side="left"
        )

        safety = ttk.Frame(self, padding=(6, 0, 6, 6))
        safety.pack(fill="x")
        ttk.Label(
            safety,
            text="UPOZORNĚNÍ: Neoficiální read-only nástroj. Používejte na vlastní nebezpečí a pracujte pouze s kopií / zálohou dat.",
            style="Warning.TLabel",
        ).pack(side="left")
        ttk.Button(safety, text="Jak udělat zálohu?", command=self.show_backup_help).pack(
            side="right"
        )
        ttk.Button(
            safety, text="Diagnostika → schránka", command=self.copy_diagnostics
        ).pack(side="right", padx=(0, 6))

        main = ttk.Panedwindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        left = ttk.Frame(main)
        center = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, weight=1)
        main.add(center, weight=3)
        main.add(right, weight=4)

        ttk.Label(left, text="Nalezené databáze").pack(anchor="w")
        self.db_tree = ttk.Treeview(left, show="tree")
        self.db_tree.pack(fill="both", expand=True)
        self.db_tree.bind("<<TreeviewSelect>>", self.on_dataset_select)

        ttk.Label(center, text="Faktury").pack(anchor="w")
        cols = ("doklad", "datum", "partner", "ico", "castka", "vs", "popis")
        self.invoice_tree = ttk.Treeview(
            center, columns=cols, show="headings", selectmode="browse"
        )
        headings = {
            "doklad": "Doklad", "datum": "Datum", "partner": "Partner",
            "ico": "IČO", "castka": "Částka", "vs": "VS", "popis": "Popis",
        }
        widths = {
            "doklad": 105, "datum": 90, "partner": 220, "ico": 90,
            "castka": 95, "vs": 100, "popis": 280,
        }
        for c in cols:
            self.invoice_tree.heading(c, text=headings[c])
            self.invoice_tree.column(c, width=widths[c], anchor="w")
        self.invoice_tree.pack(fill="both", expand=True)
        self.invoice_tree.bind("<<TreeviewSelect>>", self.on_invoice_select)

        self.detail_notebook = ttk.Notebook(right)
        self.detail_notebook.pack(fill="both", expand=True)

        self.form_tab = ttk.Frame(self.detail_notebook, padding=8)
        self.accounting_tab = ttk.Frame(self.detail_notebook, padding=8)
        self.tech_tab = ttk.Frame(self.detail_notebook, padding=8)
        self.help_tab = ttk.Frame(self.detail_notebook, padding=8)
        self.detail_notebook.add(self.form_tab, text="Faktura")
        self.detail_notebook.add(self.accounting_tab, text="Účetnictví")
        self.detail_notebook.add(self.tech_tab, text="Technická data")
        self.detail_notebook.add(self.help_tab, text="Nápověda / O programu")

        toolbar = ttk.Frame(self.form_tab)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(
            toolbar, text="Kopírovat celou fakturu", command=self.copy_all_invoice_text
        ).pack(side="right")
        ttk.Button(
            toolbar, text="Kopírovat označené pole", command=self.copy_selected_invoice_text
        ).pack(side="right", padx=(0, 6))
        ttk.Button(
            toolbar, text="Uložit nouzové PDF", command=self.save_invoice_pdf
        ).pack(side="left")
        ttk.Button(
            toolbar, text="Nouzový tisk", command=self.print_invoice
        ).pack(side="left", padx=(6, 0))

        header = ttk.Frame(self.form_tab)
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=1)

        doc_box = ttk.LabelFrame(header, text="Doklad", padding=8)
        doc_box.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        partner_box = ttk.LabelFrame(header, text="Partner", padding=8)
        partner_box.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        self.invoice_vars = {}

        for r, (key, label) in enumerate([
            ("doc_number", "Číslo dokladu"), ("ev_number", "Evidenční číslo"),
            ("date", "Datum"), ("duzp", "DUZP / plnění DPH"),
            ("due", "Splatnost"), ("paid", "Uhrazeno"),
            ("vs", "Variabilní symbol"), ("ks", "Konstantní symbol"),
            ("payment", "Úhrada"),
        ]):
            self._make_readonly_field(doc_box, r, label, key)

        for r, (key, label) in enumerate([
            ("partner", "Název"), ("street", "Ulice"), ("city", "Město"),
            ("psc", "PSČ"), ("ico", "IČO"), ("dic", "DIČ"),
            ("email", "E-mail"), ("phone", "Telefon"),
        ]):
            self._make_readonly_field(partner_box, r, label, key)

        summary = ttk.LabelFrame(self.form_tab, text="Souhrn faktury", padding=8)
        summary.pack(fill="x", pady=(8, 6))
        summary.columnconfigure(1, weight=1)
        summary.columnconfigure(3, weight=1)

        self._make_readonly_field(
            summary, 0, "Popis", "description", label_col=0, entry_col=1, span=3
        )
        self._make_readonly_field(summary, 1, "Měna", "currency", 0, 1)
        self._make_readonly_field(summary, 1, "Vystavil", "issuer", 2, 3)
        self._make_readonly_field(summary, 2, "Předkontace", "predkont", 0, 1)
        self._make_readonly_field(summary, 2, "Středisko", "stredisko", 2, 3)
        self._make_readonly_field(summary, 3, "Zakázka", "zakazka", 0, 1)
        self._make_readonly_field(summary, 3, "Činnost", "cinnost", 2, 3)

        ttk.Label(summary, text="Celkem s DPH:", style="InvoiceTotal.TLabel").grid(
            row=4, column=0, sticky="w", pady=(8, 2)
        )
        self.invoice_vars["total"] = tk.StringVar()
        ttk.Entry(
            summary, textvariable=self.invoice_vars["total"], state="readonly",
            font=("Segoe UI", 12, "bold")
        ).grid(row=4, column=1, sticky="ew", padx=(6, 14), pady=(8, 2))

        ttk.Label(summary, text="K úhradě:", style="InvoiceTotal.TLabel").grid(
            row=4, column=2, sticky="w", pady=(8, 2)
        )
        self.invoice_vars["to_pay"] = tk.StringVar()
        ttk.Entry(
            summary, textvariable=self.invoice_vars["to_pay"], state="readonly",
            font=("Segoe UI", 12, "bold")
        ).grid(row=4, column=3, sticky="ew", padx=(6, 0), pady=(8, 2))

        items_box = ttk.LabelFrame(self.form_tab, text="Položky faktury", padding=6)
        items_box.pack(fill="both", expand=True, pady=(2, 0))

        self.item_info_var = tk.StringVar(value="Vyber fakturu.")
        ttk.Label(items_box, textvariable=self.item_info_var).pack(
            fill="x", anchor="w", pady=(0, 4)
        )

        item_cols = ("popis", "poznamka", "mnozstvi", "mj", "cena", "dph", "celkem")
        self.item_tree = ttk.Treeview(
            items_box, columns=item_cols, show="headings", height=9, selectmode="browse"
        )
        item_heads = {
            "popis": "Název / popis", "poznamka": "Poznámka",
            "mnozstvi": "Počet", "mj": "MJ", "cena": "Cena / MJ",
            "dph": "DPH", "celkem": "Celkem položka",
        }
        item_widths = {
            "popis": 260, "poznamka": 260, "mnozstvi": 70, "mj": 55,
            "cena": 90, "dph": 60, "celkem": 110,
        }
        for c in item_cols:
            self.item_tree.heading(c, text=item_heads[c])
            self.item_tree.column(c, width=item_widths[c], anchor="w")
        self.item_tree.pack(fill="both", expand=True)
        self.item_tree.bind("<<TreeviewSelect>>", self.on_item_select)
        self.item_tree.bind("<Control-c>", self.copy_selected_item_row)
        self.item_tree.bind("<Control-C>", self.copy_selected_item_row)

        # Účetnictví: předkontace + experimentální vazba na účetní deník
        acc_toolbar = ttk.Frame(self.accounting_tab)
        acc_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(
            acc_toolbar,
            text="Primárně se čte UcDenik.DAT a hledá přesná vazba na číslo dokladu; ostatní vazby jsou pouze fallback."
        ).pack(side="left")
        ttk.Button(
            acc_toolbar,
            text="Najít účetní zápisy této faktury",
            command=self.scan_accounting_for_current_invoice,
        ).pack(side="right")
        ttk.Button(
            acc_toolbar,
            text="Kopírovat raw zápis",
            command=self.copy_selected_accounting_raw,
        ).pack(side="right", padx=(0, 6))
        ttk.Button(
            acc_toolbar,
            text="Kopírovat řádek",
            command=self.copy_selected_accounting_row,
        ).pack(side="right", padx=(0, 6))

        acc_top = ttk.LabelFrame(self.accounting_tab, text="Zaúčtování z dokladu / položek", padding=6)
        acc_top.pack(fill="x", pady=(0, 6))

        self.accounting_info_var = tk.StringVar(value="Vyber fakturu.")
        ttk.Label(acc_top, textvariable=self.accounting_info_var).pack(anchor="w")

        self.item_accounting_tree = ttk.Treeview(
            acc_top,
            columns=("popis", "predkont", "stredisko", "zakazka", "cinnost"),
            show="headings",
            height=5,
        )
        for c, title, width in (
            ("popis", "Položka", 280),
            ("predkont", "Předkontace", 110),
            ("stredisko", "Středisko", 90),
            ("zakazka", "Zakázka", 90),
            ("cinnost", "Činnost", 90),
        ):
            self.item_accounting_tree.heading(c, text=title)
            self.item_accounting_tree.column(c, width=width, anchor="w")
        self.item_accounting_tree.pack(fill="x", pady=(5, 0))

        acc_results = ttk.LabelFrame(
            self.accounting_tab, text="Nalezené účetní zápisy (MD / Dal)", padding=6
        )
        acc_results.pack(fill="both", expand=True)

        self.accounting_tree = ttk.Treeview(
            acc_results,
            columns=("table", "vazba", "zdroj", "datum", "doklad", "popis", "md", "dal", "ucet", "protiucet", "strana", "castka", "predkont"),
            show="headings",
            height=9,
        )
        acc_cols = (
            ("table", "Tabulka", 95),
            ("vazba", "Vazba", 90),
            ("zdroj", "Zdroj", 55),
            ("datum", "Datum", 85),
            ("doklad", "Doklad", 100),
            ("popis", "Popis", 190),
            ("md", "Má dáti", 85),
            ("dal", "Dal", 85),
            ("ucet", "Účet", 85),
            ("protiucet", "Protiúčet", 85),
            ("strana", "Strana", 55),
            ("castka", "Částka", 90),
            ("predkont", "Předkontace", 100),
        )
        for c, title, width in acc_cols:
            self.accounting_tree.heading(c, text=title)
            self.accounting_tree.column(c, width=width, anchor="w")
        self.accounting_tree.pack(fill="both", expand=True)
        self.accounting_tree.bind("<<TreeviewSelect>>", self.on_accounting_select)
        self.accounting_tree.bind("<Control-c>", self.copy_selected_accounting_row)
        self.accounting_tree.bind("<Control-C>", self.copy_selected_accounting_row)

        self.accounting_status_var = tk.StringVar(
            value="Klikni na „Najít účetní zápisy této faktury“."
        )
        ttk.Label(acc_results, textvariable=self.accounting_status_var).pack(
            fill="x", pady=(4, 0)
        )

        self.accounting_raw_tree = ttk.Treeview(
            acc_results, columns=("field", "value"), show="headings", height=6
        )
        self.accounting_raw_tree.heading("field", text="Pole účetního záznamu")
        self.accounting_raw_tree.heading("value", text="Hodnota")
        self.accounting_raw_tree.column("field", width=180, anchor="w")
        self.accounting_raw_tree.column("value", width=500, anchor="w")
        self.accounting_raw_tree.pack(fill="x", pady=(5, 0))
        self.accounting_raw_tree.bind("<Control-c>", self.copy_selected_accounting_raw)
        self.accounting_raw_tree.bind("<Control-C>", self.copy_selected_accounting_raw)

        tech_pane = ttk.Panedwindow(self.tech_tab, orient="horizontal")
        tech_pane.pack(fill="both", expand=True)
        raw_inv_frame = ttk.LabelFrame(tech_pane, text="Všechna pole faktury", padding=4)
        raw_item_frame = ttk.LabelFrame(
            tech_pane, text="Všechna pole vybrané položky", padding=4
        )
        tech_pane.add(raw_inv_frame, weight=1)
        tech_pane.add(raw_item_frame, weight=1)

        self.raw_tree = ttk.Treeview(
            raw_inv_frame, columns=("field", "value"), show="headings"
        )
        self.raw_tree.heading("field", text="Pole")
        self.raw_tree.heading("value", text="Hodnota")
        self.raw_tree.column("field", width=170, anchor="w")
        self.raw_tree.column("value", width=330, anchor="w")
        self.raw_tree.pack(fill="both", expand=True)

        self.item_raw_tree = ttk.Treeview(
            raw_item_frame, columns=("field", "value"), show="headings"
        )
        self.item_raw_tree.heading("field", text="Pole")
        self.item_raw_tree.heading("value", text="Hodnota")
        self.item_raw_tree.column("field", width=170, anchor="w")
        self.item_raw_tree.column("value", width=330, anchor="w")
        self.item_raw_tree.pack(fill="both", expand=True)

        # Nápověda / disclaimer
        help_toolbar = ttk.Frame(self.help_tab)
        help_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(
            help_toolbar, text="Kopírovat text nápovědy", command=self.copy_help_text
        ).pack(side="right")

        self.help_text = tk.Text(
            self.help_tab, wrap="word", padx=14, pady=12, font=("Segoe UI", 10), relief="flat"
        )
        self.help_text.pack(fill="both", expand=True)
        self._render_help()
        self.help_text.configure(state="disabled")

        self.status_var = tk.StringVar(value="Připraveno.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").pack(
            fill="x", side="bottom"
        )

    def _make_readonly_field(
        self, parent, row, label, key, label_col=0, entry_col=1, span=1
    ):
        ttk.Label(parent, text=label + ":").grid(
            row=row, column=label_col, sticky="w", pady=2
        )
        var = tk.StringVar()
        self.invoice_vars[key] = var
        entry = ttk.Entry(parent, textvariable=var, state="readonly")
        entry.grid(
            row=row, column=entry_col, columnspan=span,
            sticky="ew", padx=(6, 0), pady=2
        )
        parent.columnconfigure(entry_col, weight=1)
        return entry

    def _startup_checks(self):
        if os.name != "nt":
            messagebox.showerror(APP_TITLE, "Tento program je určen pro Windows.")
            return
        if not is_32bit_python():
            messagebox.showerror(
                APP_TITLE,
                "MON2KDBE.DLL je 32bitová.\n\n"
                "Spusť program v 32bit Pythonu.",
            )
            return
        if win32com is None:
            messagebox.showerror(
                APP_TITLE,
                "Chybí pywin32.\n\nNainstaluj:\npy -3.13-32 -m pip install pywin32",
            )
            return

        for p in DEFAULT_ROOTS:
            if Path(p).exists():
                self.root_var.set(p)
                self.scan()
                break

        self.after(
            300,
            lambda: self.status_var.set(
                "Read-only režim. Doporučení: otevřete kopii / zálohu dat, ne jediný originál."
            ),
        )

    def choose_root(self):
        initial = self.root_var.get() or r"C:\Users\Public\Documents"
        p = filedialog.askdirectory(initialdir=initial, title="Vyber kořen Money S3")
        if p:
            self.root_var.set(p)
            self.scan()

    def scan(self):
        root = self.root_var.get().strip()
        if not root:
            return

        self.db_tree.delete(*self.db_tree.get_children())
        self.datasets.clear()
        self.current_dataset = None
        self.current_rows = []
        self.item_cache.clear()
        self.invoice_tree.delete(*self.invoice_tree.get_children())
        self.clear_details()

        self.status_var.set("Hledám databáze…")
        self.update_idletasks()

        tables = find_invoice_tables(root)
        if not tables:
            self.status_var.set("Nenalezeny VFaktury.DAT ani PFaktury.DAT.")
            return

        agenda_nodes = {}
        for idx, ds in enumerate(tables):
            agenda = ds["agenda"] or "(bez AGENDA.xxx)"
            if agenda not in agenda_nodes:
                agenda_nodes[agenda] = self.db_tree.insert(
                    "", "end", text=agenda, open=True
                )

            pol = " + položky" if ds["items_path"] else ""
            label = f'{ds["rokdir"] or "(bez ROK.xxx)"} — {ds["title"]}{pol}'
            iid = f"ds_{idx}"
            self.datasets[iid] = ds
            self.db_tree.insert(agenda_nodes[agenda], "end", iid=iid, text=label)

        item_tables = sum(1 for x in tables if x.get("items_path"))
        self.status_var.set(
            f"Nalezeno databází faktur: {len(tables)} | databází položek: {item_tables}"
        )

    def on_dataset_select(self, _event=None):
        sel = self.db_tree.selection()
        if not sel:
            return
        iid = sel[0]
        if iid not in self.datasets:
            return

        ds = self.datasets[iid]
        self.status_var.set(f"Načítám {ds['path']}…")
        self.update_idletasks()

        try:
            with MoneyTable(ds["path"]) as table:
                rows = table.read_all()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            self.status_var.set("Chyba při čtení databáze.")
            return

        ds = dict(ds)
        ds["year"] = detect_year(rows, ds["kind"])
        ds["count"] = len(rows)
        self.current_dataset = ds
        self.current_rows = rows
        self.apply_filter()
        year_text = ds["year"] if ds["year"] else ds["rokdir"]
        item_state = "položky nalezeny" if ds["items_path"] else "bez tabulky položek"
        self.status_var.set(
            f'{ds["title"]} | {ds["agenda"]} | {year_text} | '
            f'záznamů: {len(rows)} | {item_state}'
        )

    def row_summary(self, row):
        kind = self.current_dataset["kind"] if self.current_dataset else "V"

        if kind == "V":
            partner = first_present(
                row, ["O_Nazev", "O_Jmeno", "O_NazevOb", "D_Nazev"]
            )
            ico = first_present(row, ["O_ICO", "D_ICO"])
            datum = first_present(row, ["Vystaveno", "DatVyst", "DatUcPr"])
        else:
            partner = first_present(
                row, ["D_Nazev", "D_Jmeno", "O_Nazev", "O_Jmeno", "O_NazevOb"]
            )
            ico = first_present(row, ["D_ICO", "O_ICO"])
            datum = first_present(row, ["Prijato", "DatUcPr", "Vystaveno", "DatVyst"])

        return {
            "doklad": first_present(row, ["Doklad", "EvCisDokl", "UDoklad"]),
            "datum": datum,
            "partner": partner,
            "ico": ico,
            "castka": first_present(
                row, ["CelkemSDPH", "KUhradeSDP", "Proplatit", "ValutyKUhr"]
            ),
            "vs": first_present(row, ["VarSymbol"]),
            "popis": first_present(row, ["Popis", "TextPredFa"]),
        }

    def apply_filter(self):
        self.invoice_tree.delete(*self.invoice_tree.get_children())
        needle = self.search_var.get().strip().lower()
        self.filtered_rows = []

        for row in self.current_rows:
            summary = self.row_summary(row)
            haystack = " ".join(safe_str(v).lower() for v in summary.values())
            if needle and needle not in haystack:
                continue

            idx = len(self.filtered_rows)
            self.filtered_rows.append(row)
            self.invoice_tree.insert(
                "", "end", iid=f"row_{idx}",
                values=(
                    safe_str(summary["doklad"]),
                    safe_str(summary["datum"]),
                    safe_str(summary["partner"]),
                    safe_str(summary["ico"]),
                    safe_str(summary["castka"]),
                    safe_str(summary["vs"]),
                    safe_str(summary["popis"]),
                ),
            )
        self.clear_details()

    def clear_details(self):
        if hasattr(self, "invoice_vars"):
            for var in self.invoice_vars.values():
                var.set("")
        self.raw_tree.delete(*self.raw_tree.get_children())
        self.item_tree.delete(*self.item_tree.get_children())
        self.item_raw_tree.delete(*self.item_raw_tree.get_children())
        self.current_item_rows = []
        self.current_invoice_row = None
        self.item_info_var.set("Vyber fakturu.")
        if hasattr(self, "item_accounting_tree"):
            self.item_accounting_tree.delete(*self.item_accounting_tree.get_children())
        if hasattr(self, "accounting_tree"):
            self.accounting_tree.delete(*self.accounting_tree.get_children())
        if hasattr(self, "accounting_raw_tree"):
            self.accounting_raw_tree.delete(*self.accounting_raw_tree.get_children())
        self.current_accounting_rows = []
        self.current_accounting_raw = []
        if hasattr(self, "accounting_info_var"):
            self.accounting_info_var.set("Vyber fakturu.")
        if hasattr(self, "accounting_status_var"):
            self.accounting_status_var.set("Klikni na „Najít účetní zápisy této faktury“.")

    def on_invoice_select(self, _event=None):
        sel = self.invoice_tree.selection()
        if not sel:
            return
        idx = int(sel[0].split("_", 1)[1])
        if idx >= len(self.filtered_rows):
            return
        self.show_invoice(self.filtered_rows[idx])

    def copy_selected_invoice_text(self):
        widget = self.focus_get()
        text = ""
        try:
            if isinstance(widget, (tk.Entry, ttk.Entry)) and widget.selection_present():
                text = widget.selection_get()
        except Exception:
            text = ""

        if not text:
            messagebox.showinfo(
                APP_TITLE,
                "Klikni do pole formuláře a označ v něm text, který chceš kopírovat."
            )
            return

        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set("Označený text byl zkopírován do schránky.")

    def _invoice_clipboard_text(self):
        if not getattr(self, "current_invoice_row", None):
            return ""

        v = {k: var.get() for k, var in self.invoice_vars.items()}
        lines = [
            f"Číslo dokladu: {v.get('doc_number', '')}",
            f"Evidenční číslo: {v.get('ev_number', '')}",
            f"Datum: {v.get('date', '')}",
            f"DUZP / plnění DPH: {v.get('duzp', '')}",
            f"Splatnost: {v.get('due', '')}",
            f"Uhrazeno: {v.get('paid', '')}",
            "",
            f"Partner: {v.get('partner', '')}",
            f"Ulice: {v.get('street', '')}",
            f"Město: {v.get('psc', '')} {v.get('city', '')}".strip(),
            f"IČO: {v.get('ico', '')}",
            f"DIČ: {v.get('dic', '')}",
            f"E-mail: {v.get('email', '')}",
            f"Telefon: {v.get('phone', '')}",
            "",
            f"Variabilní symbol: {v.get('vs', '')}",
            f"Konstantní symbol: {v.get('ks', '')}",
            f"Úhrada: {v.get('payment', '')}",
            f"Měna: {v.get('currency', '')}",
            f"Popis: {v.get('description', '')}",
            f"Vystavil: {v.get('issuer', '')}",
            f"Předkontace: {v.get('predkont', '')}",
            f"Středisko: {v.get('stredisko', '')}",
            f"Zakázka: {v.get('zakazka', '')}",
            f"Činnost: {v.get('cinnost', '')}",
            f"Celkem s DPH: {v.get('total', '')}",
            f"K úhradě: {v.get('to_pay', '')}",
            "",
            "Položky:",
            "Název / popis\tPoznámka\tPočet\tMJ\tCena / MJ\tDPH\tCelkem položka",
        ]
        for row in self.current_item_rows:
            s = item_summary(row)
            lines.append("\t".join([
                safe_str(s["popis"]),
                safe_str(s["poznamka"]),
                format_quantity(s["mnozstvi"]),
                safe_str(s["mj"]),
                format_amount(s["cena"]),
                safe_str(s["dph"]),
                format_amount(s["celkem"]),
            ]))
        return "\n".join(lines)

    def copy_all_invoice_text(self):
        text = self._invoice_clipboard_text()
        if not text:
            messagebox.showinfo(APP_TITLE, "Není vybraná žádná faktura.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set("Celá faktura včetně položek byla zkopírována do schránky.")

    def copy_selected_item_row(self, _event=None):
        sel = self.item_tree.selection()
        if not sel:
            return "break"
        values = self.item_tree.item(sel[0], "values")
        text = "\t".join(str(v) for v in values)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set("Vybraná položka byla zkopírována do schránky.")
        return "break"

    def load_item_table(self):
        if not self.current_dataset or not self.current_dataset.get("items_path"):
            return []
        path = str(self.current_dataset["items_path"])
        if path in self.item_cache:
            return self.item_cache[path]
        with MoneyTable(path) as table:
            rows = table.read_all()
        self.item_cache[path] = rows
        return rows

    def show_invoice_items(self, header):
        self.item_tree.delete(*self.item_tree.get_children())
        self.item_raw_tree.delete(*self.item_raw_tree.get_children())
        self.current_item_rows = []

        if not self.current_dataset or not self.current_dataset.get("items_path"):
            self.item_info_var.set("Pro tento rok nebyla nalezena tabulka položek.")
            return

        try:
            all_items = self.load_item_table()
        except Exception as exc:
            self.item_info_var.set(f"Chyba při načítání položek: {exc}")
            return

        related, relation = find_related_items(header, all_items)
        self.current_item_rows = related

        if not related:
            self.item_info_var.set(
                f"Položky nenalezeny — {relation}. Tabulka obsahuje {len(all_items)} záznamů."
            )
            return

        self.item_info_var.set(f"Položek: {len(related)} | vazba: {relation}")

        calculated = False
        for idx, row in enumerate(related):
            s = item_summary(row)
            total_text = format_amount(s["celkem"])
            if s["celkem_vypocteno"] and total_text:
                total_text += " *"
                calculated = True
            self.item_tree.insert(
                "", "end", iid=f"item_{idx}",
                values=(
                    safe_str(s["popis"]),
                    safe_str(s["poznamka"]),
                    format_quantity(s["mnozstvi"]),
                    safe_str(s["mj"]),
                    format_amount(s["cena"]),
                    safe_str(s["dph"]),
                    total_text,
                ),
            )

        if calculated:
            self.item_info_var.set(
                self.item_info_var.get() + " | * = DOPOČÍTÁNO (není to uložená celková cena položky)"
            )

    def on_item_select(self, _event=None):
        self.item_raw_tree.delete(*self.item_raw_tree.get_children())
        sel = self.item_tree.selection()
        if not sel:
            return
        idx = int(sel[0].split("_", 1)[1])
        if idx >= len(self.current_item_rows):
            return
        row = self.current_item_rows[idx]
        for key in sorted(row.keys(), key=str.lower):
            self.item_raw_tree.insert(
                "", "end", values=(key, safe_str(row.get(key)))
            )

    def show_invoice(self, row):
        self.clear_details()
        self.current_invoice_row = row
        kind = self.current_dataset["kind"] if self.current_dataset else "V"

        if kind == "V":
            partner = first_present(row, ["O_Nazev", "O_Jmeno", "O_NazevOb", "D_Nazev"])
            street = first_present(row, ["O_Ulice", "O_UliceOb", "D_Ulice"])
            city = first_present(row, ["O_Mesto", "O_MestoOb", "D_Mesto"])
            psc = first_present(row, ["O_Psc", "O_PscOb", "D_Psc"])
            ico = first_present(row, ["O_ICO", "D_ICO"])
            dic = first_present(row, ["O_DIC", "D_DIC"])
            date_main = first_present(row, ["Vystaveno", "DatVyst"])
        else:
            partner = first_present(row, ["D_Nazev", "D_Jmeno", "O_Nazev", "O_Jmeno"])
            street = first_present(row, ["D_Ulice", "O_Ulice"])
            city = first_present(row, ["D_Mesto", "O_Mesto"])
            psc = first_present(row, ["D_Psc", "O_Psc"])
            ico = first_present(row, ["D_ICO", "O_ICO"])
            dic = first_present(row, ["D_DIC", "O_DIC"])
            date_main = first_present(row, ["Prijato", "DatUcPr", "Vystaveno", "DatVyst"])

        currency = first_present(row, ["Mena"], "CZK")
        total_stored = first_present(row, ["CelkemSDPH"], "")
        to_pay_stored = first_present(row, ["KUhradeSDP", "Proplatit"], "")

        values = {
            "doc_number": first_present(row, ["Doklad", "EvCisDokl", "UDoklad"]),
            "ev_number": first_present(row, ["EvCisDokl"]),
            "date": date_main,
            "duzp": first_present(row, ["PlnenoDPH"]),
            "due": first_present(row, ["Splatno"]),
            "paid": first_present(row, ["Uhrazeno"]),
            "vs": first_present(row, ["VarSymbol"]),
            "ks": first_present(row, ["KonstSym"]),
            "payment": first_present(row, ["Uhrada"]),
            "partner": partner,
            "street": street,
            "city": city,
            "psc": psc,
            "ico": ico,
            "dic": dic,
            "email": first_present(row, ["EMail", "D_EMail"]),
            "phone": first_present(row, ["Telefon"]),
            "description": first_present(row, ["Popis"]),
            "currency": currency,
            "issuer": first_present(row, ["Vystavil"]),
            "predkont": first_present(row, ["PredKontac", "Predkontac", "PredKont", "Predkontace"]),
            "stredisko": first_present(row, ["Stredisko"]),
            "zakazka": first_present(row, ["Zakazka"]),
            "cinnost": first_present(row, ["Cinnost"]),
            "total": format_amount(total_stored) + (
                f" {currency}" if total_stored not in (None, "") else ""
            ),
            "to_pay": format_amount(to_pay_stored) + (
                f" {currency}" if to_pay_stored not in (None, "") else ""
            ),
        }

        for key, value in values.items():
            if key in self.invoice_vars:
                self.invoice_vars[key].set(safe_str(value))

        for key in sorted(row.keys(), key=str.lower):
            self.raw_tree.insert("", "end", values=(key, safe_str(row.get(key))))

        self.show_invoice_items(row)
        self.show_item_accounting_summary()

    def show_item_accounting_summary(self):
        if not hasattr(self, "item_accounting_tree"):
            return
        self.item_accounting_tree.delete(*self.item_accounting_tree.get_children())

        pred_header = ""
        if getattr(self, "current_invoice_row", None):
            pred_header = first_present(
                self.current_invoice_row,
                ["PredKontac", "Predkontac", "PredKont", "Predkontace"],
            )

        self.accounting_info_var.set(
            "Předkontace dokladu: " + (safe_str(pred_header) or "(není vyplněna / pole nebylo nalezeno)")
        )

        for i, row in enumerate(self.current_item_rows):
            self.item_accounting_tree.insert(
                "", "end", iid=f"accitem_{i}",
                values=(
                    safe_str(first_present(row, ["Popis", "Nazev"])),
                    safe_str(first_present(row, ["PredKontac", "Predkontac", "PredKont", "Predkontace"])),
                    safe_str(first_present(row, ["Stredisko"])),
                    safe_str(first_present(row, ["Zakazka"])),
                    safe_str(first_present(row, ["Cinnost"])),
                )
            )

    def _table_meta(self, path):
        key = str(path)
        if key in self.accounting_table_cache:
            return self.accounting_table_cache[key]

        try:
            with MoneyTable(path) as table:
                meta = {
                    "path": Path(path),
                    "columns": list(table.columns),
                    "count": int(table.obj.RecCount),
                }
        except Exception:
            meta = None

        self.accounting_table_cache[key] = meta
        return meta

    def scan_accounting_for_current_invoice(self):
        if not self.current_dataset or not getattr(self, "current_invoice_row", None):
            messagebox.showinfo(APP_TITLE, "Nejdřív vyber konkrétní fakturu.")
            return

        self.accounting_tree.delete(*self.accounting_tree.get_children())
        self.accounting_raw_tree.delete(*self.accounting_raw_tree.get_children())
        self.current_accounting_rows = []
        self.current_accounting_raw = []

        year_dir = Path(self.current_dataset["path"]).parent
        ucdenik = first_existing_file(year_dir, "UcDenik.DAT")
        if not ucdenik:
            self.accounting_status_var.set("UcDenik.DAT nebyl v tomto účetním roce nalezen.")
            return

        source_doc_values = []
        for key in ("Doklad", "EvCisDokl", "UDoklad"):
            v = norm(self.current_invoice_row.get(key))
            if v is not None:
                source_doc_values.append(v)

        expected_source = "FV" if self.current_dataset.get("kind") == "V" else "PF"
        source_vs = norm(self.current_invoice_row.get("VarSymbol"))

        primary = []
        related = []
        fallback = []

        self.accounting_status_var.set(
            f"Čtu UcDenik.DAT pro doklad "
            f"{safe_str(first_present(self.current_invoice_row, ['Doklad','EvCisDokl','UDoklad']))}…"
        )
        self.update_idletasks()

        try:
            with MoneyTable(ucdenik) as table:
                rows = table.read_all()
        except Exception as exc:
            self.accounting_status_var.set(f"UcDenik.DAT nelze načíst: {exc}")
            return

        for row in rows:
            displayed_doc = norm(first_by_alias(
                row, ["Doklad", "CisDokl", "CisloDokladu", "EvCisDokl", "UDoklad"]
            ))
            zdroj = safe_str(first_by_alias(row, ["Zdroj", "Source", "TypZdroj"])).strip().upper()

            # Nejspolehlivější varianta: účetní zápis je přímo zdrojová faktura.
            if displayed_doc is not None and displayed_doc in source_doc_values:
                if zdroj == expected_source:
                    primary.append((120, ucdenik.name, row, "faktura"))
                else:
                    related.append((100, ucdenik.name, row, "přímý doklad"))
                continue

            # Související řádek může odkazovat na fakturu jiným dokladovým polem.
            referenced = False
            for col, raw in row.items():
                c = str(col).lower()
                if "doklad" not in c:
                    continue
                rv = norm(raw)
                if rv is not None and rv in source_doc_values:
                    referenced = True
                    break
            if referenced:
                related.append((80, ucdenik.name, row, "související"))
                continue

            # Poslední fallback: VS / párovací symbol.
            if source_vs is not None:
                for col, raw in row.items():
                    c = str(col).lower()
                    if "symbol" in c or c in {"parsym", "varsym", "parsymbol", "varsymbol"}:
                        if norm(raw) == source_vs:
                            fallback.append((20, ucdenik.name, row, "VS fallback"))
                            break

        matches = primary + related
        if not matches:
            matches = fallback

        matches.sort(key=lambda x: (-x[0], safe_str(accounting_row_summary(x[2])["datum"])))
        self.current_accounting_rows = matches

        for idx, (mscore, table_name, row, relation) in enumerate(matches):
            s = accounting_row_summary(row)
            self.accounting_tree.insert(
                "", "end", iid=f"acc_{idx}",
                values=(
                    table_name,
                    relation,
                    safe_str(s["zdroj"]),
                    safe_str(s["datum"]),
                    safe_str(s["doklad"]),
                    safe_str(s["popis"]),
                    safe_str(s["md"]),
                    safe_str(s["dal"]),
                    safe_str(s["ucet"]),
                    safe_str(s["protiucet"]),
                    safe_str(s["strana"]),
                    format_amount(s["castka"]),
                    safe_str(s["predkont"]),
                ),
            )

        if primary:
            self.accounting_status_var.set(
                f"Nalezeno {len(primary)} účetních řádků zdrojové faktury "
                f"({expected_source}) a {len(related)} souvisejících řádků."
            )
        elif related:
            self.accounting_status_var.set(
                f"Nalezeno {len(related)} souvisejících řádků, ale žádný řádek "
                f"s kombinací Doklad + Zdroj={expected_source}. Ověř raw data."
            )
        elif fallback:
            self.accounting_status_var.set(
                f"Nalezena jen fallback vazba podle VS/párovacího symbolu: {len(fallback)} řádků. "
                "Tato část je experimentální."
            )
        else:
            self.accounting_status_var.set("V UcDenik.DAT nebyla nalezena vazba na tuto fakturu.")

    def on_accounting_select(self, _event=None):
        self.accounting_raw_tree.delete(*self.accounting_raw_tree.get_children())
        sel = self.accounting_tree.selection()
        if not sel:
            return
        idx = int(sel[0].split("_", 1)[1])
        if idx >= len(self.current_accounting_rows):
            return
        rec = self.current_accounting_rows[idx]
        table_name = rec[1]
        row = rec[2]
        self.accounting_raw_tree.insert(
            "", "end", values=("__TABULKA__", table_name)
        )
        for key in sorted(row.keys(), key=str.lower):
            self.accounting_raw_tree.insert(
                "", "end", values=(key, safe_str(row.get(key)))
            )

    def copy_selected_accounting_row(self, _event=None):
        sel = self.accounting_tree.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Nejdřív vyber účetní zápis.")
            return "break"

        values = self.accounting_tree.item(sel[0], "values")
        headers = [
            "Tabulka", "Vazba", "Zdroj", "Datum", "Doklad", "Popis", "Má dáti",
            "Dal", "Účet", "Protiúčet", "Strana", "Částka", "Předkontace"
        ]
        text = "\n".join(
            f"{h}: {v}" for h, v in zip(headers, values)
        )
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set("Vybraný účetní zápis byl zkopírován do schránky.")
        return "break"

    def copy_selected_accounting_raw(self, _event=None):
        sel = self.accounting_tree.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Nejdřív vyber účetní zápis.")
            return "break"

        idx = int(sel[0].split("_", 1)[1])
        if idx >= len(self.current_accounting_rows):
            return "break"

        rec = self.current_accounting_rows[idx]
        table_name = rec[1]
        row = rec[2]
        lines = [f"__TABULKA__: {table_name}"]
        for key in sorted(row.keys(), key=str.lower):
            lines.append(f"{key}: {safe_str(row.get(key))}")

        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.update()
        self.status_var.set("Raw data účetního zápisu byla zkopírována do schránky.")
        return "break"

    def _party_from_prefix(self, row, prefix):
        # *_Ob bývá v Money obchodní/fakturační adresa. U některých databází
        # jsou základní D_/O_ adresní hodnoty anonymizované, proto preferujeme
        # *_Ob a zároveň ignorujeme placeholder "anonymizovano".
        return {
            "name": first_real_value(
                row, [f"{prefix}_NazevOb", f"{prefix}_Nazev", f"{prefix}_Jmeno"]
            ),
            "street": first_real_value(
                row, [f"{prefix}_UliceOb", f"{prefix}_Ulice"]
            ),
            "city": first_real_value(
                row, [f"{prefix}_MestoOb", f"{prefix}_Mesto"]
            ),
            "psc": first_real_value(
                row, [f"{prefix}_PscOb", f"{prefix}_Psc"]
            ),
            "state": first_real_value(
                row, [f"{prefix}_StatOb", f"{prefix}_Stat"]
            ),
            "ico": first_real_value(row, [f"{prefix}_ICO"]),
            "dic": first_real_value(row, [f"{prefix}_DIC"]),
        }

    def _invoice_html(self):
        if not getattr(self, "current_invoice_row", None):
            raise RuntimeError("Není vybraná faktura.")

        row = self.current_invoice_row
        supplier = self._party_from_prefix(row, "D")
        customer = self._party_from_prefix(row, "O")
        v = {k: var.get() for k, var in self.invoice_vars.items()}

        def e(x):
            return html.escape(safe_str(x))

        item_rows = []
        for item in self.current_item_rows:
            s = item_summary(item)
            total = format_amount(s["celkem"])
            if s["celkem_vypocteno"] and total:
                total += " *"
            item_rows.append(
                "<tr>"
                f"<td>{e(s['popis'])}</td>"
                f"<td>{e(s['poznamka'])}</td>"
                f"<td class='num'>{e(format_quantity(s['mnozstvi']))}</td>"
                f"<td>{e(s['mj'])}</td>"
                f"<td class='num'>{e(format_amount(s['cena']))}</td>"
                f"<td class='num'>{e(s['dph'])}</td>"
                f"<td class='num'>{e(total)}</td>"
                "</tr>"
            )

        if not item_rows:
            item_rows.append("<tr><td colspan='7'><em>Položky nebyly nalezeny.</em></td></tr>")

        title = "Vydaná faktura" if self.current_dataset.get("kind") == "V" else "Přijatá faktura"
        return f"""<!doctype html>
<html lang="cs">
<head>
<meta charset="utf-8">
<title>{e(title)} {e(v.get('doc_number'))}</title>
<style>
@page {{ size: A4; margin: 14mm; }}
body {{ font-family: Arial, sans-serif; color:#111; font-size:12px; }}
h1 {{ font-size:24px; margin:0 0 4px 0; }}
.muted {{ color:#666; }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-top:18px; }}
.box {{ border:1px solid #bbb; padding:10px; min-height:110px; }}
.box h2 {{ font-size:13px; margin:0 0 8px 0; }}
.meta {{ margin-top:18px; display:grid; grid-template-columns:160px 1fr 160px 1fr; gap:4px 10px; }}
table {{ width:100%; border-collapse:collapse; margin-top:18px; }}
th, td {{ border:1px solid #bbb; padding:6px; vertical-align:top; }}
th {{ background:#eee; text-align:left; }}
.num {{ text-align:right; white-space:nowrap; }}
.total {{ margin-top:18px; text-align:right; font-size:18px; font-weight:bold; }}
.footer {{ margin-top:24px; border-top:1px solid #bbb; padding-top:8px; font-size:10px; color:#666; }}
</style>
</head>
<body>
<h1>{e(title)}</h1>
<div class="muted">Nouzová tisková sestava — Money S3 Browser {e(APP_TITLE.split()[-1])}</div>

<div class="grid">
<div class="box">
<h2>Dodavatel</h2>
<strong>{e(supplier['name'])}</strong><br>
{e(supplier['street'])}<br>
{e(supplier['psc'])} {e(supplier['city'])}<br>
{e(supplier['state'])}<br>
IČO: {e(supplier['ico'])}<br>
DIČ: {e(supplier['dic'])}
</div>
<div class="box">
<h2>Odběratel</h2>
<strong>{e(customer['name'])}</strong><br>
{e(customer['street'])}<br>
{e(customer['psc'])} {e(customer['city'])}<br>
{e(customer['state'])}<br>
IČO: {e(customer['ico'])}<br>
DIČ: {e(customer['dic'])}
</div>
</div>

<div class="meta">
<div>Číslo dokladu:</div><div><strong>{e(v.get('doc_number'))}</strong></div>
<div>Variabilní symbol:</div><div>{e(v.get('vs'))}</div>
<div>Datum vystavení:</div><div>{e(v.get('date'))}</div>
<div>Datum splatnosti:</div><div>{e(v.get('due'))}</div>
<div>DUZP:</div><div>{e(v.get('duzp'))}</div>
<div>Úhrada:</div><div>{e(v.get('payment'))}</div>
<div>Popis:</div><div>{e(v.get('description'))}</div>
<div>Předkontace:</div><div>{e(v.get('predkont'))}</div>
</div>

<table>
<thead><tr>
<th>Název / popis</th><th>Poznámka</th><th>Počet</th><th>MJ</th>
<th>Cena / MJ</th><th>DPH</th><th>Celkem</th>
</tr></thead>
<tbody>
{''.join(item_rows)}
</tbody>
</table>

<div class="total">Celkem s DPH: {e(v.get('total'))}</div>
<div class="total">K úhradě: {e(v.get('to_pay'))}</div>

<div class="footer">
Nouzový tisk vytvořený z lokálních databázových dat. Nejde o originální tiskovou sestavu Money S3.
Údaje před použitím jako účetní doklad zkontrolujte proti uloženým datům. Znak * u položky znamená dopočítanou hodnotu.
</div>
</body></html>"""

    def _write_invoice_html_temp(self):
        data = self._invoice_html()
        fd, name = tempfile.mkstemp(prefix="money_s3_invoice_", suffix=".html")
        os.close(fd)
        p = Path(name)
        p.write_text(data, encoding="utf-8")
        return p

    def print_invoice(self):
        try:
            p = self._write_invoice_html_temp()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        webbrowser.open(p.as_uri())
        messagebox.showinfo(
            "Nouzový tisk",
            "Faktura byla otevřena v prohlížeči.\n\n"
            "Použij Ctrl+P a vyber tiskárnu nebo „Microsoft Print to PDF“."
        )

    def _find_chromium_browser(self):
        candidates = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        for p in candidates:
            if Path(p).exists():
                return p
        return None

    def save_invoice_pdf(self):
        if not getattr(self, "current_invoice_row", None):
            messagebox.showinfo(APP_TITLE, "Nejdřív vyber fakturu.")
            return

        doc = safe_str(first_present(
            self.current_invoice_row, ["Doklad", "EvCisDokl", "UDoklad"]
        )) or "faktura"
        out = filedialog.asksaveasfilename(
            title="Uložit nouzovou fakturu do PDF",
            defaultextension=".pdf",
            initialfile=f"{doc}.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not out:
            return

        try:
            html_path = self._write_invoice_html_temp()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        browser = self._find_chromium_browser()
        if browser:
            try:
                proc = subprocess.run(
                    [
                        browser,
                        "--headless",
                        "--disable-gpu",
                        "--no-pdf-header-footer",
                        f"--print-to-pdf={out}",
                        html_path.as_uri(),
                    ],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=40,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if Path(out).exists() and Path(out).stat().st_size > 0:
                    self.status_var.set(f"Nouzové PDF uloženo: {out}")
                    messagebox.showinfo("PDF vytvořeno", f"PDF bylo uloženo:\n{out}")
                    return
            except Exception:
                pass

        # Fallback bez další Python knihovny.
        webbrowser.open(html_path.as_uri())
        messagebox.showwarning(
            "Automatické PDF se nepodařilo",
            "Nepodařilo se vytvořit PDF přes Edge/Chrome.\n\n"
            "Faktura byla otevřena v prohlížeči. Použij Ctrl+P → Uložit jako PDF."
        )

    def copy_diagnostics(self):
        lines = [
            "=== Money S3 Browser diagnostics ===",
            f"App: {APP_TITLE}",
            f"Python: {sys.version.replace(chr(10), ' ')}",
            f"Python bitness: {struct.calcsize('P') * 8}",
            f"OS: {platform.platform()}",
            f"Root: {self.root_var.get()}",
        ]

        if not self.current_dataset:
            lines.append("Dataset: žádný vybraný")
        else:
            ds = self.current_dataset
            lines += [
                "",
                "=== Dataset ===",
                f"Agenda: {ds.get('agenda')}",
                f"Rok dir: {ds.get('rokdir')}",
                f"Detekovaný rok: {ds.get('year')}",
                f"Typ: {'vydané faktury' if ds.get('kind') == 'V' else 'přijaté faktury'}",
                f"Header table: {ds.get('path')}",
                f"Items table: {ds.get('items_path')}",
                f"Počet faktur: {len(self.current_rows)}",
            ]

            year_dir = Path(ds["path"]).parent
            lines += ["", "=== DAT soubory v účetním roce ==="]
            try:
                for p in sorted(year_dir.glob("*.DAT"), key=lambda x: x.name.lower()):
                    st = p.stat()
                    lines.append(f"{p.name}\t{st.st_size} B")
            except Exception as exc:
                lines.append(f"CHYBA seznamu: {exc}")

        if getattr(self, "current_invoice_row", None):
            lines += ["", "=== Vybraná faktura - raw ==="]
            for key in sorted(self.current_invoice_row.keys(), key=str.lower):
                lines.append(f"{key}: {safe_str(self.current_invoice_row.get(key))}")

            lines += ["", f"=== Položky faktury ({len(self.current_item_rows)}) ==="]
            for i, row in enumerate(self.current_item_rows, 1):
                lines.append(f"--- položka {i} ---")
                for key in sorted(row.keys(), key=str.lower):
                    lines.append(f"{key}: {safe_str(row.get(key))}")

            # Bez nutnosti předchozího kliknutí zkusíme přidat přímé řádky UcDenik.
            if self.current_dataset:
                ucdenik = first_existing_file(Path(self.current_dataset["path"]).parent, "UcDenik.DAT")
                if ucdenik:
                    wanted = {
                        norm(self.current_invoice_row.get(k))
                        for k in ("Doklad", "EvCisDokl", "UDoklad")
                        if norm(self.current_invoice_row.get(k)) is not None
                    }
                    try:
                        with MoneyTable(ucdenik) as table:
                            rows = table.read_all()
                        matched = []
                        for row in rows:
                            if norm(first_by_alias(row, ["Doklad", "CisDokl", "CisloDokladu", "EvCisDokl", "UDoklad"])) in wanted:
                                matched.append(row)
                        lines += ["", f"=== Přímé UcDenik řádky ({len(matched)}) ==="]
                        for i, row in enumerate(matched, 1):
                            lines.append(f"--- UcDenik {i} ---")
                            for key in sorted(row.keys(), key=str.lower):
                                lines.append(f"{key}: {safe_str(row.get(key))}")
                    except Exception as exc:
                        lines += ["", f"UcDenik diagnostika chyba: {exc}"]

        text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set(
            "Diagnostika byla zkopírována do schránky. Můžeš ji vložit do issue nebo chatu."
        )

    def _render_help(self):
        t = self.help_text
        t.configure(state="normal")
        t.delete("1.0", "end")
        t.tag_configure("title", font=("Segoe UI", 16, "bold"), spacing3=8)
        t.tag_configure("h1", font=("Segoe UI", 11, "bold"), spacing1=12, spacing3=4)
        t.tag_configure("warning", font=("Segoe UI", 10, "bold"), foreground="red")
        t.tag_configure("body", font=("Segoe UI", 10), spacing3=3)
        t.tag_configure("code", font=("Consolas", 10), lmargin1=18, lmargin2=18, spacing3=6)
        t.tag_configure("bullet", font=("Segoe UI", 10), lmargin1=18, lmargin2=32)

        def add(text, tag="body"):
            t.insert("end", text, tag)

        add("Money S3 Browser v0.11\n", "title")
        add(
            "UPOZORNĚNÍ: Neoficiální read-only nástroj. Používejte na vlastní nebezpečí "
            "a vždy pracujte pouze s kopií / zálohou dat.\n",
            "warning",
        )

        add("\nCo tento program je\n", "h1")
        add(
            "Prohlížeč lokálních dat Money S3 pro situace, kdy má uživatel svá data "
            "na vlastním počítači, ale standardní aplikaci nemůže běžně použít.\n"
        )

        add("\nBezpečnost\n", "h1")
        add("• Program je navržen jako read-only.\n", "bullet")
        add("• Nevolá zapisovací operace Append, Update ani Delete.\n", "bullet")
        add("• Nikdy nepracujte s jedinou existující kopií účetních dat.\n", "bullet")
        add("• Projekt neobsahuje MON2KDBE.DLL ani jiné proprietární součásti Money S3.\n", "bullet")

        add("\nJak udělat zálohu\n", "h1")
        add("Pokud Money S3 funguje, použijte jeho funkci Správa dat → Vytvoření záložní kopie.\n")
        add("Pokud Money S3 nelze normálně otevřít:\n")
        add("1. Money S3 úplně ukončete.\n", "bullet")
        add("2. Zkopírujte celý datový adresář Money S3 na jiné místo.\n", "bullet")
        add("3. Browser otevřete nad touto kopií.\n", "bullet")
        add("\nBěžné umístění:\n")
        add("C:\\Users\\Public\\Documents\\Solitea\\Money S3\n", "code")
        add("Starší instalace:\n")
        add("C:\\Users\\Public\\Documents\\CIGLER SOFTWARE\\Money S3\n", "code")

        add("\nHvězdička u ceny položky\n", "h1")
        add(
            "* znamená DOPOČÍTANOU hodnotu: celková cena položky nebyla nalezena jako "
            "samostatně uložené pole a Browser ji spočítal jako Počet × Cena/MJ. "
            "Celková cena celé faktury se přednostně čte z CelkemSDPH.\n"
        )

        add("\nÚčetnictví / MD-Dal\n", "h1")
        add(
            "Zobrazení účetních vazeb je EXPERIMENTÁLNÍ. Browser čte UcDenik.DAT a u známé "
            "struktury umí pole UcMD a UcD zobrazit jako Má dáti / Dal. Různé verze Money S3 "
            "mohou mít jiné struktury nebo pomocné zápisy. Účetní výstupy proto vždy ověřte "
            "proti původním datům a nepoužívejte je bez kontroly jako jediný podklad.\n"
        )

        add("\nProč tento nástroj vznikl — poznámka autora\n", "h1")
        add(
            "Nástroj vznikl poté, co autor po letech používání bezplatné START verze narazil "
            "na licenční obrazovku a nemohl se běžným způsobem dostat ke svým historickým fakturám. "
            "Autor tuto změnu vnímal jako předem neoznámené uzamčení přístupu k vlastním datům. "
            "Cílem projektu je data přečíst, exportovat a nouzově vytisknout — nikoliv patchovat "
            "Money S3 nebo obcházet jeho aktivaci.\n\n"
            "Osobní názor autora: výkon desktopové aplikace Money S3 považuje v roce 2026 "
            "na svém výkonném hardwaru za nepřijatelně pomalý; subjektivně jej popisuje jako "
            "„slideshow“ a považuje za ostudu platit jen proto, aby znovu získal pohodlný přístup "
            "ke svým historickým datům. Toto je osobní zkušenost a hodnocení autora, nikoliv "
            "obecné tvrzení o výkonu produktu u všech uživatelů.\n"
        )

        add("\nO projektu\n", "h1")
        add(
            "Projekt není produktem společnosti Seyfor a není se společností Seyfor nijak spojen. "
            "Money S3 je uvedeno pouze pro označení kompatibility. Browser používá MON2KDBE.DLL, "
            "kterou již má uživatel ve své vlastní instalaci; knihovna není součástí projektu.\n"
        )
        t.configure(state="disabled")

    def _help_text_content(self):
        return """Money S3 Browser v0.11

CO TENTO PROGRAM JE
Neoficiální read-only prohlížeč lokálních dat Money S3. Vznikl proto, aby bylo možné zobrazit vlastní faktury a další lokálně uložená data i v situaci, kdy standardní Money S3 není možné běžně použít (např. licenční, kompatibilitní nebo servisní problém).

BEZPEČNOST A DISCLAIMER
Tento program není produktem společnosti Seyfor a není se Seyforem nijak spojen.
Použití je na vlastní nebezpečí. Účetní data jsou důležitá a jejich poškození může mít závažné následky.
Program je navržen jako read-only: používá COM objekt mon2kdbe.BFTable pouze pro otevření a čtení tabulek a nevolá zapisovací operace Append, Update ani Delete.
Přesto vždy pracujte s kopií dat, nikdy ne s jediným originálem.

JAK UDĚLAT ZÁLOHU
Varianta A – Money S3 ještě funguje:
V Money použijte Správa dat / Vytvoření záložní kopie (nebo Vytvoření záložní kopie všech agend). Pokud potřebujete úplnou zálohu, zahrňte společná data a podle potřeby dokumenty.

Varianta B – Money S3 nelze normálně otevřít:
1. Money S3 úplně ukončete.
2. Zkopírujte CELÝ datový adresář Money S3 na jiné místo a Money S3 Browser otevřete nad touto kopií.

Běžné umístění:
C:\\Users\\Public\\Documents\\Solitea\\Money S3

Starší instalace mohou používat:
C:\\Users\\Public\\Documents\\CIGLER SOFTWARE\\Money S3

HVĚZDIČKA U CENY POLOŽKY
Znak * znamená, že celková cena položky nebyla nalezena jako samostatně uložená hodnota a Browser ji dopočítal jako Počet × Cena/MJ.
Hvězdička tedy znamená DOPOČÍTANOU hodnotu, nikoliv hodnotu přímo uloženou v databázi.
Celková cena celé faktury je naopak přednostně čtena z uloženého pole CelkemSDPH.

ÚČETNICTVÍ
Zobrazení účetních vazeb je EXPERIMENTÁLNÍ. Browser primárně čte UcDenik.DAT.
U známé struktury mapuje UcMD = Má dáti a UcD = Dal. Různé verze Money S3 mohou mít odlišnou strukturu nebo pomocné zápisy. Výsledek vždy ověřte proti původním datům.

PROČ NÁSTROJ VZNIKL — POZNÁMKA AUTORA
Autor po letech používání bezplatné START verze narazil na licenční obrazovku a nemohl se běžným způsobem dostat ke svým historickým fakturám. Tuto změnu vnímal jako předem neoznámené uzamčení přístupu k vlastním datům. Projekt vznikl kvůli čtení, exportu a nouzovému tisku vlastních dat, nikoliv kvůli patchování nebo obcházení aktivace.

Osobní názor autora: výkon desktopové aplikace Money S3 považuje v roce 2026 na svém výkonném hardwaru za nepřijatelně pomalý; subjektivně jej popisuje jako „slideshow“ a považuje za ostudu platit jen proto, aby znovu získal pohodlný přístup ke svým historickým datům. Jde o osobní zkušenost autora.

TECHNICKÁ DATA
Záložka Technická data zobrazuje všechna pole hlavičky faktury a vybrané položky. Je určená hlavně k ověření mapování a podpoře dalších verzí programu.
"""

    def show_backup_help(self):
        win = tk.Toplevel(self)
        win.title("Jak bezpečně udělat zálohu dat")
        win.geometry("760x520")
        win.minsize(650, 430)
        win.transient(self)

        outer = ttk.Frame(win, padding=12)
        outer.pack(fill="both", expand=True)

        t = tk.Text(outer, wrap="word", padx=10, pady=10, relief="flat")
        sb = ttk.Scrollbar(outer, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        t.tag_configure("title", font=("Segoe UI", 15, "bold"), spacing3=10)
        t.tag_configure("warning", font=("Segoe UI", 10, "bold"), foreground="red", spacing3=10)
        t.tag_configure("h1", font=("Segoe UI", 11, "bold"), spacing1=10, spacing3=4)
        t.tag_configure("body", font=("Segoe UI", 10), spacing3=3)
        t.tag_configure("step", font=("Segoe UI", 10), lmargin1=18, lmargin2=32, spacing3=3)
        t.tag_configure("code", font=("Consolas", 10), lmargin1=22, lmargin2=22, spacing1=4, spacing3=8)

        t.insert("end", "Bezpečná záloha před použitím Browseru\n", "title")
        t.insert(
            "end",
            "Nepracujte s jedinou existující kopií účetních dat. Money S3 Browser je "
            "navržen jako read-only, ale správná záloha je stále základ.\n",
            "warning",
        )

        t.insert("end", "\nVarianta A — Money S3 se ještě normálně otevře\n", "h1")
        t.insert(
            "end",
            "Použijte přímo funkci Money S3 pro vytvoření záložní kopie. "
            "Pokud chcete úplnou zálohu, zahrňte všechny agendy, společná data "
            "a podle potřeby dokumenty.\n",
            "body",
        )

        t.insert("end", "\nVarianta B — Money S3 blokuje licenční obrazovka nebo nejde otevřít\n", "h1")
        t.insert("end", "1. Money S3 úplně ukončete.\n", "step")
        t.insert(
            "end",
            "2. Zkopírujte CELÝ datový adresář Money S3 na jiné místo. "
            "Nekopírujte jen VFaktury.DAT — vazby jsou rozdělené do více tabulek.\n",
            "step",
        )
        t.insert(
            "end",
            "3. V Money S3 Browseru zvolte jako Kořen dat právě tuto kopii.\n",
            "step",
        )

        t.insert("end", "\nBěžné umístění dat:\n", "body")
        t.insert("end", "C:\\Users\\Public\\Documents\\Solitea\\Money S3\n", "code")
        t.insert("end", "Starší instalace mohou používat:\n", "body")
        t.insert("end", "C:\\Users\\Public\\Documents\\CIGLER SOFTWARE\\Money S3\n", "code")

        t.insert("end", "\nDoporučený příklad kopie:\n", "body")
        t.insert("end", "C:\\zalohovani\\Money S3\n", "code")

        t.insert("end", "\nCo nekopírovat samostatně\n", "h1")
        t.insert(
            "end",
            "Jednotlivé .DAT soubory nejsou úplná záloha. Faktury, položky, účetní deník, "
            "číselníky a vazby jsou v několika souborech. Proto kopírujte celý datový strom.\n",
            "body",
        )

        t.configure(state="disabled")

        buttons = ttk.Frame(win, padding=(12, 0, 12, 12))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Zavřít", command=win.destroy).pack(side="right")
        ttk.Button(
            buttons,
            text="Zkopírovat návod",
            command=lambda: (
                self.clipboard_clear(),
                self.clipboard_append(
                    "1. Úplně ukončete Money S3.\n"
                    "2. Zkopírujte celý adresář C:\\Users\\Public\\Documents\\Solitea\\Money S3 "
                    "např. do C:\\zalohovani\\Money S3.\n"
                    "3. Money S3 Browser otevřete nad touto kopií."
                ),
                self.update(),
            ),
        ).pack(side="right", padx=(0, 6))

    def copy_help_text(self):
        text = self._help_text_content()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status_var.set("Text nápovědy byl zkopírován do schránky.")

    def export_current_dataset(self):
        if not self.current_rows or not self.current_dataset:
            messagebox.showinfo(APP_TITLE, "Nejdřív vyber databázi faktur.")
            return

        default_name = (
            f'{self.current_dataset["agenda"]}_'
            f'{self.current_dataset.get("year") or self.current_dataset["rokdir"]}_'
            f'{"VFaktury" if self.current_dataset["kind"] == "V" else "PFaktury"}.csv'
        ).replace(" ", "_")

        out = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV", "*.csv"), ("Všechny soubory", "*.*")],
        )
        if not out:
            return

        cols = []
        seen = set()
        for row in self.current_rows:
            for col in row.keys():
                if col not in seen:
                    seen.add(col)
                    cols.append(col)

        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=cols, delimiter=";")
            writer.writeheader()
            for row in self.current_rows:
                writer.writerow({k: safe_str(row.get(k)) for k in cols})

        self.status_var.set(f"Exportováno: {out}")


if __name__ == "__main__":
    app = MoneyBrowserApp()
    app.mainloop()
