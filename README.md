# Money S3 Browser

Neoficiální **read-only** prohlížeč lokálních dat Money S3 pro Windows.

> **UPOZORNĚNÍ:** Používejte na vlastní nebezpečí a pracujte pouze s kopií / zálohou účetních dat.

## Účel

Projekt je určen k read-only prohlížení, exportu a nouzovému tisku vlastních lokálních dat Money S3.

Není produktem společnosti Seyfor a není se společností Seyfor nijak spojen. Název Money S3 je použit pouze pro označení kompatibility.

Projekt **neobsahuje ani nedistribuuje `MON2KDBE.DLL`**. Používá COM komponentu, kterou již má uživatel nainstalovanou ve své vlastní instalaci Money S3.

## Důvod
Proč toto vlastně vzniklo: Měl jsem roky bezplatnou START verzi pr pár dokladů ročně, po změně na 3 měsíce, jsem se bez varování nedostal ke svým historickým dokladům, tak jsem za odpoledne napsal vlastní prohlížeč. A mimochodem má asi 100x rychlejší reakce na klik než ta nativní obludnost.

## Bezpečnost

Program nevolá zapisovací operace `Append`, `Update` ani `Delete`. Přesto vždy pracujte s kopií dat.

Pokud Money S3 funguje, použijte jeho standardní zálohu. Pokud jej nelze normálně otevřít:

1. Money S3 úplně ukončete.
2. Zkopírujte **celý** datový adresář.
3. Browser spusťte nad kopií.

Typická cesta:

```text
C:\Users\Public\Documents\Solitea\Money S3
```

Například kopie:

```text
C:\zalohovani\Money S3
```

## Požadavky

- Windows
- 32bit Python
- `pywin32`
- vlastní nainstalovaná instance Money S3 s registrovaným `mon2kdbe.BFTable`

```cmd
py -3.13-32 -m pip install pywin32
py -3.13-32 MoneyS3_Browser_v09.py
```

## Funkce

- vydané a přijaté faktury;
- položky faktur;
- poznámky, množství, ceny;
- uložená celková cena faktury;
- CSV export;
- kopírování dat do schránky;
- diagnostický dump aktuálního kontextu;
- předkontace, střediska, zakázky a činnosti;
- experimentální účetní vazba přes `UcDenik.DAT`;
- mapování známých polí `UcMD` = Má dáti a `UcD` = Dal;
- nouzový tisk faktury;
- nouzový export faktury do PDF přes lokálně nainstalovaný Edge/Chrome.

## Experimentální účetní data

Účetní část není náhradou účetního systému. Struktura interních `.DAT` tabulek se může mezi verzemi Money S3 měnit.

Browser se snaží:
1. najít zdrojový účetní zápis podle čísla dokladu a typu zdroje (`FV` / `PF`);
2. oddělit související účetní pohyby;
3. až jako poslední možnost použít vazbu přes variabilní/párovací symbol.

Výsledek účetního mapování vždy ověřte.

## Diagnostika

Tlačítko **Diagnostika → schránka** zkopíruje:

- verzi Browseru a Pythonu;
- cestu a identifikaci agendy/roku;
- seznam `.DAT` souborů v účetním roce;
- raw hlavičku aktuální faktury;
- raw položky;
- přímo odpovídající řádky `UcDenik.DAT`.

To je vhodný obsah pro GitHub issue bez nutnosti ručně opisovat strukturu tabulek. Před zveřejněním issue zkontrolujte, zda diagnostika neobsahuje osobní nebo účetní údaje, které nechcete publikovat.

## Licence

Licence repozitáře se vztahuje pouze na zdrojový kód tohoto projektu. Nevztahuje se na Money S3, `MON2KDBE.DLL` ani jiné komponenty třetích stran.
