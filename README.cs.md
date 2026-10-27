# SHEETDIFF

Porovná dva sešity Excel (`.xlsx`, `.xlsm`) a ukáže, co se opravdu změnilo. SHEETDIFF čte přímo
balíček souboru, takže najde i změny, které Excel (Spreadsheet Compare) nevidí: texty v textových
polích a tvarech, obrázky, grafy, komentáře a pravidla podmíněného formátování.

[English version](README.md)

![Report SHEETDIFF](docs/screenshot_report.png)

## Proč

Když sešit prochází více verzemi (`_3`, `_4`, úprava od kolegy), potřebujete vědět, jestli zůstala
čísla a vzorce stejná a co dalšího se pohnulo. SHEETDIFF dá nahoře reportu jednu odpověď, třeba
„Změněny jen texty buněk, vzorce a čísla beze změny“, a pod ní vypíše všechny změny.

Každá část souboru, která se liší, je v reportu podchycená. Pokud rozdíl nevyloží žádná vrstva,
objeví se aspoň v sekci Balíček jako „nepřiřazeno“, takže nic nezůstane skryté.

## Co se porovnává

| Vrstva | Obsah |
|---|---|
| Struktura | listy (přidané, odebrané, přejmenované, pořadí, viditelnost), pojmenované oblasti, externí odkazy, **vložené a smazané řádky a sloupce** |
| Buňky | hodnoty, texty, typ hodnoty, vzorce (i sdílené), vypočtené hodnoty |
| Formát | číselný formát, písmo, barva, výplň, ohraničení, zarovnání, zámek (seskupeno do oblastí), sloučené buňky, výška, šířka, skrytí a seskupení řádků a sloupců |
| Pravidla | podmíněné formátování (i datové pruhy a škály), ověření dat, hypertextové odkazy |
| Tvary | textová pole a tvary (text, ukotvení, poloha, formát), obrázky (podle obsahu), grafy (typ, název, osy, oblasti řad) |
| Komentáře | text a autor, klasické i vláknové |
| VBA | změněno nebo ne (otisk projektu maker) |
| Balíček | každá část ZIP balíčku, která se liší, a která vrstva ji vyložila |

Technický šum se uvádí zvlášť a nepočítá se jako změna: komprese ZIP, čas uložení, autor, pořadí
výpočtu, přečíslované sdílené řetězce a styly, aktivní buňka a lupa.

Vložené řádky SHEETDIFF najde podle obsahu a ukáže jednu změnu „Vložen řádek“ místo stovek
změněných buněk. Odkazy ve vzorcích, pojmenované oblasti, řady grafů a oblasti pravidel starého
souboru posune stejně jako Excel, takže se nehlásí jako změněné. Totéž platí pro odkazy na
přejmenovaný list.

## Stažení

Windows: `SHEETDIFF_vX.Y.Z.exe` v [Releases](https://github.com/mrSpringpeace/SHEETDIFF/releases/latest).
Jeden soubor, bez instalace a bez Pythonu.

Exe není digitálně podepsané, Windows SmartScreen proto může při prvním spuštění varovat
(„Další informace“, potom „Přesto spustit“). Pokud ho antivir blokuje, spusťte SHEETDIFF ze
zdrojového kódu (viz níže).

## Použití

### Okno

Spusťte `SHEETDIFF.exe` (nebo `SHEETDIFF.bat` při běhu ze zdrojů). Starý a nový sešit vyberte
tlačítkem „Vybrat…“ nebo je **přetáhněte myší do okna**:

* jeden soubor puštěný na pole se vloží do toho pole,
* dva soubory najednou se seřadí podle data uložení, starší bude A.

Dva soubory přetažené přímo na exe (nebo na zástupce) se porovnají hned. Report se otevře
v prohlížeči. Volba „Také Excel report“ přidá seznam změn jako `.xlsx` k filtrování.

![Okno SHEETDIFF](docs/screenshot_window.png)

Jazyk (čeština nebo angličtina) se přepíná vpravo nahoře a platí pro okno i reporty. Volba se
pamatuje v `%APPDATA%\SHEETDIFF\settings.json`.

### Příkazová řádka

```bat
python -m sheetdiff stary.xlsx novy.xlsx -o report.html --excel --otevrit
SHEETDIFF.exe --cli stary.xlsx novy.xlsx -o report.html
```

| Volba | Význam |
|---|---|
| `-o CESTA` | cesta HTML reportu (výchozí `%TEMP%\SHEETDIFF`) |
| `--excel` | vytvořit i Excel report vedle HTML |
| `--jazyk cs\|en` | jazyk reportu |
| `--bez VRSTVA` | vypnout vrstvu: `struktura`, `bunky`, `format`, `pravidla`, `tvary`, `komentare`, `balicek` |
| `-t TOL` | relativní tolerance čísel (výchozí `1e-9`) |
| `--otevrit` | otevřít report v prohlížeči |

Návratový kód: `0` bez rozdílů, `1` jen technický šum, `2` obsahové rozdíly, `3` chyba.
Exe je okenní aplikace, pro návratový kód ve skriptu použijte `start /wait SHEETDIFF.exe --cli …`.

## Jak to funguje

Soubor `.xlsx` je ZIP balíček XML částí. SHEETDIFF otevře oba balíčky, přečte části přes `lxml`
a porovná je po vrstvách. Excel není potřeba a soubory se jen čtou, nikdy se do nich nezapisuje.
Dva sešity po 4 MB s obrázky se porovnají zhruba za sekundu.

## Běh ze zdrojů

Python 3.11 nebo novější s `lxml` a `openpyxl`:

```bat
python -m pip install lxml openpyxl
python -m sheetdiff stary.xlsx novy.xlsx --otevrit
```

`SHEETDIFF.bat` spustí okno přes Python z PATH.

## Ukázkové soubory

`examples/office_budget_v1.xlsx` a `_v2.xlsx` jsou vymyšlený rozpočet kanceláře s jednou změnou
každého druhu. Vytváří je `tools/demo.py`.

## Sestavení exe

`build_exe.bat` sestaví `dist\SHEETDIFF_vX.Y.Z.exe` pomocí [Nuitka](https://nuitka.net) (onefile,
kompilátor Zig se stáhne sám). `tools/make_screenshot.py` vyrobí screenshoty z ukázkových souborů.

## Verze

Sémantické verzování, verze je v `sheetdiff/__init__.py`, vydání mají tag `vX.Y.Z`.
Přehled změn je v [CHANGELOG.md](CHANGELOG.md), `SHEETDIFF --version` verzi vypíše.

## Testy

```bat
python tests\acceptance.py
```

Testy si vyrobí vlastní sešity a kontrolují každý druh změny, šum, Excel report a to, že anglický
report neobsahuje české popisky.

## Překlady

Texty jsou v kódu česky a obalené `_()`, anglický slovník je v `sheetdiff/i18n.py`.

## Omezení

* Ovládací prvky formuláře, sparkliny, průřezy, tabulky, kontingenční tabulky a barvy motivu se
  zatím nevykládají. Jejich změny jsou v sekci Balíček jako nepřiřazené.
* Vložené řádky a sloupce se hledají podle obsahu. Při současném vložení řádků i sloupců
  s obsahem se buňky párují podle adresy.
* VBA se porovnává jako celek (změněno, nebo ne), ne po modulech.
* Sešity chráněné heslem a starý formát `.xls` nejsou podporované.

## Licence

MIT, viz [LICENSE](LICENSE).
