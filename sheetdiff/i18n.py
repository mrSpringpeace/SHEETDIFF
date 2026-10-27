"""Jazyk výstupu. Český text je klíčem, anglický překlad je ve slovníku EN.

Použití: _("Vložen řádek"), _("{n} buněk", n=5). Neznámý text se vrátí beze změny.
"""
from __future__ import annotations

LANGS = ("cs", "en")
_lang = "cs"


def set_lang(lang: str):
    global _lang
    _lang = lang if lang in LANGS else "cs"


def get_lang() -> str:
    return _lang


def _(text: str, **kw) -> str:
    if _lang == "en":
        text = EN.get(text, text)
    return text.format(**kw) if kw else text


def system_lang() -> str:
    """Čeština, je-li Windows v češtině nebo slovenštině, jinak angličtina."""
    try:
        import ctypes
        prim = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF
        return "cs" if prim in (0x05, 0x1B) else "en"
    except Exception:
        return "cs"


EN: dict[str, str] = {
    # ------------------------------------------------ příkazová řádka
    "Porovná dva sešity Excel (.xlsx, .xlsm) a vytvoří HTML report.":
        "Compares two Excel workbooks (.xlsx, .xlsm) and creates an HTML report.",
    "STARY": "OLD", "NOVY": "NEW", "VRSTVA": "LAYER",
    "starší verze (A)": "older version (A)",
    "novější verze (B)": "newer version (B)",
    "cesta HTML reportu (výchozí: %%TEMP%%\\SHEETDIFF)": "HTML report path (default: %%TEMP%%\\SHEETDIFF)",
    "relativní tolerance čísel (výchozí 1e-9)": "relative number tolerance (default 1e-9)",
    "vypnout vrstvu: {x} (lze opakovat)": "skip a layer: {x} (repeatable)",
    "vytvořit i Excel report (stejný název, přípona .xlsx)":
        "also create an Excel report (same name, .xlsx extension)",
    "jazyk reportu: cs nebo en": "report language: cs or en",
    "otevřít report v prohlížeči": "open the report in the browser",
    "Soubor neexistuje: {f}": "File does not exist: {f}",
    "Chyba: {e}": "Error: {e}",
    "Šum: ": "Noise: ",
    "Report: {p}": "Report: {p}",
    "Excel:  {p}": "Excel:  {p}",

    # ------------------------------------------------ verdikt
    "Bez rozdílů – soubory jsou totožné.": "No differences – the files are identical.",
    "Bez rozdílů.": "No differences.",
    "Jen technické rozdíly (šum) – obsah sešitu je stejný.":
        "Technical differences only (noise) – the workbook content is the same.",
    "Změněny jen texty buněk – vzorce a čísla beze změny.":
        "Only cell texts changed – formulas and numbers unchanged.",
    "Změněny jen texty v textových polích a tvarech – buňky beze změny.":
        "Only texts in text boxes and shapes changed – cells unchanged.",
    "Změněny jen texty (v buňkách a v tvarech) – vzorce a čísla beze změny.":
        "Only texts changed (in cells and shapes) – formulas and numbers unchanged.",
    "Změněna jen struktura (listy, vložené či smazané řádky a sloupce) – ostatní obsah beze změny.":
        "Only the structure changed (sheets, inserted or deleted rows and columns) – other content unchanged.",
    "Změněn jen formát – hodnoty, texty a vzorce beze změny.":
        "Only formatting changed – values, texts and formulas unchanged.",
    "Obsah buněk beze změny – liší se jen: {x}.": "Cell content unchanged – only these differ: {x}.",
    "Sešity se liší obsahem.": "The workbooks differ in content.",
    "Vzorce beze změny.": "Formulas unchanged.",
    "Čísla a vypočtené hodnoty beze změny.": "Numbers and calculated values unchanged.",
    "Buňky beze změny.": "Cells unchanged.",
    "Formát beze změny.": "Formatting unchanged.",
    "Mimo vložené a smazané řádky a sloupce:": "Outside inserted and deleted rows and columns:",
    "Pozor: některé části souboru se liší a žádná vrstva je zatím neumí vyložit – viz sekce Balíček.":
        "Note: some parts of the file differ and no layer can interpret them yet – see the Package section.",

    # ------------------------------------------------ kategorie
    "Struktura (listy, řádky, sloupce)": "Structure (sheets, rows, columns)",
    "Texty buněk": "Cell texts",
    "Čísla a konstanty": "Numbers and constants",
    "Typ hodnoty": "Value type",
    "Vzorce": "Formulas",
    "Vypočtené hodnoty": "Calculated values",
    "Formát buněk": "Cell formatting",
    "Podmíněné formátování a ověření dat": "Conditional formatting and data validation",
    "Hypertextové odkazy": "Hyperlinks",
    "Komentáře": "Comments",
    "Texty tvarů a textových polí": "Shape and text box texts",
    "Tvary (poloha, formát)": "Shapes (position, format)",
    "Obrázky": "Images",
    "Grafy": "Charts",
    "Makra VBA": "VBA macros",
    "Nepřiřazené změny": "Unassigned changes",
    "Data grafu (šum)": "Chart data (noise)",

    # ------------------------------------------------ vrstvy a stavy
    "Struktura": "Structure", "Buňky": "Cells", "Formát": "Format", "Pravidla": "Rules",
    "Odkazy": "Links", "Tvary": "Shapes", "Balíček": "Package", "VBA": "VBA",
    "vysvětleno": "explained", "šum": "noise", "nepřiřazeno": "unassigned",

    # ------------------------------------------------ struktura
    "List přejmenován": "Sheet renamed",
    "obsah listu beze změny": "sheet content unchanged",
    "List pravděpodobně přejmenován": "Sheet probably renamed",
    "stejné pořadí listu; obsah porovnán níže": "same sheet position; content compared below",
    "List odebrán": "Sheet removed",
    "List přidán": "Sheet added",
    "pořadí listů": "sheet order",
    "Pořadí listů změněno": "Sheet order changed",
    "Viditelnost listu změněna": "Sheet visibility changed",
    "viditelný": "visible", "skrytý": "hidden", "velmi skrytý": "very hidden",
    "Pojmenovaná oblast přidána": "Named range added",
    "Pojmenovaná oblast odebrána": "Named range removed",
    "Pojmenovaná oblast změněna": "Named range changed",
    "externí odkaz": "external link",
    "Externí odkaz odebrán": "External link removed",
    "Externí odkaz přidán": "External link added",
    "Vložen řádek": "Row inserted", "Vloženy řádky ({n})": "Rows inserted ({n})",
    "Smazán řádek": "Row deleted", "Smazány řádky ({n})": "Rows deleted ({n})",
    "Vložen sloupec": "Column inserted", "Vloženy sloupce ({n})": "Columns inserted ({n})",
    "Smazán sloupec": "Column deleted", "Smazány sloupce ({n})": "Columns deleted ({n})",
    "… a dalších {n}": "… and {n} more",
    "{n} buněk s obsahem": "cells with content: {n}",
    "prázdné": "empty",
    " (adresa ve starém souboru)": " (address in the old file)",

    # ------------------------------------------------ buňky
    "NEPRAVDA": "FALSE", "PRAVDA": "TRUE",
    "text": "text", "číslo": "number", "logická hodnota": "boolean", "chyba": "error", "datum": "date",
    "Číselné rozdíly pod tolerancí": "Numeric differences below tolerance",
    "{b} (dříve {a})": "{b} (was {a})",
    "Buňka vyplněna": "Cell filled",
    "Buňka vymazána": "Cell cleared",
    "hodnota: {v}": "value: {v}",
    "hodnota: {a} → {b}": "value: {a} → {b}",
    "hodnota beze změny: {v}": "value unchanged: {v}",
    "Vzorec změněn": "Formula changed",
    "Konstanta nahrazena vzorcem": "Constant replaced by formula",
    "Vzorec nahrazen konstantou": "Formula replaced by constant",
    "Vypočtená hodnota změněna": "Calculated value changed",
    "vzorec beze změny: ={f}": "formula unchanged: ={f}",
    "Typ hodnoty změněn": "Value type changed",
    "Text změněn": "Text changed",
    "Hodnota změněna": "Value changed",

    # ------------------------------------------------ formát
    "Formát změněn": "Format changed",
    "{n} buněk": "{n} cells",
    "výchozí": "default",
    "skrytí": "visibility", "seskupení": "grouping", "úroveň {n}": "level {n}",
    "výška": "height", "šířka": "width",
    "Sloučení buněk zrušeno": "Cells unmerged",
    "Buňky sloučeny": "Cells merged",
    "sloučeno": "merged",
    "Řádky skryty": "Rows hidden", "Řádky zobrazeny": "Rows unhidden",
    "Řádky: změna ({what})": "Rows: changed ({what})",
    "Sloupce skryty": "Columns hidden", "Sloupce zobrazeny": "Columns unhidden",
    "Sloupce: změna ({what})": "Columns: changed ({what})",
    "Číselný formát": "Number format", "Písmo": "Font", "Barva písma": "Font color",
    "Výplň": "Fill", "Ohraničení": "Border", "Zarovnání": "Alignment", "Zámek": "Protection",
    "automatická": "automatic", "žádná": "none", "žádné": "none",
    "plná {c}": "solid {c}", "přechod": "gradient",
    "zamčeno": "locked", "odemčeno": "unlocked", ", skrytý vzorec": ", formula hidden",
    "tučné": "bold", "kurzíva": "italic", "podtržené": "underline", "přeškrtnuté": "strikethrough",
    "vlevo": "left", "vpravo": "right", "nahoře": "top", "dole": "bottom", "úhlopříčka": "diagonal",
    "písmo": "font", "formát": "format", "výplň": "fill", "ohraničení": "border",
    "bez formátu": "no format",

    # ------------------------------------------------ pravidla a odkazy
    "Podmíněné formátování: pravidlo změněno": "Conditional formatting: rule changed",
    "Podmíněné formátování: pravidlo odebráno": "Conditional formatting: rule removed",
    "Podmíněné formátování: pravidlo přidáno": "Conditional formatting: rule added",
    "Ověření dat: pravidlo změněno": "Data validation: rule changed",
    "Ověření dat: pravidlo odebráno": "Data validation: rule removed",
    "Ověření dat: pravidlo přidáno": "Data validation: rule added",
    "typ {t}": "type {t}",
    "oblast {r}": "range {r}",
    "vzorec: ={f}": "formula: ={f}",
    "formát: {f}": "format: {f}",
    "ano": "yes",
    "bez rozbalovacího seznamu": "no drop-down list",
    "vzorec": "formula", "hodnota buňky": "cell value", "barevná škála": "color scale",
    "datový pruh": "data bar", "sada ikon": "icon set", "nejvyšší / nejnižší": "top / bottom",
    "jedinečné hodnoty": "unique values", "duplicitní hodnoty": "duplicate values",
    "obsahuje text": "contains text", "neobsahuje text": "does not contain text",
    "začíná na": "begins with", "končí na": "ends with", "neprázdné": "not blank",
    "chyby": "errors", "bez chyb": "no errors", "datum v období": "date occurring",
    "nad / pod průměrem": "above / below average",
    "operátor": "operator", "povolit prázdné": "allow blank", "chybové hlášení": "error alert",
    "vstupní zpráva": "input message", "druh chyby": "error style", "text chyby": "error text",
    "nadpis chyby": "error title", "zpráva": "message", "nadpis zprávy": "message title",
    "zastavit, je-li splněno": "stop if true", "počet": "rank", "procenta": "percent",
    "nejnižší": "bottom", "období": "period", "nad průměrem": "above average",
    "Odkaz přidán": "Link added", "Odkaz odebrán": "Link removed", "Odkaz změněn": "Link changed",

    # ------------------------------------------------ tvary a grafy
    "Tvar": "Shape", "Obrázek": "Picture", "Objekt": "Object", "Skupina": "Group",
    "Spojnice": "Connector", "Textové pole": "Text box", "Graf": "Chart", "absolutně": "absolute",
    "Odebráno: {k}": "Removed: {k}",
    "Přidáno: {k}": "Added: {k}",
    "Přejmenováno": "Renamed",
    "Přesunut na jinou buňku": "Moved to another cell",
    "Posunut nebo změněna velikost": "Moved or resized",
    "ukotvení stále v {a}": "still anchored at {a}",
    "Obrázek vyměněn": "Picture replaced",
    "obsah obrázku se liší": "picture content differs",
    "Formát nebo vlastnosti změněny": "Format or properties changed",
    "výplň, čára, písmo, propojení s buňkou apod.": "fill, line, font, cell link etc.",
    "Graf změněn": "Chart changed",
    "Graf: {k}": "Chart: {k}",
    "Graf: uložená data (cache)": "Chart: cached data",
    "graf ukazuje změněná data buněk": "the chart shows changed cell data",
    "Graf: formát nebo nastavení": "Chart: format or settings",
    "řady a osy beze změny": "series and axes unchanged",
    "Graf: odkazy posunuté vložením řádků/sloupců": "Chart: references shifted by inserted rows/columns",
    "Název": "Title", "Typ": "Type", "Řada {n}": "Series {n}", "Osa": "Axis",
    "(automatický)": "(automatic)", "(bez názvu)": "(no title)",

    # ------------------------------------------------ komentáře
    "Komentář přidán": "Comment added",
    "Komentář odebrán": "Comment removed",
    "Komentář změněn": "Comment changed",
    "Autor komentáře změněn": "Comment author changed",
    "autor: {a}": "author: {a}",

    # ------------------------------------------------ balíček
    "Nepřiřazená změna": "Unassigned change",
    "Makra VBA přidána": "VBA macros added",
    "Makra VBA odebrána": "VBA macros removed",
    "Makra VBA změněna": "VBA macros changed",
    "porovnává se jen otisk obsahu (hash)": "only a content hash is compared",
    "část přidána": "part added", "část odebrána": "part removed",
    "Vlastnosti dokumentu (čas uložení, autor)": "Document properties (save time, author)",
    "Pořadí výpočtu (calcChain)": "Calculation order (calcChain)",
    "Seznam typů částí": "Content types list",
    "Přečíslování vazeb mezi částmi": "Renumbered relationships between parts",
    "vazby změněných odkazů": "relationships of changed links",
    "změněné vazby: {d}": "changed relationships: {d}",
    "texty porovnány v buňkách": "texts compared in cells",
    "Přečíslování sdílených řetězců (texty v buňkách stejné)":
        "Renumbered shared strings (cell texts identical)",
    "Zobrazení listu (aktivní buňka, posun, lupa)": "Sheet view (active cell, scroll, zoom)",
    "buňky": "cells",
    "Technické atributy řádků a buněk (přečíslování stylů…)":
        "Technical row and cell attributes (renumbered styles…)",
    "rozšířená pravidla (x14)": "extended rules (x14)",
    "rozšíření: {x}": "extension: {x}",
    "Zobrazení listu nebo technické atributy": "Sheet view or technical attributes",
    "sešit: {x}": "workbook: {x}",
    "Nastavení okna a verze Excelu ve workbook.xml": "Window settings and Excel version in workbook.xml",
    "styly: {x}": "styles: {x}",
    "formát porovnán v buňkách a pravidlech": "formatting compared in cells and rules",
    "Styly – změny bez vlivu na použité buňky": "Styles – changes not affecting used cells",
    "motiv (barvy a písma)": "theme (colors and fonts)",
    "kresba – jiné rozdíly": "drawing – other differences",
    "VML (okna komentářů, ovládací prvky)": "VML (comment boxes, form controls)",
    "Formát komentářů (text a autoři stejní)": "Comment formatting (text and authors identical)",
    "komentáře": "comments",
    "Seznam osob vláknových komentářů": "List of threaded comment persons",
    "obsah nepřiřazen": "content not assigned",
    "Jen komprese nebo pořadí částí v ZIP": "Only ZIP compression or part order",
    "vlastnosti listu (barva záložky…)": "sheet properties (tab color…)",
    "šířky sloupců": "column widths", "sloučené buňky": "merged cells",
    "podmíněné formátování": "conditional formatting", "ověření dat": "data validation",
    "hypertextové odkazy": "hyperlinks", "okraje tisku": "print margins",
    "nastavení tisku": "page setup", "tisk": "print options", "záhlaví a zápatí": "header and footer",
    "konce stránek": "page breaks", "automatický filtr": "auto filter", "zámek listu": "sheet protection",
    "výchozí výška řádků": "default row height", "tabulky": "tables",
    "rozšíření (např. podmíněné formátování, sparkliny)": "extensions (e.g. conditional formatting, sparklines)",
    "data listu – jiné rozdíly (výšky řádků, skryté řádky…)":
        "sheet data – other differences (row heights, hidden rows…)",
    "styly podmíněného formátu": "conditional formatting styles",
    "vlastní paleta barev": "custom color palette", "styly tabulek": "table styles",
    "sparkliny": "sparklines", "průřezy (slicery)": "slicers", "časové osy": "timelines",
    "ovládací prvky": "form controls",

    # ------------------------------------------------ čtení souboru
    "(sdílený vzorec {si})": "(shared formula {si})",
    "{f} (sdílený z {a})": "{f} (shared from {a})",
    "(neplatný řetězec {v})": "(invalid string {v})",
    "Soubor není sešit Excel (chybí hlavní část workbook).":
        "The file is not an Excel workbook (main workbook part missing).",

    # ------------------------------------------------ HTML a Excel report
    "porovnání sešitů": "workbook comparison",
    "Starý (A)": "Old (A)", "Nový (B)": "New (B)",
    "Návratový kód": "Exit code", "tolerance čísel": "number tolerance", "(relativní)": "(relative)",
    "Přehled podle listů": "Overview by sheet",
    "Změny": "Changes", "Vrstva": "Layer", "List": "Sheet", "Druh změny": "Change type",
    "Všechny vrstvy": "All layers", "Všechny listy": "All sheets", "Všechny druhy": "All types",
    "Hledat v adrese, hodnotách, poznámce…": "Search address, values, notes…",
    "Hledat": "Search", "zobrazit šum": "show noise", "Místo": "Location", "Změna": "Change",
    "Zobrazit další": "Show more",
    "Technické rozdíly (šum)": "Technical differences (noise)",
    "Ignorované rozdíly": "Ignored differences",
    "Balíček (ZIP) – části, které se liší": "Package (ZIP) – parts that differ",
    "Záchranná síť: každá část souboru, která se liší, je tu uvedená. „Nepřiřazeno“ znamená, že rozdíl "
    "žádná vrstva zatím neumí vyložit – je vidět jen tady.":
        "Safety net: every part of the file that differs is listed here. “Unassigned” means that no "
        "layer can interpret the difference yet – it is visible only here.",
    "Část": "Part", "Stav": "Status", "Podrobnosti": "Details",
    "soubory byly jen čteny, nic se do nich nezapisovalo.": "the files were only read, nothing was written to them.",
    "Žádná změna neodpovídá filtru.": "No change matches the filter.",
    "zobrazeno {a} z {b}": "showing {a} of {b}",
    "{n} změn": "{n} changes",
    "(sešit)": "(workbook)",
    "vloženo": "inserted", "smazáno": "deleted",
    "{s} · uloženo {t}": "{s} · saved {t}",
    "Celkem": "Total",
    "Klikněte na list pro filtr.": "Click a sheet to filter.",
    "žádný": "none",
    "Žádná část balíčku se neliší.": "No package part differs.",
    "Listy beze změny: ": "Unchanged sheets: ",
    "Souhrn": "Summary",
    "Starý soubor (A)": "Old file (A)", "Nový soubor (B)": "New file (B)",
    "Verdikt": "Verdict", "Šum": "Noise", "Druh": "Type", "Staré (A)": "Old (A)", "Nové (B)": "New (B)",
    "Poznámka": "Note",

    # ------------------------------------------------ okno
    "SHEETDIFF – porovnání dvou sešitů": "SHEETDIFF – compare two workbooks",
    "Vyberte dva sešity nebo je sem přetáhněte myší.": "Select two workbooks or drag them here.",
    "Jazyk:": "Language:",
    "Starý soubor (A):": "Old file (A):",
    "Nový soubor (B):": "New file (B):",
    "Vybrat…": "Browse…",
    "⇅ Prohodit A a B": "⇅ Swap A and B",
    "Uložit report vedle nového souboru": "Save report next to the new file",
    "Také Excel report": "Also Excel report",
    "Porovnat": "Compare",
    "Otevřít report znovu": "Open report again",
    "Vyberte dva soubory a stiskněte Porovnat.": "Select two files and press Compare.",
    "Vyberte sešit": "Select a workbook",
    "Sešity Excel": "Excel workbooks",
    "Všechny soubory": "All files",
    "Starší soubor je A, novější B (podle data uložení). Lze prohodit.":
        "The older file is A, the newer B (by save date). You can swap them.",
    "Soubor vložen jako {x}.": "File set as {x}.",
    "Pozor: {f} není sešit .xlsx / .xlsm.": "Note: {f} is not an .xlsx / .xlsm workbook.",
    "Starý": "Old", "Nový": "New",
    "{x} soubor není vybraný.": "{x} file is not selected.",
    "{x} soubor neexistuje:\n{f}": "{x} file does not exist:\n{f}",
    "Porovnávám…": "Comparing…",
    "Porovnání selhalo:\n{e}": "Comparison failed:\n{e}",
}
