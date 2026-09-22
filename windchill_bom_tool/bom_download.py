"""BOM-Download-Tool fuer Windchill (ABB PLM).

Sucht eine Materialnummer in Windchill, oeffnet den Multilevel Report
der Baugruppenstruktur und speichert ihn als Excel-Datei im Zielordner.

Ablauf entspricht der manuell aufgenommenen Schrittfolge:
Suche -> Treffer oeffnen -> Structure-Tab -> Reports -> Multilevel Report
-> Actions -> Export List to File (Hover) -> Export List to XLSX.

Hinweis: Windchill wird direkt im Browser (Edge) angesteuert, nicht ueber
den Windchill Workgroup Manager. Der automatische Windows-SSO-Login greift
laut Test auch bei direktem Browseraufruf, dafuer muss der Rechner im
Firmennetz oder per VPN verbunden sein.
"""

import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

WINDCHILL_URL = "https://lp-global-plm.abb.com/Windchill/"
OUTPUT_DIR = Path(
    r"C:\Users\DEUFUENA\OneDrive - ABB\Dokumente\Aufgaben Ufuk\Bestandsmanagementtool\rohdaten\BOM auto"
)
EDGE_PROFILE_DIR = Path.home() / "AppData" / "Local" / "WindchillBomTool" / "edge-profile"
DEBUG_SCREENSHOT = Path(__file__).parent / "letzter_fehler.png"
DEBUG_HTML = Path(__file__).parent / "letzter_fehler.html"

DEFAULT_TIMEOUT_MS = 20_000
MATERIAL_NUMBER_PATTERN = re.compile(r"^[A-Za-z0-9]+$")


def get_material_number() -> str:
    if len(sys.argv) > 1 and sys.argv[1].strip():
        material_number = sys.argv[1].strip()
    else:
        material_number = input("Materialnummer eingeben: ").strip()

    if not material_number or not MATERIAL_NUMBER_PATTERN.match(material_number):
        print(f'Ungueltige Materialnummer: "{material_number}"')
        sys.exit(1)

    return material_number


def check_connectivity() -> None:
    try:
        urllib.request.urlopen(WINDCHILL_URL, timeout=8)
    except (urllib.error.URLError, TimeoutError) as exc:
        print(f"Windchill ist nicht erreichbar ({exc}).")
        print("Bitte pruefen, ob das VPN aktiv ist, und erneut versuchen.")
        sys.exit(1)


def target_path(material_number: str) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR / f"{material_number}.xlsx"


def locate_text(page, text, timeout: int = DEFAULT_TIMEOUT_MS):
    """Liefert das erste sichtbare Element mit einem der angegebenen
    Texte - durchsucht dabei das Hauptdokument UND alle eingebetteten
    iframes von `page`. Windchill rendert zentrale Bereiche (Structure-
    Ansicht, Multi-level-Report-Popup, ...) wiederholt in eigenen,
    dynamisch generierten iframes ohne vorhersagbare feste ID - deshalb
    wird hier nicht auf eine bestimmte iframe-Id vertraut, sondern
    generisch ueber `page.frames` gesucht (das Hauptdokument ist darin
    bereits enthalten).

    `text` kann ein einzelner String oder eine Liste mehrerer moeglicher
    Beschriftungen sein - noetig, weil Windchill je nach Spracheinstellung
    der Session unterschiedlich beschriftet ist (z. B. Tab "Structure"/
    "Struktur", "Reports"/"Berichte", "Actions"/"Aktionen", "Multilevel
    Report"/"Mehrstufiger Bericht", "Export List to File"/"Liste in Datei
    exportieren", "Export List to XLSX"/"Liste in xlsx exportieren" -
    jeweils per DevTools bzw. Screenshot verifiziert).

    Bevorzugt echte Link- bzw. Button-Elemente (role="link"/"button") -
    das vermeidet, dass rein informativer Text mit demselben Inhalt
    (z. B. die auf der Suchergebnisseite angezeigten Suchkriterien)
    faelschlich getroffen wird. Fallback auf reinen Text, falls keine der
    beiden Rollen passt (z. B. Tabs, die als <span> gerendert werden).

    Alle Kombinationen aus Frame/Sprachvariante/Rolle werden rundenweise
    mit kurzen Probes durchprobiert, bis der volle Timeout ausgeschoepft
    ist - so bekommt kein Kandidat unfair wenig Zeit, egal in welchem
    Frame oder an welcher Position in der Liste er steht.
    """
    labels = [text] if isinstance(text, str) else list(text)
    probe_timeout = 300
    deadline = time.monotonic() + timeout / 1000

    def probe(locator) -> bool:
        try:
            locator.wait_for(state="visible", timeout=probe_timeout)
            return True
        except PlaywrightTimeoutError:
            return False

    while True:
        for scope in [page] + list(page.frames):
            for label in labels:
                for role in ("link", "button"):
                    candidate = scope.get_by_role(role, name=label, exact=True).first
                    if probe(candidate):
                        return candidate
                candidate = scope.get_by_text(label, exact=True).first
                if probe(candidate):
                    return candidate
        if time.monotonic() >= deadline:
            break

    raise PlaywrightTimeoutError(f"Kein Element mit Text {labels!r} gefunden.")


def click_text(scope, text, timeout: int = DEFAULT_TIMEOUT_MS) -> None:
    """Klickt auf das erste sichtbare Element mit einem der angegebenen
    Texte. Siehe locate_text() fuer die Such-/Mehrsprachigkeitslogik."""
    locate_text(scope, text, timeout).click()


def find_search_box(page):
    """Windchill-Suchfeld oben rechts im Header.

    id="gloabalSearchField" (Tippfehler original von Windchill, kein
    Schreibfehler hier) wurde am 21.09. per DevTools am echten Feld
    bestaetigt - das ist die primaere, zuverlaessige Strategie. Die
    weiteren Kandidaten sind nur ein Sicherheitsnetz, falls sich die
    Windchill-Version/id einmal aendert. Bewusst KEIN "letztes
    sichtbares Textfeld"-Fallback mehr: das hat zuverlaessig ein
    verstecktes internes Feld des "Alle Typen"-Dropdowns getroffen statt
    des echten Suchfelds und dieses dabei geoeffnet.
    """
    candidates = [
        lambda: page.locator("#gloabalSearchField"),
        lambda: page.locator("input[placeholder*='Such' i]"),
        lambda: page.locator("input[placeholder*='Search' i]"),
        lambda: page.get_by_role("searchbox"),
        lambda: page.locator("input[type='search']"),
        lambda: page.locator("input[title*='Search' i]"),
    ]
    for build_locator in candidates:
        try:
            locator = build_locator()
            locator.wait_for(state="visible", timeout=3000)
            return locator
        except PlaywrightTimeoutError:
            continue

    # Falls das faelschlich geoeffnete "Alle Typen"-Dropdown noch offen ist
    # (aus einem vorherigen Fehlversuch), erst schliessen.
    page.keyboard.press("Escape")

    print()
    print("Das Suchfeld konnte nicht automatisch gefunden werden.")
    input(
        "Bitte im geoeffneten Browserfenster einmal manuell in das Suchfeld "
        "klicken und danach hier Enter druecken..."
    )
    return None


NOTICE_CONFIRM_LABELS = (
    "OK",
    "Weiter",
    "Akzeptieren",
    "Accept",
    "Continue",
    "I Agree",
    "Bestaetigen",
    "Bestätigen",
    "Submit",
)


def dismiss_startup_notice(page) -> None:
    """Manche Windchill-Profile zeigen beim ersten Aufruf einen CUI-Hinweis
    ("Controlled Unclassified Information") mit einer Checkbox "Diese Seite
    beim Start nicht anzeigen". Wird best-effort weggeklickt, falls
    vorhanden - kein Fehler, wenn die Seite gar nicht auftaucht.
    """
    try:
        checkbox = page.get_by_role("checkbox").first
        checkbox.wait_for(state="visible", timeout=3000)
        checkbox.check()
    except PlaywrightTimeoutError:
        return

    for label in NOTICE_CONFIRM_LABELS:
        try:
            button = page.get_by_text(label, exact=True).first
            button.wait_for(state="visible", timeout=1000)
            button.click()
            return
        except PlaywrightTimeoutError:
            continue

    # Kein passender Button gefunden - evtl. reicht das Setzen der Checkbox.
    page.keyboard.press("Enter")


def run_export(material_number: str, output_path: Path) -> None:
    EDGE_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(EDGE_PROFILE_DIR),
            channel="msedge",
            headless=False,
            accept_downloads=True,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)
        report_page = None

        try:
            print("Oeffne Windchill ...")
            page.goto(WINDCHILL_URL, wait_until="networkidle")
            dismiss_startup_notice(page)

            print(f"Suche Materialnummer {material_number} ...")
            search_box = find_search_box(page)
            if search_box is not None:
                search_box.click()
            page.keyboard.type(material_number)
            page.keyboard.press("Enter")

            print("Oeffne Suchtreffer ...")
            click_text(page, material_number)

            print("Oeffne Structure-Ansicht ...")
            click_text(page, ["Structure", "Struktur"])

            # Windchill rendert zentrale Bereiche wie die Structure-
            # Ansicht wiederholt in eigenen iframes (z. B. "#msrIFrame",
            # per Fehler-HTML-Dump bestaetigt) - locate_text()/click_text()
            # suchen deshalb generisch in page + allen ihren Frames statt
            # in einer fest angenommenen iframe-Id.
            print("Oeffne Multilevel Report ...")
            click_text(page, ["Reports", "Berichte"])
            with context.expect_page(timeout=DEFAULT_TIMEOUT_MS) as new_page_info:
                click_text(page, ["Multilevel Report", "Mehrstufiger Bericht"])
            report_page = new_page_info.value
            report_page.wait_for_load_state()

            print("Exportiere als XLSX ...")
            click_text(report_page, ["Actions", "Aktionen"])
            # "Export List to File" ist ein Untermenue, das sich erst per
            # Hover oeffnet (nicht per Klick) - erst danach ist "Export
            # List to XLSX" sichtbar.
            export_submenu = locate_text(
                report_page, ["Export List to File", "Liste in Datei exportieren"]
            )
            export_submenu.hover()
            with report_page.expect_download(timeout=DEFAULT_TIMEOUT_MS) as download_info:
                click_text(report_page, ["Export List to XLSX", "Liste in XLSX exportieren"])
            download = download_info.value
            download.save_as(str(output_path))

            print(f"Fertig. Gespeichert unter: {output_path}")

        except Exception:
            # Bei einem Fehler im Popup-Fenster (report_page) ist dessen
            # Zustand relevant, nicht der der urspruenglichen Seite -
            # deshalb beide sichern, falls vorhanden.
            pages_to_dump = [("page", page)]
            if report_page is not None:
                pages_to_dump.append(("report_page", report_page))

            for name, dump_page in pages_to_dump:
                try:
                    DEBUG_SCREENSHOT.parent.mkdir(parents=True, exist_ok=True)
                    screenshot_path = DEBUG_SCREENSHOT.with_name(f"letzter_fehler_{name}.png")
                    dump_page.screenshot(path=str(screenshot_path))
                    print(f"Screenshot des Fehlerzustands gespeichert: {screenshot_path}")
                except Exception:
                    pass
                try:
                    html_path = DEBUG_HTML.with_name(f"letzter_fehler_{name}.html")
                    html_path.write_text(dump_page.content(), encoding="utf-8")
                    print(f"HTML des Fehlerzustands gespeichert: {html_path}")
                except Exception:
                    pass
            raise
        finally:
            context.close()


def main() -> None:
    material_number = get_material_number()
    check_connectivity()

    output_path = target_path(material_number)
    if output_path.exists():
        print(f"Bestehende BOM fuer {material_number} wird aktualisiert.")
    else:
        print(f"Neue BOM fuer {material_number} wird erstellt.")

    run_export(material_number, output_path)


if __name__ == "__main__":
    main()
