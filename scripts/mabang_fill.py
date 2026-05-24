import asyncio
import subprocess
import sys
import io
import os
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

EXCEL_PATH = os.environ.get(
    "MABANG_EXCEL_PATH",
    os.path.join(os.path.expanduser("~"), "Desktop", "马帮信息.xlsx"),
)
SESSION_DIR = os.environ.get(
    "MABANG_SESSION_DIR",
    os.path.join(os.path.expanduser("~"), "mabang_session"),
)
ORDER_URL = "https://private.mabangerp.com/index.php?mod=order.add"
CDP_PORT = int(os.environ.get("MABANG_CDP_PORT", "9222"))

COUNTRY_ALIASES = {
    "USA": "United States",
    "US": "United States",
    "U.S.A.": "United States",
    "U.S.": "United States",
    "America": "United States",
    "UK": "United Kingdom",
    "U.K.": "United Kingdom",
    "England": "United Kingdom",
    "Great Britain": "United Kingdom",
    "GB": "United Kingdom",
    "UAE": "United Arab Emirates",
    "Korea": "South Korea",
    "Republic of Korea": "South Korea",
    "ROC": "Taiwan",
    "Holland": "Netherlands",
    "The Netherlands": "Netherlands",
    "Czech Republic": "Czech",
    "Czechia": "Czech",
    "Russia": "Russian Federation",
    "Россия": "Russian Federation",
    "Deutschland": "Germany",
    "Frankreich": "France",
    "Spanien": "Spain",
    "Italien": "Italy",
    "Österreich": "Austria",
    "Schweiz": "Switzerland",
    "Niederlande": "Netherlands",
    "Belgien": "Belgium",
    "Dänemark": "Denmark",
    "Schweden": "Sweden",
    "Norwegen": "Norway",
    "Polen": "Poland",
    "Türkei": "Turkey",
    "Kanada": "Canada",
    "Brasilien": "Brazil",
    "Australien": "Australia",
    "Mexiko": "Mexico",
    "Japón": "Japan",
    "Allemagne": "Germany",
    "Espagne": "Spain",
    "Italie": "Italy",
    "Suisse": "Switzerland",
    "Pays-Bas": "Netherlands",
    "Belgique": "Belgium",
    "Royaume-Uni": "United Kingdom",
    "États-Unis": "United States",
    "Alemania": "Germany",
    "Francia": "France",
    "España": "Spain",
    "Reino Unido": "United Kingdom",
    "Estados Unidos": "United States",
    "Brasil": "Brazil",
    "México": "Mexico",
    "Polska": "Poland",
    "Türkiye": "Turkey",
    "Sverige": "Sweden",
    "Norge": "Norway",
    "Danmark": "Denmark",
    "Suomi": "Finland",
    "Україна": "Ukraine",
    "美国": "United States",
    "英国": "United Kingdom",
    "德国": "Germany",
    "法国": "France",
    "西班牙": "Spain",
    "意大利": "Italy",
    "日本": "Japan",
    "韩国": "South Korea",
    "澳大利亚": "Australia",
    "加拿大": "Canada",
    "荷兰": "Netherlands",
    "瑞士": "Switzerland",
    "奥地利": "Austria",
    "波兰": "Poland",
    "土耳其": "Turkey",
    "巴西": "Brazil",
}

FIELD_RULES = [
    # buyerName - 收件人姓名
    ("Vollständiger Name des Empfängers", "buyerName"),
    ("Nom complet du destinataire", "buyerName"),
    ("Nombre completo del destinatario", "buyerName"),
    ("Nome completo del destinatario", "buyerName"),
    ("Nome completo do destinatário", "buyerName"),
    ("Volledige naam van de ontvanger", "buyerName"),
    ("Recipient Full Name", "buyerName"),
    ("Nombre completo", "buyerName"),
    ("Nome completo", "buyerName"),
    ("Volledige naam", "buyerName"),
    ("Nom du destinataire", "buyerName"),
    ("Nom destinataire", "buyerName"),
    ("Полное имя", "buyerName"),
    ("Получатель", "buyerName"),
    ("Destinataire", "buyerName"),
    ("Destinatario", "buyerName"),
    ("Empfänger", "buyerName"),
    ("Mottagare", "buyerName"),
    ("Modtager", "buyerName"),
    ("Vastaanottaja", "buyerName"),
    ("Odbiorca", "buyerName"),
    ("Příjemce", "buyerName"),
    ("Címzett", "buyerName"),
    ("Alıcı", "buyerName"),
    ("受取人", "buyerName"),
    ("氏名", "buyerName"),
    ("수취인", "buyerName"),
    ("수령인", "buyerName"),
    ("收件人", "buyerName"),
    ("客户姓名", "buyerName"),
    ("ФИО", "buyerName"),
    ("Full Name", "buyerName"),
    ("Tam ad", "buyerName"),
    ("Navn", "buyerName"),
    ("Namn", "buyerName"),
    ("Nimi", "buyerName"),
    ("Nome", "buyerName"),
    ("Nom", "buyerName"),

    # buyerUserId
    ("客户ID", "buyerUserId"),

    # phone1 - 电话
    ("Numéro de téléphone", "phone1"),
    ("Número de teléfono", "phone1"),
    ("Numero di telefono", "phone1"),
    ("Número de telefone", "phone1"),
    ("Номер телефона", "phone1"),
    ("Numer telefonu", "phone1"),
    ("Telefoonnummer", "phone1"),
    ("Telefonnummer", "phone1"),
    ("Puhelinnumero", "phone1"),
    ("Phone Number", "phone1"),
    ("電話番号", "phone1"),
    ("전화번호", "phone1"),
    ("Téléphone", "phone1"),
    ("Teléfono", "phone1"),
    ("Telefono", "phone1"),
    ("Telefone", "phone1"),
    ("Телефон", "phone1"),
    ("Telefon", "phone1"),
    ("电话", "phone1"),

    # email - 邮箱
    ("E-Mail", "email"),
    ("E-mail", "email"),
    ("Correo", "email"),
    ("Courriel", "email"),
    ("邮箱", "email"),
    ("Email", "email"),
    ("メール", "email"),
    ("이메일", "email"),
    ("Почта", "email"),

    # countryNameEN - 国家
    ("Country", "countryNameEN"),
    ("Paese", "countryNameEN"),
    ("Pays", "countryNameEN"),
    ("País", "countryNameEN"),
    ("Land", "countryNameEN"),
    ("Kraj", "countryNameEN"),
    ("Ülke", "countryNameEN"),
    ("Země", "countryNameEN"),
    ("Ország", "countryNameEN"),
    ("Страна", "countryNameEN"),
    ("国", "countryNameEN"),
    ("국가", "countryNameEN"),
    ("Maa", "countryNameEN"),

    # province - 省/州
    ("Provincie/Staat", "province"),
    ("State/Province", "province"),
    ("Bundesland/Provinz", "province"),
    ("Bundesland", "province"),
    ("Provincia", "province"),
    ("Province", "province"),
    ("Província", "province"),
    ("Région", "province"),
    ("Regione", "province"),
    ("Область", "province"),
    ("Край", "province"),
    ("Województwo", "province"),
    ("Provincie", "province"),
    ("Län", "province"),
    ("Fylke", "province"),
    ("Maakunta", "province"),
    ("Département", "province"),
    ("Comunidad", "province"),
    ("Região", "province"),
    ("İl", "province"),
    ("Kraj", "province"),
    ("Megye", "province"),
    ("都道府県", "province"),
    ("省/州", "province"),
    ("Estado", "province"),
    ("省", "province"),
    ("州", "province"),

    # city - 城市
    ("Plaats", "city"),
    ("Woonplaats", "city"),
    ("Ort", "city"),
    ("Ciudad", "city"),
    ("Cidade", "city"),
    ("Città", "city"),
    ("Località", "city"),
    ("Ville", "city"),
    ("Localité", "city"),
    ("Stadt", "city"),
    ("Miasto", "city"),
    ("Miejscowość", "city"),
    ("Город", "city"),
    ("Населённый пункт", "city"),
    ("City", "city"),
    ("Town", "city"),
    ("Stad", "city"),
    ("Şehir", "city"),
    ("İlçe", "city"),
    ("Város", "city"),
    ("Město", "city"),
    ("Obec", "city"),
    ("Localidad", "city"),
    ("Localidade", "city"),
    ("Kaupunki", "city"),
    ("Sted", "city"),
    ("市区町村", "city"),
    ("城市", "city"),
    ("도시", "city"),
    ("시", "city"),

    # postCode - 邮编
    ("ZIP/Postal Code", "postCode"),
    ("Postleitzahl", "postCode"),
    ("Code postal", "postCode"),
    ("Código postal", "postCode"),
    ("Codice postale", "postCode"),
    ("Почтовый индекс", "postCode"),
    ("Postal Code", "postCode"),
    ("Kod pocztowy", "postCode"),
    ("Posta kodu", "postCode"),
    ("Postnummer", "postCode"),
    ("Postinumero", "postCode"),
    ("Postcode", "postCode"),
    ("郵便番号", "postCode"),
    ("우편번호", "postCode"),
    ("ZIP Code", "postCode"),
    ("Zip Code", "postCode"),
    ("邮编", "postCode"),
    ("CEP", "postCode"),

    # street1 - 地址
    ("Street Address", "street1"),
    ("Street", "street1"),
    ("Indirizzo", "street1"),
    ("Dirección", "street1"),
    ("Endereço", "street1"),
    ("Adresse", "street1"),
    ("Adres", "street1"),
    ("Address", "street1"),
    ("Straße", "street1"),
    ("Straat", "street1"),
    ("Улица", "street1"),
    ("Адрес", "street1"),
    ("Ulica", "street1"),
    ("Sokak", "street1"),
    ("Gatuadress", "street1"),
    ("Osoite", "street1"),
    ("邮寄地址", "street1"),
    ("住所", "street1"),
    ("番地", "street1"),
    ("주소", "street1"),
    ("地址", "street1"),
    ("Calle", "street1"),
    ("Rue", "street1"),
    ("Via", "street1"),

    # salesRecordNumber / trackNumber2
    ("交易号", "salesRecordNumber"),
    ("内部单号", "trackNumber2"),
]


def log(msg):
    try:
        print(msg, flush=True)
    except Exception:
        print(msg.encode("ascii", "replace").decode(), flush=True)


def parse_cell(cell):
    text = str(cell)
    for sep in ["：", ":"]:
        if sep in text:
            return text.split(sep, 1)
    try:
        text2 = cell.encode("raw_unicode_escape").decode("utf-8")
        for sep in ["：", ":"]:
            if sep in text2:
                return text2.split(sep, 1)
    except Exception:
        pass
    return None


def read_excel():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        for cell in row:
            if not cell:
                continue
            parts = parse_cell(cell)
            if not parts:
                continue
            key, value = parts[0].strip(), parts[1].strip()
            key_lower = key.lower()
            for label, field_name in FIELD_RULES:
                if label.lower() in key_lower:
                    data[field_name] = value
                    break
    return data


def clean_session_locks():
    for f in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        try:
            os.remove(os.path.join(SESSION_DIR, f))
        except Exception:
            pass


def get_chromium_path():
    base = os.path.expanduser(r"~\AppData\Local\ms-playwright")
    if not os.path.isdir(base):
        return None
    for entry in sorted(os.listdir(base), reverse=True):
        if entry.startswith("chromium-"):
            for sub in ["chrome-win64", "chrome-win"]:
                exe = os.path.join(base, entry, sub, "chrome.exe")
                if os.path.isfile(exe):
                    return exe
    return None


def launch_chrome_detached(chromium_path):
    clean_session_locks()
    cmd = [
        chromium_path,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={SESSION_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "--window-size=1400,900",
    ]
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(
        cmd,
        creationflags=CREATE_NEW_PROCESS_GROUP | DETACHED_PROCESS,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
    )


async def main():
    log("Reading Excel data...")
    data = read_excel()
    if not data:
        log("ERROR: No data found in Excel file!")
        return

    log("Data parsed:")
    for field, value in data.items():
        log(f"  {field} = {value}")

    chromium_path = get_chromium_path()
    if not chromium_path:
        log("ERROR: Cannot find Playwright chromium!")
        return

    async with async_playwright() as p:
        browser = None
        try:
            browser = await p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
            log("Connected to existing browser.")
        except Exception:
            pass

        if not browser:
            log("Launching browser...")
            launch_chrome_detached(chromium_path)
            for attempt in range(20):
                await asyncio.sleep(1)
                try:
                    browser = await p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
                    log("Connected to browser.")
                    break
                except Exception:
                    continue

        if not browser:
            log("ERROR: Could not connect to browser!")
            return

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else await context.new_page()

        log("Navigating to order page...")
        await page.goto(ORDER_URL, wait_until="domcontentloaded", timeout=30000)
        await page.wait_for_timeout(5000)

        current_url = page.url
        log(f"URL: {current_url}")

        if "mod=order.add" not in current_url:
            log("NOT LOGGED IN - please login in the browser!")
            logged_in = False
            for i in range(100):
                await page.wait_for_timeout(3000)
                try:
                    cur = page.url
                    if "mod=order.add" in cur:
                        logged_in = True
                        break
                    if "mabangerp.com" in cur and "login" not in cur.lower() and "mod=" in cur:
                        logged_in = True
                        log("  Login detected via URL change, navigating to order page...")
                        await page.goto(ORDER_URL, wait_until="domcontentloaded", timeout=30000)
                        await page.wait_for_timeout(5000)
                        break
                    has_iframe = await page.evaluate(
                        "() => !!document.querySelector('iframe#iframeContent')"
                    )
                    if has_iframe:
                        logged_in = True
                        log("  Login detected via iframe, navigating to order page...")
                        await page.goto(ORDER_URL, wait_until="domcontentloaded", timeout=30000)
                        await page.wait_for_timeout(5000)
                        break
                except Exception as e:
                    log(f"  Check error: {str(e)[:80]}")
                if i % 5 == 4:
                    log(f"  Retrying navigation... ({(i + 1) * 3}s)")
                    try:
                        await page.goto(ORDER_URL, wait_until="domcontentloaded", timeout=15000)
                        await page.wait_for_timeout(3000)
                        if "mod=order.add" in page.url:
                            logged_in = True
                            log("  Login confirmed!")
                            break
                    except Exception:
                        pass
            if not logged_in:
                log("Timeout waiting for login.")
                return

        await page.wait_for_timeout(3000)
        iframe_el = await page.query_selector("iframe#iframeContent")
        if not iframe_el:
            log("ERROR: Cannot find iframeContent iframe!")
            return

        frame = await iframe_el.content_frame()
        await frame.wait_for_timeout(2000)

        log("Filling text fields...")
        for field_name, value in data.items():
            if field_name == "countryNameEN":
                continue
            try:
                el = await frame.query_selector(
                    f'input[name="{field_name}"], textarea[name="{field_name}"]'
                )
                if el:
                    await el.click()
                    await el.fill("")
                    await el.fill(value)
                    log(f"  OK: {field_name} = {value}")
                else:
                    log(f"  SKIP: {field_name} not found on form")
            except Exception as e:
                log(f"  ERR: {field_name}: {str(e)[:60]}")

        # --- Country dropdown (xm-select component) ---
        country = data.get("countryNameEN", "")
        if country:
            country = COUNTRY_ALIASES.get(country.strip(), country)
            log(f"  Selecting country: {country}")
            try:
                # Click xm-select to open it
                xm = await frame.query_selector("#selectSignleDemo xm-select")
                if xm:
                    await xm.click()
                    await frame.wait_for_timeout(500)
                    # Type in search box to filter
                    search_input = await frame.query_selector("#selectSignleDemo .xm-search-input")
                    if search_input:
                        await search_input.fill(country)
                        await frame.wait_for_timeout(500)
                    # Click matching option
                    matched = await frame.evaluate("""(country) => {
                        const options = document.querySelectorAll('#selectSignleDemo .xm-option');
                        for (const opt of options) {
                            if (opt.style.display === 'none') continue;
                            const text = opt.querySelector('.xm-option-content');
                            if (text && text.textContent.includes(country)) {
                                opt.click();
                                return text.textContent.trim();
                            }
                        }
                        return null;
                    }""", country)
                    if matched:
                        log(f"  OK: country = {matched}")
                    else:
                        log(f"  SKIP: country '{country}' not found in xm-select")
                else:
                    log("  SKIP: xm-select not found")
            except Exception as e:
                log(f"  ERR: country: {str(e)[:80]}")

        # --- Shop dropdown: select "售后(Global区域)" ---
        log("  Selecting shop: 售后(Global区域)")
        try:
            # Click caret to open dropdown
            caret = await frame.query_selector(".shop-select .caret-dropdown")
            if caret:
                await caret.click()
                await frame.wait_for_timeout(800)
            # Click the target li
            shop_matched = await frame.evaluate("""() => {
                const ul = document.querySelector('.shop-select .select_copy_control');
                if (!ul) return 'no ul';
                const lis = ul.querySelectorAll('li');
                for (const li of lis) {
                    if (li.getAttribute('oliname') === '售后(Global区域)') {
                        const a = li.querySelector('a');
                        if (a) a.click();
                        return '售后(Global区域)';
                    }
                }
                return null;
            }""")
            if shop_matched and shop_matched != 'no ul':
                log(f"  OK: shop = {shop_matched}")
                # Also set the display text
                await frame.evaluate("""() => {
                    const input = document.getElementById('shopIdname');
                    if (input) input.value = '售后(Global区域)';
                }""")
            else:
                log(f"  WARN: shop match failed: {shop_matched}")
        except Exception as e:
            log(f"  ERR: shop: {str(e)[:80]}")

        # --- Wait for salesman field to appear, then fill ---
        log("  Filling salesman: 赖丹")
        try:
            await frame.wait_for_timeout(1500)
            # Remove hide class if present
            await frame.evaluate("""() => {
                const div = document.querySelector('.sellerName');
                if (div) div.classList.remove('hide');
            }""")
            await frame.wait_for_timeout(500)
            seller_input = await frame.query_selector('input[name="sellerName"]')
            if seller_input:
                await seller_input.click()
                await seller_input.fill("赖丹")
                log("  OK: sellerName = 赖丹")
            else:
                log("  SKIP: sellerName input not found")
        except Exception as e:
            log(f"  ERR: sellerName: {str(e)[:80]}")

        # --- Custom category: check "售后(补发/换货...)" ---
        log("  Selecting category: 售后(补发/换货...)")
        try:
            checked = await frame.evaluate("""() => {
                const labels = document.querySelectorAll('label.checkbox-inline');
                for (const label of labels) {
                    if (label.textContent.includes('售后(补发/换货')) {
                        const cb = label.querySelector('input[type="checkbox"]');
                        if (cb && !cb.checked) cb.click();
                        return true;
                    }
                }
                return false;
            }""")
            if checked:
                log("  OK: category = 售后(补发/换货...)")
            else:
                log("  SKIP: category checkbox not found")
        except Exception as e:
            log(f"  ERR: category: {str(e)[:80]}")

        # --- Payment time: set to today ---
        log("  Setting payment time to today...")
        try:
            from datetime import datetime, timezone, timedelta
            tz = timezone(timedelta(hours=8))
            now = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            paid_input = await frame.query_selector("#paidTime")
            if paid_input:
                await paid_input.click()
                await frame.wait_for_timeout(500)
                # Try clicking "今天" button in WdatePicker
                today_clicked = await frame.evaluate("""() => {
                    // WdatePicker creates elements in parent document
                    return false;
                }""")
                # Directly set the value via JS
                await frame.evaluate("""(val) => {
                    const el = document.getElementById('paidTime');
                    if (el) {
                        el.value = val;
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                }""", now)
                log(f"  OK: paidTime = {now}")
            else:
                log("  SKIP: paidTime input not found")
        except Exception as e:
            log(f"  ERR: paidTime: {str(e)[:80]}")

        log("DONE - all fields filled!")


asyncio.run(main())
