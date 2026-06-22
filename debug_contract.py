"""Investiga qué muestra el portal para contratos sin 'Ver factura'."""
from browser import create_browser, close_browser
from efigas_portal import login, select_contract
from config import get_credentials

EMAIL = "ljborrero@azimutenergia.co"
_, password = get_credentials(EMAIL)

# Contratos que fallan (sin Ver factura) vs que funcionan
FAIL = ["597052", "290345"]
OK   = ["54357", "221119"]

pw, browser, page = create_browser()
login(page, EMAIL, password)

for contract in FAIL + OK:
    select_contract(page, contract)
    body = page.evaluate('() => document.body.innerText')
    btns = [b['text'] for b in page.evaluate('''() =>
        Array.from(document.querySelectorAll("button")).map(b => ({text: b.innerText.trim()}))
    ''') if b['text']]
    ver = 'Ver factura' in btns
    # Extraer las primeras lineas relevantes
    lines = [l.strip() for l in body.split('\n') if l.strip()][:15]
    print(f"\n=== Contrato {contract} ({'OK' if contract in OK else 'FALLA'}) ===")
    print("  Body:", ' | '.join(lines[:8]))
    print("  Botones:", btns)
    print("  Tiene 'Ver factura':", ver)

close_browser(pw, browser)
