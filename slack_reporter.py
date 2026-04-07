"""
slack_reporter.py
Toma un screenshot de la pestaña indicada del dashboard OTC (sin sidebar ni header)
y lo envía a un canal de Slack.

Uso:
    python3 slack_reporter.py --file /tmp/chart.xlsx --channel C0AM3K4DR6W --tab clientes
    python3 slack_reporter.py --file /tmp/chart.xlsx --channel C0AM3K4DR6W --tab clientes --section detalle
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import urllib.parse
from pathlib import Path

import requests

# ── Configuración ──────────────────────────────────────────────────────────────
SLACK_TOKEN     = os.environ.get("SLACK_TOKEN", "")
SCREENSHOT_PATH = "/tmp/koywe_tab_screenshot.png"
VIEWPORT        = {"width": 1400, "height": 900}

def _detect_app_url() -> str:
    """Detecta en qué puerto está corriendo la app de Streamlit."""
    for port in [8501, 8502, 8503, 8504, 8505]:
        try:
            r = requests.get(f"http://localhost:{port}/_stcore/health", timeout=2)
            if r.text.strip() == "ok":
                return f"http://localhost:{port}"
        except Exception:
            continue
    return "http://localhost:8501"  # fallback

# Índice (0-based) de cada pestaña en el orden declarado en app.py
TAB_INDEX = {
    "overview":  0,
    "pais":      1,
    "clientes":  2,
    "retencion": 2,   # fusionado en pestaña Clientes
    "semanal":   3,
}

# Mensajes de Slack por pestaña/sección
SECTION_META = {
    "retencion": {
        "title":   "Retención de Clientes — Dashboard OTC",
        "comment": "📊 *Retención* — Dashboard OTC Koywe",
    },
    "detalle": {
        "title":   "Análisis de Clientes — Dashboard OTC",
        "comment": "📊 *Análisis de Clientes* — Dashboard OTC Koywe",
    },
    "overview": {
        "title":   "Overview — Dashboard OTC",
        "comment": "📊 *Overview* — Dashboard OTC Koywe",
    },
    "canal": {
        "title":   "Volumen por Canal — Dashboard OTC",
        "comment": "📊 *Volumen por Canal* — Dashboard OTC Koywe",
    },
    "pais": {
        "title":   "Por País — Dashboard OTC",
        "comment": "📊 *Por País* — Dashboard OTC Koywe",
    },
    "semanal": {
        "title":   "Resumen Semanal — Dashboard OTC",
        "comment": "📊 *Semanal* — Dashboard OTC Koywe",
    },
}


# ── Screenshot con Playwright ──────────────────────────────────────────────────
def take_screenshot(
    file_path: str,
    tab: str = "clientes",
    section: str = "retencion",
    url_params: dict | None = None,
) -> str:
    from playwright.sync_api import sync_playwright

    tab_idx    = TAB_INDEX.get(tab, 2)
    app_url    = _detect_app_url()
    query_str  = ("?" + urllib.parse.urlencode(url_params)) if url_params else ""
    target_url = f"{app_url}/{query_str}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page    = browser.new_page(viewport=VIEWPORT)

        # 1. Abrir la app con los filtros replicados via URL params
        page.goto(target_url, wait_until="networkidle", timeout=30_000)
        time.sleep(2)

        # 2. Subir archivo de datos
        page.locator('input[type="file"]').set_input_files(file_path)
        page.wait_for_selector('[data-baseweb="tab"]', timeout=30_000)
        time.sleep(3)

        # 3. Navegar a la pestaña
        page.locator('[data-baseweb="tab"]').nth(tab_idx).click()
        time.sleep(2)

        # 4. Ocultar sidebar, header y decoraciones
        page.evaluate("""
            () => {
                const hide = (sel) => {
                    const el = document.querySelector(sel);
                    if (el) el.style.setProperty('display', 'none', 'important');
                };
                hide('[data-testid="stSidebar"]');
                hide('[data-testid="stHeader"]');
                hide('[data-testid="stToolbar"]');
                hide('[data-testid="stStatusWidget"]');
                const main = document.querySelector('[data-testid="stMain"]');
                if (main) {
                    main.style.setProperty('margin-left', '0', 'important');
                    main.style.setProperty('padding-left', '2rem', 'important');
                }
                const block = document.querySelector('.block-container');
                if (block) block.style.setProperty('padding-top', '1rem', 'important');
            }
        """)
        # Ampliar el viewport para que todo entre en pantalla sin scroll
        # → elimina lazy rendering de Streamlit (tablas, gráficas, etc.)
        full_height = page.evaluate("document.body.scrollHeight")
        page.set_viewport_size({"width": VIEWPORT["width"], "height": max(full_height, 2000)})
        time.sleep(2)  # dar tiempo a que re-rendericen los componentes ahora visibles

        # Esperar que los dataframes terminen de renderizar
        try:
            page.wait_for_selector('[data-testid="stDataFrame"]', timeout=8_000)
        except Exception:
            pass
        time.sleep(0.5)

        # 5. Clip según sección
        if section == "canal":
            clip = page.evaluate("""
                () => {
                    const start = document.getElementById('inicio-canal');
                    const end   = document.getElementById('fin-canal');
                    if (!start || !end) return null;
                    const r1 = start.getBoundingClientRect();
                    const r2 = end.getBoundingClientRect();
                    const scrollY = window.scrollY || 0;
                    return {
                        x:      0,
                        y:      Math.max(0, r1.top + scrollY - 16),
                        width:  window.innerWidth,
                        height: r2.top + scrollY - (r1.top + scrollY) + 40,
                    };
                }
            """)
            if clip and clip["height"] > 0:
                page.screenshot(path=SCREENSHOT_PATH, full_page=True, clip=clip)
            else:
                page.screenshot(path=SCREENSHOT_PATH, full_page=True)
        elif section == "detalle":
            clip = page.evaluate("""
                () => {
                    const start = document.getElementById('seccion-analisis');
                    const end   = document.getElementById('fin-graficas');
                    if (!start || !end) return null;
                    const r1 = start.getBoundingClientRect();
                    const r2 = end.getBoundingClientRect();
                    const scrollY = window.scrollY || 0;
                    return {
                        x:      0,
                        y:      r1.top + scrollY,
                        width:  window.innerWidth,
                        height: r2.top + scrollY - (r1.top + scrollY),
                    };
                }
            """)
            if clip and clip["height"] > 0:
                page.screenshot(path=SCREENSHOT_PATH, full_page=True, clip=clip)
            else:
                # Fallback: scroll hasta la sección y captura el viewport
                page.evaluate("""
                    () => {
                        const el = document.getElementById('seccion-analisis');
                        if (el) el.scrollIntoView({ behavior: 'instant', block: 'start' });
                    }
                """)
                time.sleep(0.5)
                page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        else:
            # Screenshot completo del área principal
            page.screenshot(path=SCREENSHOT_PATH, full_page=True)

        browser.close()

    return SCREENSHOT_PATH


# ── Envío a Slack ──────────────────────────────────────────────────────────────
def send_to_slack(img_path: str, channel: str, section: str = "retencion") -> dict:
    meta      = SECTION_META.get(section, SECTION_META["retencion"])
    img_bytes = Path(img_path).read_bytes()

    r1 = requests.post(
        "https://slack.com/api/files.getUploadURLExternal",
        headers={"Authorization": f"Bearer {SLACK_TOKEN}"},
        data={"filename": "koywe_dashboard.png", "length": len(img_bytes)},
        timeout=30,
    )
    r1_json = r1.json()
    if not r1_json.get("ok"):
        return r1_json

    requests.post(
        r1_json["upload_url"],
        data=img_bytes,
        headers={"Content-Type": "image/png"},
        timeout=60,
    )

    r3 = requests.post(
        "https://slack.com/api/files.completeUploadExternal",
        headers={"Authorization": f"Bearer {SLACK_TOKEN}", "Content-Type": "application/json"},
        json={
            "files":           [{"id": r1_json["file_id"], "title": meta["title"]}],
            "channel_id":      channel,
            "initial_comment": meta["comment"],
        },
        timeout=30,
    )
    return r3.json()


# ── Flujo completo ─────────────────────────────────────────────────────────────
def report(
    file_path: str,
    channel: str,
    tab: str = "clientes",
    section: str = "retencion",
    url_params: dict | None = None,
) -> bool:
    img_path = take_screenshot(file_path, tab, section, url_params)
    result   = send_to_slack(img_path, channel, section)
    if result.get("ok"):
        print(f"✅  Enviado ({section}) — file_id={result.get('files', [{}])[0].get('id', '?')}")
        return True
    else:
        print(f"❌  Error Slack: {result.get('error', result)}", file=sys.stderr)
        return False


# ── CLI ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Envía screenshot del dashboard a Slack")
    parser.add_argument("--file",       required=True,                               help="Ruta al .xlsx de datos")
    parser.add_argument("--channel",    default="C0AM3K4DR6W",                       help="Slack channel ID")
    parser.add_argument("--tab",        default="clientes", choices=list(TAB_INDEX.keys()))
    parser.add_argument("--section",    default="retencion", choices=list(SECTION_META.keys()))
    # Filtros de vista (para replicar lo que el usuario ve)
    parser.add_argument("--year",       default="")
    parser.add_argument("--view",       default="")
    parser.add_argument("--mes",        default="")
    parser.add_argument("--trimestre",  default="")
    parser.add_argument("--countries",  default="")
    parser.add_argument("--ret-desde",  default="", dest="ret_desde")
    parser.add_argument("--ret-hasta",  default="", dest="ret_hasta")
    args = parser.parse_args()

    # Construir dict de params URL (solo los que tengan valor)
    params = {}
    for k, v in [("year", args.year), ("view", args.view), ("mes", args.mes),
                 ("trimestre", args.trimestre), ("countries", args.countries),
                 ("ret_desde", args.ret_desde), ("ret_hasta", args.ret_hasta)]:
        if v:
            params[k] = v

    ok = report(args.file, args.channel, args.tab, args.section, params or None)
    sys.exit(0 if ok else 1)
