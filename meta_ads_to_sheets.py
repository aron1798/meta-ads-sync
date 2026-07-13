"""
Meta Ads → Google Sheets (Versión MENSUAL con período de gracia)
=================================================================
Descarga los insights de los conjuntos de anuncios (Ad Sets) de Meta Ads
y los guarda en una Google Sheet, AGRUPADOS POR MES.

LÓGICA:
- Una fila por (conjunto de anuncios × mes).
- Meses con más de GRACE_DAYS días cerrados quedan CONGELADOS (no se vuelven a pedir a Meta).
- Mes actual + meses con menos de GRACE_DAYS días cerrados se actualizan en cada corrida.
- Empieza desde START_YEAR_MONTH (configurable, default enero del año en curso).

MEJORAS EN ESTA VERSIÓN:
- El "Resultado" ahora se elige según el OBJETIVO de cada anuncio (igual que Meta),
  cubriendo mensajes, tráfico (clics / vistas de landing), leads, ventas y alcance.
- Se agrega la columna "ID del conjunto de anuncios" para no confundir dos adsets
  que tengan el mismo nombre.
- La columna "Indicador de resultado" ahora se llena correctamente.

Variables de entorno necesarias (configuradas como Secrets en GitHub):
- META_ACCESS_TOKEN     : Token de larga duración de Meta (Graph API)
- META_AD_ACCOUNT_ID    : ID de la cuenta publicitaria SIN el "act_"
- GOOGLE_SHEET_ID       : ID de la Google Sheet
- GOOGLE_SERVICE_ACCOUNT_JSON : Contenido completo del JSON de la cuenta de servicio
- START_YEAR_MONTH      : (opcional) Mes desde el que empezar (formato YYYY-MM). Default: 2026-01
- GRACE_DAYS            : (opcional) Días de gracia tras cerrar un mes. Default: 7

Autor: Generado para Aron - Refriperu
"""

import os
import sys
import json
import calendar
from datetime import datetime, timedelta, timezone, date

import requests
import gspread
from google.oauth2.service_account import Credentials


# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------
META_API_VERSION = "v19.0"
META_BASE_URL = f"https://graph.facebook.com/{META_API_VERSION}"

WORKSHEET_NAME = "Meta_Ads_Adsets"

# Las mismas columnas que aparecen en tu Excel exportado de Ads Manager
# (se agregó "ID del conjunto de anuncios" al inicio para identificar cada adset de forma única)
COLUMNS = [
    "ID del conjunto de anuncios",
    "Inicio del informe",
    "Fin del informe",
    "Nombre del conjunto de anuncios",
    "Entrega del conjunto de anuncios",
    "Resultados",
    "Indicador de resultado",
    "Costo por resultados",
    "Presupuesto del conjunto de anuncios",
    "Tipo de presupuesto del conjunto de anuncios",
    "Importe gastado (USD)",
    "Impresiones",
    "Alcance",
    "Finalización",
    "Inicio",
    "Configuración de atribución",
    "Puja",
    "Tipo de puja",
    "Último cambio significativo",
    "Contactos de mensajes totales",
    "Nuevos contactos de mensajes",
    "Compras",
    "Costo por compra (USD)",
    "Última actualización",
]

# Etiquetas legibles para cada tipo de acción (para la columna "Indicador de resultado")
RESULT_LABELS = {
    "onsite_conversion.messaging_conversation_started_7d": "Conversaciones con mensajes iniciadas",
    "leadgen.other": "Clientes potenciales",
    "lead": "Clientes potenciales",
    "link_click": "Clics en el enlace",
    "omni_landing_page_view": "Visualizaciones de la página de destino",
    "landing_page_view": "Visualizaciones de la página de destino",
    "post_engagement": "Interacciones con la publicación",
    "page_engagement": "Interacciones con la página",
    "omni_purchase": "Compras",
    "purchase": "Compras",
    "reach": "Alcance",
}

# Para cada OBJETIVO de Meta, cuál es el resultado principal (en orden de preferencia).
# Esto replica cómo Meta decide qué mostrar en la columna "Resultados" al exportar.
OBJECTIVE_RESULT = {
    # Objetivos "Outcome" (nuevos)
    "OUTCOME_LEADS":      ["onsite_conversion.messaging_conversation_started_7d", "leadgen.other", "lead"],
    "OUTCOME_ENGAGEMENT": ["onsite_conversion.messaging_conversation_started_7d", "post_engagement", "page_engagement"],
    "OUTCOME_TRAFFIC":    ["link_click", "omni_landing_page_view", "landing_page_view"],
    "OUTCOME_SALES":      ["onsite_conversion.messaging_conversation_started_7d", "omni_purchase", "purchase"],
    "OUTCOME_AWARENESS":  ["reach"],
    "OUTCOME_APP_PROMOTION": ["omni_app_install", "app_install"],
    # Objetivos antiguos (por compatibilidad)
    "MESSAGES":           ["onsite_conversion.messaging_conversation_started_7d"],
    "LEAD_GENERATION":    ["leadgen.other", "lead"],
    "LINK_CLICKS":        ["link_click", "omni_landing_page_view"],
    "CONVERSIONS":        ["omni_purchase", "purchase", "onsite_conversion.messaging_conversation_started_7d"],
    "REACH":              ["reach"],
    "BRAND_AWARENESS":    ["reach"],
    "POST_ENGAGEMENT":    ["post_engagement", "page_engagement"],
}

# Orden de respaldo si el objetivo no está en el mapa o no encuentra su acción principal.
FALLBACK_PRIORITY = [
    "onsite_conversion.messaging_conversation_started_7d",
    "leadgen.other",
    "lead",
    "omni_purchase",
    "purchase",
    "link_click",
    "omni_landing_page_view",
    "landing_page_view",
    "post_engagement",
    "page_engagement",
]


# ----------------------------------------------------------------------
# Helpers de fechas
# ----------------------------------------------------------------------
def month_start_end(year, month):
    """Devuelve (primer día, último día) de un mes como strings YYYY-MM-DD."""
    first = date(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    last = date(year, month, last_day)
    return first.strftime("%Y-%m-%d"), last.strftime("%Y-%m-%d")


def is_month_active(year, month, today, grace_days):
    """
    Devuelve True si el mes debe actualizarse (no está congelado).
    Un mes está activo si:
    - Es el mes actual, O
    - Han pasado <= grace_days desde su último día.
    """
    last_day = calendar.monthrange(year, month)[1]
    month_end = date(year, month, last_day)
    if today <= month_end:
        return True
    days_since_close = (today - month_end).days
    return days_since_close <= grace_days


def list_months(start_ym, today):
    """Devuelve lista de tuplas (year, month) desde start_ym hasta el mes de today, inclusive."""
    start_year, start_month = start_ym
    months = []
    y, m = start_year, start_month
    while (y, m) <= (today.year, today.month):
        months.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


# ----------------------------------------------------------------------
# Helpers Meta Graph API
# ----------------------------------------------------------------------
def meta_get(path, params):
    """GET a Graph API con manejo de paginación."""
    url = f"{META_BASE_URL}/{path}"
    items = []
    while url:
        r = requests.get(url, params=params, timeout=60)
        if r.status_code != 200:
            print(f"ERROR Meta API ({r.status_code}): {r.text[:500]}", file=sys.stderr)
            r.raise_for_status()
        data = r.json()
        items.extend(data.get("data", []))
        paging = data.get("paging", {})
        url = paging.get("next")
        params = None
    return items


def fetch_adsets(account_id, token):
    """Trae los metadatos de todos los conjuntos de anuncios."""
    fields = ",".join([
        "id", "name", "status", "effective_status",
        "daily_budget", "lifetime_budget", "budget_remaining",
        "start_time", "end_time",
        "attribution_spec",
        "bid_amount", "bid_strategy",
        "updated_time",
    ])
    params = {
        "fields": fields,
        "limit": 200,
        "access_token": token,
    }
    return meta_get(f"act_{account_id}/adsets", params)


def fetch_insights_for_month(account_id, token, year, month):
    """Trae insights del mes (year, month) — un total mensual por adset."""
    since, until = month_start_end(year, month)

    fields = ",".join([
        "adset_id", "adset_name",
        "date_start", "date_stop",
        "impressions", "reach", "spend",
        "actions", "cost_per_action_type",
        "objective",
    ])
    params = {
        "level": "adset",
        "fields": fields,
        "time_range": json.dumps({"since": since, "until": until}),
        "limit": 200,
        "access_token": token,
    }
    return meta_get(f"act_{account_id}/insights", params), since, until


# ----------------------------------------------------------------------
# Procesamiento
# ----------------------------------------------------------------------
def get_action_value(actions, action_type):
    if not actions:
        return None
    for a in actions:
        if a.get("action_type") == action_type:
            try:
                return float(a.get("value", 0))
            except (ValueError, TypeError):
                return a.get("value")
    return None


def pick_result(insight):
    """
    Elige el RESULTADO PRINCIPAL igual que Meta, según el objetivo del anuncio.
    Devuelve (valor, action_type, etiqueta).
    """
    actions = insight.get("actions", []) or []
    objective = (insight.get("objective") or "").upper()

    # 1) Intentar según el objetivo del anuncio
    candidates = OBJECTIVE_RESULT.get(objective, [])
    for action_type in candidates:
        v = get_action_value(actions, action_type)
        if v is not None:
            return v, action_type, RESULT_LABELS.get(action_type, action_type)

    # 2) Respaldo: recorrer prioridad general
    for action_type in FALLBACK_PRIORITY:
        v = get_action_value(actions, action_type)
        if v is not None:
            return v, action_type, RESULT_LABELS.get(action_type, action_type)

    # 3) Último respaldo: alcance
    reach = insight.get("reach")
    if reach:
        return float(reach), "reach", RESULT_LABELS["reach"]

    return None, "", ""


def attribution_label(spec):
    if not spec:
        return ""
    parts = []
    for item in spec:
        click = item.get("event_type", "")
        click_window = item.get("window_days", "")
        parts.append(f"{click_window} días {click}")
    return " | ".join(parts)


def budget_label(adset):
    if adset.get("daily_budget"):
        return float(adset["daily_budget"]) / 100, "Diario"
    if adset.get("lifetime_budget"):
        return float(adset["lifetime_budget"]) / 100, "Total"
    return "", "Con el presupuesto de la campaña"


def build_row(adset, insight, since, until, now_iso):
    """Construye una fila combinando metadata del adset + insights del mes."""
    ins = insight or {}
    budget_value, budget_type = budget_label(adset)

    cpa = ins.get("cost_per_action_type", [])

    # --- Resultado principal según el objetivo (igual que Meta) ---
    result_value, result_action_type, result_indicator = pick_result(ins)
    cost_per_result = get_action_value(cpa, result_action_type) if result_action_type else None
    # Si el resultado fue alcance calculado, costo = gasto / alcance
    if result_action_type == "reach" and result_value:
        spend = float(ins.get("spend", 0) or 0)
        cost_per_result = (spend / result_value) if result_value else None

    actions = ins.get("actions", [])
    msg_total = get_action_value(actions, "onsite_conversion.total_messaging_connection")
    msg_new = get_action_value(actions, "onsite_conversion.messaging_first_reply")
    purchases = get_action_value(actions, "purchase")
    cost_purchase = get_action_value(cpa, "purchase")

    return [
        adset.get("id", ""),
        since,
        until,
        adset.get("name", ""),
        adset.get("effective_status", adset.get("status", "")),
        result_value if result_value is not None else "",
        result_indicator,
        cost_per_result if cost_per_result is not None else "",
        budget_value,
        budget_type,
        float(ins.get("spend", 0) or 0),
        int(ins.get("impressions", 0) or 0),
        int(ins.get("reach", 0) or 0),
        adset.get("end_time", "En curso"),
        adset.get("start_time", ""),
        attribution_label(adset.get("attribution_spec", [])),
        (float(adset["bid_amount"]) / 100) if adset.get("bid_amount") else 0,
        adset.get("bid_strategy", ""),
        adset.get("updated_time", ""),
        msg_total if msg_total is not None else "",
        msg_new if msg_new is not None else "",
        purchases if purchases is not None else "",
        cost_purchase if cost_purchase is not None else "",
        now_iso,
    ]


# ----------------------------------------------------------------------
# Google Sheets
# ----------------------------------------------------------------------
def get_worksheet(sheet_id, sa_json):
    creds_info = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=5000, cols=len(COLUMNS))
    return ws


def read_existing_rows(ws):
    """Lee las filas existentes (sin encabezado) de la Sheet."""
    try:
        all_values = ws.get_all_values()
    except Exception:
        return []
    if len(all_values) <= 1:
        return []
    return all_values[1:]  # sin encabezado


def last_col_letter(n):
    """Convierte número de columna (1-based) en letra(s) de Excel/Sheets."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def write_to_sheet(ws, rows):
    ws.clear()
    payload = [COLUMNS] + rows
    ws.update(values=payload, range_name="A1")
    header_range = f"A1:{last_col_letter(len(COLUMNS))}1"
    ws.format(header_range, {"textFormat": {"bold": True}})
    print(f"✅ Escritas {len(rows)} filas en '{WORKSHEET_NAME}'.")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    token = os.environ["META_ACCESS_TOKEN"]
    account_id = os.environ["META_AD_ACCOUNT_ID"]
    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

    # Configuración de fechas
    start_ym_str = os.environ.get("START_YEAR_MONTH", "2026-01")
    sy, sm = start_ym_str.split("-")
    start_ym = (int(sy), int(sm))
    grace_days = int(os.environ.get("GRACE_DAYS", "7"))

    today = datetime.now(timezone.utc).date()
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    print(f"📅 Hoy: {today.isoformat()} | Inicio histórico: {start_ym_str} | Gracia: {grace_days} días")

    # 1. Determinar qué meses procesar
    all_months = list_months(start_ym, today)
    active_months = [(y, m) for (y, m) in all_months if is_month_active(y, m, today, grace_days)]
    frozen_months = [(y, m) for (y, m) in all_months if (y, m) not in active_months]

    print(f"📊 Meses totales: {len(all_months)} | Activos: {len(active_months)} | Congelados: {len(frozen_months)}")
    print(f"   → Activos: {active_months}")
    print(f"   → Congelados: {frozen_months}")

    # 2. Leer datos existentes de la Sheet (para conservar los meses congelados)
    ws = get_worksheet(sheet_id, sa_json)
    existing = read_existing_rows(ws)
    print(f"📖 Filas existentes en Sheet: {len(existing)}")

    # Detectar si la Sheet existente ya tiene el formato NUEVO (con ID en la col 0).
    # Si el encabezado guardado no coincide, forzamos backfill de todo (no preservamos filas viejas).
    same_format = True
    try:
        header = ws.row_values(1)
        if header and header[0] != COLUMNS[0]:
            same_format = False
    except Exception:
        same_format = False

    # En el formato nuevo, las fechas están en las columnas 1 y 2 (no 0 y 1).
    date_idx = (1, 2) if same_format else (0, 1)

    existing_keys = set()
    if same_format:
        for row in existing:
            if len(row) >= 3:
                existing_keys.add((row[date_idx[0]], row[date_idx[1]]))

    preserved_rows = []
    months_to_backfill = []
    for (y, m) in frozen_months:
        s, e = month_start_end(y, m)
        if same_format and (s, e) in existing_keys:
            for row in existing:
                if len(row) >= 3 and (row[date_idx[0]], row[date_idx[1]]) == (s, e):
                    preserved_rows.append(row)
        else:
            months_to_backfill.append((y, m))

    print(f"❄️  Filas congeladas conservadas: {len(preserved_rows)}")
    if not same_format:
        print("⚠️  El formato de la Sheet cambió (se agregó columna ID). Se recalculará todo el histórico.")
    if months_to_backfill:
        print(f"🔄 Meses a rellenar (backfill): {months_to_backfill}")

    # 3. Descargar metadatos de adsets
    print("📥 Descargando lista de conjuntos de anuncios...")
    adsets = fetch_adsets(account_id, token)
    print(f"   → {len(adsets)} adsets encontrados.")

    # 4. Descargar insights de los meses activos + backfill, construir filas nuevas
    months_to_fetch = months_to_backfill + active_months
    new_rows = []
    for (y, m) in months_to_fetch:
        print(f"📥 Descargando insights de {y}-{m:02d}...")
        insights, since, until = fetch_insights_for_month(account_id, token, y, m)
        print(f"   → {len(insights)} insights")

        insights_by_id = {i["adset_id"]: i for i in insights}

        for ads in adsets:
            ins = insights_by_id.get(ads["id"])
            if ins is None:
                continue
            row = build_row(ads, ins, since, until, now_iso)
            new_rows.append(row)

    print(f"🆕 Filas nuevas construidas: {len(new_rows)}")

    # 5. Combinar: preservadas + nuevas, y ordenar
    all_rows = preserved_rows + new_rows
    # Ordenar por nombre del adset (col 3) y luego por fecha de inicio (col 1)
    all_rows.sort(key=lambda r: (r[3] if len(r) > 3 else "", r[1] if len(r) > 1 else ""))

    print(f"📤 Total filas a escribir: {len(all_rows)}")

    # 6. Escribir a la Sheet
    write_to_sheet(ws, all_rows)

    print("🎉 Listo.")


if __name__ == "__main__":
    main()
