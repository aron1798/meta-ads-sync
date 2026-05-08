"""
Meta Ads → Google Sheets (Versión MENSUAL con período de gracia)
=================================================================
Descarga los insights de los conjuntos de anuncios (Ad Sets) de Meta Ads
y los guarda en una Google Sheet, AGRUPADOS POR MES.
 
LÓGICA:
- Una fila por (conjunto de anuncios × mes).
- Meses con más de 7 días cerrados quedan CONGELADOS (no se vuelven a pedir a Meta).
- Mes actual + meses con menos de 7 días cerrados se actualizan en cada corrida.
- Empieza desde START_YEAR_MONTH (configurable, default enero del año en curso).
 
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
COLUMNS = [
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
        # Mes actual o futuro
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
 
    actions = ins.get("actions", [])
    cpa = ins.get("cost_per_action_type", [])
 
    result_value = None
    result_indicator = ""
    cost_per_result = None
    priority_actions = [
        ("onsite_conversion.messaging_conversation_started_7d", "Conversaciones con mensajes iniciadas"),
        ("lead", "Clientes potenciales"),
        ("leadgen.other", "Clientes potenciales"),
        ("purchase", "Compras"),
        ("link_click", "Clics en el enlace"),
        ("post_engagement", "Interacciones con la publicación"),
        ("page_engagement", "Interacciones con la página"),
        ("reach", "Alcance"),
    ]
    for action_type, label in priority_actions:
        v = get_action_value(actions, action_type)
        if v is not None:
            result_value = v
            result_indicator = label
            cost_per_result = get_action_value(cpa, action_type)
            break
 
    if result_value is None and ins.get("reach"):
        result_value = float(ins.get("reach", 0))
        result_indicator = "reach"
        spend = float(ins.get("spend", 0) or 0)
        cost_per_result = (spend / result_value) if result_value else None
 
    msg_total = get_action_value(actions, "onsite_conversion.total_messaging_connection")
    msg_new = get_action_value(actions, "onsite_conversion.messaging_first_reply")
    purchases = get_action_value(actions, "purchase")
    cost_purchase = get_action_value(cpa, "purchase")
 
    return [
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
 
 
def write_to_sheet(ws, rows):
    ws.clear()
    payload = [COLUMNS] + rows
    ws.update(values=payload, range_name="A1")
    ws.format("A1:W1", {"textFormat": {"bold": True}})
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
 
    # Filtrar filas congeladas: las que pertenezcan a meses congelados se conservan
    frozen_keys = set()
    for (y, m) in frozen_months:
        s, e = month_start_end(y, m)
        frozen_keys.add((s, e))
 
    preserved_rows = []
    for row in existing:
        if len(row) < 2:
            continue
        key = (row[0], row[1])  # (Inicio, Fin)
        if key in frozen_keys:
            preserved_rows.append(row)
 
    print(f"❄️  Filas congeladas conservadas: {len(preserved_rows)}")
 
    # 3. Descargar metadatos de adsets
    print(f"📥 Descargando lista de conjuntos de anuncios...")
    adsets = fetch_adsets(account_id, token)
    print(f"   → {len(adsets)} adsets encontrados.")
 
    # 4. Descargar insights de los meses activos y construir filas nuevas
    new_rows = []
    for (y, m) in active_months:
        print(f"📥 Descargando insights de {y}-{m:02d}...")
        insights, since, until = fetch_insights_for_month(account_id, token, y, m)
        print(f"   → {len(insights)} insights")
 
        insights_by_id = {i["adset_id"]: i for i in insights}
 
        for ads in adsets:
            ins = insights_by_id.get(ads["id"])
            # Solo agregamos la fila si hay datos del adset en ese mes
            # (si no hubo actividad, no creamos fila vacía)
            if ins is None:
                continue
            row = build_row(ads, ins, since, until, now_iso)
            new_rows.append(row)
 
    print(f"🆕 Filas nuevas construidas: {len(new_rows)}")
 
    # 5. Combinar: preservadas + nuevas, y ordenar
    all_rows = preserved_rows + new_rows
    # Ordenar por nombre del adset (col 2) y luego por fecha de inicio (col 0)
    all_rows.sort(key=lambda r: (r[2] if len(r) > 2 else "", r[0] if len(r) > 0 else ""))
 
    print(f"📤 Total filas a escribir: {len(all_rows)}")
 
    # 6. Escribir a la Sheet
    write_to_sheet(ws, all_rows)
 
    print("🎉 Listo.")
 
 
if __name__ == "__main__":
    main()
 
