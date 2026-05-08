"""
Meta Ads → Google Sheets
=========================
Descarga los insights de los conjuntos de anuncios (Ad Sets) de Meta Ads
y los guarda en una Google Sheet. Pensado para correr automáticamente
en GitHub Actions varias veces al día.

Variables de entorno necesarias (configuradas como Secrets en GitHub):
- META_ACCESS_TOKEN     : Token de larga duración de Meta (Graph API)
- META_AD_ACCOUNT_ID    : ID de la cuenta publicitaria SIN el "act_" (ej. 554289509584085)
- GOOGLE_SHEET_ID       : ID de la Google Sheet (sale en la URL)
- GOOGLE_SERVICE_ACCOUNT_JSON : Contenido completo del JSON de la cuenta de servicio
- DAYS_BACK             : (opcional) Cuántos días hacia atrás traer. Default 30.

Autor: Generado para Aron - Refriperu
"""

import os
import sys
import json
from datetime import datetime, timedelta, timezone

import requests
import gspread
from google.oauth2.service_account import Credentials


# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------
META_API_VERSION = "v19.0"
META_BASE_URL = f"https://graph.facebook.com/{META_API_VERSION}"

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
    "Última actualización",  # extra: timestamp de cuándo se trajo la fila
]


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
        # Paginación
        paging = data.get("paging", {})
        next_url = paging.get("next")
        url = next_url
        params = None  # el next URL ya tiene los params
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


def fetch_insights(account_id, token, days_back=30):
    """Trae los insights (métricas de rendimiento) por adset en el rango de fechas."""
    since = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime("%Y-%m-%d")
    until = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    fields = ",".join([
        "adset_id", "adset_name",
        "date_start", "date_stop",
        "impressions", "reach", "spend",
        "actions", "action_values",
        "cost_per_action_type",
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
    """Extrae un valor específico de la lista de actions de Meta."""
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
    """Convierte attribution_spec en texto legible (similar a Ads Manager)."""
    if not spec:
        return ""
    parts = []
    for item in spec:
        click = item.get("event_type", "")
        click_window = item.get("window_days", "")
        parts.append(f"{click_window} días {click}")
    return " | ".join(parts)


def budget_label(adset):
    """Devuelve ('valor', 'tipo de presupuesto') igual que el Excel."""
    if adset.get("daily_budget"):
        # Meta devuelve los presupuestos en centavos (cents)
        return float(adset["daily_budget"]) / 100, "Diario"
    if adset.get("lifetime_budget"):
        return float(adset["lifetime_budget"]) / 100, "Total"
    return "", "Con el presupuesto de la campaña"


def build_rows(adsets, insights, since, until):
    """Combina metadatos de adsets con insights y arma las filas finales."""
    insights_by_id = {i["adset_id"]: i for i in insights}
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    rows = []
    for ads in adsets:
        ins = insights_by_id.get(ads["id"], {})
        budget_value, budget_type = budget_label(ads)

        # Métrica principal de "Resultados" según el objetivo
        actions = ins.get("actions", [])
        cpa = ins.get("cost_per_action_type", [])

        # Heurística: tomar la action más relevante si existe
        # Ads Manager muestra: messaging_conversation_started_7d, leadgen.other, purchase, etc.
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

        # Si no se encontró ninguna acción, usar reach como fallback
        if result_value is None and ins.get("reach"):
            result_value = float(ins.get("reach", 0))
            result_indicator = "reach"
            spend = float(ins.get("spend", 0) or 0)
            cost_per_result = (spend / result_value) if result_value else None

        msg_total = get_action_value(actions, "onsite_conversion.total_messaging_connection")
        msg_new = get_action_value(actions, "onsite_conversion.messaging_first_reply")
        purchases = get_action_value(actions, "purchase")
        cost_purchase = get_action_value(cpa, "purchase")

        row = [
            since,                                              # Inicio del informe
            until,                                              # Fin del informe
            ads.get("name", ""),                                # Nombre del conjunto
            ads.get("effective_status", ads.get("status", "")), # Entrega
            result_value if result_value is not None else "",   # Resultados
            result_indicator,                                   # Indicador de resultado
            cost_per_result if cost_per_result is not None else "",  # Costo por resultados
            budget_value,                                       # Presupuesto
            budget_type,                                        # Tipo de presupuesto
            float(ins.get("spend", 0) or 0),                    # Importe gastado USD
            int(ins.get("impressions", 0) or 0),                # Impresiones
            int(ins.get("reach", 0) or 0),                      # Alcance
            ads.get("end_time", "En curso"),                    # Finalización
            ads.get("start_time", ""),                          # Inicio
            attribution_label(ads.get("attribution_spec", [])), # Configuración de atribución
            (float(ads["bid_amount"]) / 100) if ads.get("bid_amount") else 0,  # Puja
            ads.get("bid_strategy", ""),                        # Tipo de puja
            ads.get("updated_time", ""),                        # Último cambio significativo
            msg_total if msg_total is not None else "",         # Contactos mensajes total
            msg_new if msg_new is not None else "",             # Nuevos contactos
            purchases if purchases is not None else "",         # Compras
            cost_purchase if cost_purchase is not None else "", # Costo por compra
            now_iso,                                            # Última actualización
        ]
        rows.append(row)
    return rows


# ----------------------------------------------------------------------
# Google Sheets
# ----------------------------------------------------------------------
def write_to_sheet(sheet_id, sa_json, rows):
    """Escribe los datos en la primera hoja, sobrescribiendo todo el contenido."""
    creds_info = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet("Meta_Ads_Adsets")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Meta_Ads_Adsets", rows=2000, cols=len(COLUMNS))

    ws.clear()
    payload = [COLUMNS] + rows
    ws.update(values=payload, range_name="A1")

    # Formato simple: encabezado en negrita
    ws.format("A1:W1", {"textFormat": {"bold": True}})
    print(f"✅ Escritas {len(rows)} filas en la hoja 'Meta_Ads_Adsets'.")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    token = os.environ["META_ACCESS_TOKEN"]
    account_id = os.environ["META_AD_ACCOUNT_ID"]
    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    days_back = int(os.environ.get("DAYS_BACK", "30"))

    print(f"📊 Descargando datos de Meta Ads (cuenta act_{account_id})...")
    adsets = fetch_adsets(account_id, token)
    print(f"   → {len(adsets)} conjuntos de anuncios encontrados.")

    insights, since, until = fetch_insights(account_id, token, days_back)
    print(f"   → {len(insights)} filas de insights ({since} a {until}).")

    rows = build_rows(adsets, insights, since, until)

    print("📤 Subiendo a Google Sheets...")
    write_to_sheet(sheet_id, sa_json, rows)

    print("🎉 Listo.")


if __name__ == "__main__":
    main()
