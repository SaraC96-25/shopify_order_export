import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime, date
import time

# --- CONFIGURAZIONE ---
SHOP_URL = st.secrets["SHOP_URL"]
API_VERSION = st.secrets["API_VERSION"]
ACCESS_TOKEN = st.secrets["ACCESS_TOKEN"]

headers = {
    "X-Shopify-Access-Token": ACCESS_TOKEN,
    "Content-Type": "application/json"
}

# --- FUNZIONI ---
def get_orders():
    orders = []
    url = f"{SHOP_URL}/admin/api/{API_VERSION}/orders.json?status=any&limit=50"

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            st.error(f"Errore {response.status_code}: {response.text}")
            break

        data = response.json().get("orders", [])
        orders.extend(data)

        # Estrai link alla pagina successiva dal header
        link_header = response.headers.get("Link")
        if link_header and 'rel="next"' in link_header:
            parts = link_header.split(",")
            next_url = None
            for part in parts:
                if 'rel="next"' in part:
                    next_url = part.split(";")[0].strip().strip("<>")
                    break
            url = next_url
        else:
            url = None

    return orders

def get_events(order_id):
    url = f"{SHOP_URL}/admin/api/{API_VERSION}/orders/{order_id}/events.json"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json().get("events", [])
        else:
            print(f"[ERRORE] Eventi non trovati per ordine {order_id}: {response.status_code}")
            return []
    except requests.exceptions.RequestException as e:
        print(f"[FALLITO] Errore di rete per ordine {order_id}: {e}")
        return []

def estrai_commenti_con_ca(data_inizio, data_fine):
    orders = get_orders()
    dati_filtrati = []
    progress = st.progress(0)

    for i, order in enumerate(orders):
        order_id = order["id"]
        order_name = order["name"]
        created_at_str = order["created_at"]
        created_at = datetime.fromisoformat(created_at_str.replace("Z", "+00:00")).date()

        # Filtra per intervallo di date
        if not (data_inizio <= created_at <= data_fine):
            continue

        eventi = get_events(order_id)
        time.sleep(0.3)

        if not eventi:
            print(f"[INFO] Nessun evento per ordine {order_name}")

        for ev in eventi:
            autore = ev.get("author", "").strip().lower()
            messaggio = ev.get("message", "").strip().lower()
            print(f"[DEBUG] Ordine: {order_name}, Autore: {autore}, Messaggio: {messaggio}")
            if (
                "ca" in messaggio and
                "chiara" in autore and "azzaretto" in autore
            ):
                dati_filtrati.append({
                    "Numero Ordine": order_name,
                    "Data Ordine": created_at_str,
                    "Commento": ev["message"],
                    "Data Commento": ev["created_at"]
                })
            elif "ca" in messaggio:
                print(f"[SCARTATO] Autore non compatibile: {autore} → ordine {order_name}")

        progress.progress((i + 1) / len(orders))

    return pd.DataFrame(dati_filtrati)

# --- STREAMLIT UI ---
st.title("Esporta ordini con commento 'ca' da Chiara Azzaretto")
st.write("Filtra per data e genera un file Excel con gli ordini che contengono un commento con 'ca' inserito da Chiara Azzaretto.")

# --- Filtro data ---
col1, col2 = st.columns(2)
with col1:
    data_inizio = st.date_input("Data inizio", value=date.today().replace(day=1))
with col2:
    data_fine = st.date_input("Data fine", value=date.today())

# Validazione: data_inizio deve essere <= data_fine
if data_inizio > data_fine:
    st.error("La data di inizio non può essere successiva alla data di fine.")
else:
    if st.button("Esporta Excel"):
        with st.spinner("Recupero ordini..."):
            df = estrai_commenti_con_ca(data_inizio, data_fine)

        if df.empty:
            st.warning("Nessun ordine con commento contenente 'ca' da Chiara Azzaretto trovato.")
        else:
            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            st.download_button(
                label="Scarica Excel",
                data=buffer,
                file_name="ordini_con_ca.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
