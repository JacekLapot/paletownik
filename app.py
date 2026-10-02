import io
import json
import re
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openai import OpenAI

st.set_page_config(
    page_title="Paletownik AI",
    page_icon="📦",
    layout="wide",
)

FIELDS = [
    "Lp.", "Ilość", "Kategoria", "Marka", "Produkt", "Model", "Stan",
    "Kompletność", "Cena nowego", "Cena używanego",
    "Realna cena sprzedaży", "Cena wystawienia", "Źródło ceny",
    "Link do oferty", "Uwagi", "Priorytet"
]

SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "kategoria": {"type": "string"},
        "marka": {"type": "string"},
        "produkt": {"type": "string"},
        "model": {"type": "string"},
        "stan": {"type": "string"},
        "kompletnosc": {"type": "string"},
        "cena_nowego": {"type": "number"},
        "cena_uzywanego": {"type": "number"},
        "realna_cena_sprzedazy": {"type": "number"},
        "cena_wystawienia": {"type": "number"},
        "zrodlo_ceny": {"type": "string"},
        "link_do_oferty": {"type": "string"},
        "uwagi": {"type": "string"},
        "priorytet": {
            "type": "string",
            "enum": ["Wysoki", "Normalny", "Niski"]
        },
        "pewnosc_ident": {
            "type": "string",
            "enum": ["Wysoka", "Średnia", "Niska"]
        }
    },
    "required": [
        "kategoria", "marka", "produkt", "model", "stan", "kompletnosc",
        "cena_nowego", "cena_uzywanego", "realna_cena_sprzedazy",
        "cena_wystawienia", "zrodlo_ceny", "link_do_oferty", "uwagi",
        "priorytet", "pewnosc_ident"
    ]
}

PROMPT = """
Jesteś asystentem do wyceny produktów z palet zwrotów konsumenckich w Polsce.

Przeanalizuj zdjęcie produktu. Zidentyfikuj markę, produkt i model możliwie
dokładnie. Jeśli modelu nie da się potwierdzić, wpisz pusty string zamiast
zgadywać.

Następnie użyj wyszukiwania internetowego i sprawdź AKTUALNE ceny w Polsce.
Priorytet źródeł: Allegro, OLX, Ceneo, polskie sklepy, oficjalny producent.
Szukaj przede wszystkim dokładnego modelu. Jeśli go nie ma, użyj porównywalnych
ofert i wyraźnie zaznacz to w uwagach.

Zasady wyceny:
- Cena nowego = realna aktualna cena nowego egzemplarza, nie cena MSRP.
- Cena używanego = typowa cena kompletnego i sprawnego używanego egzemplarza.
- Realna cena sprzedaży = konserwatywna kwota, którą realnie można uzyskać
  przy sprzedaży z palety w Polsce.
- Cena wystawienia = cena początkowa trochę wyższa od realnej ceny sprzedaży.
- Nie zawyżaj wartości przez pojedynczą drogą ofertę.
- Jeśli rynek jest nasycony tanimi ofertami, obniż wycenę.
- Nie zakładaj, że produkt jest kompletny tylko dlatego, że wygląda na nowy.
- Stan z jednego zdjęcia oznacz jako "Nowy / do sprawdzenia" lub
  "Nieznany / do sprawdzenia".
- Nie uwzględniaj ilości; ilość zawsze ustala użytkownik.
- Link do oferty ma być prawdziwym linkiem znalezionym w wyszukiwaniu, jeśli
  jest dostępny. Jeśli nie ma dokładnego linku, zostaw pusty.
- W uwagach wyjaśnij niepewności i czy cena opiera się na dokładnym modelu
  czy na porównywalnych ofertach.
- Zwróć wyłącznie JSON zgodny ze schematem.
"""


def normalize(value):
    value = str(value or "").lower().strip()
    return re.sub(r"[^a-z0-9ąćęłńóśźż]+", "", value)


def get_client():
    try:
        key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        key = ""
    if not key:
        return None
    return OpenAI(api_key=key)


def find_duplicate(data):
    new_model = normalize(data.get("model"))
    new_brand = normalize(data.get("marka"))
    new_product = normalize(data.get("produkt"))

    if not new_model and not new_product:
        return None

    for idx, row in enumerate(st.session_state.products):
        old_brand = normalize(row.get("Marka"))
        old_product = normalize(row.get("Produkt"))
        old_model = normalize(row.get("Model"))

        if new_model and old_model and new_model == old_model:
            if new_brand and old_brand and new_brand != old_brand:
                continue
            return idx

        if (
            new_brand and old_brand and new_brand == old_brand
            and new_product and old_product and new_product == old_product
        ):
            return idx

    return None


def analyze_image(image_bytes, mime_type):
    client = get_client()
    if client is None:
        raise RuntimeError(
            "Brak OPENAI_API_KEY. Dodaj go w Streamlit Cloud → Settings → Secrets."
        )

    b64 = __import__("base64").b64encode(image_bytes).decode("utf-8")
    data_url = f"data:{mime_type};base64,{b64}"

    response = client.responses.create(
        model="gpt-5.6-luna",
        tools=[{"type": "web_search"}],
        input=[{
            "role": "user",
            "content": [
                {"type": "input_text", "text": PROMPT},
                {
                    "type": "input_image",
                    "image_url": data_url,
                    "detail": "high"
                }
            ]
        }],
        text={
            "format": {
                "type": "json_schema",
                "name": "product_valuation",
                "strict": True,
                "schema": SCHEMA
            }
        }
    )
    return json.loads(response.output_text)


def add_new_product(data, quantity=1):
    row = {
        "Lp.": len(st.session_state.products) + 1,
        "Ilość": int(quantity),
        "Kategoria": data["kategoria"],
        "Marka": data["marka"],
        "Produkt": data["produkt"],
        "Model": data["model"],
        "Stan": data["stan"],
        "Kompletność": data["kompletnosc"],
        "Cena nowego": data["cena_nowego"],
        "Cena używanego": data["cena_uzywanego"],
        "Realna cena sprzedaży": data["realna_cena_sprzedazy"],
        "Cena wystawienia": data["cena_wystawienia"],
        "Źródło ceny": data["zrodlo_ceny"],
        "Link do oferty": data["link_do_oferty"],
        "Uwagi": (
            data["uwagi"]
            + f" | Pewność identyfikacji: {data['pewnosc_ident']}"
        ),
        "Priorytet": data["priorytet"],
    }
    st.session_state.products.append(row)


def update_lps():
    for i, row in enumerate(st.session_state.products, 1):
        row["Lp."] = i


def total_value():
    total = 0.0
    for row in st.session_state.products:
        total += float(row["Ilość"]) * float(row["Realna cena sprzedaży"])
    return total


def make_excel(cost):
    wb = Workbook()
    ws = wb.active
    ws.title = "Produkty"
    ws.append(FIELDS)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in st.session_state.products:
        ws.append([row.get(f, "") for f in FIELDS])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = [
        7, 9, 22, 18, 30, 18, 22, 25, 16, 18, 24, 18, 25, 45, 45, 16
    ]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[__import__("openpyxl").utils.get_column_letter(i)].width = width

    summary = wb.create_sheet("Podsumowanie")
    summary["A1"] = "PODSUMOWANIE PALETY"
    summary["A1"].font = Font(bold=True, size=16)
    summary["A2"] = "Nazwa palety"
    summary["B2"] = st.session_state.current_pallet_name
    summary["A3"] = "Liczba pozycji"
    summary["B3"] = len(st.session_state.products)
    summary["A4"] = "Liczba sztuk"
    summary["B4"] = sum(int(r["Ilość"]) for r in st.session_state.products)
    summary["A5"] = "Koszt palety"
    summary["B5"] = cost
    summary["A6"] = "Realna wartość sprzedaży"
    summary["B6"] = total_value()
    summary["A7"] = "Nadwyżka przed kosztami sprzedaży"
    summary["B7"] = total_value() - cost
    summary.column_dimensions["A"].width = 40
    summary.column_dimensions["B"].width = 20

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ---------- Session state ----------
if "products" not in st.session_state:
    st.session_state.products = []
if "pending" not in st.session_state:
    st.session_state.pending = None
if "camera_enabled" not in st.session_state:
    st.session_state.camera_enabled = True
if "last_analyzed_hash" not in st.session_state:
    st.session_state.last_analyzed_hash = None
if "pallet_history" not in st.session_state:
    st.session_state.pallet_history = []
if "current_pallet_name" not in st.session_state:
    st.session_state.current_pallet_name = "Nowa paleta"
if "current_pallet_cost" not in st.session_state:
    st.session_state.current_pallet_cost = 360.0
if "pallet_id" not in st.session_state:
    st.session_state.pallet_id = 0


def save_current_pallet():
    if not st.session_state.products:
        return False

    # Zapisujemy kopię danych, żeby późniejsza edycja bieżącej palety
    # nie zmieniała historycznego zapisu.
    snapshot = [dict(row) for row in st.session_state.products]
    entry = {
        "name": st.session_state.current_pallet_name.strip() or "Bez nazwy",
        "cost": float(st.session_state.current_pallet_cost),
        "products": snapshot,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    st.session_state.pallet_history.insert(0, entry)
    return True


def start_new_pallet():
    save_current_pallet()
    st.session_state.products = []
    st.session_state.pending = None
    st.session_state.last_analyzed_hash = None
    st.session_state.current_pallet_name = "Nowa paleta"
    st.session_state.current_pallet_cost = 360.0
    st.session_state.pallet_id += 1


def load_history_pallet(index):
    entry = st.session_state.pallet_history[index]
    st.session_state.products = [dict(row) for row in entry["products"]]
    st.session_state.pending = None
    st.session_state.last_analyzed_hash = None
    st.session_state.current_pallet_name = entry["name"]
    st.session_state.current_pallet_cost = float(entry["cost"])
    st.session_state.pallet_id += 1


# ---------- Sidebar ----------
with st.sidebar:
    st.header("⚙️ Palety")

    # Nie używamy klucza widgetu jako zmiennej stanu palety.
    # Dzięki temu funkcja „Nowa paleta” może bezpiecznie zmienić nazwę/koszt.
    pallet_name_input = st.text_input(
        "Nazwa bieżącej palety",
        value=st.session_state.current_pallet_name,
        placeholder="np. Paleta 01 - elektronika"
    )
    st.session_state.current_pallet_name = pallet_name_input

    pallet_cost = st.number_input(
        "Koszt palety (zł)",
        min_value=0.0,
        value=float(st.session_state.current_pallet_cost),
        step=10.0
    )
    st.session_state.current_pallet_cost = float(pallet_cost)

    if st.button("🆕 Nowa paleta", use_container_width=True, type="primary"):
        start_new_pallet()
        st.rerun()

    st.divider()
    st.subheader("📚 Historia palet")

    if st.session_state.pallet_history:
        history_labels = [
            f"{i + 1}. {entry['name']} — {entry['created_at']}"
            for i, entry in enumerate(st.session_state.pallet_history)
        ]
        selected_history = st.selectbox(
            "Wybierz zapisaną paletę",
            history_labels,
            key="history_selector"
        )
        history_index = history_labels.index(selected_history)

        if st.button("📂 Wczytaj wybraną paletę", use_container_width=True):
            load_history_pallet(history_index)
            st.rerun()

        entry = st.session_state.pallet_history[history_index]
        st.caption(
            f"{len(entry['products'])} pozycji • "
            f"{sum(int(r['Ilość']) for r in entry['products'])} szt. • "
            f"wartość {sum(float(r['Ilość']) * float(r['Realna cena sprzedaży']) for r in entry['products']):.0f} zł"
        )
    else:
        st.info("Brak zapisanych palet.")

    st.divider()
    st.write("**Status API:**")
    if get_client():
        st.success("OPENAI_API_KEY ustawiony")
    else:
        st.error("Brak OPENAI_API_KEY")

# ---------- Header ----------
st.title("📦 Paletownik AI")
st.subheader(f"🗂️ {st.session_state.current_pallet_name}")
st.caption("Zdjęcie → rozpoznanie → ceny → duplikaty → wartość palety")

total = total_value()
items = len(st.session_state.products)
units = sum(int(r["Ilość"]) for r in st.session_state.products)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Pozycje", items)
m2.metric("Sztuki", units)
m3.metric("Wartość palety", f"{total:,.0f} zł")
m4.metric("Nadwyżka", f"{total - pallet_cost:,.0f} zł")

st.divider()

# ---------- Photo input ----------
c1, c2 = st.columns(2)

with c1:
    st.subheader("📸 Zrób zdjęcie")
    camera = st.camera_input(
        "Aparat",
        key=f"camera_{st.session_state.pallet_id}",
        resolution="720p"
    )

with c2:
    st.subheader("📁 Wybierz z urządzenia")
    upload = st.file_uploader(
        "Zdjęcie produktu",
        type=["jpg", "jpeg", "png", "webp"],
        key=f"uploader_{st.session_state.pallet_id}"
    )

image_file = camera if camera is not None else upload

if image_file is not None:
    image_bytes = image_file.getvalue()
    image_hash = hashlib.sha256(image_bytes).hexdigest()
    st.image(image_file, caption="Wybrane zdjęcie", width="stretch")

    # Automatyczna analiza: nowe zdjęcie uruchamia AI bez klikania przycisku.
    # Hash zapobiega ponownemu wywołaniu API przy każdym rerunie Streamlit.
    if image_hash != st.session_state.last_analyzed_hash:
        st.session_state.last_analyzed_hash = image_hash
        with st.spinner("🤖 AI rozpoznaje produkt i sprawdza aktualne ceny..."):
            try:
                result = analyze_image(
                    image_bytes,
                    getattr(image_file, "type", "image/jpeg")
                )
                st.session_state.pending = result
                st.rerun()
            except Exception as exc:
                st.session_state.last_analyzed_hash = None
                st.error(f"Nie udało się przeanalizować zdjęcia: {exc}")
    elif st.session_state.pending is None:
        st.info("Zdjęcie jest gotowe. AI analizuje je automatycznie po wybraniu nowego zdjęcia.")

# ---------- Pending result ----------
if st.session_state.pending:
    data = st.session_state.pending
    st.divider()
    st.subheader("🔎 Wynik AI")

    p1, p2 = st.columns([2, 1])

    with p1:
        st.write(
            f"**{data['marka']} {data['produkt']}**"
            + (f" — **{data['model']}**" if data["model"] else "")
        )
        st.write(f"Kategoria: **{data['kategoria']}**")
        st.write(f"Stan: **{data['stan']}**")
        st.write(f"Kompletność: **{data['kompletnosc']}**")
        st.write(f"Pewność identyfikacji: **{data['pewnosc_ident']}**")
        if data["uwagi"]:
            st.info(data["uwagi"])
        if data["link_do_oferty"]:
            st.markdown(f"[Przykładowa oferta]({data['link_do_oferty']})")

    with p2:
        st.metric("Cena nowego", f"{data['cena_nowego']:.0f} zł")
        st.metric("Cena używanego", f"{data['cena_uzywanego']:.0f} zł")
        st.metric(
            "Realna sprzedaż / szt.",
            f"{data['realna_cena_sprzedazy']:.0f} zł"
        )
        st.metric(
            "Cena wystawienia",
            f"{data['cena_wystawienia']:.0f} zł"
        )

    duplicate_idx = find_duplicate(data)

    if duplicate_idx is not None:
        existing = st.session_state.products[duplicate_idx]
        st.warning(
            f"⚠️ **Ten produkt już istnieje w palecie.** "
            f"{existing['Marka']} {existing['Produkt']} "
            f"{existing['Model']} — obecnie **{existing['Ilość']} szt.**"
        )

        d1, d2, d3 = st.columns([1, 1, 2])
        with d1:
            add_qty = st.number_input(
                "Dodaj sztuk",
                min_value=1,
                value=1,
                step=1,
                key="duplicate_add_qty"
            )
        with d2:
            if st.button("➕ Dodaj", type="primary", use_container_width=True):
                existing["Ilość"] += int(add_qty)
                st.session_state.pending = None
                st.rerun()
        with d3:
            if st.button("➕ Dodaj jako osobną pozycję", use_container_width=True):
                add_new_product(data, int(add_qty))
                st.session_state.pending = None
                st.rerun()

    else:
        qty = st.number_input(
            "Ilość sztuk",
            min_value=1,
            value=1,
            step=1,
            key="new_product_qty"
        )
        st.caption(
            f"Wartość pozycji: "
            f"**{int(qty) * data['realna_cena_sprzedazy']:.0f} zł**"
        )

        a1, a2 = st.columns(2)
        with a1:
            if st.button(
                "✅ Dodaj do palety",
                type="primary",
                use_container_width=True
            ):
                add_new_product(data, int(qty))
                st.session_state.pending = None
                st.rerun()
        with a2:
            if st.button("❌ Odrzuć wynik AI", use_container_width=True):
                st.session_state.pending = None
                st.rerun()

# ---------- Inventory ----------
st.divider()
st.subheader("📋 Zawartość palety")

if st.session_state.products:
    df = pd.DataFrame(st.session_state.products)
    df["Wartość pozycji"] = (
        df["Ilość"].astype(float) *
        df["Realna cena sprzedaży"].astype(float)
    )
    display_cols = [
        "Lp.", "Ilość", "Kategoria", "Marka", "Produkt", "Model",
        "Realna cena sprzedaży", "Wartość pozycji", "Cena wystawienia",
        "Priorytet"
    ]
    st.dataframe(
        df[display_cols],
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### 🛠️ Ręczna edycja")
    options = [
        f"{r['Lp.']}. {r['Marka']} {r['Produkt']} {r['Model']}"
        for r in st.session_state.products
    ]
    selected = st.selectbox("Wybierz produkt", options)

    idx = options.index(selected)
    row = st.session_state.products[idx]

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        new_qty = st.number_input(
            "Ilość",
            min_value=1,
            value=int(row["Ilość"]),
            step=1,
            key=f"edit_qty_{idx}"
        )
    with e2:
        new_real = st.number_input(
            "Realna sprzedaż / szt.",
            min_value=0.0,
            value=float(row["Realna cena sprzedaży"]),
            step=5.0,
            key=f"edit_real_{idx}"
        )
    with e3:
        new_listing = st.number_input(
            "Cena wystawienia",
            min_value=0.0,
            value=float(row["Cena wystawienia"]),
            step=5.0,
            key=f"edit_listing_{idx}"
        )
    with e4:
        priority = st.selectbox(
            "Priorytet",
            ["Wysoki", "Normalny", "Niski"],
            index=["Wysoki", "Normalny", "Niski"].index(row["Priorytet"]),
            key=f"edit_priority_{idx}"
        )

    if st.button("💾 Zapisz zmiany", use_container_width=True):
        row["Ilość"] = int(new_qty)
        row["Realna cena sprzedaży"] = float(new_real)
        row["Cena wystawienia"] = float(new_listing)
        row["Priorytet"] = priority
        st.rerun()

    if st.button("🗑️ Usuń wybraną pozycję", use_container_width=True):
        st.session_state.products.pop(idx)
        update_lps()
        st.rerun()
else:
    st.info("Paleta jest pusta. Zrób pierwsze zdjęcie produktu.")

# ---------- Export ----------
if st.session_state.products:
    st.divider()
    st.subheader("📥 Eksport")

    excel_bytes = make_excel(pallet_cost)
    safe_name = re.sub(r"[^a-zA-Z0-9ąćęłńóśźżĄĆĘŁŃÓŚŹŻ _-]+", "", st.session_state.current_pallet_name).strip().replace(" ", "_") or "paleta"
    filename = f"{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.download_button(
        "📊 Pobierz Excel",
        data=excel_bytes,
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.caption(
        "Dane palety są przechowywane w bieżącej sesji Streamlit. "
        "Pobierz Excel przed zamknięciem/odświeżeniem strony."
    )
