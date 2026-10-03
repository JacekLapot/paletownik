import io
import json
import re
import hashlib
import os
from html import escape
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openai import OpenAI
from sqlalchemy import text
from PIL import Image

st.set_page_config(page_title="Paletownia", page_icon="📦", layout="wide")

# Paletownia PWA metadata
st.markdown(
    """
    <link rel="manifest" href="app/static/manifest.json">
    <meta name="application-name" content="Paletownia">
    <meta name="apple-mobile-web-app-title" content="Paletownia">
    <meta name="theme-color" content="#2563eb">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <link rel="apple-touch-icon" href="app/static/icon-192.png">
    """,
    unsafe_allow_html=True,
)


FIELDS = [
    "Lp.", "Ilość", "Kategoria", "Marka", "Produkt", "Model", "Stan",
    "Kompletność", "Cena nowego", "Cena używanego", "Realna cena sprzedaży",
    "Cena wystawienia", "Źródło ceny", "Link do oferty", "Uwagi", "Priorytet",
    "Status sprzedaży", "Cena sprzedaży", "Data wystawienia", "Data sprzedaży"
]
PRIORITIES = ["🟡 Priorytet", "🟢 Ważne", "🟠 Mogą poczekać", "🔴 Badziew"]

LEVELS = {
    1: "🟡 Priorytet",
    2: "🟢 Ważne",
    3: "🟠 Mogą poczekać",
    4: "🔴 Badziew",
}

def level_label(price):
    price = float(price or 0)
    if price >= 250: return LEVELS[1]
    if price >= 150: return LEVELS[2]
    if price >= 50: return LEVELS[3]
    return LEVELS[4]
OLD_PRIORITY_MAP = {
    "Wysoki":"🟡 Priorytet", "Normalny":"🟢 Ważne", "Niski":"🟠 Mogą poczekać",
    "Najważniejsze do sprzedaży":"🟡 Priorytet", "Ważne":"🟢 Ważne",
    "Mogą poczekać":"🟠 Mogą poczekać", "Drobnica / badziew":"🔴 Badziew",
    "Poziom 1":"🟡 Priorytet", "Poziom 2":"🟢 Ważne", "Poziom 3":"🟠 Mogą poczekać", "Poziom 4":"🔴 Badziew"
}

def priority_from_price(price):
    """Automatyczne oznaczenie na podstawie realnej ceny sprzedaży za sztukę."""
    return level_label(price)


SCHEMA = {
    "type":"object", "additionalProperties":False,
    "properties": {
        "kategoria":{"type":"string"}, "marka":{"type":"string"}, "produkt":{"type":"string"},
        "model":{"type":"string"}, "stan":{"type":"string"}, "kompletnosc":{"type":"string"},
        "cena_nowego":{"type":"number"}, "cena_uzywanego":{"type":"number"},
        "realna_cena_sprzedazy":{"type":"number"}, "cena_wystawienia":{"type":"number"},
        "zrodlo_ceny":{"type":"string"}, "link_do_oferty":{"type":"string"}, "uwagi":{"type":"string"},
        "priorytet":{"type":"string","enum":PRIORITIES},
        "pewnosc_ident":{"type":"string","enum":["Wysoka","Średnia","Niska"]}
    },
    "required":["kategoria","marka","produkt","model","stan","kompletnosc","cena_nowego","cena_uzywanego",
                 "realna_cena_sprzedazy","cena_wystawienia","zrodlo_ceny","link_do_oferty","uwagi","priorytet","pewnosc_ident"]
}

PROMPT = """
Jesteś asystentem do wyceny produktów z palet zwrotów konsumenckich w Polsce.
Przeanalizuj zdjęcie produktu. Zidentyfikuj markę, produkt i model możliwie dokładnie. Jeśli modelu nie da się potwierdzić, wpisz pusty string zamiast zgadywać.
Użyj wyszukiwania internetowego i sprawdź AKTUALNE ceny w Polsce. Priorytet źródeł: Allegro, OLX, Ceneo, polskie sklepy, oficjalny producent. Szukaj przede wszystkim dokładnego modelu. Jeśli go nie ma, użyj porównywalnych ofert i zaznacz to w uwagach.
Zasady: cena nowego = realna aktualna cena, nie MSRP; cena używanego = typowa cena kompletnego sprawnego egzemplarza; realna cena sprzedaży = konserwatywna kwota możliwa do uzyskania w Polsce; cena wystawienia = trochę wyższa; nie zawyżaj pojedynczą drogą ofertą; przy dużej konkurencji obniż wycenę; nie zakładaj kompletności ze zdjęcia; stan ze zdjęcia oznacz jako Nowy / do sprawdzenia lub Nieznany / do sprawdzenia; link ma być prawdziwy, jeśli znaleziony.
Poziom wartości: wyznacz go WYŁĄCZNIE na podstawie realnej ceny sprzedaży za sztukę: >=250 zł = Poziom 1; >=150 zł i <250 zł = Poziom 2; >=50 zł i <150 zł = Poziom 3; <50 zł = Poziom 4. Nie wybieraj poziomu na podstawie atrakcyjności produktu, marki ani łatwości sprzedaży.
Zwróć wyłącznie JSON zgodny ze schematem.
"""

DB_DDL = """
CREATE TABLE IF NOT EXISTS pallets (
    id BIGSERIAL PRIMARY KEY,
    name TEXT NOT NULL DEFAULT 'Nowa paleta',
    cost NUMERIC(12,2) NOT NULL DEFAULT 360,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS products (
    id BIGSERIAL PRIMARY KEY,
    pallet_id BIGINT NOT NULL REFERENCES pallets(id) ON DELETE CASCADE,
    position INTEGER NOT NULL,
    quantity INTEGER NOT NULL DEFAULT 1,
    category TEXT NOT NULL DEFAULT '',
    brand TEXT NOT NULL DEFAULT '',
    product TEXT NOT NULL DEFAULT '',
    model TEXT NOT NULL DEFAULT '',
    state TEXT NOT NULL DEFAULT '',
    completeness TEXT NOT NULL DEFAULT '',
    new_price NUMERIC(12,2) NOT NULL DEFAULT 0,
    used_price NUMERIC(12,2) NOT NULL DEFAULT 0,
    real_sale_price NUMERIC(12,2) NOT NULL DEFAULT 0,
    listing_price NUMERIC(12,2) NOT NULL DEFAULT 0,
    price_source TEXT NOT NULL DEFAULT '',
    offer_link TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    priority TEXT NOT NULL DEFAULT 'Ważne',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    image_thumb TEXT NOT NULL DEFAULT '',
    sale_status TEXT NOT NULL DEFAULT 'Na stanie',
    sold_price NUMERIC(12,2) NOT NULL DEFAULT 0,
    listed_at TIMESTAMPTZ NULL,
    sold_at TIMESTAMPTZ NULL
);
ALTER TABLE products ADD COLUMN IF NOT EXISTS image_thumb TEXT NOT NULL DEFAULT '';
ALTER TABLE products ADD COLUMN IF NOT EXISTS sale_status TEXT NOT NULL DEFAULT 'Na stanie';
ALTER TABLE products ADD COLUMN IF NOT EXISTS sold_price NUMERIC(12,2) NOT NULL DEFAULT 0;
ALTER TABLE products ADD COLUMN IF NOT EXISTS listed_at TIMESTAMPTZ NULL;
ALTER TABLE products ADD COLUMN IF NOT EXISTS sold_at TIMESTAMPTZ NULL;
CREATE INDEX IF NOT EXISTS idx_products_pallet_id ON products(pallet_id);
"""


def db_conn():
    try:
        return st.connection("postgresql", type="sql")
    except Exception as exc:
        st.error("Brak połączenia z bazą PostgreSQL. Dodaj [connections.postgresql] do Secrets w Streamlit.")
        st.code(str(exc))
        return None


def db_init(conn):
    if st.session_state.get("_db_initialized"):
        return
    with conn.session as s:
        for statement in [x.strip() for x in DB_DDL.split(';') if x.strip()]:
            s.execute(text(statement))
        s.commit()
    st.session_state["_db_initialized"] = True


def _cache_version():
    # Zmiana tej wersji unieważnia tylko dane bieżącej sesji po zapisie.
    return int(st.session_state.get("_data_version", 0))


def invalidate_data_cache():
    st.session_state["_data_version"] = _cache_version() + 1


def db_query(conn, sql, params=None, ttl=5):
    # SQLConnection.query cache'uje wynik. Dodajemy wersję sesji do komentarza SQL,
    # dzięki czemu po zapisie użytkownik od razu dostaje świeże dane.
    versioned_sql = f"{sql.rstrip()}\n/* paletownia_data_version={_cache_version()} */"
    return conn.query(versioned_sql, params=params or {}, ttl=ttl)


def ensure_first_pallet(conn):
    df = db_query(conn, "SELECT id FROM pallets ORDER BY id LIMIT 1")
    if df.empty:
        with conn.session as s:
            result = s.execute(text("INSERT INTO pallets (name, cost) VALUES (:name, :cost) RETURNING id"),
                               {"name":"Nowa paleta", "cost":360})
            pid = result.scalar_one()
            s.commit()
        invalidate_data_cache()
        return int(pid)
    return int(df.iloc[0]["id"])


def load_pallets(conn):
    return db_query(conn, "SELECT id, name, cost, created_at, updated_at FROM pallets ORDER BY updated_at DESC, id DESC")


def load_products(conn, pallet_id):
    sql = """
    SELECT id, position AS \"Lp.\", quantity AS \"Ilość\", category AS \"Kategoria\", brand AS \"Marka\",
           product AS \"Produkt\", model AS \"Model\", state AS \"Stan\", completeness AS \"Kompletność\",
           new_price AS \"Cena nowego\", used_price AS \"Cena używanego\", real_sale_price AS \"Realna cena sprzedaży\",
           listing_price AS \"Cena wystawienia\", price_source AS \"Źródło ceny\", offer_link AS \"Link do oferty\",
           notes AS \"Uwagi\", priority AS \"Priorytet\", image_thumb AS \"Miniatura\",
           sale_status AS \"Status sprzedaży\", sold_price AS \"Cena sprzedaży\", listed_at AS \"Data wystawienia\", sold_at AS \"Data sprzedaży\", created_at AS \"Data dodania\"
    FROM products WHERE pallet_id = :pallet_id ORDER BY position, id
    """
    df = db_query(conn, sql, {"pallet_id": int(pallet_id)})
    rows = df.to_dict("records")
    for r in rows:
        for k in ["Cena nowego","Cena używanego","Realna cena sprzedaży","Cena wystawienia"]:
            r[k] = float(r[k] or 0)
        r["Priorytet"] = priority_from_price(r["Realna cena sprzedaży"])
        r["Ilość"] = int(r["Ilość"] or 1)
        r["Cena sprzedaży"] = float(r.get("Cena sprzedaży") or 0)
        r["Status sprzedaży"] = str(r.get("Status sprzedaży") or "Na stanie")
        r["Lp."] = int(r["Lp."])
        r["_db_id"] = int(r["id"])
        r.pop("id", None)
    return rows


def make_thumbnail_data_url(image_bytes, max_size=1000):
    """Tworzy wysokiej jakości podgląd JPEG do przechowywania w PostgreSQL.
    Obraz jest ograniczany do 1000 px na dłuższym boku, dzięki czemu
    pozostaje wyraźny w katalogu i na telefonie, ale nie robi się niepotrzebnie ciężki.
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=90, optimize=True, progressive=True)
        b64 = __import__("base64").b64encode(out.getvalue()).decode("ascii")
        return "data:image/jpeg;base64," + b64
    except Exception:
        return ""


def save_product(conn, pallet_id, data, quantity=1, image_thumb=""):
    row = {
        "position": 1,
        "quantity": int(quantity), "category": data["kategoria"], "brand": data["marka"],
        "product": data["produkt"], "model": data["model"], "state": data["stan"],
        "completeness": data["kompletnosc"], "new_price": data["cena_nowego"], "used_price": data["cena_uzywanego"],
        "real_sale_price": data["realna_cena_sprzedazy"], "listing_price": data["cena_wystawienia"],
        "price_source": data["zrodlo_ceny"], "offer_link": data["link_do_oferty"],
        "notes": data["uwagi"] + f" | Pewność identyfikacji: {data['pewnosc_ident']}", "priority": priority_from_price(data["realna_cena_sprzedazy"]),
        "image_thumb": image_thumb or "", "sale_status": "Na stanie"
    }
    with conn.session as s:
        max_position = s.execute(
            text("SELECT COALESCE(MAX(position), 0) + 1 FROM products WHERE pallet_id=:p"),
            {"p": int(pallet_id)},
        ).scalar_one()
        row["position"] = int(max_position)
        result = s.execute(text("""
            INSERT INTO products (pallet_id, position, quantity, category, brand, product, model, state, completeness,
                new_price, used_price, real_sale_price, listing_price, price_source, offer_link, notes, priority, image_thumb, sale_status)
            VALUES (:pallet_id,:position,:quantity,:category,:brand,:product,:model,:state,:completeness,
                :new_price,:used_price,:real_sale_price,:listing_price,:price_source,:offer_link,:notes,:priority,:image_thumb,:sale_status)
            RETURNING id
        """), {"pallet_id":int(pallet_id), **row})
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=:id"), {"id":int(pallet_id)})
        s.commit()
    invalidate_data_cache()
    return int(result.scalar_one())


def update_product(conn, product_id, quantity, real_price, listing_price, priority=None, offer_link=None, product_name=None):
    priority = priority_from_price(real_price)
    with conn.session as s:
        params={"q":int(quantity),"r":float(real_price),"l":float(listing_price),"p":priority,"id":int(product_id)}
        sets="quantity=:q, real_sale_price=:r, listing_price=:l, priority=:p"
        if offer_link is not None:
            sets += ", offer_link=:o"
            params["o"] = str(offer_link or "").strip()
        if product_name is not None:
            sets += ", product=:product"
            params["product"] = str(product_name or "").strip()
        s.execute(text(f"UPDATE products SET {sets}, updated_at=NOW() WHERE id=:id"), params)
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=(SELECT pallet_id FROM products WHERE id=:id)"), {"id":int(product_id)})
        s.commit()
    invalidate_data_cache()


def update_product_thumbnail(conn, product_id, image_thumb, replace=False):
    if not image_thumb:
        return
    with conn.session as s:
        if replace:
            s.execute(text("UPDATE products SET image_thumb=:img, updated_at=NOW() WHERE id=:id"),
                      {"img": image_thumb, "id": int(product_id)})
        else:
            s.execute(text("UPDATE products SET image_thumb=:img, updated_at=NOW() WHERE id=:id AND COALESCE(image_thumb, '')=''"),
                      {"img": image_thumb, "id": int(product_id)})
        s.commit()
    invalidate_data_cache()


def update_sale_status(conn, product_id, status, sold_price=0):
    """Zmienia status sprzedaży produktu i zapisuje cenę sprzedaży."""
    allowed = {"Na stanie", "Wystawiony", "Sprzedany"}
    status = status if status in allowed else "Na stanie"
    with conn.session as s:
        params = {"id": int(product_id), "status": status, "sold_price": float(sold_price or 0)}
        if status == "Wystawiony":
            sql = """UPDATE products SET sale_status=:status, sold_price=0, listed_at=COALESCE(listed_at, NOW()), sold_at=NULL, updated_at=NOW() WHERE id=:id"""
        elif status == "Sprzedany":
            sql = """UPDATE products SET sale_status=:status, sold_price=:sold_price, sold_at=NOW(), updated_at=NOW() WHERE id=:id"""
        else:
            sql = """UPDATE products SET sale_status=:status, sold_price=0, sold_at=NULL, updated_at=NOW() WHERE id=:id"""
        s.execute(text(sql), params)
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=(SELECT pallet_id FROM products WHERE id=:id)"), {"id": int(product_id)})
        s.commit()
    invalidate_data_cache()


def delete_product(conn, product_id):
    with conn.session as s:
        row = s.execute(text("SELECT pallet_id FROM products WHERE id=:id"), {"id":int(product_id)}).first()
        if not row: return
        pallet_id = int(row[0])
        s.execute(text("DELETE FROM products WHERE id=:id"), {"id":int(product_id)})
        s.execute(text("""WITH numbered AS (SELECT id, ROW_NUMBER() OVER (ORDER BY position,id) AS rn FROM products WHERE pallet_id=:p)
                        UPDATE products SET position=numbered.rn FROM numbered WHERE products.id=numbered.id"""), {"p":pallet_id})
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=:p"), {"p":pallet_id})
        s.commit()
    invalidate_data_cache()


def update_pallet(conn, pallet_id, name, cost):
    with conn.session as s:
        s.execute(text("UPDATE pallets SET name=:name, cost=:cost, updated_at=NOW() WHERE id=:id"),
                  {"name":name.strip() or "Bez nazwy", "cost":float(cost), "id":int(pallet_id)})
        s.commit()
    invalidate_data_cache()


def create_pallet(conn, name="Nowa paleta", cost=360):
    with conn.session as s:
        result = s.execute(text("INSERT INTO pallets (name,cost) VALUES (:name,:cost) RETURNING id"),
                           {"name":name.strip() or "Nowa paleta", "cost":float(cost)})
        pid = int(result.scalar_one())
        s.commit()
    invalidate_data_cache()
    return pid


def delete_pallet(conn, pallet_id):
    with conn.session as s:
        s.execute(text("DELETE FROM pallets WHERE id=:id"), {"id":int(pallet_id)})
        s.commit()
    invalidate_data_cache()


def total_value(products):
    return sum(float(r["Ilość"])*float(r["Realna cena sprzedaży"]) for r in products)


def normalize(v):
    return re.sub(r"[^a-z0-9ąćęłńóśźż]+", "", str(v or "").lower().strip())


def find_duplicate(products, data):
    nm, nb, np = normalize(data.get("model")), normalize(data.get("marka")), normalize(data.get("produkt"))
    for i, row in enumerate(products):
        om, ob, op = normalize(row.get("Model")), normalize(row.get("Marka")), normalize(row.get("Produkt"))
        if nm and om and nm == om and (not nb or not ob or nb == ob): return i
        if nb and ob and nb == ob and np and op and np == op: return i
    return None


@st.cache_resource
def get_client():
    try:
        key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        key = ""
    return OpenAI(api_key=key) if key else None


def analyze_image(image_bytes, mime_type):
    client = get_client()
    if not client: raise RuntimeError("Brak OPENAI_API_KEY w Secrets.")
    b64 = __import__("base64").b64encode(image_bytes).decode()
    response = client.responses.create(
        model="gpt-5.6-luna", tools=[{"type":"web_search"}],
        input=[{"role":"user","content":[{"type":"input_text","text":PROMPT},
            {"type":"input_image","image_url":f"data:{mime_type};base64,{b64}","detail":"high"}]}],
        text={"format":{"type":"json_schema","name":"product_valuation","strict":True,"schema":SCHEMA}}
    )
    return json.loads(response.output_text)


def make_excel(products, pallet_name, cost):
    wb=Workbook(); ws=wb.active; ws.title="Produkty"; ws.append(FIELDS)
    fill=PatternFill("solid",fgColor="D9EAF7")
    for c in ws[1]: c.font=Font(bold=True); c.fill=fill; c.alignment=Alignment(horizontal="center")
    for row in products: ws.append([row.get(f,"") for f in FIELDS])
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for i,w in enumerate([7,9,22,18,30,18,22,25,16,18,24,18,25,45,45,28,18,18,20,20],1): ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width=w
    s=wb.create_sheet("Podsumowanie"); s["A1"]="PODSUMOWANIE PALETY"; s["A1"].font=Font(bold=True,size=16)
    vals=[("Nazwa palety",pallet_name),("Liczba pozycji",len(products)),("Liczba sztuk",sum(int(r['Ilość']) for r in products)),("Koszt palety",cost),("Realna wartość sprzedaży",total_value(products)),("Nadwyżka przed kosztami sprzedaży",total_value(products)-cost)]
    for i,(a,b) in enumerate(vals,2): s.cell(i,1,a); s.cell(i,2,b)
    s.column_dimensions['A'].width=40; s.column_dimensions['B'].width=25
    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()


def load_public_products(conn):
    """Publiczny katalog: tylko dane potrzebne klientowi."""
    sql = """
    SELECT p.id, p.quantity, p.brand, p.product, p.model,
           p.real_sale_price, p.image_thumb, p.created_at
    FROM products p
    WHERE p.quantity > 0
      AND p.real_sale_price > 0
      AND COALESCE(p.sale_status, 'Na stanie') <> 'Sprzedany'
    ORDER BY p.updated_at DESC, p.id DESC
    """
    df = db_query(conn, sql)
    rows = df.to_dict("records")
    for r in rows:
        r["id"] = int(r["id"])
        r["quantity"] = int(r["quantity"] or 0)
        r["real_sale_price"] = float(r["real_sale_price"] or 0)
    return rows


def public_product_name(r):
    return " ".join(
        str(x).strip()
        for x in [r.get("brand", ""), r.get("product", ""), r.get("model", "")]
        if str(x).strip()
    ) or "Produkt"


def _secret_value(name, default=""):
    """Odczytuje sekret z Streamlit Secrets lub środowiska."""
    try:
        value = st.secrets.get(name, None)
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    value = os.environ.get(name, "")
    if str(value).strip():
        return str(value).strip()

    # Dopuszczamy również sekcję [resend] w Secrets, gdyby ktoś tak ją skonfigurował.
    if name == "RESEND_API_KEY":
        try:
            section = st.secrets.get("resend", {})
            if isinstance(section, dict):
                value = section.get("api_key", "")
                if str(value).strip():
                    return str(value).strip()
        except Exception:
            pass
    return str(default).strip()


def send_product_inquiry(product, customer_name, customer_email, customer_phone="", customer_message=""):
    """Wysyła zapytanie klienta przez Gmail SMTP z użyciem hasła aplikacji Google."""
    import smtplib
    from email.message import EmailMessage

    smtp_username = _secret_value("SMTP_USERNAME", "jacek.lapot@gmail.com")
    smtp_password = _secret_value("SMTP_PASSWORD")
    recipient = _secret_value("CONTACT_EMAIL", "alfa.alwaysfair@gmail.com")
    smtp_host = _secret_value("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(_secret_value("SMTP_PORT", "587"))

    if not smtp_password:
        raise RuntimeError(
            "Brak SMTP_PASSWORD w Streamlit Secrets. Dodaj 16-znakowe hasło aplikacji Google "
            "dla konta jacek.lapot@gmail.com."
        )

    product_name = public_product_name(product)
    price = float(product.get("real_sale_price", 0) or 0)
    phone_line = customer_phone.strip() if customer_phone else "Nie podano"
    message_text = customer_message.strip() or "Nie podano"

    msg = EmailMessage()
    msg["Subject"] = f"Zapytanie o produkt — {product_name}"
    msg["From"] = smtp_username
    msg["To"] = recipient
    msg["Reply-To"] = customer_email.strip()
    msg.set_content(
        "Nowe zapytanie z Paletownii\n\n"
        f"Produkt: {product_name}\n"
        f"Cena: {price:.2f} zł\n"
        f"Imię: {customer_name.strip()}\n"
        f"E-mail: {customer_email.strip()}\n"
        f"Telefon: {phone_line}\n\n"
        f"Wiadomość od kupującego:\n{message_text}\n\n"
        "Wiadomość została wysłana z publicznego katalogu Paletownii."
    )

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
    except smtplib.SMTPAuthenticationError as exc:
        raise RuntimeError(
            "Gmail odrzucił logowanie. Sprawdź SMTP_USERNAME oraz czy SMTP_PASSWORD "
            "jest 16-znakowym hasłem aplikacji Google."
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Gmail SMTP: {exc}") from exc

    return True


def render_inquiry_form(product):
    """Formularz zapytania o konkretny produkt."""
    st.markdown("### ✉️ Zapytaj o produkt")
    st.markdown(f"**{escape(public_product_name(product))}** · **{float(product.get('real_sale_price', 0) or 0):,.0f} zł**")

    with st.form("product_inquiry_form", clear_on_submit=True):
        customer_name = st.text_input("Imię *", placeholder="Np. Jan")
        customer_email = st.text_input("Adres e-mail *", placeholder="Np. jan@example.com")
        customer_phone = st.text_input("Telefon (opcjonalnie)", placeholder="Np. 500 600 700")
        customer_message = st.text_area("Wiadomość", placeholder="Napisz, o co chcesz zapytać…", height=140)
        submitted = st.form_submit_button("📨 Wyślij zapytanie", type="primary", use_container_width=True)

        if submitted:
            name_ok = bool(customer_name.strip())
            email_ok = bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", customer_email.strip()))
            if not name_ok:
                st.error("Podaj imię.")
            elif not email_ok:
                st.error("Podaj poprawny adres e-mail.")
            else:
                try:
                    send_product_inquiry(product, customer_name, customer_email, customer_phone, customer_message)
                    st.session_state.inquiry_sent = True
                except Exception as exc:
                    st.error(f"Nie udało się wysłać wiadomości. {exc}")

    if st.session_state.get("inquiry_sent"):
        st.success("✅ Dziękujemy! Zapytanie zostało wysłane. Skontaktujemy się z Tobą.")
        st.session_state.inquiry_sent = False

    if st.button("← Wróć do katalogu", key="back_to_catalog"):
        st.session_state.selected_inquiry_product = None
        st.rerun()


def render_public_catalog(conn):
    """Domyślny, niezalogowany widok katalogu."""
    public_products = load_public_products(conn)

    st.markdown("""
    <style>
    .public-hero {
        padding: 18px 20px;
        border-radius: 18px;
        background: linear-gradient(135deg, #eff6ff 0%, #ffffff 70%);
        border: 1px solid #dbeafe;
        margin-bottom: 18px;
    }
    .public-card {
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 12px;
        background: #fff;
        box-shadow: 0 1px 4px rgba(0,0,0,.05);
        height: 100%;
    }
    .public-name { font-weight: 700; font-size: 1rem; line-height: 1.3; margin-top: 8px; }
    .public-price { font-size: 1.25rem; font-weight: 800; margin-top: 7px; }
    .public-meta { color: #6b7280; font-size: .86rem; margin-top: 3px; }
    </style>
    """, unsafe_allow_html=True)

    top_left, top_right = st.columns([5, 1])
    with top_left:
        st.markdown('<div class="public-hero"><h1 style="margin:0">📦 Paletownia</h1><div style="color:#64748b;margin-top:4px">Produkty dostępne w sprzedaży</div></div>', unsafe_allow_html=True)
    with top_right:
        st.write("")
        st.caption("")
        with st.popover("🔐 Zaloguj", use_container_width=True):
            st.markdown("**Panel właściciela**")
            pin = st.text_input("PIN", type="password", key="admin_pin_input", label_visibility="collapsed", placeholder="Wpisz PIN")
            if st.button("Zaloguj", type="primary", use_container_width=True, key="public_login"):
                configured_pin = str(st.secrets.get("ADMIN_PIN", "")).strip()
                if configured_pin and pin == configured_pin:
                    st.session_state.admin_logged_in = True
                    st.session_state.admin_login_error = False
                    st.rerun()
                else:
                    st.session_state.admin_login_error = True
                    st.error("Nieprawidłowy PIN.")

    if not public_products:
        st.info("Aktualnie nie ma produktów dostępnych w katalogu.")
        if not str(st.secrets.get("ADMIN_PIN", "")).strip():
            st.caption("Panel właściciela nie jest jeszcze skonfigurowany.")
        return

    total_units = sum(r["quantity"] for r in public_products)
    c1, c2 = st.columns(2)
    c1.metric("Dostępne pozycje", len(public_products))
    c2.metric("Dostępne sztuki", total_units)

    search = st.text_input("🔎 Szukaj produktu", placeholder="np. Tikom, słuchawki, EZVIZ...", key="public_search")
    sort_public = st.selectbox(
        "Sortowanie",
        ["Najnowsze", "Nazwa A–Z", "Cena — od najwyższej", "Cena — od najniższej"],
        key="public_sort",
    )

    filtered = public_products
    q = search.strip().lower()
    if q:
        filtered = [r for r in filtered if q in public_product_name(r).lower()]

    if sort_public == "Nazwa A–Z":
        filtered = sorted(filtered, key=lambda r: public_product_name(r).lower())
    elif sort_public == "Cena — od najwyższej":
        filtered = sorted(filtered, key=lambda r: r["real_sale_price"], reverse=True)
    elif sort_public == "Cena — od najniższej":
        filtered = sorted(filtered, key=lambda r: r["real_sale_price"])

    st.markdown(f"**Wyniki: {len(filtered)}**")
    if not filtered:
        st.info("Nie znaleziono produktu.")
        return

    # 4 mniejsze karty w jednym wierszu na desktopie.
    for start in range(0, len(filtered), 4):
        row = filtered[start:start + 4]
        cols = st.columns(4)
        for col, r in zip(cols, row):
            with col:
                thumb = str(r.get("image_thumb") or "")
                if thumb:
                    st.image(thumb, width="stretch")
                else:
                    st.markdown('<div style="height:130px;display:flex;align-items:center;justify-content:center;border:1px dashed #cbd5e1;border-radius:12px;color:#94a3b8;font-size:42px;">📦</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="public-name">{escape(public_product_name(r))}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="public-price">{r["real_sale_price"]:,.0f} zł</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="public-meta">Dostępne: <b>{r["quantity"]} szt.</b></div>', unsafe_allow_html=True)
                if st.button("Zapytaj o produkt", key=f"ask_product_{r["id"]}", use_container_width=True, type="secondary"):
                    st.session_state.selected_inquiry_product = int(r["id"])
                    st.rerun()
                st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    st.caption("Ceny dotyczą produktów widocznych jako dostępne w Paletownii.")

# Session/UI state only stores current selection and temporary AI result. Actual data is in PostgreSQL.
if "_data_version" not in st.session_state: st.session_state._data_version = 0
if "_db_initialized" not in st.session_state: st.session_state._db_initialized = False
if "current_pallet_id" not in st.session_state: st.session_state.current_pallet_id=None
if "pending" not in st.session_state: st.session_state.pending=None
if "last_analyzed_hash" not in st.session_state: st.session_state.last_analyzed_hash=None
if "last_added" not in st.session_state: st.session_state.last_added=None

conn=db_conn()
if conn:
    try:
        db_init(conn)
        if st.session_state.current_pallet_id is None:
            st.session_state.current_pallet_id=ensure_first_pallet(conn)
    except Exception as exc:
        st.error("Nie udało się zainicjalizować bazy danych.")
        st.code(str(exc)); st.stop()

# Domyślnie każdy użytkownik trafia do katalogu publicznego.
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False
if "selected_inquiry_product" not in st.session_state:
    st.session_state.selected_inquiry_product = None
if "inquiry_sent" not in st.session_state:
    st.session_state.inquiry_sent = False

if not st.session_state.admin_logged_in:
    if st.session_state.selected_inquiry_product is not None:
        public_products_for_form = load_public_products(conn)
        selected_product = next((r for r in public_products_for_form if r["id"] == int(st.session_state.selected_inquiry_product)), None)
        if selected_product is None:
            st.session_state.selected_inquiry_product = None
            st.rerun()
        else:
            render_inquiry_form(selected_product)
    else:
        render_public_catalog(conn)
    st.stop()

# Tryb administratora — pełny dotychczasowy interfejs.
admin_left, admin_right = st.columns([5, 1])
with admin_left:
    st.markdown("### 🔐 Paletownia — panel właściciela")
with admin_right:
    if st.button("Wyloguj", use_container_width=True, key="admin_logout"):
        st.session_state.admin_logged_in = False
        st.rerun()

pallets=load_pallets(conn)
if pallets.empty:
    st.session_state.current_pallet_id=ensure_first_pallet(conn); pallets=load_pallets(conn)

current_row=pallets[pallets["id"]==st.session_state.current_pallet_id]
if current_row.empty:
    st.session_state.current_pallet_id=int(pallets.iloc[0]["id"]); current_row=pallets[pallets["id"]==st.session_state.current_pallet_id]
current=current_row.iloc[0]
products=load_products(conn, int(st.session_state.current_pallet_id))

with st.sidebar:
    st.header("⚙️ Palety")
    name=st.text_input("Nazwa bieżącej palety", value=str(current["name"]), key="pallet_name_input")
    cost=st.number_input("Koszt palety (zł)", min_value=0.0, value=float(current["cost"]), step=10.0, key="pallet_cost_input")
    if st.button("💾 Zapisz nazwę i koszt", use_container_width=True):
        update_pallet(conn, st.session_state.current_pallet_id, name, cost); st.rerun()
    if st.button("🆕 Nowa paleta", use_container_width=True, type="primary"):
        pid=create_pallet(conn, "Nowa paleta", 360); st.session_state.current_pallet_id=pid; st.session_state.pending=None; st.session_state.last_analyzed_hash=None; st.session_state.last_added=None; st.rerun()
    st.divider(); st.subheader("📚 Historia palet")
    labels=[f"{int(r.id)} — {r['name']} — {str(r['updated_at'])[:16]}" for _,r in pallets.iterrows()]
    selected=st.selectbox("Wybierz paletę", labels, index=max(0,next((i for i,r in pallets.iterrows() if int(r.id)==st.session_state.current_pallet_id),0)), key="pallet_selector")
    selected_id=int(selected.split(" — ",1)[0])
    if selected_id != st.session_state.current_pallet_id:
        st.session_state.current_pallet_id=selected_id; st.session_state.pending=None; st.session_state.last_analyzed_hash=None; st.session_state.last_added=None; st.rerun()
    hist_row=pallets[pallets["id"]==selected_id].iloc[0]
    hist_products=load_products(conn, selected_id)
    st.caption(f"{len(hist_products)} pozycji • {sum(int(r['Ilość']) for r in hist_products)} szt. • wartość {total_value(hist_products):.0f} zł")
    if st.button("🗑️ Usuń wybraną paletę", use_container_width=True):
        if len(pallets)>1:
            delete_pallet(conn, selected_id); st.session_state.current_pallet_id=int(pallets[pallets['id']!=selected_id].iloc[0]['id']); st.session_state.pending=None; st.rerun()
        else: st.warning("Nie można usunąć jedynej palety.")
    st.divider(); st.write("**Status API:**")
    if get_client():
        st.success("OPENAI_API_KEY ustawiony")
    else:
        st.error("Brak OPENAI_API_KEY")

st.title("📦 Paletownia")
st.subheader(f"🗂️ {current['name']}")
total=total_value(products); units=sum(int(r['Ilość']) for r in products)
m1,m2,m3,m4=st.columns(4); m1.metric("Pozycje",len(products)); m2.metric("Sztuki",units); m3.metric("Wartość palety",f"{total:,.0f} zł"); m4.metric("Nadwyżka",f"{total-float(current['cost']):,.0f} zł")
st.divider()

c1,c2=st.columns(2)
with c1:
    st.subheader("📸 Zrób zdjęcie")
    camera=st.camera_input("Aparat", key=f"camera_{st.session_state.current_pallet_id}", resolution="720p")
with c2:
    st.subheader("📁 Wybierz z urządzenia")
    upload=st.file_uploader("Zdjęcie produktu", type=["jpg","jpeg","png","webp"], key=f"uploader_{st.session_state.current_pallet_id}")
image_file=camera if camera is not None else upload
if image_file is not None:
    b=image_file.getvalue(); h=hashlib.sha256(b).hexdigest(); st.image(image_file, caption="Wybrane zdjęcie", width="stretch")
    if h != st.session_state.last_analyzed_hash:
        st.session_state.last_analyzed_hash=h
        with st.spinner("🤖 AI rozpoznaje produkt i sprawdza aktualne ceny..."):
            try:
                data=analyze_image(b,getattr(image_file,"type","image/jpeg"))
                # Automatyczne dodanie do palety bez akceptacji.
                products_now=load_products(conn, int(st.session_state.current_pallet_id))
                thumb = make_thumbnail_data_url(b)
                duplicate_idx=find_duplicate(products_now,data)
                if duplicate_idx is not None:
                    existing=products_now[duplicate_idx]
                    new_qty=int(existing["Ilość"])+1
                    update_product(conn, existing["_db_id"], new_qty, existing["Realna cena sprzedaży"], existing["Cena wystawienia"], existing["Priorytet"])
                    if not existing.get("Miniatura"):
                        update_product_thumbnail(conn, existing["_db_id"], thumb)
                    st.session_state.last_added={
                        "kind":"duplicate", "name":f"{existing['Marka']} {existing['Produkt']} {existing['Model']}",
                        "quantity":new_qty, "price":float(existing["Realna cena sprzedaży"])
                    }
                else:
                    save_product(conn, int(st.session_state.current_pallet_id), data, 1, thumb)
                    st.session_state.last_added={
                        "kind":"new", "name":f"{data['marka']} {data['produkt']} {data['model']}".strip(),
                        "quantity":1, "price":float(data["realna_cena_sprzedazy"]), "link":data.get("link_do_oferty", "")
                    }
                st.rerun()
            except Exception as exc:
                st.session_state.last_analyzed_hash=None
                st.error(f"Nie udało się przeanalizować zdjęcia: {exc}")

if st.session_state.last_added:
    added=st.session_state.last_added
    if added.get("kind")=="duplicate":
        st.success(f"✅ Produkt rozpoznany jako duplikat i automatycznie dodany jako kolejna sztuka: **{added['name']}**. Łącznie: **{added['quantity']} szt.**")
    else:
        st.success(f"✅ Produkt automatycznie dodany do palety: **{added['name']}** — realna sprzedaż: **{added['price']:.0f} zł**")
        if added.get("link"):
            st.markdown(f"[Przykładowa oferta]({added['link']})")

st.divider(); st.subheader("📋 Zawartość palety")

sort_choice_mobile = st.selectbox(
    "Sortowanie",
    [
        "Data dodania — najnowsze",
        "Data dodania — najstarsze",
        "Nazwa A–Z",
        "Nazwa Z–A",
        "Wartość — od najwyższej",
        "Wartość — od najniższej",
    ],
    key="sort_choice_mobile",
)

def _product_name_for_sort(r):
    return " ".join(
        str(x).strip()
        for x in [r.get("Marka", ""), r.get("Produkt", ""), r.get("Model", "")]
        if str(x).strip()
    ).lower()

def _product_value_for_sort(r):
    return float(r.get("Ilość", 1) or 1) * float(r.get("Realna cena sprzedaży", 0) or 0)

def _product_date_for_sort(r):
    # load_products returns created_at when available; fallback keeps stable order.
    return str(r.get("created_at", r.get("Data dodania", "")) or "")

if products:
    if sort_choice_mobile == "Nazwa A–Z":
        products = sorted(products, key=_product_name_for_sort)
    elif sort_choice_mobile == "Nazwa Z–A":
        products = sorted(products, key=_product_name_for_sort, reverse=True)
    elif sort_choice_mobile == "Wartość — od najwyższej":
        products = sorted(products, key=_product_value_for_sort, reverse=True)
    elif sort_choice_mobile == "Wartość — od najniższej":
        products = sorted(products, key=_product_value_for_sort)
    elif sort_choice_mobile == "Data dodania — najstarsze":
        products = sorted(products, key=_product_date_for_sort)
    else:
        products = sorted(products, key=_product_date_for_sort, reverse=True)
if products:
    sold_rows = [r for r in products if str(r.get("Status sprzedaży") or "Na stanie") == "Sprzedany"]
    listed_rows = [r for r in products if str(r.get("Status sprzedaży") or "Na stanie") == "Wystawiony"]
    available_rows = [r for r in products if str(r.get("Status sprzedaży") or "Na stanie") != "Sprzedany"]
    sm1, sm2, sm3 = st.columns(3)
    with sm1: st.metric("🟢 Wystawione", len(listed_rows))
    with sm2: st.metric("🔴 Sprzedane", len(sold_rows))
    with sm3: st.metric("💰 Przychód ze sprzedanych", f"{sum(float(r.get('Cena sprzedaży',0) or 0) for r in sold_rows):.0f} zł")
    # Mobile-first card view: much easier to scan on a phone than a wide dataframe.
    st.markdown("""
    <style>
    .product-card { border:1px solid #e5e7eb; border-radius:16px; padding:12px; margin:8px 0; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.05); }
    .product-name { font-size:1.02rem; font-weight:700; line-height:1.25; margin-bottom:5px; }
    .product-meta { color:#6b7280; font-size:.86rem; line-height:1.35; }
    .product-price { font-size:1.18rem; font-weight:800; margin-top:5px; }
    .level-badge { display:inline-block; padding:4px 9px; border-radius:999px; font-size:.78rem; font-weight:700; margin:2px 0 4px; }
    .lvl1 { background:#fff1bf; color:#6b5200; }
    .lvl2 { background:#dcfce7; color:#166534; }
    .lvl3 { background:#ffedd5; color:#9a3412; }
    .lvl4 { background:#fee2e2; color:#991b1b; }
    @media (max-width: 640px) {
      .product-card { padding:10px; border-radius:14px; }
      .product-name { font-size:.98rem; }
      .product-price { font-size:1.1rem; }
    }
    </style>
    """, unsafe_allow_html=True)

    def level_css(label):
        if "Priorytet" in label: return "lvl1"
        if "Ważne" in label: return "lvl2"
        if "Mogą poczekać" in label: return "lvl3"
        return "lvl4"

    for r in products:
        level = level_label(float(r["Realna cena sprzedaży"]))
        cls = level_css(level)
        name = " ".join(str(x).strip() for x in [r.get("Marka", ""), r.get("Produkt", ""), r.get("Model", "")] if str(x).strip())
        thumb = str(r.get("Miniatura") or "")

        with st.container(border=True):
            photo_col, details_col = st.columns([1.05, 2.6], vertical_alignment="center")
            with photo_col:
                if thumb:
                    st.image(thumb, width="stretch")
                    with st.popover("🖼️ Zmień zdjęcie", help="Zmień zdjęcie przypisane do tej oferty"):
                        st.write(f"**Zmień zdjęcie:** {name}")
                        photo = st.file_uploader(
                            "Wybierz nowe zdjęcie",
                            type=["jpg", "jpeg", "png", "webp"],
                            key=f"card_replace_photo_{r['_db_id']}",
                            label_visibility="collapsed"
                        )
                        if photo is not None:
                            try:
                                new_thumb = make_thumbnail_data_url(photo.getvalue())
                                if new_thumb:
                                    update_product_thumbnail(conn, r["_db_id"], new_thumb, replace=True)
                                    st.success("Zdjęcie zostało zmienione.")
                                    st.rerun()
                            except Exception as exc:
                                st.error(f"Nie udało się zapisać zdjęcia: {exc}")
                else:
                    st.markdown('<div style="height:120px;display:flex;align-items:center;justify-content:center;border:1px dashed #cbd5e1;border-radius:12px;color:#94a3b8;font-size:32px;">📷</div>', unsafe_allow_html=True)
                    with st.popover("➕", help="Dodaj zdjęcie do tego produktu"):
                        st.write(f"**Dodaj zdjęcie:** {name}")
                        photo = st.file_uploader(
                            "Wybierz zdjęcie",
                            type=["jpg", "jpeg", "png", "webp"],
                            key=f"card_add_photo_{r['_db_id']}",
                            label_visibility="collapsed"
                        )
                        if photo is not None:
                            try:
                                new_thumb = make_thumbnail_data_url(photo.getvalue())
                                if new_thumb:
                                    update_product_thumbnail(conn, r["_db_id"], new_thumb, replace=True)
                                    st.success("Zdjęcie dodane.")
                                    st.rerun()
                            except Exception as exc:
                                st.error(f"Nie udało się zapisać zdjęcia: {exc}")
            with details_col:
                st.markdown(f'<div class="product-name">{int(r["Lp."])}. {name}</div>', unsafe_allow_html=True)
                st.markdown(f'<span class="level-badge {cls}">{level}</span>', unsafe_allow_html=True)
                st.markdown(f'<div class="product-price">{float(r["Realna cena sprzedaży"]):.0f} zł</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="product-meta">Ilość: <b>{int(r["Ilość"])}</b> · Wartość pozycji: <b>{float(r["Ilość"])*float(r["Realna cena sprzedaży"]):.0f} zł</b></div>', unsafe_allow_html=True)
                if str(r.get("Link do oferty") or "").strip():
                    try:
                        st.link_button("🔗 Zobacz ofertę", str(r["Link do oferty"]), use_container_width=True)
                    except Exception:
                        st.markdown(f'[🔗 Zobacz ofertę]({r["Link do oferty"]})')
                status = str(r.get("Status sprzedaży") or "Na stanie")
                status_icon = {"Na stanie":"⚪", "Wystawiony":"🟢", "Sprzedany":"🔴"}.get(status, "⚪")
                if status == "Sprzedany":
                    st.markdown(f"**{status_icon} Sprzedany** · cena sprzedaży: **{float(r.get('Cena sprzedaży',0)):.0f} zł**")
                else:
                    st.markdown(f"**{status_icon} {status}**")

            with st.expander("💰 Status sprzedaży", expanded=False):
                current_status = str(r.get("Status sprzedaży") or "Na stanie")
                new_status = st.selectbox(
                    "Status", ["Na stanie", "Wystawiony", "Sprzedany"],
                    index=["Na stanie", "Wystawiony", "Sprzedany"].index(current_status) if current_status in ["Na stanie", "Wystawiony", "Sprzedany"] else 0,
                    key=f"sale_status_{r['_db_id']}"
                )
                if new_status == "Sprzedany":
                    sold_price = st.number_input(
                        "Za ile sprzedano?", min_value=0.0, value=float(r.get("Cena sprzedaży",0) or 0), step=5.0,
                        key=f"sold_price_{r['_db_id']}"
                    )
                else:
                    sold_price = 0.0
                if st.button("💾 Zapisz status", key=f"save_sale_{r['_db_id']}", use_container_width=True):
                    update_sale_status(conn, r["_db_id"], new_status, sold_price)
                    st.success("Status sprzedaży zapisany.")
                    st.rerun()

            with st.expander("🛠️ Edycja produktu", expanded=False):
                new_product_name = st.text_input(
                    "Nazwa produktu",
                    value=str(r.get("Produkt", "") or ""),
                    key=f"card_edit_product_{r['_db_id']}"
                )
                e1, e2 = st.columns(2)
                with e1:
                    new_qty = st.number_input("Ilość", min_value=1, value=int(r["Ilość"]), step=1, key=f"card_edit_qty_{r['_db_id']}")
                with e2:
                    new_real = st.number_input("Realna sprzedaż / szt.", min_value=0.0, value=float(r["Realna cena sprzedaży"]), step=5.0, key=f"card_edit_real_{r['_db_id']}")
                e3, e4 = st.columns(2)
                with e3:
                    new_listing = st.number_input("Cena wystawienia", min_value=0.0, value=float(r["Cena wystawienia"]), step=5.0, key=f"card_edit_listing_{r['_db_id']}")
                with e4:
                    st.metric("Poziom", level_label(new_real))
                new_offer = st.text_input("Link do przykładowej oferty", value=str(r.get("Link do oferty", "") or ""), key=f"card_edit_offer_{r['_db_id']}")
                b1, b2 = st.columns(2)
                with b1:
                    if st.button("💾 Zapisz zmiany", key=f"card_save_{r['_db_id']}", use_container_width=True):
                        update_product(conn, r["_db_id"], new_qty, new_real, new_listing, None, new_offer, new_product_name)
                        st.rerun()
                with b2:
                    if st.button("🗑️ Usuń produkt", key=f"card_delete_{r['_db_id']}", use_container_width=True):
                        delete_product(conn, r["_db_id"])
                        st.rerun()

    st.caption("Poziom jest liczony automatycznie z realnej ceny sprzedaży za sztukę: 🟡 Priorytet ≥250 zł • 🟢 Ważne 150–249,99 zł • 🟠 Mogą poczekać 50–149,99 zł • 🔴 Badziew <50 zł.")

    with st.expander("🖼️ Zdjęcia produktów", expanded=False):
        st.caption("Miniatury są zapisane w bazie. Produkty bez zdjęcia możesz uzupełnić przyciskiem ➕ na karcie produktu.")
        for r in products:
            photo_col, name_col, info_col = st.columns([0.7, 4.8, 2.5])
            with photo_col:
                thumb = str(r.get("Miniatura") or "")
                if thumb:
                    st.image(thumb, width=58)
                else:
                    st.caption("brak")
            with name_col:
                st.markdown(f"**{r['Lp.']}. {r['Marka']} {r['Produkt']} {r['Model']}**")
            with info_col:
                st.caption(f"Realna sprzedaż: {float(r['Realna cena sprzedaży']):.0f} zł • Ilość: {int(r['Ilość'])}")

    with st.expander("🛠️ Ręczna edycja", expanded=False):
        options=[f"{r['Lp.']}. {r['Marka']} {r['Produkt']} {r['Model']}" for r in products]
        selected_product=st.selectbox("Wybierz produkt",options); idx=options.index(selected_product); row=products[idx]
        new_product_name=st.text_input("Nazwa produktu",value=str(row.get("Produkt","") or ""),key=f"edit_product_{row['_db_id']}")
        e1,e2,e3,e4=st.columns(4)
        with e1: new_qty=st.number_input("Ilość",min_value=1,value=int(row['Ilość']),step=1,key=f"edit_qty_{row['_db_id']}")
        with e2: new_real=st.number_input("Realna sprzedaż / szt.",min_value=0.0,value=float(row['Realna cena sprzedaży']),step=5.0,key=f"edit_real_{row['_db_id']}")
        with e3: new_listing=st.number_input("Cena wystawienia",min_value=0.0,value=float(row['Cena wystawienia']),step=5.0,key=f"edit_listing_{row['_db_id']}")
        with e4: st.metric("Poziom", level_label(new_real))
        new_offer=st.text_input("Link do przykładowej oferty",value=str(row.get("Link do oferty","") or ""),key=f"edit_offer_{row['_db_id']}")
        st.markdown("**💰 Status sprzedaży**")
        current_status = str(row.get("Status sprzedaży") or "Na stanie")
        manual_status = st.selectbox("Status", ["Na stanie", "Wystawiony", "Sprzedany"], index=["Na stanie", "Wystawiony", "Sprzedany"].index(current_status) if current_status in ["Na stanie", "Wystawiony", "Sprzedany"] else 0, key=f"manual_sale_status_{row['_db_id']}")
        manual_sold_price = st.number_input("Za ile sprzedano?", min_value=0.0, value=float(row.get("Cena sprzedaży",0) or 0), step=5.0, key=f"manual_sold_price_{row['_db_id']}") if manual_status == "Sprzedany" else 0.0
        st.caption("Poziom jest automatycznie wyliczany z realnej ceny sprzedaży: 🟡 Priorytet ≥250 zł • 🟢 Ważne 150–249,99 zł • 🟠 Mogą poczekać 50–149,99 zł • 🔴 Badziew <50 zł.")
        if st.button("💾 Zapisz zmiany",use_container_width=True):
            update_product(conn,row['_db_id'],new_qty,new_real,new_listing,None,new_offer,new_product_name)
            update_sale_status(conn,row['_db_id'],manual_status,manual_sold_price)
            st.rerun()
        if st.button("🗑️ Usuń wybraną pozycję",use_container_width=True): delete_product(conn,row['_db_id']); st.rerun()
else: st.info("Paleta jest pusta. Zrób pierwsze zdjęcie produktu.")

if products:
    st.divider(); st.subheader("📥 Eksport")
    excel=make_excel(products,str(current['name']),float(current['cost']))
    safe=re.sub(r"[^a-zA-Z0-9ąćęłńóśźżĄĆĘŁŃÓŚŹŻ _-]+","",str(current['name'])).strip().replace(' ','_') or 'paleta'
    st.download_button("📊 Pobierz Excel",data=excel,file_name=f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.caption("Dane są zapisane w zewnętrznej bazie PostgreSQL i są wspólne dla urządzeń.")
