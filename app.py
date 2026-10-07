import io
import json
import re
import hashlib
import os
from html import escape
from datetime import datetime, timezone, timedelta
import hmac

import pandas as pd
import streamlit as st
import extra_streamlit_components as stx
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openai import OpenAI
from sqlalchemy import text
from PIL import Image

st.set_page_config(page_title="Paletownia", page_icon="📦", layout="wide")

# CookieManager jest widgetem Streamlit — musi być tworzony poza @st.cache_* .
# Jedna instancja na uruchomienie aplikacji wystarcza do odczytu/zapisu trwałego logowania.
cookie_manager = stx.CookieManager()

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
    "Status sprzedaży", "Cena sprzedaży", "Data wystawienia", "Data sprzedaży",
    "Długość paczki", "Szerokość paczki", "Wysokość paczki", "Waga paczki"
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
    sold_at TIMESTAMPTZ NULL,
    package_l NUMERIC(8,2) NOT NULL DEFAULT 0,
    package_w NUMERIC(8,2) NOT NULL DEFAULT 0,
    package_h NUMERIC(8,2) NOT NULL DEFAULT 0,
    package_weight NUMERIC(8,2) NOT NULL DEFAULT 0
);
ALTER TABLE products ADD COLUMN IF NOT EXISTS image_thumb TEXT NOT NULL DEFAULT '';
ALTER TABLE products ADD COLUMN IF NOT EXISTS sale_status TEXT NOT NULL DEFAULT 'Na stanie';
ALTER TABLE products ADD COLUMN IF NOT EXISTS sold_price NUMERIC(12,2) NOT NULL DEFAULT 0;
ALTER TABLE products ADD COLUMN IF NOT EXISTS listed_at TIMESTAMPTZ NULL;
ALTER TABLE products ADD COLUMN IF NOT EXISTS sold_at TIMESTAMPTZ NULL;
ALTER TABLE products ADD COLUMN IF NOT EXISTS package_l NUMERIC(8,2) NOT NULL DEFAULT 0;
ALTER TABLE products ADD COLUMN IF NOT EXISTS package_w NUMERIC(8,2) NOT NULL DEFAULT 0;
ALTER TABLE products ADD COLUMN IF NOT EXISTS package_h NUMERIC(8,2) NOT NULL DEFAULT 0;
ALTER TABLE products ADD COLUMN IF NOT EXISTS package_weight NUMERIC(8,2) NOT NULL DEFAULT 0;
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


def load_all_packaging(conn):
    """Zwraca wszystkie produkty z zapisanymi wymiarami paczek ze wszystkich palet."""
    sql = """
    SELECT p.id, p.pallet_id, p.position, p.quantity, p.category, p.brand, p.product, p.model,
           p.package_l, p.package_w, p.package_h, p.package_weight, pl.name AS pallet_name
    FROM products p
    JOIN pallets pl ON pl.id = p.pallet_id
    WHERE COALESCE(p.package_l,0) > 0
      AND COALESCE(p.package_w,0) > 0
      AND COALESCE(p.package_h,0) > 0
    ORDER BY p.package_l, p.package_w, p.package_h, p.id
    """
    df = db_query(conn, sql, ttl=5)
    return df.to_dict("records")


def load_products(conn, pallet_id):
    sql = """
    SELECT id, position AS \"Lp.\", quantity AS \"Ilość\", category AS \"Kategoria\", brand AS \"Marka\",
           product AS \"Produkt\", model AS \"Model\", state AS \"Stan\", completeness AS \"Kompletność\",
           new_price AS \"Cena nowego\", used_price AS \"Cena używanego\", real_sale_price AS \"Realna cena sprzedaży\",
           listing_price AS \"Cena wystawienia\", price_source AS \"Źródło ceny\", offer_link AS \"Link do oferty\",
           notes AS \"Uwagi\", priority AS \"Priorytet\", image_thumb AS \"Miniatura\",
           sale_status AS \"Status sprzedaży\", sold_price AS \"Cena sprzedaży\", listed_at AS \"Data wystawienia\", sold_at AS \"Data sprzedaży\", created_at AS \"Data dodania\",
           package_l AS \"Długość paczki\", package_w AS \"Szerokość paczki\", package_h AS \"Wysokość paczki\", package_weight AS \"Waga paczki\"
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
        for k in ["Długość paczki","Szerokość paczki","Wysokość paczki","Waga paczki"]:
            r[k] = float(r.get(k) or 0)
        r["Lp."] = int(r["Lp."])
        r["_db_id"] = int(r["id"])
        r.pop("id", None)
    return rows


def load_product_by_id(conn, product_id):
    """Pobiera jeden produkt bez przeładowywania całej listy."""
    sql = """
    SELECT id, position AS "Lp.", quantity AS "Ilość", category AS "Kategoria", brand AS "Marka",
           product AS "Produkt", model AS "Model", state AS "Stan", completeness AS "Kompletność",
           new_price AS "Cena nowego", used_price AS "Cena używanego",
           real_sale_price AS "Realna cena sprzedaży", listing_price AS "Cena wystawienia",
           price_source AS "Źródło ceny", offer_link AS "Link do oferty", notes AS "Uwagi",
           priority AS "Priorytet", image_thumb AS "Miniatura",
           sale_status AS "Status sprzedaży", sold_price AS "Cena sprzedaży",
           listed_at AS "Data wystawienia", sold_at AS "Data sprzedaży",
           created_at AS "Data dodania",
           package_l AS "Długość paczki", package_w AS "Szerokość paczki",
           package_h AS "Wysokość paczki", package_weight AS "Waga paczki"
    FROM products
    WHERE id = :product_id
    LIMIT 1
    """
    df = db_query(conn, sql, {"product_id": int(product_id)}, ttl=0)
    if df.empty:
        return None
    r = df.iloc[0].to_dict()
    for k in ["Cena nowego", "Cena używanego", "Realna cena sprzedaży", "Cena wystawienia"]:
        r[k] = float(r.get(k) or 0)
    r["Priorytet"] = priority_from_price(r["Realna cena sprzedaży"])
    r["Ilość"] = int(r.get("Ilość") or 1)
    r["Cena sprzedaży"] = float(r.get("Cena sprzedaży") or 0)
    r["Status sprzedaży"] = str(r.get("Status sprzedaży") or "Na stanie")
    for k in ["Długość paczki", "Szerokość paczki", "Wysokość paczki", "Waga paczki"]:
        r[k] = float(r.get(k) or 0)
    r["_db_id"] = int(r.pop("id"))
    r["Lp."] = int(r["Lp."])
    return r


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


CARRIER_LIMITS = {
    "InPost": {"A": (8, 38, 64, 25), "B": (19, 38, 64, 25), "C": (41, 38, 64, 25)},
    "DPD automat": {"A": (11, 44, 59, 20), "B": (24, 44, 59, 20), "C": (50, 44, 59, 20)},
    "ORLEN Paczka": {"A": (8, 38, 60, 20), "B": (19, 38, 60, 20), "C": (41, 38, 60, 20)},
}

def package_with_padding(l, w, h):
    return tuple(round(float(x) + 2, 1) for x in (l, w, h))

def fits_box(dims, limits):
    if not dims or any(float(x) <= 0 for x in dims):
        return False
    a = sorted(float(x) for x in dims)
    b = sorted(float(x) for x in limits[:3])
    return all(x <= y + 1e-9 for x, y in zip(a, b))

def carrier_matches(l, w, h, weight=0):
    dims = package_with_padding(l, w, h)
    result = []
    for carrier, sizes in CARRIER_LIMITS.items():
        for name, lim in sizes.items():
            if fits_box(dims, lim) and (float(weight or 0) <= 0 or float(weight) <= lim[3]):
                result.append((carrier, name))
    return result

def best_shipping_options(l, w, h, weight=0):
    dims = package_with_padding(l, w, h)
    out = {}
    for carrier, sizes in CARRIER_LIMITS.items():
        hit = None
        for name, lim in sizes.items():
            if fits_box(dims, lim) and (float(weight or 0) <= 0 or float(weight) <= lim[3]):
                hit = name
                break
        out[carrier] = hit
    return dims, out

def save_product(conn, pallet_id, data, quantity=1, image_thumb=""):
    row = {
        "position": 1,
        "quantity": int(quantity), "category": data["kategoria"], "brand": data["marka"],
        "product": data["produkt"], "model": data["model"], "state": data["stan"],
        "completeness": data["kompletnosc"], "new_price": data["cena_nowego"], "used_price": data["cena_uzywanego"],
        "real_sale_price": data["realna_cena_sprzedazy"], "listing_price": data["cena_wystawienia"],
        "price_source": data["zrodlo_ceny"], "offer_link": data["link_do_oferty"],
        "notes": data["uwagi"] + f" | Pewność identyfikacji: {data['pewnosc_ident']}", "priority": priority_from_price(data["realna_cena_sprzedazy"]),
        "image_thumb": image_thumb or "", "sale_status": "Na stanie",
        "package_l": data.get("package_l", 0), "package_w": data.get("package_w", 0),
        "package_h": data.get("package_h", 0), "package_weight": data.get("package_weight", 0)
    }
    with conn.session as s:
        max_position = s.execute(
            text("SELECT COALESCE(MAX(position), 0) + 1 FROM products WHERE pallet_id=:p"),
            {"p": int(pallet_id)},
        ).scalar_one()
        row["position"] = int(max_position)
        result = s.execute(text("""
            INSERT INTO products (pallet_id, position, quantity, category, brand, product, model, state, completeness,
                new_price, used_price, real_sale_price, listing_price, price_source, offer_link, notes, priority, image_thumb, sale_status, package_l, package_w, package_h, package_weight)
            VALUES (:pallet_id,:position,:quantity,:category,:brand,:product,:model,:state,:completeness,
                :new_price,:used_price,:real_sale_price,:listing_price,:price_source,:offer_link,:notes,:priority,:image_thumb,:sale_status,:package_l,:package_w,:package_h,:package_weight)
            RETURNING id
        """), {"pallet_id":int(pallet_id), **row})
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=:id"), {"id":int(pallet_id)})
        s.commit()
    invalidate_data_cache()
    return int(result.scalar_one())


def update_product(conn, product_id, quantity, real_price, listing_price, priority=None, offer_link=None, product_name=None, package_l=None, package_w=None, package_h=None, package_weight=None):
    priority = priority_from_price(real_price)
    with conn.session as s:
        params={"q":int(quantity),"r":float(real_price),"l":float(listing_price),"p":priority,"id":int(product_id)}
        sets="quantity=:q, real_sale_price=:r, listing_price=:l, priority=:p"
        if package_l is not None:
            sets += ", package_l=:pl, package_w=:pw, package_h=:ph, package_weight=:pwt"
            params.update({"pl":float(package_l or 0),"pw":float(package_w or 0),"ph":float(package_h or 0),"pwt":float(package_weight or 0)})
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


def excel_safe_value(value):
    """Zamienia daty z PostgreSQL TIMESTAMPTZ na wartości akceptowane przez Excel."""
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    # Obsługa obiektów datetime/time z tzinfo, jeśli takie pojawią się w danych.
    try:
        from datetime import time as datetime_time
        if isinstance(value, datetime_time):
            return value.replace(tzinfo=None)
    except Exception:
        pass
    return value


def admin_session_token():
    """Stały, nieodwracalny token dla przeglądarki administratora."""
    secret = str(st.secrets.get("ADMIN_SESSION_SECRET", "")).strip()
    if not secret:
        # Fallback: działa bez dodatkowego wpisu w Secrets, ale docelowo
        # warto ustawić osobny, długi ADMIN_SESSION_SECRET.
        secret = str(st.secrets.get("ADMIN_PIN", "")).strip()
    if not secret:
        return ""
    return hmac.new(secret.encode("utf-8"), b"paletownia-admin-v1", hashlib.sha256).hexdigest()


def persistent_admin_login():
    """Odtwarza logowanie z trwałego cookie przeglądarki."""
    try:
        manager = cookie_manager
        token = manager.get(cookie="paletownia_admin")
        expected = admin_session_token()
        if token and expected and hmac.compare_digest(str(token), expected):
            st.session_state.admin_logged_in = True
    except Exception:
        # Brak/awaria komponentu cookie nie blokuje zwykłego logowania PIN-em.
        pass


def persist_admin_login():
    manager = cookie_manager
    token = admin_session_token()
    if token:
        manager.set(
            "paletownia_admin",
            token,
            expires_at=datetime.now(timezone.utc) + timedelta(days=30),
        )


def clear_persistent_admin_login():
    try:
        cookie_manager.delete("paletownia_admin")
    except Exception:
        pass


def make_excel(products, pallet_name, cost):
    wb=Workbook(); ws=wb.active; ws.title="Produkty"; ws.append(FIELDS)
    fill=PatternFill("solid",fgColor="D9EAF7")
    for c in ws[1]: c.font=Font(bold=True); c.fill=fill; c.alignment=Alignment(horizontal="center")
    for row in products:
        ws.append([excel_safe_value(row.get(f,"")) for f in FIELDS])
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for i,w in enumerate([7,9,22,18,30,18,22,25,16,18,24,18,25,45,45,28,18,18,20,20],1): ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width=w
    s=wb.create_sheet("Podsumowanie"); s["A1"]="PODSUMOWANIE PALETY"; s["A1"].font=Font(bold=True,size=16)
    vals=[("Nazwa palety",pallet_name),("Liczba pozycji",len(products)),("Liczba sztuk",sum(int(r['Ilość']) for r in products)),("Koszt palety",cost),("Realna wartość sprzedaży",total_value(products)),("Nadwyżka przed kosztami sprzedaży",total_value(products)-cost)]
    for i,(a,b) in enumerate(vals,2): s.cell(i,1,a); s.cell(i,2,b)
    s.column_dimensions['A'].width=40; s.column_dimensions['B'].width=25
    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()


# Session/UI state only stores current selection and temporary AI result. Actual data is in PostgreSQL.
if "_data_version" not in st.session_state: st.session_state._data_version = 0
if "_db_initialized" not in st.session_state: st.session_state._db_initialized = False
if "current_pallet_id" not in st.session_state: st.session_state.current_pallet_id=None
if "pending" not in st.session_state: st.session_state.pending=None
if "last_analyzed_hash" not in st.session_state: st.session_state.last_analyzed_hash=None
if "last_added" not in st.session_state: st.session_state.last_added=None
if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False
if "products_page" not in st.session_state: st.session_state.products_page = 1
if "app_page" not in st.session_state: st.session_state.app_page = "home"

# Trwałe logowanie: po pierwszym poprawnym PIN-ie przeglądarka dostaje
# cookie ważne 30 dni. Dzięki temu zamknięcie karty/aplikacji nie wylogowuje admina.
if not st.session_state.admin_logged_in:
    persistent_admin_login()

# Paletownia jest wyłącznie panelem administracyjnym.
# Po wejściu na stronę użytkownik od razu dostaje ekran PIN-u.
if not st.session_state.admin_logged_in:
    st.markdown("""
    <style>
    .login-wrap {
        max-width: 430px;
        margin: 12vh auto 0 auto;
        text-align: center;
    }
    .login-title { font-size: 2rem; font-weight: 800; margin-bottom: 4px; }
    .login-subtitle { color: #64748b; margin-bottom: 24px; }
    </style>
    <div class="login-wrap">
        <div class="login-title">📦 Paletownia</div>
        <div class="login-subtitle">Panel administracyjny</div>
    </div>
    """, unsafe_allow_html=True)
    with st.form("admin_login_form"):
        pin = st.text_input("PIN administratora", type="password", placeholder="Wpisz PIN", label_visibility="collapsed")
        submitted = st.form_submit_button("🔐 Zaloguj", type="primary", use_container_width=True)
        if submitted:
            configured_pin = str(st.secrets.get("ADMIN_PIN", "")).strip()
            if configured_pin and pin == configured_pin:
                st.session_state.admin_logged_in = True
                st.session_state.admin_login_error = False
                persist_admin_login()
                st.rerun()
            else:
                st.session_state.admin_login_error = True
                st.error("Nieprawidłowy PIN.")
    if not str(st.secrets.get("ADMIN_PIN", "")).strip():
        st.warning("Brak ADMIN_PIN w Secrets Streamlit.")
    st.stop()

conn=db_conn()
if conn:
    try:
        db_init(conn)
        if st.session_state.current_pallet_id is None:
            st.session_state.current_pallet_id=ensure_first_pallet(conn)
    except Exception as exc:
        st.error("Nie udało się zainicjalizować bazy danych.")
        st.code(str(exc)); st.stop()

# Tryb administratora — pełny dotychczasowy interfejs.
admin_left, admin_right = st.columns([5, 1])
with admin_left:
    st.markdown("### 🔐 Paletownia — panel właściciela")
with admin_right:
    if st.button("Wyloguj", use_container_width=True, key="admin_logout"):
        clear_persistent_admin_login()
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



# Główna nawigacja aplikacji
nav_col1, nav_col2 = st.columns(2)
with nav_col1:
    if st.button("🏠 Strona główna", use_container_width=True, type="primary" if st.session_state.get("app_page", "home") == "home" else "secondary", key="nav_home"):
        st.session_state.app_page = "home"
        st.rerun()
with nav_col2:
    if st.button("📦 Produkty", use_container_width=True, type="primary" if st.session_state.get("app_page", "home") == "products" else "secondary", key="nav_products"):
        st.session_state.app_page = "products"
        st.rerun()

app_page = st.session_state.get("app_page", "home")


if app_page == "home":
    st.title("📦 Paletownia")
    st.subheader("📸 Skanowanie produktu")
    st.caption("Zrób zdjęcie produktu lub wybierz zdjęcie z urządzenia. Paletownia rozpozna produkt, sprawdzi aktualne ceny i zapisze go do bieżącej palety.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📸 Zrób zdjęcie")
        camera = st.camera_input("Aparat", key=f"camera_home_{st.session_state.current_pallet_id}", resolution="720p")
    with c2:
        st.subheader("📁 Wybierz z urządzenia")
        upload = st.file_uploader("Zdjęcie produktu", type=["jpg", "jpeg", "png", "webp"], key=f"uploader_home_{st.session_state.current_pallet_id}")

    image_file = camera if camera is not None else upload
    if image_file is not None:
        b = image_file.getvalue()
        h = hashlib.sha256(b).hexdigest()
        st.image(image_file, caption="Wybrane zdjęcie", width="stretch")
        if h != st.session_state.last_analyzed_hash:
            st.session_state.last_analyzed_hash = h
            with st.spinner("🤖 AI rozpoznaje produkt i sprawdza aktualne ceny..."):
                try:
                    data = analyze_image(b, getattr(image_file, "type", "image/jpeg"))
                    products_now = load_products(conn, int(st.session_state.current_pallet_id))
                    thumb = make_thumbnail_data_url(b)
                    duplicate_idx = find_duplicate(products_now, data)
                    if duplicate_idx is not None:
                        existing = products_now[duplicate_idx]
                        new_qty = int(existing["Ilość"]) + 1
                        update_product(conn, existing["_db_id"], new_qty, existing["Realna cena sprzedaży"], existing["Cena wystawienia"], existing["Priorytet"])
                        if not existing.get("Miniatura"):
                            update_product_thumbnail(conn, existing["_db_id"], thumb)
                        st.session_state.last_added = {"kind":"duplicate", "name":f"{existing['Marka']} {existing['Produkt']} {existing['Model']}", "quantity":new_qty, "price":float(existing["Realna cena sprzedaży"])}
                    else:
                        save_product(conn, int(st.session_state.current_pallet_id), data, 1, thumb)
                        st.session_state.last_added = {"kind":"new", "name":f"{data['marka']} {data['produkt']} {data['model']}".strip(), "quantity":1, "price":float(data["realna_cena_sprzedazy"]), "link":data.get("link_do_oferty", "")}
                    st.rerun()
                except Exception as exc:
                    st.session_state.last_analyzed_hash = None
                    st.error(f"Nie udało się przeanalizować zdjęcia: {exc}")

    if st.session_state.last_added:
        added = st.session_state.last_added
        if added.get("kind") == "duplicate":
            st.success(f"✅ Produkt rozpoznany jako duplikat i automatycznie dodany jako kolejna sztuka: **{added['name']}**. Łącznie: **{added['quantity']} szt.**")
        else:
            st.success(f"✅ Produkt automatycznie dodany do palety: **{added['name']}** — realna sprzedaż: **{added['price']:.0f} zł**")
            if added.get("link"):
                st.markdown(f"[Przykładowa oferta]({added['link']})")

else:

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

        # Kompaktowy widok: karty w siatce + paginacja. Szczegóły edycji są zwinięte.
        st.markdown("""
        <style>
        .product-card-compact {
            border:1px solid #e5e7eb; border-radius:14px; padding:10px;
            margin:5px 0 10px 0; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.05);
            min-height:250px;
        }
        .compact-name { font-size:.96rem; font-weight:750; line-height:1.2; min-height:42px; }
        .compact-price { font-size:1.12rem; font-weight:800; margin:3px 0; }
        .compact-meta { color:#6b7280; font-size:.78rem; line-height:1.3; }
        .level-badge { display:inline-block; padding:3px 8px; border-radius:999px; font-size:.72rem; font-weight:700; margin:2px 0; }
        .lvl1 { background:#fff1bf; color:#6b5200; }
        .lvl2 { background:#dcfce7; color:#166534; }
        .lvl3 { background:#ffedd5; color:#9a3412; }
        .lvl4 { background:#fee2e2; color:#991b1b; }
        @media (max-width: 900px) {
          .compact-name { font-size:.9rem; }
          .product-card-compact { min-height:235px; }
        }
        </style>
        """, unsafe_allow_html=True)

        def level_css(label):
            if "Priorytet" in label: return "lvl1"
            if "Ważne" in label: return "lvl2"
            if "Mogą poczekać" in label: return "lvl3"
            return "lvl4"

        @st.fragment
        def render_product_card(product_id):
            r = load_product_by_id(conn, product_id)
            if r is None:
                return

            level = level_label(float(r["Realna cena sprzedaży"]))
            cls = level_css(level)
            name = " ".join(
                str(x).strip()
                for x in [r.get("Marka", ""), r.get("Produkt", ""), r.get("Model", "")]
                if str(x).strip()
            )
            thumb = str(r.get("Miniatura") or "")
            status = str(r.get("Status sprzedaży") or "Na stanie")
            status_icon = {"Na stanie":"⚪", "Wystawiony":"🟢", "Sprzedany":"🔴"}.get(status, "⚪")

            with st.container(border=True):
                if thumb:
                    st.image(thumb, width="stretch")
                else:
                    st.markdown(
                        '<div style="height:125px;display:flex;align-items:center;justify-content:center;'
                        'border:1px dashed #cbd5e1;border-radius:10px;color:#94a3b8;font-size:30px;">📷</div>',
                        unsafe_allow_html=True
                    )

                st.markdown(f'<div class="compact-name">{int(r["Lp."])}. {escape(name)}</div>', unsafe_allow_html=True)
                st.markdown(f'<span class="level-badge {cls}">{level}</span>', unsafe_allow_html=True)
                st.markdown(f'<div class="compact-price">{float(r["Realna cena sprzedaży"]):.0f} zł</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="compact-meta">Ilość: <b>{int(r["Ilość"])}</b> · {status_icon} {escape(status)}</div>',
                    unsafe_allow_html=True
                )
                if status == "Sprzedany":
                    st.caption(f'Sprzedano za {float(r.get("Cena sprzedaży",0) or 0):.0f} zł')

                if str(r.get("Link do oferty") or "").strip():
                    try:
                        st.link_button("🔗 Oferta", str(r["Link do oferty"]), use_container_width=True)
                    except Exception:
                        st.markdown(f'[🔗 Oferta]({r["Link do oferty"]})')

                with st.expander("✏️ Edytuj", expanded=False):
                    with st.form(key=f"product_edit_form_{product_id}", clear_on_submit=False):
                        new_product_name = st.text_input(
                            "Nazwa produktu", value=str(r.get("Produkt", "") or ""),
                            key=f"card_edit_product_{product_id}"
                        )
                        e1, e2 = st.columns(2)
                        with e1:
                            new_qty = st.number_input("Ilość", min_value=1, value=int(r["Ilość"]), step=1, key=f"card_edit_qty_{product_id}")
                        with e2:
                            new_real = st.number_input("Realna sprzedaż / szt.", min_value=0.0, value=float(r["Realna cena sprzedaży"]), step=5.0, key=f"card_edit_real_{product_id}")
                        new_listing = st.number_input("Cena wystawienia", min_value=0.0, value=float(r["Cena wystawienia"]), step=5.0, key=f"card_edit_listing_{product_id}")

                        st.markdown("**📦 Wymiary paczki / wysyłka**")
                        p1, p2 = st.columns(2)
                        with p1:
                            pkg_l = st.number_input("Długość (cm)", min_value=0.0, value=float(r.get("Długość paczki", 0)), step=0.5, key=f"pkg_l_{product_id}")
                            pkg_h = st.number_input("Wysokość (cm)", min_value=0.0, value=float(r.get("Wysokość paczki", 0)), step=0.5, key=f"pkg_h_{product_id}")
                        with p2:
                            pkg_w = st.number_input("Szerokość (cm)", min_value=0.0, value=float(r.get("Szerokość paczki", 0)), step=0.5, key=f"pkg_w_{product_id}")
                            pkg_weight = st.number_input("Waga (kg)", min_value=0.0, value=float(r.get("Waga paczki", 0)), step=0.1, key=f"pkg_weight_{product_id}")

                        if pkg_l > 0 and pkg_w > 0 and pkg_h > 0:
                            pdims, pops = best_shipping_options(pkg_l, pkg_w, pkg_h, pkg_weight)
                            st.caption(f"📐 Po dodaniu +2 cm: **{pdims[0]:g} × {pdims[1]:g} × {pdims[2]:g} cm**")
                            st.caption(" • ".join(f"{k}: **{v or 'poza automatem'}**" for k, v in pops.items()))

                        new_offer = st.text_input(
                            "Link do przykładowej oferty", value=str(r.get("Link do oferty", "") or ""),
                            key=f"card_edit_offer_{product_id}"
                        )
                        current_status = str(r.get("Status sprzedaży") or "Na stanie")
                        manual_status = st.selectbox(
                            "Status", ["Na stanie", "Wystawiony", "Sprzedany"],
                            index=["Na stanie", "Wystawiony", "Sprzedany"].index(current_status) if current_status in ["Na stanie", "Wystawiony", "Sprzedany"] else 0,
                            key=f"manual_sale_status_{product_id}"
                        )
                        manual_sold_price = (
                            st.number_input("Za ile sprzedano?", min_value=0.0, value=float(r.get("Cena sprzedaży", 0) or 0), step=5.0, key=f"manual_sold_price_{product_id}")
                            if manual_status == "Sprzedany" else 0.0
                        )
                        save_clicked = st.form_submit_button("💾 Zapisz zmiany", use_container_width=True, type="primary")

                    if save_clicked:
                        update_product(conn, product_id, new_qty, new_real, new_listing, None, new_offer, new_product_name, pkg_l, pkg_w, pkg_h, pkg_weight)
                        update_sale_status(conn, product_id, manual_status, manual_sold_price)
                        st.success("✅ Zapisano zmiany.")
                        st.rerun(scope="fragment")

                with st.popover("🖼️ Zdjęcie", use_container_width=True):
                    st.write(f"**{name}**")
                    photo = st.file_uploader(
                        "Wybierz nowe zdjęcie", type=["jpg", "jpeg", "png", "webp"],
                        key=f"card_photo_{product_id}", label_visibility="collapsed"
                    )
                    if photo is not None:
                        new_thumb = make_thumbnail_data_url(photo.getvalue())
                        if new_thumb:
                            update_product_thumbnail(conn, product_id, new_thumb, replace=True)
                            st.success("Zdjęcie zapisane.")
                            st.rerun(scope="fragment")

                if st.button("🗑️ Usuń", key=f"card_delete_{product_id}", use_container_width=True):
                    delete_product(conn, product_id)
                    st.rerun()

        # Paginacja: na jednej stronie maksymalnie 20 produktów.
        page_size_options = [12, 24, 48]
        pc1, pc2, pc3 = st.columns([1.3, 1.3, 2.4])
        with pc1:
            page_size = st.selectbox("Na stronę", page_size_options, index=1, key="products_page_size")
        total_products = len(products)
        total_pages = max(1, (total_products + page_size - 1) // page_size)
        current_page = min(max(int(st.session_state.get("products_page", 1)), 1), total_pages)
        with pc2:
            page_selected = st.number_input("Strona", min_value=1, max_value=total_pages, value=current_page, step=1, key="products_page_input")
            if int(page_selected) != current_page:
                st.session_state.products_page = int(page_selected)
                st.rerun()
        with pc3:
            start_idx = (current_page - 1) * page_size
            end_idx = min(start_idx + page_size, total_products)
            st.caption(f"Wyświetlam **{start_idx + 1}–{end_idx}** z **{total_products}** produktów")

        page_products = products[start_idx:end_idx]
        grid_cols = st.columns(4)
        for idx, r in enumerate(page_products):
            with grid_cols[idx % 4]:
                render_product_card(r["_db_id"])

        nav1, nav2, nav3, nav4, nav5 = st.columns([1, 1, 2, 1, 1])
        with nav1:
            if st.button("⏮️", disabled=current_page <= 1, key="page_first", use_container_width=True):
                st.session_state.products_page = 1; st.rerun()
        with nav2:
            if st.button("◀️", disabled=current_page <= 1, key="page_prev", use_container_width=True):
                st.session_state.products_page = current_page - 1; st.rerun()
        with nav3:
            st.markdown(f"<div style='text-align:center;padding:8px;font-weight:700;'>Strona {current_page} z {total_pages}</div>", unsafe_allow_html=True)
        with nav4:
            if st.button("▶️", disabled=current_page >= total_pages, key="page_next", use_container_width=True):
                st.session_state.products_page = current_page + 1; st.rerun()
        with nav5:
            if st.button("⏭️", disabled=current_page >= total_pages, key="page_last", use_container_width=True):
                st.session_state.products_page = total_pages; st.rerun()

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

    else: st.info("Paleta jest pusta. Zrób pierwsze zdjęcie produktu.")

    if products:
        st.divider()
        with st.expander("📦 Kartony i wysyłka — analiza całej bazy", expanded=False):
            all_pack = load_all_packaging(conn)
            if not all_pack:
                st.info("Dodaj wymiary paczek przy produktach, aby Paletownia mogła policzyć gabaryty i zapotrzebowanie na kartony.")
            else:
                st.caption("Wpisujesz rzeczywiste wymiary mierzonego pakunku/produktu. Paletownia automatycznie dodaje **2 cm do każdego wymiaru** jako zapas na wypełnienie i dopiero tak powiększone wymiary porównuje z limitami przewoźników.")

                rows=[]
                carrier_counts={k:{"A":0,"B":0,"C":0} for k in CARRIER_LIMITS}
                courier_needed=0
                for r in all_pack:
                    dims,opts=best_shipping_options(r["package_l"],r["package_w"],r["package_h"],r.get("package_weight",0))
                    qty=max(1,int(r.get("quantity") or 1))
                    rows.extend([dims] * qty)
                    for carrier, gab in opts.items():
                        if gab:
                            carrier_counts[carrier][gab] += qty
                    if not any(opts.values()):
                        courier_needed += qty

                st.markdown("### 📊 Jakich kartonów potrzebujesz najczęściej?")
                counts=pd.Series([tuple(round(x,1) for x in d) for d in rows]).value_counts()
                for dims,count in counts.head(10).items():
                    st.write(f"📦 **{dims[0]:g} × {dims[1]:g} × {dims[2]:g} cm** — **{int(count)} szt.**")

                st.markdown("### 🚚 W jakich automatach zmieszczą się paczki?")
                c1,c2,c3,c4=st.columns(4)
                c1.metric("InPost", sum(carrier_counts["InPost"].values()))
                c2.metric("DPD automat", sum(carrier_counts["DPD automat"].values()))
                c3.metric("ORLEN Paczka", sum(carrier_counts["ORLEN Paczka"].values()))
                c4.metric("Poza automatami", courier_needed)

                st.markdown("**Najmniejszy dostępny gabaryt dla każdego przewoźnika:**")
                for carrier in CARRIER_LIMITS:
                    counts_c=carrier_counts[carrier]
                    best=next((g for g in ["A","B","C"] if counts_c[g]), "—")
                    st.caption(f"{carrier}: **{best}** — A: {counts_c['A']} szt. • B: {counts_c['B']} szt. • C: {counts_c['C']} szt.")

                st.markdown("### 📋 Produkty wymagające największych kartonów / kuriera")
                for r in all_pack:
                    dims,opts=best_shipping_options(r["package_l"],r["package_w"],r["package_h"],r.get("package_weight",0))
                    if not any(opts.values()):
                        name=f"{r['brand']} {r['product']} {r['model']}".strip()
                        st.warning(f"{name} — {dims[0]:g} × {dims[1]:g} × {dims[2]:g} cm po dodaniu zapasu. **Nie mieści się w automatach InPost, DPD ani ORLEN Paczka** — potrzebny kurier / inna usługa.")

                st.caption("Dopasowanie uwzględnia obrót prostopadłościanu, tak aby wymiary mogły zostać ustawione w najbardziej korzystnej orientacji. Limity wagowe są sprawdzane, jeśli podasz wagę paczki.")

        st.divider(); st.subheader("📥 Eksport")
        excel=make_excel(products,str(current['name']),float(current['cost']))
        safe=re.sub(r"[^a-zA-Z0-9ąćęłńóśźżĄĆĘŁŃÓŚŹŻ _-]+","",str(current['name'])).strip().replace(' ','_') or 'paleta'
        st.download_button("📊 Pobierz Excel",data=excel,file_name=f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.caption("Dane są zapisane w zewnętrznej bazie PostgreSQL i są wspólne dla urządzeń.")
