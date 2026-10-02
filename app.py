import io
import json
import re
import hashlib
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openai import OpenAI
from sqlalchemy import text

st.set_page_config(page_title="Paletownik AI", page_icon="📦", layout="wide")

FIELDS = [
    "Lp.", "Ilość", "Kategoria", "Marka", "Produkt", "Model", "Stan",
    "Kompletność", "Cena nowego", "Cena używanego", "Realna cena sprzedaży",
    "Cena wystawienia", "Źródło ceny", "Link do oferty", "Uwagi", "Priorytet"
]
PRIORITIES = [
    "Najważniejsze do sprzedaży", "Ważne", "Mogą poczekać", "Drobnica / badziew"
]
OLD_PRIORITY_MAP = {"Wysoki":"Najważniejsze do sprzedaży", "Normalny":"Ważne", "Niski":"Mogą poczekać"}

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
Priorytet: Najważniejsze do sprzedaży = wartościowy i łatwy do sprzedaży; Ważne = sensowny produkt o dobrej wartości/popycie; Mogą poczekać = mniej atrakcyjny lub wymagający czasu; Drobnica / badziew = tani, no-name, mało atrakcyjny lub trudny do sprzedaży.
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
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
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
    with conn.session as s:
        for statement in [x.strip() for x in DB_DDL.split(';') if x.strip()]:
            s.execute(text(statement))
        s.commit()


def db_query(conn, sql, params=None):
    return conn.query(sql, params=params or {}, ttl=0)


def ensure_first_pallet(conn):
    df = db_query(conn, "SELECT id FROM pallets ORDER BY id LIMIT 1")
    if df.empty:
        with conn.session as s:
            result = s.execute(text("INSERT INTO pallets (name, cost) VALUES (:name, :cost) RETURNING id"),
                               {"name":"Nowa paleta", "cost":360})
            pid = result.scalar_one()
            s.commit()
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
           notes AS \"Uwagi\", priority AS \"Priorytet\"
    FROM products WHERE pallet_id = :pallet_id ORDER BY position, id
    """
    df = db_query(conn, sql, {"pallet_id": int(pallet_id)})
    rows = df.to_dict("records")
    for r in rows:
        r["Priorytet"] = OLD_PRIORITY_MAP.get(r.get("Priorytet"), r.get("Priorytet", "Ważne"))
        for k in ["Cena nowego","Cena używanego","Realna cena sprzedaży","Cena wystawienia"]:
            r[k] = float(r[k] or 0)
        r["Ilość"] = int(r["Ilość"] or 1)
        r["Lp."] = int(r["Lp."])
        r["_db_id"] = int(r["id"])
        r.pop("id", None)
    return rows


def save_product(conn, pallet_id, data, quantity=1):
    row = {
        "position": len(load_products(conn, pallet_id)) + 1,
        "quantity": int(quantity), "category": data["kategoria"], "brand": data["marka"],
        "product": data["produkt"], "model": data["model"], "state": data["stan"],
        "completeness": data["kompletnosc"], "new_price": data["cena_nowego"], "used_price": data["cena_uzywanego"],
        "real_sale_price": data["realna_cena_sprzedazy"], "listing_price": data["cena_wystawienia"],
        "price_source": data["zrodlo_ceny"], "offer_link": data["link_do_oferty"],
        "notes": data["uwagi"] + f" | Pewność identyfikacji: {data['pewnosc_ident']}", "priority": data["priorytet"]
    }
    with conn.session as s:
        result = s.execute(text("""
            INSERT INTO products (pallet_id, position, quantity, category, brand, product, model, state, completeness,
                new_price, used_price, real_sale_price, listing_price, price_source, offer_link, notes, priority)
            VALUES (:pallet_id,:position,:quantity,:category,:brand,:product,:model,:state,:completeness,
                :new_price,:used_price,:real_sale_price,:listing_price,:price_source,:offer_link,:notes,:priority)
            RETURNING id
        """), {"pallet_id":int(pallet_id), **row})
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=:id"), {"id":int(pallet_id)})
        s.commit()
    return int(result.scalar_one())


def update_product(conn, product_id, quantity, real_price, listing_price, priority):
    with conn.session as s:
        s.execute(text("""UPDATE products SET quantity=:q, real_sale_price=:r, listing_price=:l, priority=:p, updated_at=NOW() WHERE id=:id"""),
                  {"q":int(quantity),"r":float(real_price),"l":float(listing_price),"p":priority,"id":int(product_id)})
        s.execute(text("UPDATE pallets SET updated_at=NOW() WHERE id=(SELECT pallet_id FROM products WHERE id=:id)"), {"id":int(product_id)})
        s.commit()


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


def update_pallet(conn, pallet_id, name, cost):
    with conn.session as s:
        s.execute(text("UPDATE pallets SET name=:name, cost=:cost, updated_at=NOW() WHERE id=:id"),
                  {"name":name.strip() or "Bez nazwy", "cost":float(cost), "id":int(pallet_id)})
        s.commit()


def create_pallet(conn, name="Nowa paleta", cost=360):
    with conn.session as s:
        result = s.execute(text("INSERT INTO pallets (name,cost) VALUES (:name,:cost) RETURNING id"),
                           {"name":name.strip() or "Nowa paleta", "cost":float(cost)})
        pid = int(result.scalar_one())
        s.commit()
    return pid


def delete_pallet(conn, pallet_id):
    with conn.session as s:
        s.execute(text("DELETE FROM pallets WHERE id=:id"), {"id":int(pallet_id)})
        s.commit()


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


def get_client():
    try: key = st.secrets["OPENAI_API_KEY"]
    except Exception: key = ""
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
    for i,w in enumerate([7,9,22,18,30,18,22,25,16,18,24,18,25,45,45,28],1): ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width=w
    s=wb.create_sheet("Podsumowanie"); s["A1"]="PODSUMOWANIE PALETY"; s["A1"].font=Font(bold=True,size=16)
    vals=[("Nazwa palety",pallet_name),("Liczba pozycji",len(products)),("Liczba sztuk",sum(int(r['Ilość']) for r in products)),("Koszt palety",cost),("Realna wartość sprzedaży",total_value(products)),("Nadwyżka przed kosztami sprzedaży",total_value(products)-cost)]
    for i,(a,b) in enumerate(vals,2): s.cell(i,1,a); s.cell(i,2,b)
    s.column_dimensions['A'].width=40; s.column_dimensions['B'].width=25
    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

# Session/UI state only stores current selection and temporary AI result. Actual data is in PostgreSQL.
if "current_pallet_id" not in st.session_state: st.session_state.current_pallet_id=None
if "pending" not in st.session_state: st.session_state.pending=None
if "last_analyzed_hash" not in st.session_state: st.session_state.last_analyzed_hash=None

conn=db_conn()
if conn:
    try:
        db_init(conn)
        if st.session_state.current_pallet_id is None:
            st.session_state.current_pallet_id=ensure_first_pallet(conn)
    except Exception as exc:
        st.error("Nie udało się zainicjalizować bazy danych.")
        st.code(str(exc)); st.stop()

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
        pid=create_pallet(conn, "Nowa paleta", 360); st.session_state.current_pallet_id=pid; st.session_state.pending=None; st.session_state.last_analyzed_hash=None; st.rerun()
    st.divider(); st.subheader("📚 Historia palet")
    labels=[f"{int(r.id)} — {r['name']} — {str(r['updated_at'])[:16]}" for _,r in pallets.iterrows()]
    selected=st.selectbox("Wybierz paletę", labels, index=max(0,next((i for i,r in pallets.iterrows() if int(r.id)==st.session_state.current_pallet_id),0)), key="pallet_selector")
    selected_id=int(selected.split(" — ",1)[0])
    if selected_id != st.session_state.current_pallet_id:
        st.session_state.current_pallet_id=selected_id; st.session_state.pending=None; st.session_state.last_analyzed_hash=None; st.rerun()
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

st.title("📦 Paletownik AI")
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
            try: st.session_state.pending=analyze_image(b,getattr(image_file,"type","image/jpeg")); st.rerun()
            except Exception as exc: st.session_state.last_analyzed_hash=None; st.error(f"Nie udało się przeanalizować zdjęcia: {exc}")

if st.session_state.pending:
    data=st.session_state.pending; st.divider(); st.subheader("🔎 Wynik AI")
    p1,p2=st.columns([2,1])
    with p1:
        st.write(f"**{data['marka']} {data['produkt']}**" + (f" — **{data['model']}**" if data['model'] else ""))
        st.write(f"Kategoria: **{data['kategoria']}**"); st.write(f"Stan: **{data['stan']}**"); st.write(f"Kompletność: **{data['kompletnosc']}**"); st.write(f"Pewność identyfikacji: **{data['pewnosc_ident']}**")
        if data['uwagi']: st.info(data['uwagi'])
        if data['link_do_oferty']: st.markdown(f"[Przykładowa oferta]({data['link_do_oferty']})")
    with p2:
        st.metric("Cena nowego",f"{data['cena_nowego']:.0f} zł"); st.metric("Cena używanego",f"{data['cena_uzywanego']:.0f} zł"); st.metric("Realna sprzedaż / szt.",f"{data['realna_cena_sprzedazy']:.0f} zł"); st.metric("Cena wystawienia",f"{data['cena_wystawienia']:.0f} zł")
    duplicate_idx=find_duplicate(products,data)
    if duplicate_idx is not None:
        existing=products[duplicate_idx]; st.warning(f"⚠️ Ten produkt już istnieje: {existing['Marka']} {existing['Produkt']} {existing['Model']} — {existing['Ilość']} szt.")
        d1,d2=st.columns(2)
        with d1: add_qty=st.number_input("Dodaj sztuk",min_value=1,value=1,step=1,key="duplicate_add_qty")
        with d2:
            if st.button("➕ Dodaj",type="primary",use_container_width=True):
                update_product(conn, existing['_db_id'], int(existing['Ilość'])+int(add_qty), existing['Realna cena sprzedaży'], existing['Cena wystawienia'], existing['Priorytet']); st.session_state.pending=None; st.rerun()
        if st.button("➕ Dodaj jako osobną pozycję",use_container_width=True):
            save_product(conn,st.session_state.current_pallet_id,data,int(add_qty)); st.session_state.pending=None; st.rerun()
    else:
        qty=st.number_input("Ilość sztuk",min_value=1,value=1,step=1,key="new_product_qty")
        st.caption(f"Wartość pozycji: **{int(qty)*data['realna_cena_sprzedazy']:.0f} zł**")
        a1,a2=st.columns(2)
        with a1:
            if st.button("✅ Dodaj do palety",type="primary",use_container_width=True): save_product(conn,st.session_state.current_pallet_id,data,int(qty)); st.session_state.pending=None; st.rerun()
        with a2:
            if st.button("❌ Odrzuć wynik AI",use_container_width=True): st.session_state.pending=None; st.rerun()

st.divider(); st.subheader("📋 Zawartość palety")
if products:
    df=pd.DataFrame(products); df["Wartość pozycji"]=df["Ilość"].astype(float)*df["Realna cena sprzedaży"].astype(float)
    display_cols=["Lp.","Ilość","Kategoria","Marka","Produkt","Model","Realna cena sprzedaży","Wartość pozycji","Cena wystawienia","Priorytet"]
    st.dataframe(df[display_cols],use_container_width=True,hide_index=True)
    st.markdown("### 🛠️ Ręczna edycja")
    options=[f"{r['Lp.']}. {r['Marka']} {r['Produkt']} {r['Model']}" for r in products]
    selected_product=st.selectbox("Wybierz produkt",options); idx=options.index(selected_product); row=products[idx]
    e1,e2,e3,e4=st.columns(4)
    with e1: new_qty=st.number_input("Ilość",min_value=1,value=int(row['Ilość']),step=1,key=f"edit_qty_{row['_db_id']}")
    with e2: new_real=st.number_input("Realna sprzedaż / szt.",min_value=0.0,value=float(row['Realna cena sprzedaży']),step=5.0,key=f"edit_real_{row['_db_id']}")
    with e3: new_listing=st.number_input("Cena wystawienia",min_value=0.0,value=float(row['Cena wystawienia']),step=5.0,key=f"edit_listing_{row['_db_id']}")
    with e4:
        cp=OLD_PRIORITY_MAP.get(row.get('Priorytet'),row.get('Priorytet','Ważne')); cp=cp if cp in PRIORITIES else 'Ważne'
        priority=st.selectbox("Priorytet",PRIORITIES,index=PRIORITIES.index(cp),key=f"edit_priority_{row['_db_id']}")
    if st.button("💾 Zapisz zmiany",use_container_width=True): update_product(conn,row['_db_id'],new_qty,new_real,new_listing,priority); st.rerun()
    if st.button("🗑️ Usuń wybraną pozycję",use_container_width=True): delete_product(conn,row['_db_id']); st.rerun()
else: st.info("Paleta jest pusta. Zrób pierwsze zdjęcie produktu.")

if products:
    st.divider(); st.subheader("📥 Eksport")
    excel=make_excel(products,str(current['name']),float(current['cost']))
    safe=re.sub(r"[^a-zA-Z0-9ąćęłńóśźżĄĆĘŁŃÓŚŹŻ _-]+","",str(current['name'])).strip().replace(' ','_') or 'paleta'
    st.download_button("📊 Pobierz Excel",data=excel,file_name=f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.caption("Dane są zapisane w zewnętrznej bazie PostgreSQL i są wspólne dla urządzeń.")
