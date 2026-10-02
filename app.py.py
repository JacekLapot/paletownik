import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import base64
import json
import os
import threading
import time

from openpyxl import Workbook

FIELDS = [
    "Lp.", "Ilość", "Kategoria", "Marka", "Produkt", "Model", "Stan", "Kompletność",
    "Cena nowego", "Cena używanego", "Realna cena sprzedaży",
    "Cena wystawienia", "Źródło ceny", "Link do oferty", "Uwagi", "Priorytet"
]

MODEL = "gpt-5.6-luna"

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
        "priorytet": {"type": "string", "enum": ["Wysoki", "Normalny", "Niski"]},
        "pewnosc_ident": {"type": "string", "enum": ["Wysoka", "Średnia", "Niska"]}
    },
    "required": [
        "kategoria", "marka", "produkt", "model", "stan", "kompletnosc",
        "cena_nowego", "cena_uzywanego", "realna_cena_sprzedazy",
        "cena_wystawienia", "zrodlo_ceny", "link_do_oferty", "uwagi",
        "priorytet", "pewnosc_ident"
    ]
}


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Wycena palet AI – v2")
        self.root.geometry("1250x760")
        self.rows = []
        self.photo_path = None
        self.api_key = os.getenv("OPENAI_API_KEY", "")
        self.busy = False

        self.build_ui()

    def build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="📷 Wybierz zdjęcie", command=self.choose_photo).pack(side="left", padx=4)
        ttk.Button(top, text="📸 Zrób zdjęcie", command=self.take_photo).pack(side="left", padx=4)
        self.ai_btn = ttk.Button(top, text="🤖 AI: rozpoznaj + wyceń", command=self.ai_analyze)
        self.ai_btn.pack(side="left", padx=4)
        ttk.Button(top, text="➕ Dodaj ręcznie", command=self.add_product).pack(side="left", padx=4)
        ttk.Button(top, text="✏ Edytuj", command=self.edit_selected).pack(side="left", padx=4)
        ttk.Button(top, text="🗑 Usuń", command=self.delete_selected).pack(side="left", padx=4)
        ttk.Button(top, text="💾 Zapisz Excel", command=self.save_excel).pack(side="right", padx=4)
        ttk.Button(top, text="🔑 Klucz API", command=self.api_dialog).pack(side="right", padx=4)

        info = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        info.pack(fill="x")

        ttk.Label(info, text="Koszt palety (zł):").pack(side="left")
        self.cost_var = tk.StringVar(value="360")
        ttk.Entry(info, textvariable=self.cost_var, width=10).pack(side="left", padx=5)

        ttk.Label(info, text="Wartość realna:").pack(side="left", padx=(30, 5))
        self.total_var = tk.StringVar(value="0 zł")
        ttk.Label(info, textvariable=self.total_var, font=("Arial", 12, "bold")).pack(side="left")

        ttk.Label(info, text="Zdjęcie:").pack(side="left", padx=(30, 5))
        self.photo_var = tk.StringVar(value="brak")
        ttk.Label(info, textvariable=self.photo_var).pack(side="left")

        cols = ["Lp.", "Kategoria", "Marka", "Produkt", "Model",
                "Realna cena sprzedaży", "Cena wystawienia", "Priorytet"]
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings", height=23)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=135, anchor="center")
        self.tree.column("Produkt", width=260)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill="x")
        self.status_var = tk.StringVar(
            value="Gotowe. Wybierz zdjęcie lub zrób je kamerą, potem kliknij „AI: rozpoznaj + wyceń”."
        )
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")

    def choose_photo(self):
        p = filedialog.askopenfilename(
            title="Wybierz zdjęcie produktu",
            filetypes=[("Obrazy", "*.jpg *.jpeg *.png *.webp"), ("Wszystkie pliki", "*.*")]
        )
        if p:
            self.photo_path = p
            self.photo_var.set(Path(p).name)
            self.status_var.set("Zdjęcie wybrane. Możesz uruchomić AI.")

    def take_photo(self):
        """Open the default webcam. SPACE captures; ESC cancels."""
        try:
            import cv2
        except ImportError:
            messagebox.showerror(
                "Brak biblioteki",
                "Do zdjęć z kamery potrzebny jest pakiet opencv-python.\n"
                "Uruchom ponownie plik uruchom_wycene_AI_v3.bat, aby go zainstalować."
            )
            return

        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            cap.release()
            # Fallback for systems where CAP_DSHOW is unavailable.
            cap = cv2.VideoCapture(0)

        if not cap.isOpened():
            messagebox.showerror(
                "Kamera niedostępna",
                "Nie udało się otworzyć kamery. Sprawdź, czy kamera jest podłączona "
                "i czy Windows zezwala Pythonowi na dostęp do kamery."
            )
            return

        # Try a useful resolution without requiring it.
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

        window_name = "Wycena palet AI – kamera | SPACE = zdjęcie | ESC = anuluj"
        saved = False

        while True:
            ok, frame = cap.read()
            if not ok:
                break

            cv2.putText(
                frame,
                "SPACE = zrob zdjecie    ESC = anuluj",
                (20, 40),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.8,
                (255, 255, 255),
                2,
                cv2.LINE_AA
            )
            cv2.imshow(window_name, frame)
            key = cv2.waitKey(1) & 0xFF

            if key == 27:  # ESC
                break

            if key == 32:  # SPACE
                photo_dir = Path.home() / "WycenaPalet_AI" / "zdjecia"
                photo_dir.mkdir(parents=True, exist_ok=True)
                filename = f"produkt_{time.strftime('%Y%m%d_%H%M%S')}.jpg"
                photo_path = photo_dir / filename

                if cv2.imwrite(str(photo_path), frame):
                    self.photo_path = str(photo_path)
                    self.photo_var.set(photo_path.name)
                    self.status_var.set(
                        "Zdjęcie wykonane. Możesz uruchomić AI."
                    )
                    saved = True
                break

        cap.release()
        cv2.destroyAllWindows()

        if saved:
            # Bring the main window back to the front.
            self.root.lift()
            self.root.focus_force()

    def api_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Klucz OpenAI API")
        win.geometry("620x180")
        ttk.Label(
            win,
            text="Klucz jest używany tylko podczas działania programu i nie jest zapisywany do pliku."
        ).pack(padx=15, pady=(15, 8))

        entry = ttk.Entry(win, width=75, show="•")
        entry.insert(0, self.api_key)
        entry.pack(padx=15, pady=5)

        def save():
            self.api_key = entry.get().strip()
            if self.api_key:
                self.status_var.set("Klucz API ustawiony dla tej sesji.")
            else:
                self.status_var.set("Klucz API usunięty.")
            win.destroy()

        ttk.Button(win, text="Zapisz na tę sesję", command=save).pack(pady=12)

    def add_product(self):
        self.product_window()

    def edit_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Wycena palet", "Zaznacz produkt.")
            return
        idx = int(self.tree.item(sel[0], "values")[0]) - 1
        self.product_window(self.rows[idx], idx)

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(self.tree.item(sel[0], "values")[0]) - 1
        del self.rows[idx]
        self.renumber()
        self.refresh()
        self.status_var.set("Produkt usunięty.")

    def product_window(self, data=None, idx=None):
        win = tk.Toplevel(self.root)
        win.title("Produkt")
        win.geometry("700x700")

        values = dict(zip(FIELDS, data)) if data else {f: "" for f in FIELDS}
        if not data:
            values["Lp."] = str(len(self.rows) + 1)
            values["Ilość"] = "1"

        entries = {}
        form = ttk.Frame(win, padding=12)
        form.pack(fill="both", expand=True)

        for r, field in enumerate(FIELDS):
            ttk.Label(form, text=field + ":").grid(row=r, column=0, sticky="w", padx=5, pady=4)
            ent = ttk.Entry(form, width=68)
            ent.insert(0, values.get(field, ""))
            ent.grid(row=r, column=1, sticky="ew", padx=5, pady=4)
            entries[field] = ent

        ttk.Label(
            form,
            text="Ilość = liczba identycznych sztuk. Wartość palety będzie liczona jako ilość × realna cena sprzedaży.",
            foreground="gray"
        ).grid(row=len(FIELDS), column=1, sticky="w", padx=5, pady=(2, 8))

        def save():
            row = [entries[f].get().strip() for f in FIELDS]
            try:
                qty = int(row[1])
                if qty < 1:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Błędna ilość", "Ilość musi być liczbą całkowitą większą od 0.")
                return
            if not row[3]:
                messagebox.showwarning("Brak danych", "Podaj nazwę produktu.")
                return
            if idx is None:
                self.rows.append(row)
            else:
                self.rows[idx] = row
            self.renumber()
            self.refresh()
            win.destroy()

        ttk.Button(form, text="Zapisz produkt", command=save).grid(
            row=len(FIELDS), column=1, sticky="e", pady=12
        )

    def renumber(self):
        for i, row in enumerate(self.rows, 1):
            row[0] = str(i)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.rows:
            vals = [row[0], row[1], row[2], row[4], row[5], row[10], row[11], row[15]]
            self.tree.insert("", "end", values=vals)
        self.update_totals()

    def update_totals(self):
        total = 0.0
        for r in self.rows:
            try:
                qty = float(str(r[1]).replace(",", ".").strip() or "1")
                unit = float(str(r[10]).replace(",", ".").replace("zł", "").strip())
                total += qty * unit
            except Exception:
                pass
        self.total_var.set(f"{total:,.2f} zł".replace(",", " "))

    def ai_analyze(self):
        if self.busy:
            return
        if not self.photo_path:
            messagebox.showinfo("AI", "Najpierw wybierz zdjęcie produktu.")
            return
        if not self.api_key:
            messagebox.showinfo(
                "AI",
                "Najpierw kliknij „🔑 Klucz API” i wklej swój klucz OpenAI."
            )
            return

        self.busy = True
        self.ai_btn.config(state="disabled")
        self.status_var.set("AI analizuje zdjęcie i sprawdza aktualne ceny...")
        threading.Thread(target=self._ai_worker, daemon=True).start()

    def _ai_worker(self):
        try:
            from openai import OpenAI

            client = OpenAI(api_key=self.api_key)

            mime = {
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".png": "image/png",
                ".webp": "image/webp"
            }.get(Path(self.photo_path).suffix.lower(), "image/jpeg")

            with open(self.photo_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("utf-8")

            prompt = """
Jesteś asystentem do wyceny produktów z palet zwrotów konsumenckich w Polsce.

Przeanalizuj zdjęcie produktu. Zidentyfikuj markę, produkt i model możliwie dokładnie.
Jeżeli modelu nie da się potwierdzić, wpisz pusty string zamiast zgadywać.

Następnie użyj wyszukiwania internetowego, aby sprawdzić aktualne ceny w Polsce.
Priorytet źródeł: Allegro, OLX, Ceneo, sklepy polskie, oficjalny sklep producenta.
Nie traktuj ceny katalogowej jako realnej ceny sprzedaży.

Zasady wyceny:
- "Cena nowego" = realna aktualna cena nowego egzemplarza, nie MSRP.
- "Cena używanego" = typowa cena kompletnego, sprawnego używanego egzemplarza.
- "Realna cena sprzedaży" = konserwatywna kwota, którą można realnie uzyskać przy sprzedaży z palety.
- "Cena wystawienia" = cena początkowa, trochę wyższa od realnej ceny sprzedaży.
- Jeśli rynek jest nasycony tanimi ofertami, obniż wycenę.
- Nie zawyżaj wartości przez pojedynczą drogą ofertę.
- Jeśli nie masz pewności co do identyfikacji lub cen, zaznacz to w uwagach.
- Stan i kompletność na podstawie jednego zdjęcia oznacz jako "Nowy / do sprawdzenia" lub "Nieznany / do sprawdzenia", jeżeli nie da się tego potwierdzić.
- Zwróć wyłącznie dane zgodne ze schematem JSON.
"""

            response = client.responses.create(
                model=MODEL,
                tools=[{"type": "web_search"}],
                input=[{
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {
                            "type": "input_image",
                            "image_url": f"data:{mime};base64,{b64}",
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

            data = json.loads(response.output_text)
            self.root.after(0, lambda: self._apply_ai(data))

        except Exception as e:
            msg = str(e)
            self.root.after(0, lambda: self._ai_error(msg))

    def _norm(self, value):
        import re
        value = str(value or "").lower().strip()
        value = re.sub(r"[^a-z0-9ąćęłńóśźż]+", "", value)
        return value

    def _find_duplicate(self, data):
        """
        Returns index of an existing row if AI identification is sufficiently
        strong. Model is the strongest signal; otherwise brand + product.
        """
        new_model = self._norm(data.get("model"))
        new_brand = self._norm(data.get("marka"))
        new_product = self._norm(data.get("produkt"))

        if not new_model and not new_product:
            return None

        for idx, row in enumerate(self.rows):
            # V5 layout: Lp, Ilość, Kategoria, Marka, Produkt, Model...
            old_brand = self._norm(row[3])
            old_product = self._norm(row[4])
            old_model = self._norm(row[5])

            if new_model and old_model and new_model == old_model:
                # Model match; require brand match when both are known.
                if new_brand and old_brand and new_brand != old_brand:
                    continue
                return idx

            if (new_brand and old_brand and new_brand == old_brand and
                    new_product and old_product and new_product == old_product):
                return idx

        return None

    def _increment_duplicate(self, idx, data):
        row = self.rows[idx]
        try:
            old_qty = int(row[1] or "1")
        except ValueError:
            old_qty = 1

        row[1] = str(old_qty + 1)

        # If the new AI result has better price/source information, don't
        # silently overwrite the user's existing values. We only add a note.
        note = row[13] if len(row) > 13 else ""
        extra = f" | Kolejne rozpoznanie produktu: +1 szt. (łącznie {old_qty + 1})"
        if extra not in note:
            row[13] = note + extra

        self.refresh()
        self.status_var.set(
            f"Duplikat: {row[3]} {row[5]} → ilość zwiększona do {row[1]} szt."
        )

    def _apply_ai(self, data):
        self.busy = False
        self.ai_btn.config(state="normal")

        duplicate_idx = self._find_duplicate(data)
        if duplicate_idx is not None:
            existing = self.rows[duplicate_idx]
            try:
                old_qty = int(existing[1] or "1")
            except ValueError:
                old_qty = 1

            answer = messagebox.askyesno(
                "Wykryto duplikat",
                f"AI rozpoznało produkt, który już jest na liście:\n\n"
                f"{existing[3]} {existing[5]}\n"
                f"Obecna ilość: {old_qty} szt.\n\n"
                f"Czy dodać 1 kolejną sztukę?"
            )

            if answer:
                self._increment_duplicate(duplicate_idx, data)
            else:
                self.status_var.set("Duplikat znaleziony — nic nie zmieniono.")
            return

        row = [
            str(len(self.rows) + 1),
            "1",
            data["kategoria"],
            data["marka"],
            data["produkt"],
            data["model"],
            data["stan"],
            data["kompletnosc"],
            str(data["cena_nowego"]),
            str(data["cena_uzywanego"]),
            str(data["realna_cena_sprzedazy"]),
            str(data["cena_wystawienia"]),
            data["zrodlo_ceny"],
            data["link_do_oferty"],
            data["uwagi"] + f" | Pewność identyfikacji: {data['pewnosc_ident']}",
            data["priorytet"]
        ]

        self.rows.append(row)
        self.refresh()
        self.status_var.set(
            f"AI dodało produkt: {data['marka']} {data['model']} | "
            f"realna sprzedaż: {data['realna_cena_sprzedazy']} zł"
        )

        # Otwórz edycję, żeby użytkownik mógł zatwierdzić/poprawić wynik.
        self.root.after(100, lambda: self.product_window(self.rows[-1], len(self.rows) - 1))

    def _ai_error(self, msg):
        self.busy = False
        self.ai_btn.config(state="normal")
        self.status_var.set("Błąd AI.")
        messagebox.showerror(
            "AI – błąd",
            "Nie udało się wykonać analizy.\n\n"
            + msg
            + "\n\nSprawdź klucz API, połączenie z internetem i instalację pakietu openai."
        )

    def save_excel(self):
        if not self.rows:
            messagebox.showinfo("Wycena palet", "Brak produktów.")
            return

        path = filedialog.asksaveasfilename(
            title="Zapisz arkusz",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Produkty"
        ws.append(FIELDS)

        from openpyxl.styles import Font, PatternFill, Alignment
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="D9EAF7")
            c.alignment = Alignment(horizontal="center")

        for row in self.rows:
            ws.append(row)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        summary = wb.create_sheet("Podsumowanie")
        summary["A1"] = "PODSUMOWANIE PALETY"
        summary["A1"].font = ws["A1"].font.copy(bold=True, size=16)
        summary["A3"] = "Liczba produktów"
        summary["B3"] = len(self.rows)

        try:
            cost = float(self.cost_var.get().replace(",", "."))
        except Exception:
            cost = 0

        total = 0
        for r in self.rows:
            try:
                qty = float(str(r[1]).replace(",", ".").strip() or "1")
                unit = float(str(r[10]).replace(",", ".").replace("zł", "").strip())
                total += qty * unit
            except Exception:
                pass

        summary["A4"] = "Koszt palety"
        summary["B4"] = cost
        summary["A5"] = "Realna wartość sprzedaży"
        summary["B5"] = total
        summary["A6"] = "Nadwyżka przed kosztami sprzedaży"
        summary["B6"] = total - cost

        for col, width in {"A": 38, "B": 20}.items():
            summary.column_dimensions[col].width = width

        wb.save(path)
        messagebox.showinfo("Zapisano", f"Zapisano plik:\n{path}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
