import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Dict, List, Tuple, Optional
from tksheet import Sheet

# ------------------------------------------------------------
# Editor JSON Mapping — v0.3 (Tkinter + tksheet)
# ------------------------------------------------------------
# - "Tabella" a tutta larghezza (pane ridimensionabile) + colonne richieste
# - Editing direttamente in tabella
# - Filtri per colonna (stile Excel: combobox per domini, testo/operatore per numeri)
# - Salvataggio con backup .bak
# - Mapping campi trigger[0] (estendibile a N triggers)
# ------------------------------------------------------------
APP_TITLE = "Editor JSON Mapping — Dragflow (v0.3)"

# ----------------- Utility nested path -----------------

def safe_get(d: Dict[str, Any], path: List[str]):
    cur: Any = d
    for p in path:
        if p.endswith("]"):
            key, idx = p[:-1].split("[")
            cur = cur.get(key, []) if isinstance(cur, dict) else []
            try:
                cur = cur[int(idx)]
            except (ValueError, IndexError, TypeError):
                return None
        else:
            if not isinstance(cur, dict):
                return None
            cur = cur.get(p)
        if cur is None:
            return None
    return cur


def safe_set(d: Dict[str, Any], path: List[str], value: Any):
    cur: Any = d
    for i, p in enumerate(path):
        last = i == len(path) - 1
        if p.endswith("]"):
            key, idx = p[:-1].split("[")
            idx = int(idx)
            if key not in cur or not isinstance(cur[key], list):
                cur[key] = []
            while len(cur[key]) <= idx:
                cur[key].append({})
            if last:
                cur[key][idx] = value
                return True
            cur = cur[key][idx]
        else:
            if last:
                cur[p] = value
                return True
            if p not in cur or not isinstance(cur[p], dict):
                cur[p] = {}
            cur = cur[p]
    return False

# ----------------- App -----------------

HEADERS = [
    "type",                # top-level type
    "label",
    "unit",
    "trigger type",        # triggers[0].type
    "level",               # triggers[0].level
    "mode",                # triggers[0].mode
    "min interval ms",     # triggers[0].minIntervalMs
    "skip first n changes",# triggers[0].skipFirstNChanges
    "change mask",         # triggers[0].changeMask (string domain)
    "deadband",            # triggers[0].deadband OR deadbandPercent
    "deadband type"        # ABS (deadband) | PERC (deadbandPercent)
]

IDX = {h:i for i,h in enumerate(HEADERS)}

class MappingEditor(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill=tk.BOTH, expand=True)

        self.file_path: Optional[str] = None
        self.data: Optional[Dict[str, Any]] = None
        self.property_items: List[Tuple[str, Dict[str, Any]]] = []

        # Domini
        self.domain_types: List[str] = []          # top-level type
        self.domain_units: List[str] = []
        self.domain_trig_types: List[str] = []     # trigger type
        self.domain_trig_modes: List[str] = []     # trigger mode
        self.domain_change_masks: List[str] = []   # changeMask strings

        # dataset completo (tutte le righe) e dataset filtrato per la tabella
        self.rows_all: List[List[Any]] = []
        self.rows_view: List[List[Any]] = []
        self.row_to_path: List[str] = []  # mappa indice rows_all -> path
        self.view_index_map: List[int] = []  # mappa indice rows_view -> indice in rows_all

        self._build_ui()

    # ---------------- UI -----------------
    def _build_ui(self):
        if Sheet is None:
            f = ttk.Frame(self)
            f.pack(fill="both", expand=True, padx=12, pady=12)
            ttk.Label(f, text=(
                "Manca la dipendenza 'tksheet'.\n"
                "Installa con: pip install tksheet",
            )).pack()
            return

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Toolbar
        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=6, pady=(6,3))
        ttk.Button(toolbar, text="Apri…", command=self.on_open).pack(side=tk.LEFT)
        self.btn_save = ttk.Button(toolbar, text="Salva", command=self.on_save, state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=(6, 0))
        self.btn_save_as = ttk.Button(toolbar, text="Salva come…", command=self.on_save_as, state=tk.DISABLED)
        self.btn_save_as.pack(side=tk.LEFT, padx=(6, 0))

        # Filter row (per colonna)
        self.filter_bar = ttk.Frame(self)
        self.filter_bar.grid(row=1, column=0, sticky="ew", padx=6, pady=(0,3))
        self._build_filters(self.filter_bar)

        # Paned (tabella = principale)
        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0,6))

        # left: sheet
        self.left = ttk.Frame(paned)
        self.left.rowconfigure(0, weight=1)
        self.left.columnconfigure(0, weight=1)
        self.sheet = Sheet(self.left, headers=HEADERS)
        self.sheet.enable_bindings((
            "single_select",
            "row_select",
            "column_select",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "edit_cell",
            "copy",
            "cut",
            "paste",
            "delete",
            "undo",
            "redo",
            "drag_select",
            "drop_select",
            "column_width_resize",
            "double_click_column_resize",
            "sort",
        ))
        self.sheet.grid(row=0, column=0, sticky="nsew")

        # right: dettagli minimi (opzionale)
        self.right = ttk.Frame(paned)
        ttk.Label(self.right, text="Dettagli (opz.)", foreground="#666").pack(anchor="w", padx=8, pady=8)
        ttk.Label(self.right, text="La maggior parte dell'editing avviene direttamente in tabella.\n"
                                  "A destra puoi inserire funzioni future (diff, preset, ecc.)",
                  wraplength=260, foreground="#666").pack(anchor="w", padx=8)

        paned.add(self.left, weight=4)
        paned.add(self.right, weight=1)
        try:
            self.after(50, lambda: paned.sashpos(0, max(500, self.winfo_width() - 360)))
        except Exception:
            pass

        # Statusbar
        self.status = tk.StringVar(value="Apri un file JSON di mapping…")
        ttk.Label(self, textvariable=self.status, anchor="w").grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 6))

    def _build_filters(self, parent: ttk.Frame):
        # widget per ogni colonna
        self.filters: Dict[str, tk.Variable] = {}
        grid = ttk.Frame(parent)
        grid.pack(fill="x")
        # costruiamo una riga di widget con proporzioni
        coldefs = [
            ("type", "combo"),
            ("label", "text"),
            ("unit", "combo"),
            ("trigger type", "combo"),
            ("level", "text"),             # supporta >, >=, <, <=, =
            ("mode", "combo"),
            ("min interval ms", "text"),    # supporta operatori
            ("skip first n changes", "text"),
            ("change mask", "combo"),
            ("deadband", "text"),           # supporta operatori
            ("deadband type", "combo"),
        ]
        for i, (name, kind) in enumerate(coldefs):
            grid.columnconfigure(i, weight=1)
            lbl = ttk.Label(grid, text=name)
            lbl.grid(row=0, column=i, sticky="w", padx=2)
            if kind == "combo":
                var = tk.StringVar()
                cb = ttk.Combobox(grid, textvariable=var, state="readonly")
                cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
                cb.grid(row=1, column=i, sticky="ew", padx=2, pady=2)
                self.filters[name] = var
            else:
                var = tk.StringVar()
                ent = ttk.Entry(grid, textvariable=var)
                ent.bind("<KeyRelease>", lambda e: self.apply_filters())
                ent.grid(row=1, column=i, sticky="ew", padx=2, pady=2)
                self.filters[name] = var

    # -------------- File ops ---------------
    def on_open(self):
        path = filedialog.askopenfilename(title="Apri mapping JSON", filetypes=[("File JSON", "*.json"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
            self.file_path = path
            self.btn_save.config(state=tk.NORMAL)
            self.btn_save_as.config(state=tk.NORMAL)
            self.status.set(f"Caricato: {os.path.basename(path)}")
            self._reindex()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Errore apertura file:\n{e}")

    def on_save(self):
        if not (self.file_path and self.data):
            return
        try:
            # Commit dalla tabella al JSON (con validazioni base)
            self._commit_table_to_json()

            bak = self.file_path + ".bak"
            try:
                if os.path.exists(bak):
                    os.remove(bak)
                if os.path.exists(self.file_path):
                    os.replace(self.file_path, bak)
            except Exception:
                pass
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
            self.status.set(f"Salvato: {self.file_path}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Errore salvataggio:\n{e}")

    def on_save_as(self):
        if not self.data:
            return
        path = filedialog.asksaveasfilename(title="Salva come", defaultextension=".json", filetypes=[("File JSON", "*.json")])
        if not path:
            return
        self.file_path = path
        self.on_save()

    # -------------- Index & domains ---------------
    def _reindex(self):
        self.property_items.clear()
        if not self.data:
            return
        root = self.data.get("json") if isinstance(self.data, dict) else None
        if not root:
            root = self.data
        props = (root or {}).get("properties", {})
        for path, obj in props.items():
            if isinstance(obj, dict):
                self.property_items.append((path, obj))

        # domini
        def collect_top(key: str) -> List[str]:
            vals: List[str] = []
            for _, o in self.property_items:
                v = o.get(key)
                if isinstance(v, str) and v not in vals:
                    vals.append(v)
            return sorted(vals, key=str.lower)

        def collect_trig(key: str) -> List[str]:
            vals: List[str] = []
            for _, o in self.property_items:
                for t in (o.get("sendPolicy",{}).get("triggers",[]) or []):
                    v = t.get(key)
                    if isinstance(v, str) and v not in vals:
                        vals.append(v)
            return sorted(vals, key=str.lower)

        self.domain_types = collect_top("type") or ["boolean", "integer", "double", "string"]
        self.domain_units = collect_top("unit")
        self.domain_trig_types = collect_trig("type") or ["onchange", "periodic", "mixed"]
        self.domain_trig_modes = collect_trig("mode")
        self.domain_change_masks = collect_trig("changeMask") or [""]

        # popola tabella dati completa
        self._build_rows_all()
        # inizializza filtri
        self._refresh_filter_widgets()
        # applica filtri (iniziale: nessuno) e mostra tabella
        self.apply_filters()

    def _build_rows_all(self):
        self.rows_all.clear()
        self.row_to_path.clear()
        for path, obj in self.property_items:
            tr = (obj.get("sendPolicy",{}).get("triggers",[]) or [{}])[0]
            # deadband & type
            db_val: Optional[float] = None
            db_type = ""
            if "deadbandPercent" in tr and tr.get("deadbandPercent") is not None:
                db_val = tr.get("deadbandPercent")
                db_type = "PERC"
            elif "deadband" in tr and tr.get("deadband") is not None:
                db_val = tr.get("deadband")
                db_type = "ABS"

            row = [
                obj.get("type", ""),               # type (top)
                obj.get("label", ""),               # label
                obj.get("unit", ""),                # unit
                tr.get("type", ""),                 # trigger type
                tr.get("level", ""),                # level
                tr.get("mode", ""),                 # mode
                tr.get("minIntervalMs", ""),        # min interval
                tr.get("skipFirstNChanges", ""),    # skip first
                tr.get("changeMask", ""),           # change mask
                "" if db_val is None else db_val,    # deadband
                db_type,                               # deadband type
            ]
            self.rows_all.append(row)
            self.row_to_path.append(path)

    def _refresh_filter_widgets(self):
        # carica domini nei combo
        def set_combo(name: str, values: List[str]):
            w = self._find_filter_widget(name)
            if isinstance(w, ttk.Combobox):
                w["values"] = [""] + values
                w.set("")
        set_combo("type", self.domain_types)
        set_combo("unit", self.domain_units)
        set_combo("trigger type", self.domain_trig_types)
        set_combo("mode", self.domain_trig_modes)
        set_combo("change mask", self.domain_change_masks)
        set_combo("deadband type", ["ABS", "PERC"]) 

    def _find_filter_widget(self, name: str):
        # cerca in self.filter_bar
        for child in self.filter_bar.winfo_children():
            for sub in child.winfo_children():
                if isinstance(sub, ttk.Label) and sub.cget("text") == name:
                    # l'entry/combo è riga successiva stessa colonna
                    info = sub.grid_info()
                    for peer in child.winfo_children():
                        if peer.grid_info().get("row") == info["row"] + 1 and peer.grid_info().get("column") == info["column"]:
                            return peer
        return None

    # -------------- Filtering ---------------
    def apply_filters(self):
        def match_text(val: Any, query: str) -> bool:
            if query.strip() == "":
                return True
            s = str(val).lower()
            q = query.lower()
            # operatori numerici semplici
            for op in (">=", "<=", ">", "<", "="):
                if q.startswith(op):
                    try:
                        num = float(q[len(op):].strip())
                        v = float(val)
                        return eval(f"v {op} num")
                    except Exception:
                        return False
            return q in s

        # leggi valori filtro
        fv = {h: self._get_filter_value(h) for h in HEADERS}

        self.rows_view.clear()
        self.view_index_map.clear()
        for i, row in enumerate(self.rows_all):
            ok = True
            if not match_text(row[IDX["type"]], fv["type"]): ok = False
            if not match_text(row[IDX["label"]], fv["label"]): ok = False
            if fv["unit"] and row[IDX["unit"]] != fv["unit"]: ok = False
            if fv["trigger type"] and row[IDX["trigger type"]] != fv["trigger type"]: ok = False
            if not match_text(row[IDX["level"]], fv["level"]): ok = False
            if fv["mode"] and row[IDX["mode"]] != fv["mode"]: ok = False
            if not match_text(row[IDX["min interval ms"]], fv["min interval ms"]): ok = False
            if not match_text(row[IDX["skip first n changes"]], fv["skip first n changes"]): ok = False
            if fv["change mask"] and row[IDX["change mask"]] != fv["change mask"]: ok = False
            if not match_text(row[IDX["deadband"]], fv["deadband"]): ok = False
            if fv["deadband type"] and row[IDX["deadband type"]] != fv["deadband type"]: ok = False
            if ok:
                self.rows_view.append(list(row))
                self.view_index_map.append(i)

        # aggiorna sheet
        self.sheet.set_sheet_data(self.rows_view, reset_col_positions=True, reset_row_positions=True)
        self.sheet.headers(HEADERS)

    def _get_filter_value(self, name: str) -> str:
        w = self._find_filter_widget(name)
        if isinstance(w, ttk.Combobox) or isinstance(w, ttk.Entry):
            return w.get()
        return ""

    # -------------- Commit ---------------
    def _commit_table_to_json(self):
        if not self.data:
            return
        # sincronizza rows_view -> rows_all (l'utente può aver editato nella vista filtrata)
        # leggiamo i dati attuali della sheet
        current = self.sheet.get_sheet_data(return_copy=True)
        # aggiorna rows_view con current e propaga a rows_all per le righe visibili
        for idx_view, row_vals in enumerate(current):
            idx_all = self.view_index_map[idx_view]
            self.rows_all[idx_all] = row_vals

        # applica rows_all nel JSON
        root = self.data.get("json") or self.data
        props = root.get("properties", {})

        errors: List[str] = []
        for i, row in enumerate(self.rows_all):
            path = self.row_to_path[i]
            obj = props.get(path, {})

            # top-level
            obj_type = str(row[IDX["type"]]).strip()
            obj_label = str(row[IDX["label"]]).strip()
            obj_unit  = str(row[IDX["unit"]]).strip()

            if obj_type: obj["type"] = obj_type
            if obj_label != "": obj["label"] = obj_label
            if obj_unit != "": obj["unit"] = obj_unit
            elif "unit" in obj: # consenti vuoto -> rimuovi
                obj.pop("unit", None)

            # trigger[0]
            tr_list = obj.setdefault("sendPolicy", {}).setdefault("triggers", [])
            if not tr_list:
                tr_list.append({})
            tr = tr_list[0]

            trig_type = str(row[IDX["trigger type"]]).strip()
            level = row[IDX["level"]]
            mode  = str(row[IDX["mode"]]).strip()
            minint = row[IDX["min interval ms"]]
            skipn  = row[IDX["skip first n changes"]]
            cmask  = str(row[IDX["change mask"]]).strip()
            db     = row[IDX["deadband"]]
            dbt    = str(row[IDX["deadband type"]]).strip().upper()

            if trig_type: tr["type"] = trig_type
            if mode != "": tr["mode"] = mode

            # level (int)
            try:
                if str(level).strip() == "":
                    tr.pop("level", None)
                else:
                    lv = int(level)
                    tr["level"] = lv
            except Exception:
                errors.append(f"{path}: 'level' non valido")

            # min interval (>=0)
            try:
                if str(minint).strip() == "":
                    tr.pop("minIntervalMs", None)
                else:
                    mi = int(minint)
                    if mi < 0: raise ValueError
                    tr["minIntervalMs"] = mi
            except Exception:
                errors.append(f"{path}: 'min interval ms' non valido (>=0)")

            # skip first n changes (>=0)
            try:
                if str(skipn).strip() == "":
                    tr.pop("skipFirstNChanges", None)
                else:
                    sk = int(skipn)
                    if sk < 0: raise ValueError
                    tr["skipFirstNChanges"] = sk
            except Exception:
                errors.append(f"{path}: 'skip first n changes' non valido (>=0)")

            # change mask (string dominio, consenti vuoto = rimuovi)
            if cmask == "":
                tr.pop("changeMask", None)
            else:
                tr["changeMask"] = cmask

            # deadband + tipo
            try:
                if str(dbt) == "":
                    tr.pop("deadband", None)
                    tr.pop("deadbandPercent", None)
                elif dbt == "ABS":
                    # float >= 0
                    if str(db).strip() == "":
                        tr.pop("deadband", None)
                        tr.pop("deadbandPercent", None)
                    else:
                        v = float(db)
                        if v < 0: raise ValueError
                        tr["deadband"] = v
                        tr.pop("deadbandPercent", None)
                elif dbt == "PERC":
                    if str(db).strip() == "":
                        tr.pop("deadband", None)
                        tr.pop("deadbandPercent", None)
                    else:
                        v = float(db)
                        if not (0.0 <= v <= 100.0):
                            raise ValueError
                        tr["deadbandPercent"] = v
                        tr.pop("deadband", None)
                else:
                    errors.append(f"{path}: 'deadband type' deve essere ABS o PERC")
            except Exception:
                errors.append(f"{path}: 'deadband' non valido")

            # riassegna obj nel props (non strettamente necessario in-place)
            props[path] = obj

        if errors:
            raise ValueError("\n".join(errors))

    # ---------------- Form helpers (non usati in v0.3) ----------------

# ---------------- main ----------------

def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1300x720")  # finestra più ampia, ridimensionabile
    root.minsize(1024, 600)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = MappingEditor(root)

    # auto-load se il file è a fianco dello script
    default_path = os.path.join(os.path.dirname(__file__), "2500053_Mapping.json")
    if os.path.exists(default_path):
        try:
            with open(default_path, "r", encoding="utf-8") as f:
                app.data = json.load(f)
            app.file_path = default_path
            if Sheet is None:
                pass
            else:
                app.btn_save.config(state=tk.NORMAL)
                app.btn_save_as.config(state=tk.NORMAL)
                app.status.set(f"Caricato: {os.path.basename(default_path)}")
                app._reindex()
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"Apertura iniziale fallita: {e}")

    root.mainloop()


if __name__ == "__main__":
    main()
