# -*- coding: utf-8 -*-
"""
Editor JSON Mapping — Dragflow (v0.4)
Requisiti:  pip install tksheet
"""
import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Dict, List, Tuple, Optional

from tksheet import Sheet

APP_TITLE = "Editor JSON Mapping — Dragflow (v0.4)"

# ----------------- Utility (non usate ovunque, ma comode se servono) -----------------
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

# ----------------- Tabella -----------------
HEADERS = [
    "type",                # top-level type
    "label",
    "unit",
    "trigger type",        # triggers[0].type
    "level",               # triggers[0].level
    "mode",                # triggers[0].mode
    "min interval ms",     # triggers[0].minIntervalMs
    "skip first n changes",# triggers[0].skipFirstNChanges
    "change mask",         # triggers[0].changeMask
    "deadband",            # triggers[0].deadband OR deadbandPercent
    "deadband type"        # ABS | PERC
]
IDX = {h: i for i, h in enumerate(HEADERS)}
NUMERIC_COLS = {IDX["level"], IDX["min interval ms"], IDX["skip first n changes"], IDX["deadband"]}


class MappingEditor(ttk.Frame):
    # ---------- helper: frecce header ----------
    def _refresh_headers_with_arrow(self):
        hdrs = HEADERS.copy()
        if self._last_sort_col is not None:
            arrow = "▲" if self._last_sort_asc else "▼"
            hdrs[self._last_sort_col] = f"{hdrs[self._last_sort_col]} {arrow}"
        self.sheet.headers(hdrs)

    # ---------- helper: estrai colonna da evento (se mai servisse) ----------
    def _event_to_col(self, event) -> Optional[int]:
        if isinstance(event, dict):
            for key in ("column", "col", "c", "index", "selected", "selected_column"):
                v = event.get(key, None)
                if isinstance(v, int):
                    return v
                if isinstance(v, (list, tuple)) and v and isinstance(v[0], int):
                    return v[0]
            for v in event.values():
                if isinstance(v, int):
                    return v
                if isinstance(v, (list, tuple)) and v and isinstance(v[0], int):
                    return v[0]
        try:
            for attr in ("column", "col", "c", "index"):
                if hasattr(event, attr):
                    v = getattr(event, attr)
                    if isinstance(v, int):
                        return v
        except Exception:
            pass
        if isinstance(event, (tuple, list)):
            for item in event:
                if isinstance(item, int):
                    return item
                if isinstance(item, (list, tuple)) and item and isinstance(item[0], int):
                    return item[0]
        try:
            cols = []
            if hasattr(self.sheet, "get_selected_columns"):
                cols = self.sheet.get_selected_columns() or []
            if not cols and hasattr(self.sheet, "get_selected_cols"):
                cols = self.sheet.get_selected_cols() or []
            if cols:
                return cols[0]
        except Exception:
            pass
        return None

    # ---------- init ----------
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill=tk.BOTH, expand=True)

        self.file_path: Optional[str] = None
        self.data: Optional[Dict[str, Any]] = None
        self.property_items: List[Tuple[str, Dict[str, Any]]] = []

        # Domini per i combo filtro
        self.domain_types: List[str] = []
        self.domain_units: List[str] = []
        self.domain_trig_types: List[str] = []
        self.domain_trig_modes: List[str] = []
        self.domain_change_masks: List[str] = []

        # Dataset completo e vista filtrata
        self.rows_all: List[List[Any]] = []
        self.rows_view: List[List[Any]] = []
        self.row_to_path: List[str] = []
        self.view_index_map: List[int] = []

        # Stato ordinamento globale
        self._last_sort_col: Optional[int] = None
        self._last_sort_asc: bool = True

        # Overlay combobox per deadband type
        self._overlay_combo: Optional[ttk.Combobox] = None
        self._overlay_cell: Optional[Tuple[int, int]] = None
        self._prev_cell_value: Optional[str] = None

        # Stato sort per colonna (pulsanti)
        self.sort_dir_by_col: Dict[int, bool] = {}
        self.sort_buttons: Dict[int, ttk.Button] = {}

        # Meta (GUI)
        self.var_name = tk.StringVar(value="")
        self.var_instance = tk.StringVar(value="")

        self._build_ui()

    # ---------------- UI -----------------
    def _build_ui(self):
        if Sheet is None:
            f = ttk.Frame(self)
            f.pack(fill="both", expand=True, padx=12, pady=12)
            ttk.Label(f, text=(
                """Manca la dipendenza 'tksheet'.
Installa con: pip install tksheet"""
            )).pack()
            return

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Toolbar
        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 3))
        ttk.Button(toolbar, text="Apri…", command=self.on_open).pack(side=tk.LEFT)
        self.btn_save = ttk.Button(toolbar, text="Salva", command=self.on_save, state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=(6, 0))
        self.btn_save_as = ttk.Button(toolbar, text="Salva come…", command=self.on_save_as, state=tk.DISABLED)
        self.btn_save_as.pack(side=tk.LEFT, padx=(6, 0))

        # Stili compatti
        style = ttk.Style(self)
        style.configure(".", font=("Segoe UI", 9))
        style.configure("Small.TEntry", padding=(2, 1))
        style.configure("Small.TCombobox", padding=(2, 1))
        style.configure("Small.TButton", padding=(2, 0))

        # Filtro per colonna
        self.filter_bar = ttk.Frame(self)
        self.filter_bar.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 3))
        self._build_filters(self.filter_bar)

        # Paned (tabella principale)
        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 6))

        # left: tksheet
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
        ))
        # Font compatti per tksheet (usa tuple a 3 elementi)
        try:
            self.sheet.set_options(
                table_font=("Segoe UI", 9, "normal"),
                header_font=("Segoe UI", 9, "normal"),
            )
            self.sheet.set_header_height(20)  # opzionale, per header più basso
        except Exception:
            # Fallback sicuro (se la tua versione di tksheet vuole un altro font)
            try:
                self.sheet.set_options(
                    table_font=("Arial", 9, "normal"),
                    header_font=("Arial", 9, "normal"),
                )
            except Exception:
                pass

        self.sheet.grid(row=0, column=0, sticky="nsew")

        # Bind di editing (NO bind header: usiamo i pulsanti sort)
        try:
            self.sheet.extra_bind("begin_edit_cell", self._on_begin_edit_cell)
            self.sheet.extra_bind("end_edit_cell", self._on_end_edit_cell)
            self.sheet.extra_bind("double_click_cell", self._on_double_click_cell)
        except Exception:
            pass

        # right: pannello meta + note
        self.right = ttk.Frame(paned)

        meta = ttk.LabelFrame(self.right, text="Meta")
        meta.pack(fill="x", padx=8, pady=(8, 4))
        meta.columnconfigure(1, weight=1)
        ttk.Label(meta, text="name").grid(row=0, column=0, sticky="w", padx=(4, 2), pady=2)
        ttk.Entry(meta, textvariable=self.var_name, width=28, style="Small.TEntry").grid(
            row=0, column=1, sticky="ew", padx=(2, 4), pady=2
        )
        ttk.Label(meta, text="instanceOf").grid(row=1, column=0, sticky="w", padx=(4, 2), pady=2)
        ttk.Entry(meta, textvariable=self.var_instance, width=28, style="Small.TEntry").grid(
            row=1, column=1, sticky="ew", padx=(2, 4), pady=2
        )

        ttk.Label(self.right, text="Dettagli (opz.)", foreground="#666").pack(anchor="w", padx=8, pady=8)
        ttk.Label(
            self.right,
            text=(
                """La maggior parte dell'editing avviene direttamente in tabella.
A destra puoi inserire funzioni future (diff, preset, ecc.)"""
            ),
            wraplength=260,
            foreground="#666",
        ).pack(anchor="w", padx=8)

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
        self.filters: Dict[str, tk.Variable] = {}
        grid = ttk.Frame(parent)
        grid.pack(fill="x")

        coldefs = [
            ("type", "combo"),
            ("label", "text"),
            ("unit", "combo"),
            ("trigger type", "combo"),
            ("level", "text"),
            ("mode", "combo"),
            ("min interval ms", "text"),
            ("skip first n changes", "text"),
            ("change mask", "combo"),
            ("deadband", "text"),
            ("deadband type", "combo"),
        ]

        for i, (name, kind) in enumerate(coldefs):
            grid.columnconfigure(i, weight=1)
            ttk.Label(grid, text=name).grid(row=0, column=i, sticky="w", padx=2)

            cell = ttk.Frame(grid)
            cell.grid(row=1, column=i, sticky="ew", padx=2, pady=2)
            cell.columnconfigure(0, weight=1)

            if kind == "combo":
                var = tk.StringVar()
                inp = ttk.Combobox(cell, textvariable=var, state="readonly", width=10, style="Small.TCombobox")
                inp.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
            else:
                var = tk.StringVar()
                inp = ttk.Entry(cell, textvariable=var, width=12, style="Small.TEntry")
                inp.bind("<KeyRelease>", lambda e: self.apply_filters())
            inp.grid(row=0, column=0, sticky="ew")

            self.filters[name] = var

            col_index = i
            btn = ttk.Button(cell, text="A→Z", width=3, style="Small.TButton",
                             command=lambda c=col_index: self._on_sort_button(c))
            btn.grid(row=0, column=1, sticky="e", padx=(4, 0))
            self.sort_dir_by_col[col_index] = True
            self.sort_buttons[col_index] = btn

    # -------------- File ops ---------------
    def on_open(self):
        path = filedialog.askopenfilename(
            title="Apri mapping JSON",
            filetypes=[("File JSON", "*.json"), ("Tutti i file", "*.*")]
        )
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
            messagebox.showerror(APP_TITLE, f"""Errore apertura file:
{e}""")

    def on_save(self):
        if not (self.file_path and self.data):
            return
        try:
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
            messagebox.showerror(APP_TITLE, f"""Errore salvataggio:
{e}""")

    def on_save_as(self):
        if not self.data:
            return
        path = filedialog.asksaveasfilename(
            title="Salva come",
            defaultextension=".json",
            filetypes=[("File JSON", "*.json")]
        )
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

        # Precompila META
        self.var_name.set(str(self.data.get("name", "")))
        self.var_instance.set(str((root or {}).get("instanceOf", "")))

        props = (root or {}).get("properties", {})
        for path, obj in props.items():
            if isinstance(obj, dict):
                self.property_items.append((path, obj))

        # Domini
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
                for t in (o.get("sendPolicy", {}).get("triggers", []) or []):
                    v = t.get(key)
                    if isinstance(v, str) and v not in vals:
                        vals.append(v)
            return sorted(vals, key=str.lower)

        self.domain_types = collect_top("type") or ["boolean", "integer", "double", "string"]
        self.domain_units = collect_top("unit")
        self.domain_trig_types = collect_trig("type") or ["onchange", "periodic", "mixed"]
        self.domain_trig_modes = collect_trig("mode")
        self.domain_change_masks = collect_trig("changeMask") or [""]

        self._build_rows_all()
        self._refresh_filter_widgets()
        self.apply_filters()

    def _build_rows_all(self):
        self.rows_all.clear()
        self.row_to_path.clear()
        for path, obj in self.property_items:
            tr = (obj.get("sendPolicy", {}).get("triggers", []) or [{}])[0]
            db_val: Optional[float] = None
            db_type = ""
            if "deadbandPercent" in tr and tr.get("deadbandPercent") is not None:
                db_val = tr.get("deadbandPercent")
                db_type = "PERC"
            elif "deadband" in tr and tr.get("deadband") is not None:
                db_val = tr.get("deadband")
                db_type = "ABS"

            row = [
                obj.get("type", ""),
                obj.get("label", ""),
                obj.get("unit", ""),
                tr.get("type", ""),
                tr.get("level", ""),
                tr.get("mode", ""),
                tr.get("minIntervalMs", ""),
                tr.get("skipFirstNChanges", ""),
                tr.get("changeMask", ""),
                "" if db_val is None else db_val,
                db_type,
            ]
            self.rows_all.append(row)
            self.row_to_path.append(path)

    def _refresh_filter_widgets(self):
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
        for child in self.filter_bar.winfo_children():
            for sub in child.winfo_children():
                if isinstance(sub, ttk.Label) and sub.cget("text") == name:
                    info = sub.grid_info()
                    for peer in child.winfo_children():
                        if peer.grid_info().get("row") == info["row"] + 1 and peer.grid_info().get("column") == info["column"]:
                            for inner in peer.winfo_children():
                                if isinstance(inner, ttk.Combobox) or isinstance(inner, ttk.Entry):
                                    return inner
        return None

    # -------------- Filtering ---------------
    def apply_filters(self):
        def match_text(val: Any, query: str) -> bool:
            q = (query or "").strip()
            if q == "":
                return True
            s = str(val).lower()
            ql = q.lower()
            for op in (">=", "<=", ">", "<", "="):
                if ql.startswith(op):
                    try:
                        num = float(q[len(op):].strip())
                        v = float(val)
                    except Exception:
                        return False
                    if op == ">":
                        return v > num
                    if op == ">=":
                        return v >= num
                    if op == "<":
                        return v < num
                    if op == "<=":
                        return v <= num
                    if op == "=":
                        return v == num
            return ql in s

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

        self.sheet.set_sheet_data(self.rows_view, reset_col_positions=True, reset_row_positions=True)
        self._refresh_headers_with_arrow()
        if self._last_sort_col is not None:
            self._sort_view_by(self._last_sort_col, self._last_sort_asc)

    def _get_filter_value(self, name: str) -> str:
        w = self._find_filter_widget(name)
        if isinstance(w, ttk.Combobox) or isinstance(w, ttk.Entry):
            return w.get()
        return ""

    # -------------- Sorting (da pulsanti accanto ai filtri) ---------------
    def _on_sort_button(self, col: int):
        asc = self.sort_dir_by_col.get(col, True)
        self._last_sort_col = col
        self._last_sort_asc = asc
        self._sort_view_by(col, asc)
        self.sort_dir_by_col[col] = not asc
        self._update_sort_button_label(col)

    def _update_sort_button_label(self, col: int):
        btn = self.sort_buttons.get(col)
        if not btn:
            return
        next_is_asc = self.sort_dir_by_col.get(col, True)
        btn.configure(text=("A→Z" if next_is_asc else "Z→A"))

    def _sort_view_by(self, col: int, ascending: bool):
        pairs = list(zip(self.rows_view, self.view_index_map))

        def key_fn(item):
            row, _ = item
            v = row[col]
            if col in NUMERIC_COLS:
                try:
                    return (0, float(v))
                except Exception:
                    return (1, str(v).lower())
            return (0, str(v).lower())

        pairs.sort(key=key_fn, reverse=not ascending)
        self.rows_view = [r for r, _ in pairs]
        self.view_index_map = [i for _, i in pairs]
        self.sheet.set_sheet_data(self.rows_view, reset_col_positions=False, reset_row_positions=True)
        self._refresh_headers_with_arrow()

    # -------------- In-cell editing helpers ---------------
    def _on_begin_edit_cell(self, _event=None):
        sel = None
        try:
            sel = self.sheet.get_currently_selected()
        except Exception:
            return
        row, col = None, None
        if isinstance(sel, dict) and sel.get("type") == "cell":
            row, col = sel.get("row"), sel.get("column")
        elif isinstance(sel, tuple) and len(sel) >= 2:
            row, col = sel[0], sel[1]
        else:
            return
        try:
            self._prev_cell_value = self.sheet.get_cell_data(row, col)
        except Exception:
            self._prev_cell_value = None
        if col == IDX["deadband type"]:
            self.after(1, lambda r=row, c=col: self._open_deadband_combo(r, c))

    def _on_double_click_cell(self, _event=None):
        sel = None
        try:
            sel = self.sheet.get_currently_selected()
        except Exception:
            return
        if isinstance(sel, dict) and sel.get("type") == "cell":
            row, col = sel.get("row"), sel.get("column")
        elif isinstance(sel, tuple) and len(sel) >= 2:
            row, col = sel[0], sel[1]
        else:
            return
        if col == IDX["deadband type"]:
            self._open_deadband_combo(row, col)

    def _on_end_edit_cell(self, _event=None):
        sel = None
        try:
            sel = self.sheet.get_currently_selected()
        except Exception:
            return
        row, col = None, None
        if isinstance(sel, dict) and sel.get("type") == "cell":
            row, col = sel.get("row"), sel.get("column")
        elif isinstance(sel, tuple) and len(sel) >= 2:
            row, col = sel[0], sel[1]
        else:
            return
        if col == IDX["deadband type"]:
            try:
                val = str(self.sheet.get_cell_data(row, col)).upper().strip()
            except Exception:
                return
            if val not in ("ABS", "PERC"):
                prev = self._prev_cell_value
                if prev in ("ABS", "PERC"):
                    self.sheet.set_cell_data(row, col, prev)
                else:
                    self.sheet.set_cell_data(row, col, "")
            else:
                self.sheet.set_cell_data(row, col, val)

    def _open_deadband_combo(self, row: int, col: int):
        if self._overlay_combo is not None:
            try:
                self._overlay_combo.destroy()
            except Exception:
                pass
            self._overlay_combo = None
            self._overlay_cell = None
        try:
            x, y, w, h = self.sheet.get_cell_bbox(row, col, include_text=True)
        except Exception:
            return
        combo = ttk.Combobox(self.sheet, values=["ABS", "PERC"], state="readonly")
        try:
            current = str(self.sheet.get_cell_data(row, col)).upper().strip()
            if current in ("ABS", "PERC"):
                combo.set(current)
        except Exception:
            pass
        combo.place(x=x, y=y, width=max(60, w), height=h)
        combo.focus_set()
        combo.bind("<<ComboboxSelected>>", lambda e: self._commit_deadband_combo(row, col, combo))
        combo.bind("<Return>", lambda e: self._commit_deadband_combo(row, col, combo))
        combo.bind("<Escape>", lambda e: self._destroy_overlay_combo())
        combo.bind("<FocusOut>", lambda e: self._destroy_overlay_combo())
        self._overlay_combo = combo
        self._overlay_cell = (row, col)

    def _commit_deadband_combo(self, row: int, col: int, combo: ttk.Combobox):
        val = combo.get().strip().upper()
        if val not in ("ABS", "PERC"):
            return
        self.sheet.set_cell_data(row, col, val)
        self._destroy_overlay_combo()

    def _destroy_overlay_combo(self):
        if self._overlay_combo is not None:
            try:
                self._overlay_combo.destroy()
            except Exception:
                pass
        self._overlay_combo = None
        self._overlay_cell = None

    # -------------- Commit ---------------
    def _commit_table_to_json(self):
        if not self.data:
            return

        # sincr. vista -> all (solo righe visibili)
        current = [list(r) for r in self.sheet.get_sheet_data()]
        for idx_view, row_vals in enumerate(current):
            idx_all = self.view_index_map[idx_view]
            self.rows_all[idx_all] = row_vals

        root = self.data.get("json") or self.data

        # commit meta
        nm = self.var_name.get().strip()
        if nm != "":
            self.data["name"] = nm
        inst = self.var_instance.get().strip()
        if inst != "":
            root["instanceOf"] = inst

        props = root.get("properties", {})
        errors: List[str] = []

        for i, row in enumerate(self.rows_all):
            path = self.row_to_path[i]
            obj = props.get(path, {})

            obj_type = str(row[IDX["type"]]).strip()
            obj_label = str(row[IDX["label"]]).strip()
            obj_unit = str(row[IDX["unit"]]).strip()
            if obj_type: obj["type"] = obj_type
            if obj_label != "": obj["label"] = obj_label
            if obj_unit != "": obj["unit"] = obj_unit
            elif "unit" in obj:
                obj.pop("unit", None)

            tr_list = obj.setdefault("sendPolicy", {}).setdefault("triggers", [])
            if not tr_list:
                tr_list.append({})
            tr = tr_list[0]

            trig_type = str(row[IDX["trigger type"]]).strip()
            level = row[IDX["level"]]
            mode = str(row[IDX["mode"]]).strip()
            minint = row[IDX["min interval ms"]]
            skipn = row[IDX["skip first n changes"]]
            cmask = str(row[IDX["change mask"]]).strip()
            db = row[IDX["deadband"]]
            dbt = str(row[IDX["deadband type"]]).strip().upper()

            if trig_type: tr["type"] = trig_type
            if mode != "": tr["mode"] = mode

            # level (int)
            try:
                if str(level).strip() == "":
                    tr.pop("level", None)
                else:
                    tr["level"] = int(level)
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

            # change mask
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
                        if not (0.0 <= v <= 100.0): raise ValueError
                        tr["deadbandPercent"] = v
                        tr.pop("deadband", None)
                else:
                    errors.append(f"{path}: 'deadband type' deve essere ABS o PERC")
            except Exception:
                errors.append(f"{path}: 'deadband' non valido")

            props[path] = obj

        if errors:
            raise ValueError("\n".join(errors))

# ---------------- main ----------------
def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1300x720")
    root.minsize(1024, 600)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = MappingEditor(root)

    # auto-load se il file è a fianco dello script
    try:
        base = os.path.dirname(__file__)
    except NameError:
        base = os.getcwd()
    default_path = os.path.join(base, "2500053_Mapping.json")

    if os.path.exists(default_path):
        try:
            with open(default_path, "r", encoding="utf-8") as f:
                app.data = json.load(f)
            app.file_path = default_path
            app.btn_save.config(state=tk.NORMAL)
            app.btn_save_as.config(state=tk.NORMAL)
            app.status.set(f"Caricato: {os.path.basename(default_path)}")
            app._reindex()
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"""Apertura iniziale fallita:
{e}""")

    root.mainloop()


if __name__ == "__main__":
    main()
