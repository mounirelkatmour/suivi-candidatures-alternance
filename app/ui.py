import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
from datetime import date, datetime, timedelta
import os
import csv

from .core import *
from .dialogs import Formulaire, DialogNouveauSuivi

# ── Vue Suivi ────────────────────────────────────────────────────────────────
class VueSuivi(ctk.CTkFrame):
    def __init__(self, parent, suivi_nom, excel_file, on_back):
        super().__init__(parent, corner_radius=0, fg_color=UI["bg"])
        self.suivi_nom = suivi_nom
        self.excel_file = excel_file
        self.on_back = on_back
        self.rows = []
        self.rows_filtered = []  # list[(real_idx, row)]
        self.search_cache = []
        self.row_checked = {}  # dict: {real_idx: bool}
        self.sort_col = None
        self.sort_asc = True
        self.batch_edit_shown = False
        self.select_all_var = ctk.BooleanVar(value=False)
        self.pending_click = None  # For double-click detection
        self.hovered_item = None
        self.shortcut_sequences = ["<Control-f>", "<Control-n>", "<Delete>"]
        settings = load_settings()
        self.weekly_goal = int(settings.get("weekly_goal", 15) or 15)
        self.build_ui()
        self.charger()

    def destroy(self):
        if self.pending_click:
            self.after_cancel(self.pending_click)
        self.unbind_shortcuts()
        super().destroy()

    def build_ui(self):
        # ── Header ──
        header = ctk.CTkFrame(self, height=72, corner_radius=0, fg_color=UI["surface"])
        header.pack(fill="x")
        ctk.CTkButton(header, text="← Retour", command=self.on_back,
                      width=110, height=38, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"], font=BODY_FONT).pack(side="left", padx=14, pady=14)
        ctk.CTkLabel(header, text=f"📋  {self.suivi_nom}",
                     font=H2_FONT, text_color=UI["text"]).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(header, text="➕ Ajouter", command=self.ajouter,
                      width=130, height=38, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                      font=BODY_FONT).pack(side="right", padx=14, pady=14)
        ctk.CTkButton(header, text="🔄 Actualiser", command=self.charger,
                      width=130, height=38, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"], font=BODY_FONT).pack(side="right", padx=6, pady=14)

        # ── Stats ──
        self.stats_bar = ctk.CTkFrame(self, height=74, corner_radius=0, fg_color=UI["bg"])
        self.stats_bar.pack(fill="x", padx=14, pady=(10, 0))
        self.kpi_total = ctk.CTkLabel(self.stats_bar, text="Total\n0", justify="left", font=BODY_FONT, text_color=UI["text"])
        self.kpi_pending = ctk.CTkLabel(self.stats_bar, text="En attente\n0", justify="left", font=BODY_FONT, text_color="#ffd166")
        self.kpi_reply = ctk.CTkLabel(self.stats_bar, text="Reponses\n0", justify="left", font=BODY_FONT, text_color="#80ed99")
        self.kpi_meet = ctk.CTkLabel(self.stats_bar, text="Entretiens\n0", justify="left", font=BODY_FONT, text_color="#8ecae6")
        self.kpi_rate = ctk.CTkLabel(self.stats_bar, text="Taux reponse\n0%", justify="left", font=BODY_FONT, text_color="#b8f2e6")
        self.kpi_total.pack(side="left", padx=(8, 18), pady=8)
        self.kpi_pending.pack(side="left", padx=10, pady=8)
        self.kpi_reply.pack(side="left", padx=10, pady=8)
        self.kpi_meet.pack(side="left", padx=10, pady=8)
        self.kpi_rate.pack(side="left", padx=10, pady=8)

        goal_block = ctk.CTkFrame(self.stats_bar, fg_color="transparent")
        goal_block.pack(side="right", padx=(8, 0), pady=4)
        self.goal_progress = ctk.CTkProgressBar(goal_block, width=170, height=8, progress_color=UI["primary"])
        self.goal_progress.set(0)
        self.goal_progress.pack(side="top", anchor="e", pady=(2, 4))
        goal_row = ctk.CTkFrame(goal_block, fg_color="transparent")
        goal_row.pack(side="top", anchor="e")
        self.goal_label = ctk.CTkLabel(goal_row, text="Objectif hebdo", font=SMALL_FONT, text_color=UI["muted"])
        self.goal_label.pack(side="left", padx=(0, 6))
        self.goal_var = ctk.StringVar(value=str(self.weekly_goal))
        self.goal_entry = ctk.CTkEntry(goal_row, width=48, height=28, textvariable=self.goal_var,
                           fg_color=UI["surface_alt"], border_color=UI["line"])
        self.goal_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(goal_row, text="OK", width=36, height=28, font=SMALL_FONT,
                  fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                  border_width=1, border_color=UI["line"],
                  command=self.update_weekly_goal).pack(side="left")

        self.lbl_stats = ctk.CTkLabel(self.stats_bar, text="", font=SMALL_FONT, text_color=UI["muted"])
        self.lbl_stats.pack(side="right", padx=12, pady=8)

        # ── Filtres & Recherche ──
        filter_frame = ctk.CTkFrame(self, fg_color=UI["surface"], corner_radius=14, border_width=1, border_color=UI["line"])
        filter_frame.pack(fill="x", padx=14, pady=(10, 10))

        # Ligne unique : recherche + filtres
        row1 = ctk.CTkFrame(filter_frame, fg_color="transparent")
        row1.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(row1, text="🔍", font=ctk.CTkFont(size=14), text_color=UI["muted"]).pack(side="left", padx=(0, 4))
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.appliquer_filtres())
        self.search_entry = ctk.CTkEntry(row1, textvariable=self.search_var,
                         placeholder_text="Rechercher dans entreprise, contact, date, notes, ville...",
                         width=420, height=36, fg_color=UI["surface_alt"], border_color=UI["line"])
        self.search_entry.pack(side="left", padx=4)
        ctk.CTkButton(row1, text="✕ Effacer", command=self.effacer_filtres,
                  width=100, height=36, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                  border_width=1, border_color=UI["line"], font=BODY_FONT).pack(side="left", padx=8)

        self.filtres = {}
        filtres_def = [
            ("Canal", ["Tous"] + CANAUX),
            ("Statut", ["Tous"] + STATUTS),
            ("Priorité", ["Tous"] + PRIORITES),
        ]
        for label, vals in filtres_def:
            ctk.CTkLabel(row1, text=label + " :", font=SMALL_FONT, text_color=UI["muted"]).pack(side="left", padx=(10, 2))
            cb = ctk.CTkComboBox(row1, values=vals, width=122, height=32,
                                  command=lambda v, l=label: self.appliquer_filtres(),
                                  fg_color=UI["surface_alt"], border_color=UI["line"],
                                  button_color=UI["primary"], button_hover_color=UI["primary_hover"])
            cb.set("Tous")
            cb.pack(side="left", padx=(0, 8))
            self.filtres[label] = cb

        # ── Tableau ──
        table_frame = ctk.CTkFrame(self, corner_radius=14, fg_color=UI["surface"], border_width=1, border_color=UI["line"])
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Modern.Treeview",
                    background=UI["surface"], foreground=UI["text"],
                    fieldbackground=UI["surface"], rowheight=34,
                    font=("Segoe UI", 10), borderwidth=0)
        style.configure("Modern.Treeview.Heading",
                    background=UI["surface_alt"], foreground=UI["text"],
                    font=("Segoe UI Semibold", 10, "bold"), relief="flat")
        style.map("Modern.Treeview",
            background=[("selected", UI["surface"])],
                foreground=[("selected", UI["text"])])
        style.map("Modern.Treeview.Heading",
                foreground=[("active", UI["surface_alt"])])

        cols = ["☐"] + COLONNES
        self.tree = ttk.Treeview(table_frame, columns=cols,
                                  show="headings", selectmode="none",
                                  style="Modern.Treeview")

        self.tree.heading("☐", text="☐")
        self.tree.column("☐", width=40, anchor="center", minwidth=40)
        
        for col, w in zip(COLONNES, [140, 140, 150, 85, 105, 110, 210, 85, 72]):
            self.tree.heading(col, text=col,
                               command=lambda c=col: self.trier(c))
            self.tree.column(col, width=w, anchor="w", minwidth=60)

        sy = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        for tag, (name, color) in STATUS_TAGS_UI.items():
            self.tree.tag_configure(name, background=color)
            self.tree.tag_configure(f"hover_{name}", background=darken_hex(color, 0.15))
        self.tree.tag_configure("hover_default", background=darken_hex(UI["surface"], 0.12))

        # Bind double-click to edit
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        # Bind checkbox clicks
        self.tree.bind("<Button-1>", self.on_tree_click)
        # Bind row hover effects
        self.tree.bind("<Motion>", self.on_tree_motion)
        self.tree.bind("<Leave>", self.on_tree_leave)

        # ── Boutons bas ──
        btn_bar = ctk.CTkFrame(self, height=58, corner_radius=0, fg_color=UI["bg"])
        btn_bar.pack(fill="x", padx=14, pady=(0, 12))
        ctk.CTkButton(btn_bar, text="✏️ Modifier", command=self.modifier,
                  width=160, height=40, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                  font=BODY_FONT).pack(side="left", padx=10, pady=8)
        ctk.CTkButton(btn_bar, text="🗑️ Supprimer", command=self.supprimer,
                  width=160, height=40, fg_color=UI["danger"],
                  hover_color=UI["danger_hover"], font=BODY_FONT).pack(side="left", padx=4, pady=8)

        # UX: raccourcis clavier
        self.bind_shortcuts()

    def bind_shortcuts(self):
        root = self.winfo_toplevel()
        root.bind("<Control-f>", self.shortcut_focus_search)
        root.bind("<Control-n>", self.shortcut_add)
        root.bind("<Delete>", self.shortcut_delete)

    def unbind_shortcuts(self):
        root = self.winfo_toplevel()
        for sequence in self.shortcut_sequences:
            root.unbind(sequence)

    def charger(self):
        self.rows = lire_donnees(self.excel_file)
        self.row_checked = {}  # Reset all checkboxes
        self.search_cache = []
        for row in self.rows:
            padded = (row + [""] * 9)[:9]
            self.search_cache.append(" ".join(padded[c] for c in [0, 1, 4, 6, 7] if padded[c]).lower())
        self.appliquer_filtres()

    def appliquer_filtres(self):
        query = self.search_var.get().lower()
        canal_f = self.filtres["Canal"].get()
        statut_f = self.filtres["Statut"].get()
        priorite_f = self.filtres["Priorité"].get()

        filtered = []
        for real_idx, row in enumerate(self.rows):
            padded = (row + [""] * 9)[:9]

            if query and query not in self.search_cache[real_idx]:
                    continue
            if canal_f != "Tous" and padded[3] != canal_f:
                continue
            if statut_f != "Tous" and padded[5] != statut_f:
                continue
            if priorite_f != "Tous" and padded[8] != priorite_f:
                continue

            filtered.append((real_idx, padded))

        if self.sort_col is not None:
            col_idx = COLONNES.index(self.sort_col)
            filtered.sort(key=lambda r: r[1][col_idx] or "", reverse=not self.sort_asc)

        self.rows_filtered = filtered
        self.refresh_tree()

    def refresh_tree(self):
        self.hovered_item = None
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, (_real_idx, row) in enumerate(self.rows_filtered):
            real_idx = _real_idx
            is_checked = self.row_checked.get(real_idx, False)
            checkbox_text = "☑" if is_checked else "☐"
            status = row[5] or ""
            tag_name = STATUS_TAGS_UI.get(status, ("", ""))[0]
            
            values = [checkbox_text] + list(row)
            self.tree.insert("", "end", iid=str(i), values=values,
                              tags=(tag_name,) if tag_name else ())

        # Update header checkbox based on all rows being checked
        checked_count = sum(1 for real_idx, _row in self.rows_filtered if self.row_checked.get(real_idx, False))
        total_filtered = len(self.rows_filtered)
        header_checkbox = "☑" if (checked_count == total_filtered and total_filtered > 0) else "☐"
        self.tree.heading("☐", text=header_checkbox)

        total = len(self.rows)
        shown = len(self.rows_filtered)
        en_attente = sum(1 for r in self.rows if len(r) > 5 and r[5] == "En attente")
        reponses = sum(1 for r in self.rows if len(r) > 5 and r[5] == "✅ Réponse")
        refus = sum(1 for r in self.rows if len(r) > 5 and r[5] == "❌")
        entretiens = sum(1 for r in self.rows if len(r) > 5 and r[5] == "Entretien")
        sent_last_7 = self.count_sent_last_days(7)
        response_rate = (reponses / total * 100) if total else 0
        interview_rate = (entretiens / total * 100) if total else 0
        progress = min(1, sent_last_7 / max(self.weekly_goal, 1))
        self.kpi_total.configure(text=f"Total\n{total}")
        self.kpi_pending.configure(text=f"En attente\n{en_attente}")
        self.kpi_reply.configure(text=f"Reponses\n{reponses}")
        self.kpi_meet.configure(text=f"Entretiens\n{entretiens}")
        self.kpi_rate.configure(text=f"Taux reponse\n{response_rate:.0f}%")
        self.goal_progress.set(progress)
        self.lbl_stats.configure(
            text=f"Affiches : {shown} | Refus : {refus} | Taux entretien : {interview_rate:.0f}%"
        )

    def trier(self, col):
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = True
        self.appliquer_filtres()

        for c in COLONNES:
            arrow = ""
            if c == self.sort_col:
                arrow = " ▲" if self.sort_asc else " ▼"
            self.tree.heading(c, text=c + arrow)

    def effacer_filtres(self):
        self.search_var.set("")
        for cb in self.filtres.values():
            cb.set("Tous")
        self.sort_col = None
        self.sort_asc = True
        for c in COLONNES:
            self.tree.heading(c, text=c)
        self.appliquer_filtres()

    def update_weekly_goal(self):
        try:
            value = int(self.goal_var.get().strip())
            if value <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Attention", "Objectif invalide. Entrez un nombre entier positif.")
            self.goal_var.set(str(self.weekly_goal))
            return
        self.weekly_goal = value
        settings = load_settings()
        settings["weekly_goal"] = value
        save_settings(settings)
        self.refresh_tree()

    def count_sent_last_days(self, days):
        today = date.today()
        start = today - timedelta(days=days - 1)
        count = 0
        for row in self.rows:
            sent_date = parse_date_fr(row[4] if len(row) > 4 else "")
            if sent_date and start <= sent_date <= today:
                count += 1
        return count

    def weekly_trend_text(self):
        pass
    def on_tree_double_click(self, event):
        """Handle double-click to edit only the clicked row - cancels pending single click"""
        # Cancel any pending checkbox toggle
        if self.pending_click:
            self.after_cancel(self.pending_click)
            self.pending_click = None
        
        item = self.tree.identify("item", event.x, event.y)
        col = self.tree.identify_column(event.x)
        
        # Don't edit if clicking on checkbox column or header
        if col == "#1" or not item or item == "":
            return
        
        # Only edit the double-clicked row, ignore any other selections
        try:
            display_idx = int(item)
            if display_idx < len(self.rows_filtered):
                real_idx = self.rows_filtered[display_idx][0]
                def callback(row):
                    self.rows[real_idx] = row
                    ecrire_donnees(self.excel_file, self.rows)
                    self.charger()
                Formulaire(self, callback, donnees=self.rows[real_idx])
        except (ValueError, IndexError):
            pass

    def _set_item_hover(self, item, enabled):
        tags = list(self.tree.item(item, "tags"))
        if enabled:
            if tags and tags[0].startswith("hover_"):
                return
            base_tag = tags[0] if tags else ""
            hover_tag = f"hover_{base_tag}" if base_tag else "hover_default"
            self.tree.item(item, tags=(hover_tag,))
            return

        if tags and tags[0].startswith("hover_"):
            base_tag = tags[0][6:]
            if base_tag == "default":
                self.tree.item(item, tags=())
            else:
                self.tree.item(item, tags=(base_tag,))

    def on_tree_motion(self, event):
        item = self.tree.identify_row(event.y)
        if item == self.hovered_item:
            return

        if self.hovered_item and self.hovered_item in self.tree.get_children(""):
            self._set_item_hover(self.hovered_item, False)

        self.hovered_item = item if item else None
        if self.hovered_item:
            self._set_item_hover(self.hovered_item, True)

    def on_tree_leave(self, _event):
        if self.hovered_item and self.hovered_item in self.tree.get_children(""):
            self._set_item_hover(self.hovered_item, False)
        self.hovered_item = None

    def on_tree_click(self, event):
        """Handle checkbox clicks - delayed to detect double-click"""
        item = self.tree.identify("item", event.x, event.y)
        col = self.tree.identify_column(event.x)
        
        if col != "#1":  # Only handle checkbox column
            return
        
        # Cancel any pending single click
        if self.pending_click:
            self.after_cancel(self.pending_click)
            self.pending_click = None
        
        # Schedule checkbox toggle with delay
        # If double-click happens, this will be cancelled
        self.pending_click = self.after(200, self._do_checkbox_toggle, item)
    
    def _do_checkbox_toggle(self, item):
        """Actually toggle the checkbox (called after delay if not double-clicked)"""
        self.pending_click = None
        
        if not item or item == "":
            # Select all header click
            checked_count = sum(1 for real_idx, row in self.rows_filtered if self.row_checked.get(real_idx, False))
            total_filtered = len(self.rows_filtered)
            new_state = not (checked_count == total_filtered and total_filtered > 0)
            for real_idx, _row in self.rows_filtered:
                self.row_checked[real_idx] = new_state
            self.select_all_var.set(new_state)
            self.refresh_tree()
        else:
            # Toggle checkbox for this specific row
            try:
                display_idx = int(item)
                if display_idx < len(self.rows_filtered):
                    real_idx = self.rows_filtered[display_idx][0]
                    self.row_checked[real_idx] = not self.row_checked.get(real_idx, False)
                    self.refresh_tree()
            except (ValueError, IndexError):
                pass

    def get_checked_indices(self):
        """Get indices of rows that are checked"""
        return [idx for idx, checked in self.row_checked.items() if checked]


    def edition_lot(self, real_indices=None):
        if real_indices is None:
            real_indices = self.get_checked_indices()
        if len(real_indices) < 2:
            messagebox.showwarning("Attention", "Sélectionnez au moins 2 candidatures.")
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title(f"✏️ Édition en lot ({len(real_indices)} sélectionnées)")
        dlg.geometry("500x320")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(fg_color=UI["bg"])

        main_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        header = ctk.CTkFrame(main_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(header, text=f"🧩 Modification en masse", font=H2_FONT, text_color=UI["text"]).pack(side="left")
        ctk.CTkLabel(header, text=f"{len(real_indices)} entrées", font=SMALL_FONT, text_color=UI["muted"]).pack(side="left", padx=(12, 0))

        fields_frame = ctk.CTkFrame(main_frame, fg_color=UI["surface"], corner_radius=12, border_width=1, border_color=UI["line"])
        fields_frame.pack(fill="both", expand=True, padx=0, pady=(0, 16))

        for label_text, values in [("Statut", ["Ne pas changer"] + STATUTS), 
                                     ("Priorité", ["Ne pas changer"] + PRIORITES),
                                     ("Canal", ["Ne pas changer"] + CANAUX)]:
            row = ctk.CTkFrame(fields_frame, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=10)
            ctk.CTkLabel(row, text=label_text, font=BODY_FONT, text_color=UI["text"], width=80).pack(side="left")
            cb = ctk.CTkComboBox(row, values=values, width=300,
                                fg_color=UI["surface_alt"], border_color=UI["line"],
                                button_color=UI["primary"], button_hover_color=UI["primary_hover"])
            cb.set("Ne pas changer")
            cb.pack(side="left")
            setattr(dlg, f"cb_{label_text.lower()}", cb)

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        def appliquer():
            statut_v = dlg.cb_statut.get()
            priorite_v = dlg.cb_priorité.get()
            canal_v = dlg.cb_canal.get()
            changes = 0
            for idx in real_indices:
                if statut_v != "Ne pas changer":
                    self.rows[idx][5] = statut_v
                    changes += 1
                if priorite_v != "Ne pas changer":
                    self.rows[idx][8] = priorite_v
                    changes += 1
                if canal_v != "Ne pas changer":
                    self.rows[idx][3] = canal_v
                    changes += 1
            if changes == 0:
                messagebox.showwarning("Attention", "Sélectionnez au moins un champ à modifier.")
                return
            ecrire_donnees(self.excel_file, self.rows)
            self.charger()
            dlg.destroy()
            messagebox.showinfo("Succès", f"Modifications appliquées à {len(real_indices)} candidatures.")

        ctk.CTkButton(btn_frame, text="✅ Appliquer", width=160, height=38, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                      font=BODY_FONT, command=appliquer).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Annuler", width=160, height=38, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"], font=BODY_FONT, command=dlg.destroy).pack(side="left", padx=6)

    def export_csv(self):
        if not self.rows_filtered:
            messagebox.showwarning("Attention", "Aucune ligne à exporter.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"export_{timestamp}.csv"
        path = os.path.join(DATA_DIR, filename)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLONNES)
            for _idx, row in self.rows_filtered:
                writer.writerow(row)
        messagebox.showinfo("Export", f"✅ {len(self.rows_filtered)} lignes exportées dans:\n{filename}")

    def export_weekly_report(self):
        total = len(self.rows)
        reponses = sum(1 for r in self.rows if len(r) > 5 and r[5] == "✅ Réponse")
        entretiens = sum(1 for r in self.rows if len(r) > 5 and r[5] == "Entretien")
        refus = sum(1 for r in self.rows if len(r) > 5 and r[5] == "❌")
        sent_last_7 = self.count_sent_last_days(7)
        response_rate = (reponses / total * 100) if total else 0
        interview_rate = (entretiens / total * 100) if total else 0
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"rapport_hebdo_{safe_filename(self.suivi_nom)}_{stamp}.txt"
        path = os.path.join(DATA_DIR, filename)
        content = (
            f"Rapport hebdomadaire - {self.suivi_nom}\n"
            f"Date : {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
            f"Total candidatures : {total}\n"
            f"Envoyees sur 7 jours : {sent_last_7}\n"
            f"Objectif hebdo : {self.weekly_goal}\n"
            f"Progression : {min(100, int(sent_last_7 * 100 / max(self.weekly_goal, 1)))}%\n"
            f"Reponses : {reponses} ({response_rate:.1f}%)\n"
            f"Entretiens : {entretiens} ({interview_rate:.1f}%)\n"
            f"Refus : {refus}\n"
        )
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Rapport généré", f"Fichier créé :\n{path}")

    def shortcut_focus_search(self, _event=None):
        self.search_entry.focus_set()
        return "break"

    def shortcut_add(self, _event=None):
        self.ajouter()
        return "break"

    def shortcut_delete(self, _event=None):
        if self.tree.focus_get() == self.tree or self.tree.selection():
            self.supprimer()
            return "break"
        return None

    def ajouter(self):
        def callback(row):
            self.rows.append(row)
            ecrire_donnees(self.excel_file, self.rows)
            self.charger()
        Formulaire(self, callback)

    def modifier(self):
        checked = self.get_checked_indices()
        if len(checked) == 0:
            messagebox.showwarning("Attention", "Sélectionnez au moins une candidature.")
            return
        elif len(checked) == 1:
            # Single edit
            real_idx = checked[0]
            def callback(row):
                self.rows[real_idx] = row
                ecrire_donnees(self.excel_file, self.rows)
                self.charger()
            Formulaire(self, callback, donnees=self.rows[real_idx])
        else:
            # Batch edit
            self.edition_lot(checked)

    def supprimer(self):
        checked = self.get_checked_indices()
        if len(checked) == 0:
            messagebox.showwarning("Attention", "Sélectionnez au moins une candidature.")
            return
        if len(checked) == 1:
            nom = self.rows[checked[0]][0] if self.rows[checked[0]] else "cette entrée"
            msg = f"Supprimer la candidature chez {nom} ?"
        else:
            msg = f"Supprimer {len(checked)} candidatures ?"
        if messagebox.askyesno("Confirmer", msg):
            # Delete in reverse order to maintain indices
            for idx in sorted(checked, reverse=True):
                self.rows.pop(idx)
            ecrire_donnees(self.excel_file, self.rows)
            self.charger()

# ── Menu Principal ────────────────────────────────────────────────────────────
class MenuPrincipal(ctk.CTkFrame):
    def __init__(self, parent, on_open):
        super().__init__(parent, corner_radius=0, fg_color=UI["bg"])
        self.on_open = on_open
        self.build_ui()
        self.rafraichir()

    def build_ui(self):
        # Header
        header = ctk.CTkFrame(self, height=86, corner_radius=0, fg_color=UI["surface"])
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📁  Mes Suivis d'Alternance",
                     font=TITLE_FONT, text_color=UI["text"]).pack(side="left", padx=24, pady=22)
        ctk.CTkButton(header, text="➕ Nouveau suivi", command=self.nouveau_suivi,
                      width=170, height=40, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                      font=BODY_FONT).pack(side="right", padx=16, pady=20)

        # Sous-titre
        ctk.CTkLabel(self, text="Sélectionnez un suivi pour commencer",
                     font=BODY_FONT, text_color=UI["muted"]).pack(pady=(22, 10))

        # Liste
        self.liste_frame = ctk.CTkScrollableFrame(self, corner_radius=8,
                                                   fg_color=UI["surface"], border_width=1, border_color=UI["line"])
        self.liste_frame.pack(fill="both", expand=True, padx=40, pady=(0, 40))

    def rafraichir(self):
        for w in self.liste_frame.winfo_children():
            w.destroy()

        suivis = load_suivis()
        if not suivis:
            ctk.CTkLabel(self.liste_frame,
                         text="Aucun suivi pour l'instant.\nCliquez sur '➕ Nouveau suivi' pour commencer.",
                         font=BODY_FONT, text_color=UI["muted"]).pack(pady=40)
            return

        for suivi in suivis:
            self.creer_carte(suivi)

    def creer_carte(self, suivi):
        card = ctk.CTkFrame(self.liste_frame, corner_radius=10,
                             fg_color=UI["card"], height=76, border_width=1, border_color=UI["line"])
        card.pack(fill="x", pady=6, padx=8)
        card.pack_propagate(False)

        path = excel_path(suivi["nom"])
        nb = len(lire_donnees(path)) if os.path.exists(path) else 0

        ctk.CTkLabel(card, text=f"📋  {suivi['nom']}",
                     font=H2_FONT,
                     text_color=UI["text"],
                     anchor="w").pack(side="left", padx=16, pady=8)
        ctk.CTkLabel(card, text=f"{nb} candidature{'s' if nb > 1 else ''}",
                     font=SMALL_FONT, text_color=UI["muted"]).pack(side="left", padx=4)

        buttons_frame = ctk.CTkFrame(card, fg_color="transparent")
        buttons_frame.pack(side="right", padx=8, pady=8)

        ctk.CTkButton(buttons_frame, text="Ouvrir →", width=100, height=36,
                  fg_color=UI["success"], hover_color="#00bf84",
                  font=BODY_FONT, anchor="center",
                  command=lambda s=suivi: self.on_open(s)).pack(side="left", padx=4)

        def show_menu():
            menu_dlg = ctk.CTkToplevel(self)
            menu_dlg.title(f"Options - {suivi['nom']}")
            menu_dlg.geometry("320x240")
            menu_dlg.resizable(False, False)
            menu_dlg.grab_set()
            menu_dlg.configure(fg_color=UI["bg"])

            def open_graphs_safely():
                # Hide immediately for UX, destroy later to avoid Tk focus callbacks on a dead window.
                try:
                    menu_dlg.withdraw()
                except Exception:
                    pass
                self.show_graphs_dialog(suivi, menu_dlg)
                try:
                    menu_dlg.after(220, menu_dlg.destroy)
                except Exception:
                    pass

            frame = ctk.CTkFrame(menu_dlg, fg_color=UI["surface"], corner_radius=12, border_width=1, border_color=UI["line"])
            frame.pack(fill="both", expand=True, padx=16, pady=16)

            ctk.CTkLabel(frame, text="Options", font=H2_FONT, text_color=UI["text"]).pack(pady=(12, 16))

            buttons = [
                ("📊 Voir statistiques", lambda: self.show_stats_dialog(suivi, menu_dlg)),
                ("📈 Graphiques", open_graphs_safely),
                ("📝 Rapport hebdo", lambda: self.export_weekly_report_for_suivi(suivi, menu_dlg)),
            ]
            for text, cmd in buttons:
                ctk.CTkButton(frame, text=text, width=240, height=36,
                            fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                            border_width=1, border_color=UI["line"], font=BODY_FONT,
                            command=cmd).pack(pady=4)

        ctk.CTkButton(buttons_frame, text="➕ Options", width=100, height=36,
                  fg_color=UI["primary"], hover_color=UI["primary_hover"],
                  font=BODY_FONT, anchor="center",
                  command=show_menu).pack(side="left", padx=4)
        ctk.CTkButton(buttons_frame, text="🗑️ Supprimer", width=100, height=36,
                  fg_color=UI["danger"], hover_color=UI["danger_hover"],
                  font=BODY_FONT, anchor="center",
                  command=lambda s=suivi: self.supprimer_suivi(s)).pack(side="left", padx=4)

    def show_stats_dialog(self, suivi, parent_dlg):
        path = excel_path(suivi["nom"])
        rows = lire_donnees(path)
        if not rows:
            messagebox.showwarning("Attention", "Aucune donnée.")
            return

        total = len(rows)
        reponses = sum(1 for r in rows if len(r) > 5 and r[5] == "✅ Réponse")
        refus = sum(1 for r in rows if len(r) > 5 and r[5] == "❌")
        en_attente = sum(1 for r in rows if len(r) > 5 and r[5] == "En attente")
        relances = sum(1 for r in rows if len(r) > 5 and r[5] == "Relancé")
        entretiens = sum(1 for r in rows if len(r) > 5 and r[5] == "Entretien")

        stats_dlg = ctk.CTkToplevel(parent_dlg)
        stats_dlg.title(f"Statistiques - {suivi['nom']}")
        stats_dlg.geometry("420x380")
        stats_dlg.resizable(False, False)
        stats_dlg.configure(fg_color=UI["bg"])

        frame = ctk.CTkFrame(stats_dlg, fg_color=UI["surface"], corner_radius=12, border_width=1, border_color=UI["line"])
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(frame, text="📊 Statistiques détaillées", font=H2_FONT, text_color=UI["text"]).pack(pady=(12, 16))

        stats_text = (
            f"Total candidatures : {total}\n"
            f"En attente : {en_attente} ({en_attente*100//total if total else 0}%)\n"
            f"Réponses positives : {reponses} ({reponses*100//total if total else 0}%)\n"
            f"Entretiens : {entretiens} ({entretiens*100//total if total else 0}%)\n"
            f"Relances : {relances}\n"
            f"Refus : {refus} ({refus*100//total if total else 0}%)"
        )
        ctk.CTkLabel(frame, text=stats_text, font=BODY_FONT, text_color=UI["text"], justify="left").pack(padx=16, pady=12, anchor="w")

    def show_graphs_dialog(self, suivi, parent_dlg):
        path = excel_path(suivi["nom"])
        rows = lire_donnees(path)
        if not rows:
            messagebox.showwarning("Attention", "Aucune donnée.")
            return
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib import patheffects as pe
            from matplotlib.patches import FancyBboxPatch
        except ImportError:
            messagebox.showerror("Erreur", "matplotlib n'est pas installé.\npip install matplotlib")
            return
        graph_dlg = ctk.CTkToplevel(self)
        graph_dlg.title(f"Graphiques - {suivi['nom']}")
        graph_dlg.geometry("1120x760")
        graph_dlg.minsize(980, 680)
        graph_dlg.configure(fg_color=UI["bg"])
        graph_dlg.withdraw()

        def maximize_graph_dialog():
            try:
                graph_dlg.state("zoomed")
            except Exception:
                sw = graph_dlg.winfo_screenwidth()
                sh = graph_dlg.winfo_screenheight()
                graph_dlg.geometry(f"{sw}x{sh}+0+0")

        shell = ctk.CTkFrame(graph_dlg, fg_color=UI["surface"], corner_radius=12,
                             border_width=1, border_color=UI["line"])
        shell.pack(fill="both", expand=True, padx=14, pady=14)

        header = ctk.CTkFrame(shell, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 2))
        ctk.CTkLabel(header, text=f"Tableau de bord - {suivi['nom']}",
                 font=H2_FONT, text_color=UI["text"]).pack(side="left")
        ctk.CTkLabel(header, text=f"Mise à jour: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                 font=SMALL_FONT, text_color=UI["muted"]).pack(side="right")
        subtitle_lbl = ctk.CTkLabel(shell,
                 text="Vue analytique structurée: performance, volumes et tendances",
                 font=SMALL_FONT, text_color=UI["muted"]).pack(anchor="w", padx=16, pady=(0, 8))
        subtitle_lbl = shell.winfo_children()[-1]

        controls = ctk.CTkFrame(shell, fg_color=UI["surface_alt"], corner_radius=10,
                    border_width=1, border_color=UI["line"], height=52)
        controls.pack(fill="x", padx=18, pady=(0, 10))
        controls.pack_propagate(False)
        ctk.CTkLabel(controls, text="Période", font=SMALL_FONT, text_color=UI["muted"]).pack(side="left", padx=(10, 8))

        period_map = {
            "7j": 7,
            "30j": 30,
            "90j": 90,
            "Tout": None,
        }
        period_var = ctk.StringVar(value="30j")

        kpi_row = ctk.CTkFrame(shell, fg_color="transparent")
        kpi_row.pack(fill="x", padx=14, pady=(0, 8))

        kpi_cards = []
        kpi_accents = [UI["primary"], UI["success"], "#8ecae6", "#ffb347"]
        for i in range(4):
            card = ctk.CTkFrame(kpi_row, fg_color=UI["card"], corner_radius=10,
                                border_width=1, border_color=UI["line"], height=86)
            card.pack(side="left", fill="x", expand=True, padx=5)
            card.pack_propagate(False)
            ctk.CTkFrame(card, fg_color=kpi_accents[i], height=4, corner_radius=8).pack(fill="x", padx=8, pady=(8, 2))
            title_lbl = ctk.CTkLabel(card, text="", font=SMALL_FONT, text_color=UI["muted"])
            title_lbl.pack(anchor="w", padx=12, pady=(2, 2))
            value_lbl = ctk.CTkLabel(card, text="", font=H2_FONT, text_color=UI["text"])
            value_lbl.pack(anchor="w", padx=12, pady=(0, 8))
            kpi_cards.append((title_lbl, value_lbl))

        plot_wrap = ctk.CTkFrame(shell, fg_color=darken_hex(UI["surface_alt"], 0.08), corner_radius=12,
                                 border_width=2, border_color=UI["line"])
        plot_wrap.pack(fill="both", expand=True, padx=14, pady=(4, 12))

        canvas_holder = ctk.CTkFrame(plot_wrap, fg_color=UI["surface_alt"], corner_radius=10,
                                     border_width=1, border_color=UI["line"])
        canvas_holder.pack(fill="both", expand=True, padx=10, pady=10)

        chart_state = {"canvas": None, "figure": None, "kpi_token": 0}
        after_jobs = {"kpi": None}

        def export_dashboard_png():
            if chart_state["figure"] is None:
                messagebox.showwarning("Attention", "Aucun graphique à exporter.")
                return
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suivi_name = safe_filename(suivi["nom"]).replace(" ", "_")
            filename = f"dashboard_{suivi_name}_{timestamp}.png"
            filepath = os.path.join(DATA_DIR, filename)
            chart_state["figure"].savefig(filepath, dpi=170, bbox_inches="tight", facecolor=UI["surface_alt"])
            messagebox.showinfo("Export PNG", f"Graphique exporté : {filename}")

        def set_kpis_immediate(kpi_specs):
            chart_state["kpi_token"] += 1
            if after_jobs["kpi"] is not None:
                try:
                    graph_dlg.after_cancel(after_jobs["kpi"])
                except Exception:
                    pass
                after_jobs["kpi"] = None

            for i, (title, target, kind) in enumerate(kpi_specs):
                kpi_cards[i][0].configure(text=title)
                if kind == "pct":
                    kpi_cards[i][1].configure(text=f"{target:.1f}%")
                else:
                    kpi_cards[i][1].configure(text=str(int(round(target))))

        def style_axis(ax):
            panel_bg = darken_hex(UI["surface_alt"], 0.10)
            ax.set_facecolor(panel_bg)
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_color(UI["line"])
                spine.set_linewidth(1.4)
            ax.tick_params(axis="both", colors=UI["text"], labelsize=9)
            ax.grid(axis="y", color=UI["line"], alpha=0.30, linestyle="--", linewidth=0.8)

        def add_axis_group_card(fig, ax, pad_x=0.012, pad_bottom=0.010, pad_top=0.048):
            x, y, w, h = ax.get_position().bounds
            card = FancyBboxPatch(
                (x - pad_x, y - pad_bottom),
                w + 2 * pad_x,
                h + pad_bottom + pad_top,
                boxstyle="round,pad=0.006,rounding_size=0.012",
                transform=fig.transFigure,
                facecolor="none",
                edgecolor=UI["line"],
                linewidth=1.8,
                zorder=0,
            )
            fig.add_artist(card)

        def draw_dashboard(period_key):
            days = period_map.get(period_key)
            today = date.today()

            if days is None:
                filtered_rows = rows[:]
                period_text = "Toutes les données"
            else:
                start_day = today - timedelta(days=days - 1)
                filtered_rows = []
                for r in rows:
                    sent_date = parse_date_fr(r[4] if len(r) > 4 else "")
                    if sent_date and sent_date >= start_day:
                        filtered_rows.append(r)
                period_text = f"{days} derniers jours"

            if not filtered_rows:
                filtered_rows = rows[:]
                period_text = "Toutes les données (aucune entrée sur la période)"

            total = len(filtered_rows)
            statuts_count = {
                "En attente": sum(1 for r in filtered_rows if len(r) > 5 and r[5] == "En attente"),
                "Réponse": sum(1 for r in filtered_rows if len(r) > 5 and r[5] == "✅ Réponse"),
                "Refus": sum(1 for r in filtered_rows if len(r) > 5 and r[5] == "❌"),
                "Relancé": sum(1 for r in filtered_rows if len(r) > 5 and r[5] == "Relancé"),
                "Entretien": sum(1 for r in filtered_rows if len(r) > 5 and r[5] == "Entretien"),
            }

            reponses = statuts_count["Réponse"]
            entretiens = statuts_count["Entretien"]
            refus = statuts_count["Refus"]
            relances = statuts_count["Relancé"]
            tx_reponse = (reponses / total * 100) if total else 0
            tx_entretien = (entretiens / total * 100) if total else 0

            kpi_specs = [
                ("Candidatures", float(total), "int"),
                ("Taux réponse", tx_reponse, "pct"),
                ("Taux entretien", tx_entretien, "pct"),
                ("Relances", float(relances), "int"),
            ]
            set_kpis_immediate(kpi_specs)

            canaux_count = {}
            priorites_count = {}
            per_day = {}
            for r in filtered_rows:
                canal = r[3] if len(r) > 3 and r[3] else "Inconnu"
                prio = r[8] if len(r) > 8 and r[8] else "-"
                canaux_count[canal] = canaux_count.get(canal, 0) + 1
                priorites_count[prio] = priorites_count.get(prio, 0) + 1

                sent_date = parse_date_fr(r[4] if len(r) > 4 else "")
                if sent_date:
                    per_day[sent_date] = per_day.get(sent_date, 0) + 1

            status_labels = ["En attente", "Réponse", "Relancé", "Entretien", "Refus"]
            status_values = [
                statuts_count["En attente"],
                statuts_count["Réponse"],
                statuts_count["Relancé"],
                statuts_count["Entretien"],
                statuts_count["Refus"],
            ]
            status_colors = ["#ffd166", "#80ed99", "#fcb69f", "#8ecae6", "#f06252"]

            sorted_days = sorted(per_day.items(), key=lambda kv: kv[0])
            x_days = [d.strftime("%d/%m") for d, _v in sorted_days]
            y_days = [_v for _d, _v in sorted_days]

            fig = plt.figure(figsize=(11, 6.6), facecolor=UI["surface_alt"])
            gs = fig.add_gridspec(2, 3, width_ratios=[1.05, 1.3, 1.1], height_ratios=[1.15, 1],
                                  wspace=0.42, hspace=0.45)

            ax_status = fig.add_subplot(gs[:, 0])
            ax_timeline = fig.add_subplot(gs[0, 1:])
            ax_canal = fig.add_subplot(gs[1, 1])
            ax_prio = fig.add_subplot(gs[1, 2])

            bars = ax_status.barh(status_labels, status_values, color=status_colors,
                                  edgecolor=darken_hex(UI["line"], 0.05), linewidth=1.1)
            ax_status.invert_yaxis()
            ax_status.set_title("Répartition des statuts", color=UI["text"], pad=10, fontsize=11,
                                bbox=dict(boxstyle="round,pad=0.28", fc=UI["card"], ec=UI["line"], lw=1.0))
            style_axis(ax_status)
            for bar in bars:
                width = bar.get_width()
                ax_status.text(width + 0.1, bar.get_y() + bar.get_height() / 2,
                               f"{int(width)}", va="center", color=UI["text"], fontsize=9)

            if x_days:
                line = ax_timeline.plot(x_days, y_days, color=UI["primary"], linewidth=2.4,
                                        marker="o", markersize=5,
                                        markerfacecolor=UI["surface"], markeredgewidth=1.3)[0]
                line.set_path_effects([pe.Stroke(linewidth=4.5, foreground=darken_hex(UI["primary"], 0.35), alpha=0.35),
                                       pe.Normal()])
                ax_timeline.fill_between(x_days, y_days, color=UI["primary"], alpha=0.2)
            ax_timeline.set_title(f"Évolution des envois - {period_text}", color=UI["text"], pad=10, fontsize=11,
                                  bbox=dict(boxstyle="round,pad=0.28", fc=UI["card"], ec=UI["line"], lw=1.0))
            style_axis(ax_timeline)
            ax_timeline.tick_params(axis="x", rotation=30)

            canal_labels = list(canaux_count.keys()) or ["Inconnu"]
            canal_values = list(canaux_count.values()) or [0]
            ax_canal.pie(canal_values, labels=canal_labels, autopct="%1.0f%%", startangle=90,
                         colors=[UI["primary"], "#8ecae6", "#ffd166", "#fcb69f"],
                         wedgeprops={"width": 0.52, "edgecolor": UI["surface_alt"], "linewidth": 1.2},
                         pctdistance=0.78)
            ax_canal.set_title("Canaux", color=UI["text"], pad=6, y=0.98, fontsize=11)
            ax_canal.set_facecolor(darken_hex(UI["surface_alt"], 0.10))
            for spine in ax_canal.spines.values():
                spine.set_visible(True)
                spine.set_color(UI["line"])
                spine.set_linewidth(2.0)
            for txt in ax_canal.texts:
                txt.set_color(UI["text"])
            add_axis_group_card(fig, ax_canal)

            prio_label_map = {
                "⭐": "★",
                "⭐⭐": "★★",
                "⭐⭐⭐": "★★★",
            }
            prio_labels = [prio_label_map.get(k, k) for k in list(priorites_count.keys())] or ["-"]
            prio_values = list(priorites_count.values()) or [0]
            prio_bars = ax_prio.bar(prio_labels, prio_values, color=["#ffd166", "#ffb347", "#ff8c42"],
                                    edgecolor=darken_hex("#ff8c42", 0.25), linewidth=1.1)
            ax_prio.set_title("Priorités", color=UI["text"], pad=8, fontsize=11,
                              bbox=dict(boxstyle="round,pad=0.28", fc=UI["card"], ec=UI["line"], lw=1.0))
            style_axis(ax_prio)
            for bar in prio_bars:
                height = bar.get_height()
                ax_prio.text(bar.get_x() + bar.get_width() / 2, height + 0.05, f"{int(height)}",
                             ha="center", va="bottom", color=UI["text"], fontsize=9)

            fig.text(0.015, 0.015,
                     f"Total: {total} | Réponses: {reponses} | Entretiens: {entretiens} | Refus: {refus}",
                     color=UI["muted"], fontsize=9)

            if chart_state["canvas"] is not None:
                chart_state["canvas"].get_tk_widget().destroy()
            if chart_state["figure"] is not None:
                plt.close(chart_state["figure"])

            canvas = FigureCanvasTkAgg(fig, master=canvas_holder)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            chart_state["canvas"] = canvas
            chart_state["figure"] = fig

        period_switch = ctk.CTkSegmentedButton(
            controls,
            values=["7j", "30j", "90j", "Tout"],
            variable=period_var,
            command=draw_dashboard,
            fg_color=UI["surface_alt"],
            selected_color=UI["primary"],
            selected_hover_color=UI["primary_hover"],
            unselected_color=UI["surface_alt"],
            unselected_hover_color=UI["card_hover"],
            text_color=UI["text"],
        )
        period_switch.pack(side="left", padx=(0, 10))

        ctk.CTkButton(controls, text="Actualiser", width=110, height=32,
                      fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"],
                      command=lambda: draw_dashboard(period_var.get())).pack(side="right", padx=(0, 10))
        ctk.CTkButton(controls, text="Exporter PNG", width=125, height=32,
                  fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                  border_width=1, border_color=UI["line"],
                  command=export_dashboard_png).pack(side="right", padx=(0, 8))

        def cleanup_graph_dialog(event=None):
            if event is not None and event.widget is not graph_dlg:
                return
            if after_jobs["kpi"] is not None:
                try:
                    graph_dlg.after_cancel(after_jobs["kpi"])
                except Exception:
                    pass
                after_jobs["kpi"] = None
            if chart_state["figure"] is not None:
                try:
                    plt.close(chart_state["figure"])
                except Exception:
                    pass
                chart_state["figure"] = None

        graph_dlg.bind("<Destroy>", cleanup_graph_dialog)

        draw_dashboard(period_var.get())
        graph_dlg.update_idletasks()
        graph_dlg.deiconify()
        graph_dlg.lift()
        graph_dlg.focus_force()
        graph_dlg.after(20, maximize_graph_dialog)

    def export_weekly_report_for_suivi(self, suivi, parent_dlg):
        path = excel_path(suivi["nom"])
        rows = lire_donnees(path)
        total = len(rows)
        reponses = sum(1 for r in rows if len(r) > 5 and r[5] == "✅ Réponse")
        entretiens = sum(1 for r in rows if len(r) > 5 and r[5] == "Entretien")
        refus = sum(1 for r in rows if len(r) > 5 and r[5] == "❌")
        response_rate = (reponses / total * 100) if total else 0
        interview_rate = (entretiens / total * 100) if total else 0
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"rapport_{timestamp}.txt"
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        default_dir = downloads_dir if os.path.isdir(downloads_dir) else DATA_DIR
        filepath = filedialog.asksaveasfilename(
            title="Enregistrer le rapport hebdomadaire",
            initialdir=default_dir,
            initialfile=filename,
            defaultextension=".txt",
            filetypes=[("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*")],
            parent=parent_dlg,
        )
        if not filepath:
            return
        content = (
            f"Rapport hebdomadaire - {suivi['nom']}\n"
            f"Date : {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
            f"Total candidatures : {total}\n"
            f"Réponses : {reponses} ({response_rate:.1f}%)\n"
            f"Entretiens : {entretiens} ({interview_rate:.1f}%)\n"
            f"Refus : {refus}\n"
        )
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Rapport généré", f"Fichier créé : {os.path.basename(filepath)}")
        parent_dlg.destroy()

    def nouveau_suivi(self):
        def callback(nom):
            suivis = load_suivis()
            if any(s["nom"] == nom for s in suivis):
                messagebox.showwarning("Attention", "Un suivi avec ce nom existe déjà.")
                return
            path = excel_path(nom)
            init_excel(path)
            suivis.append({"nom": nom, "fichier": os.path.basename(path)})
            save_suivis(suivis)
            self.rafraichir()
        DialogNouveauSuivi(self, callback)

    def supprimer_suivi(self, suivi):
        if not messagebox.askyesno("Confirmer",
                                    f"Vous êtes sur de vouloir supprimer le suivi « {suivi['nom']} » ?"):
            return
        path = excel_path(suivi["nom"])
        if os.path.exists(path):
            os.remove(path)
        suivis = [s for s in load_suivis() if s["nom"] != suivi["nom"]]
        save_suivis(suivis)
        self.rafraichir()

# ── Application principale ────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📋 Suivi Candidatures Alternance")
        self.geometry("1350x750")
        self.minsize(900, 600)
        self.configure(fg_color=UI["bg"])
        ensure_data_dir()
        self.current_frame = None
        self.footer = ctk.CTkLabel(
            self,
            text="Created by : Mounir ELKATMOUR",
            font=SMALL_FONT,
            text_color=UI["muted"],
        )
        self.footer.pack(side="bottom", fill="x", pady=(0, 8))
        self.show_menu()
        self.after(20, self.maximize_main_window)

    def maximize_main_window(self):
        try:
            self.state("zoomed")
        except Exception:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            self.geometry(f"{sw}x{sh}+0+0")

    def show_menu(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = MenuPrincipal(self, on_open=self.open_suivi)
        self.current_frame.pack(fill="both", expand=True)

    def open_suivi(self, suivi):
        if self.current_frame:
            self.current_frame.destroy()
        path = excel_path(suivi["nom"])
        if not os.path.exists(path):
            init_excel(path)
        self.current_frame = VueSuivi(
            self, suivi["nom"], path, on_back=self.show_menu
        )
        self.current_frame.pack(fill="both", expand=True)

if __name__ == "__main__":
    app = App()
    app.mainloop()
