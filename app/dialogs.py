import customtkinter as ctk
from tkinter import messagebox
from datetime import date

from .core import *

# ── Formulaire ────────────────────────────────────────────────────────────────
class Formulaire(ctk.CTkToplevel):
    def __init__(self, parent, callback, donnees=None):
        super().__init__(parent)
        self.title("✏️ Modifier" if donnees else "➕ Nouvelle candidature")
        self.geometry("560x700")
        self.resizable(False, False)
        self.grab_set()
        self.callback = callback
        self.champs = {}
        self.configure(fg_color=UI["bg"])

        card = ctk.CTkFrame(self, fg_color=UI["surface"], corner_radius=14, border_width=1, border_color=UI["line"])
        card.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            card,
            text="✏️ Modifier" if donnees else "➕ Nouvelle candidature",
            font=H2_FONT,
            text_color=UI["text"],
        ).grid(row=0, column=0, columnspan=2, pady=(18, 12), padx=20, sticky="w")

        for i, label in enumerate(COLONNES, 1):
            ctk.CTkLabel(
                card,
                text=label,
                anchor="w",
                font=BODY_FONT,
                text_color=UI["muted"],
            ).grid(row=i, column=0, padx=20, pady=6, sticky="w")

            val = donnees[i - 1] if donnees and len(donnees) >= i else ""

            if label == "Canal":
                widget = ctk.CTkComboBox(
                    card,
                    values=CANAUX,
                    width=310,
                    height=34,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                    button_color=UI["primary"],
                    button_hover_color=UI["primary_hover"],
                )
                widget.set(val if val else CANAUX[0])
            elif label == "Statut":
                widget = ctk.CTkComboBox(
                    card,
                    values=STATUTS,
                    width=310,
                    height=34,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                    button_color=UI["primary"],
                    button_hover_color=UI["primary_hover"],
                )
                widget.set(val if val else STATUTS[0])
            elif label == "Priorité":
                widget = ctk.CTkComboBox(
                    card,
                    values=PRIORITES,
                    width=310,
                    height=34,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                    button_color=UI["primary"],
                    button_hover_color=UI["primary_hover"],
                )
                widget.set(val if val else PRIORITES[1])
            elif label == "Notes":
                widget = ctk.CTkTextbox(
                    card,
                    width=310,
                    height=90,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                    border_width=1,
                    font=BODY_FONT,
                )
                if val:
                    widget.insert("0.0", val)
            elif label == "Date envoi":
                widget = ctk.CTkEntry(
                    card,
                    width=310,
                    height=34,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                    placeholder_text=f"Ex: {date.today().strftime('%d/%m/%Y')}",
                )
                if val:
                    widget.insert(0, val)
            else:
                widget = ctk.CTkEntry(
                    card,
                    width=310,
                    height=34,
                    fg_color=UI["surface_alt"],
                    border_color=UI["line"],
                )
                if val:
                    widget.insert(0, val)

            widget.grid(row=i, column=1, padx=20, pady=6, sticky="w")
            self.champs[label] = widget

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=len(COLONNES) + 1, column=0, columnspan=2, pady=20)
        ctk.CTkButton(btn_frame, text="💾 Enregistrer", command=self.enregistrer,
                      width=170, height=38, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                      font=BODY_FONT).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Annuler", command=self.destroy,
                      width=170, height=38, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"], font=BODY_FONT).pack(side="left", padx=10)

    def enregistrer(self):
        row = []
        for label in COLONNES:
            widget = self.champs[label]
            if isinstance(widget, ctk.CTkTextbox):
                val = widget.get("0.0", "end").strip()
            elif isinstance(widget, ctk.CTkComboBox):
                val = widget.get()
            else:
                val = widget.get().strip()
            row.append(val)

        if not row[0].strip():
            messagebox.showwarning("Attention", "Le champ 'Entreprise' est obligatoire.")
            return

        if row[4].strip() and parse_date_fr(row[4].strip()) is None:
            messagebox.showwarning("Attention", "Date invalide. Format attendu: JJ/MM/AAAA")
            return

        self.callback(row)
        self.destroy()


class DialogNouveauSuivi(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.title("➕ Nouveau suivi")
        self.geometry("460x240")
        self.resizable(False, False)
        self.grab_set()
        self.callback = callback
        self.configure(fg_color=UI["bg"])

        body = ctk.CTkFrame(self, fg_color=UI["surface"], corner_radius=14, border_width=1, border_color=UI["line"])
        body.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(body, text="Nom du suivi",
                     font=H2_FONT, text_color=UI["text"]).pack(pady=(24, 8))
        self.entry = ctk.CTkEntry(body, width=360, height=38,
                                   fg_color=UI["surface_alt"], border_color=UI["line"],
                                   placeholder_text="Ex: Alternance Polytech Paris 2026")
        self.entry.pack(pady=4)
        self.entry.bind("<Return>", lambda e: self.confirmer())

        btn_frame = ctk.CTkFrame(body, fg_color="transparent")
        btn_frame.pack(pady=16)
        ctk.CTkButton(btn_frame, text="✅ Créer", command=self.confirmer,
                      width=150, height=38, fg_color=UI["primary"], hover_color=UI["primary_hover"],
                      font=BODY_FONT).pack(side="left", padx=8)
        ctk.CTkButton(btn_frame, text="Annuler", command=self.destroy,
                      width=150, height=38, fg_color=UI["surface_alt"], hover_color=UI["card_hover"],
                      border_width=1, border_color=UI["line"], font=BODY_FONT).pack(side="left", padx=8)

    def confirmer(self):
        nom = self.entry.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Le nom ne peut pas être vide.")
            return
        self.callback(nom)
        self.destroy()
