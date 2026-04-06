# 📋 Suivi Candidatures Alternance

Une application desktop locale pour gérer et suivre ses candidatures d'alternance, avec sauvegarde automatique dans un fichier Excel.

---

## ✨ Fonctionnalités

- 📊 Tableau de suivi avec code couleur par statut
- ➕ Ajouter une candidature via un formulaire
- ✏️ Modifier une candidature existante
- 🗑️ Supprimer une candidature
- 💾 Sauvegarde automatique dans un fichier Excel
- 📈 Statistiques en temps réel (total, en attente, réponses, refus)

---

## 🚀 Lancement

### Windows
Double-cliquez simplement sur **`Suivi Alternance.exe`**

> ⚠️ Le fichier Excel `Suivi Candidatures Alternance CNAM Reims.xlsx` sera créé automatiquement dans le même dossier au premier lancement.

### Mac / Linux
**1. Installer les dépendances :**
```bash
pip install customtkinter openpyxl
```

**2. Lancer l'application :**
```bash
python main.py
```

---

## 🗂️ Structure du projet

```
📁 Suivi d'alternance/
├── main.py                                          # Code source
├── Suivi Alternance.exe                             # Exécutable Windows
├── Suivi Candidatures Alternance CNAM Reims.xlsx    # Données (créé automatiquement)
└── README.md
```

---

## 📋 Colonnes du tableau

| Colonne | Description |
|--------|-------------|
| Entreprise | Nom de l'entreprise |
| Contact | Nom de la personne contactée |
| Poste | Intitulé du poste |
| Canal | LinkedIn / Email / Site web |
| Date envoi | Date de la candidature |
| Statut | En attente / ✅ Réponse / ❌ / Relancé / Entretien |
| Notes | Informations complémentaires |
| Ville | Ville de l'entreprise |
| Priorité | ⭐ / ⭐⭐ / ⭐⭐⭐ |

---

## 🛠️ Technologies

- Python 3
- CustomTkinter — interface graphique moderne
- openpyxl — lecture/écriture Excel
- PyInstaller — génération du .exe

---

## 👨‍💻 Auteur

**Mounir ELKATMOUR**  
[LinkedIn](https://www.linkedin.com/in/mounir-elkatmour-703aa8269/)
