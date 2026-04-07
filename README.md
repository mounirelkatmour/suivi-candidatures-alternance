# 📋 Suivi Candidatures Alternance

Application desktop locale pour gérer et suivre ses candidatures d'alternance, avec sauvegarde automatique dans un fichier Excel.

---

## ✨ Fonctionnalités

- Tableau de suivi avec code couleur par statut
- Ajout, modification et suppression de candidatures
- Sauvegarde automatique dans un fichier Excel
- Statistiques et graphiques intégrés
- Export de rapport hebdomadaire

---

## 🚀 Lancement rapide

### Windows
1. Installer Python 3 si besoin.
2. Ouvrir un terminal dans le dossier du projet.
3. Installer les dépendances :
```bash
pip install -r requirements.txt
```
4. Lancer l'application :
```bash
python main.py
```

Tu peux aussi double-cliquer sur `run.bat`.

### Linux / macOS
```bash
pip install -r requirements.txt
python main.py
```

---

## 🗂️ Structure du projet

```text
📁 Suivi d'alternance/
├── main.py                 # Lanceur principal
├── run.bat                 # Lancement rapide Windows
├── requirements.txt        # Dépendances Python
├── app/                    # Code modulaire de l'application
│   ├── __init__.py
│   ├── core.py
│   ├── dialogs.py
│   ├── ui.py
│   ├── charts.py
│   └── legacy.py
├── data/                   # Fichiers Excel / JSON créés automatiquement
├── README.md
└── Suivi Alternance.exe    # Version exécutable optionnelle
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
- CustomTkinter
- openpyxl
- matplotlib

---

## 👨‍💻 Auteur

**Mounir ELKATMOUR**  
[LinkedIn](https://www.linkedin.com/in/mounir-elkatmour-703aa8269/)
