# Terminal L/D Predictor — Flask App

Estimation du Load et Discharge par lane terminale, basée sur les 3 derniers mois d'activité historique.

---

## Structure des fichiers

```
terminal_predictor/
│
├── app.py                  # Application Flask principale (routes & API)
├── requirements.txt        # Dépendances Python
├── .env.example            # Variables d'environnement (à copier en .env)
├── README.md
│
├── utils/
│   ├── __init__.py
│   ├── predictor.py        # Logique de calcul L/D (ratios + prédiction)
│   └── export.py           # Export Excel formaté (openpyxl)
│
├── templates/
│   └── index.html          # Template HTML Jinja2
│
└── static/
    ├── css/
    │   └── style.css       # Styles de l'application
    └── js/
        └── app.js          # JavaScript frontend (fetch API, UI)
```

---

## Installation locale

```bash
# 1. Créer un environnement virtuel
python -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows

# 2. Installer les dépendances
pip install -r requirements.txt

# 3. Configurer les variables d'environnement
cp .env.example .env
# Éditer .env et changer SECRET_KEY

# 4. Lancer le serveur de développement
python app.py
```

L'app sera disponible sur **http://localhost:5000**

---

## Déploiement en production

### Option A — Gunicorn (Linux / serveur dédié)

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

### Option B — Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8000
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:8000", "app:app"]
```

```bash
docker build -t terminal-predictor .
docker run -p 8000:8000 terminal-predictor
```

### Option C — PythonAnywhere (gratuit)

1. Créer un compte sur pythonanywhere.com
2. Uploader les fichiers via le gestionnaire de fichiers
3. Dans "Web" → "Add a new web app" → Flask → Python 3.11
4. Configurer le chemin vers `app.py`

---

## API Endpoints

| Méthode | Route | Description |
|---------|-------|-------------|
| `GET`  | `/` | Interface principale |
| `GET`  | `/api/lanes` | Liste de toutes les lanes |
| `POST` | `/api/predict` | Prédiction pour 1 lane |
| `POST` | `/api/predict/batch` | Prédiction depuis fichier Excel/CSV |
| `POST` | `/api/export` | Export des résultats en .xlsx |

### Exemple `/api/predict`

**Request:**
```json
POST /api/predict
{ "lane": "FAL1WB", "volume": 5000 }
```

**Response:**
```json
{
  "lane": "FAL1WB",
  "volume": 5000,
  "pred_L": 1640,
  "pred_D": 3360,
  "pct_L": 32.79,
  "pct_D": 67.21,
  "period_start": "2025-09-29",
  "last_date": "2025-12-29",
  "n_voyages": 10,
  "avg_L": 1221.5,
  "avg_D": 2503.4,
  "outdated": false
}
```

---

## Mettre à jour les ratios

Les ratios L/D sont dans `utils/predictor.py` dans le dictionnaire `RATIOS`.

Pour les recalculer avec de nouvelles données, relancer le script Python d'analyse et remplacer le dictionnaire `RATIOS`.
