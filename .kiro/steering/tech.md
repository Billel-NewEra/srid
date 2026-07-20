---
inclusion: always
---

# Stack technique — SRID COM

## Composants
| Composant | Détail |
|-----------|--------|
| Backend | Python 3 / **Flask 3.1.1** |
| ORM | **Flask-SQLAlchemy 3.1.1** |
| Base de données | **SQLite** (`database.db`, gitignored) |
| Formulaires | Flask-WTF 1.2.2 / WTForms 3.2.1 |
| Excel I/O | openpyxl 3.1.5 |
| Frontend CSS | TailwindCSS (CDN) + **DaisyUI 4.12.22** |
| Interactivité | **HTMX 2.0.4** |
| Charts | Chart.js 4.4.7 |
| Icônes | Font Awesome 6.5.1 |
| Hébergement | cPanel (Passenger WSGI) |
| PWA | Service Worker + manifest.json |

Dépendances Python : voir `requirements.txt`. Config Flask : `config.py`
(`SECRET_KEY`, `SQLALCHEMY_DATABASE_URI` SQLite, session 7 jours).

## Commandes utiles
Interpréteur Windows : `"C:/Program Files/Python312/python.exe"`. Shell : Git Bash
(cwd `/c/WorkSpace_AWS/SRID`).

```bash
# Valider la syntaxe après édition backend
cd /c/WorkSpace_AWS/SRID && python -m py_compile app.py

# Lancer en local
python app.py
```

Smoke-test via le test-client Flask (rendu de page sans navigateur) :
```python
with app.test_client() as c:
    with c.session_transaction() as s:
        s['user_id'] = 1; s['user_nom'] = 'dm3595'; s['user_role'] = 'admin'
    print(c.get('/operations').status_code)
```

## Contraintes techniques critiques

### Déploiement sous sous-chemin (SCRIPT_ROOT)
- L'app tourne sous un sous-chemin (ex. `mondomaine.com/app`), pas à la racine.
- En Jinja : **toujours** `url_for()` (gère le préfixe automatiquement).
- En JS : `var SCRIPT_ROOT = {{ request.script_root|tojson }};` puis
  **toujours** préfixer les `fetch()`, `form.action`, `<img src>` construits en JS
  par `SCRIPT_ROOT + '/path'`.

### SQLite / migrations
- `database.db` est gitignored ; chaque environnement a sa propre base.
- `db.create_all()` au démarrage crée les tables manquantes mais **ne modifie pas**
  les tables existantes.
- **Pas d'Alembic** : pour ajouter une colonne → `ALTER TABLE ... ADD COLUMN` en SQL direct.
- Toujours **sauvegarder** la base avant une opération destructive
  (ex. `database.db.bak-YYYYMMDD-HHMMSS`).

### HTMX
- Tables paginées via `hx-get` + `hx-target` + `hx-swap="innerHTML"`.
- Formulaires : `hx-post` avec réponse partielle.
- Badge notifications : `hx-trigger="load, every 60s"`.
- Suppression : intercepter `htmx:confirm` → modal DaisyUI.
- `login_required` renvoie 401 JSON quand l'en-tête `HX-Request` est présent.

### Pagination
- `REF_PER_PAGE = 25`, `LOG_PER_PAGE = 25`, `BON_PER_PAGE = 25` (OFFSET/LIMIT).
- Exception : le filtre par statut logistique charge tout en mémoire
  (statut logistique = propriété Python calculée, pas une colonne).

## Performance en place
- ~13 index SQL sur les colonnes filtrées/triées fréquemment.
- Dashboard : requêtes `GROUP BY` agrégées plutôt que des centaines de requêtes.
- KPIs logistique : `COUNT` SQL reproduisant la logique de statut.
- Rejets : batch `IN(...)` au lieu de N+1.

## Constantes clés (app.py)
```python
CHECK_TYPE_CHOICES = ['Garantie', 'À encaisser', 'À échéance']
STATUS_CHOICES     = ['Encaissé', 'Rejeté', 'Échéance', 'En cours', 'Arrive à échéance', 'Échu']
BON_STATUTS        = ['Brouillon', 'En attente', 'Approuvé', 'Envoyé', 'Reçu']
LOG_STATUTS        = ['EN COURS', 'D10', 'ARRIVÉ', 'ÉCHÉANCE', 'ARRIVE À ÉCHÉANCE', 'ÉCHU', 'PAIEMENT EN COURS', 'PAYÉ']
```

## Règles de développement
- Conserver l'architecture **Flask + Jinja + HTMX** ; ne pas introduire de framework front.
- Préférer des changements minimes et locaux aux gros refactors.
- Ne pas dupliquer la logique métier entre routes page complète et routes partielles HTMX
  (réutiliser un helper).
- Respecter le langage visuel DaisyUI/Tailwind existant.
- Valider la compilation Python et le rendu des templates après édition.
