# SRID COM — Documentation Technique Complète

## 1. Vue d'ensemble

**SRID COM** est une application web de gestion financière et logistique pour les sociétés **SRID** et **SRID Genetics** (société algérienne spécialisée en produits phytosanitaires/agrochimiques).

**Utilisateurs** : Équipe administrative (Mehdi = admin, Sabrina = saisie) + consultants en lecture seule.

**Objectif** : Suivi des opérations financières (chèques, virements, versements) et de la chaîne logistique (commandes fournisseurs, suivi d'import avec dates douanières D10, arrivée, échéances de paiement).

**Développeur** : MobiBenz (https://mobibenz.dz/fr)

---

## 2. Stack technique

| Composant | Détail |
|-----------|--------|
| Backend | Python 3 / Flask 3.1.1 |
| ORM | Flask-SQLAlchemy 3.1.1 |
| Base de données | SQLite (`database.db`, gitignored) |
| Formulaires | Flask-WTF 1.2.2 / WTForms 3.2.1 |
| Excel I/O | openpyxl 3.1.5 |
| Frontend CSS | TailwindCSS (CDN) + DaisyUI 4.12.22 |
| Interactivité | HTMX 2.0.4 |
| Charts | Chart.js 4.4.7 |
| Icônes | Font Awesome 6.5.1 |
| Hébergement | cPanel (Passenger WSGI) |
| PWA | Service Worker + manifest.json |

---

## 3. Architecture fichiers

```
SRID/
├── app.py                    # Application Flask principale (~2270 lignes)
├── models.py                 # Modèles SQLAlchemy (~260 lignes)
├── config.py                 # Configuration Flask
├── passenger_wsgi.py         # Point d'entrée cPanel/Passenger
├── database.db               # SQLite (gitignored)
├── requirements.txt
├── seed_fournisseurs.py      # Script seed fournisseurs
├── generate_test_data.py     # Génération données test
├── load_products.py          # Import produits depuis Excel
├── .gitignore
├── static/
│   ├── icon-192.png, icon-512.png
│   ├── manifest.json         # PWA manifest
│   ├── sw.js                 # Service Worker
│   └── img/
│       ├── entete-srid.png
│       └── entete-srid-genetics.png
├── templates/
│   ├── base.html             # Layout principal (sidebar + header)
│   ├── login.html
│   ├── dashboard.html        # Dashboard KPIs + graphiques
│   ├── saisie.html           # Formulaire nouvelle opération
│   ├── edit.html             # Page édition standalone
│   ├── consultation.html     # Liste opérations + filtres
│   ├── import.html           # Import Excel
│   ├── referentiels.html     # Gestion clients/remettants
│   ├── historique.html       # Audit log
│   ├── users.html            # Gestion utilisateurs (admin)
│   ├── logistique_gestion.html   # Tableau commandes logistiques
│   ├── logistique_bons.html      # Bons de commande
│   ├── logistique_referentiels.html  # Fournisseurs par société
│   ├── bon_print.html        # Impression bon de commande
│   └── partials/
│       ├── operations_table.html
│       ├── operation_detail.html
│       ├── edit_form.html
│       ├── success_message.html
│       ├── notifications_panel.html
│       ├── logistique_notifications_panel.html
│       ├── global_notifications_badge.html
│       ├── rejections_panel.html
│       ├── logistique_gestion_table.html
│       ├── logistique_bons_table.html
│       ├── log_form_fields.html
│       ├── ref_clients_list.html
│       ├── ref_remettants_list.html
│       └── ref_fournisseurs_list.html
```

---

## 4. Modèles (models.py)

### User
| Champ | Type | Notes |
|-------|------|-------|
| id | Integer PK | |
| username | String(80) | unique |
| password_hash | String(200) | werkzeug |
| nom_complet | String(150) | |
| role | String(20) | `admin`, `saisie`, `consultation` |

### Operation
| Champ | Type | Notes |
|-------|------|-------|
| id | Integer PK | |
| type_operation | String(20) | `Chèque`, `Virement`, `Versement`, `Transfer`, `Autre` |
| societe | String(100) | `SRID` ou `Genetics`, **indexé** |
| famille | String(100) | |
| date_operation | Date | **indexé** |
| date_reception | Date | chèques uniquement |
| date_encaissement | Date | **ATTENTION: stocke la date d'échéance** pour chèques, **indexé** |
| date_sortie | Date | |
| client | String(200) | |
| remettant | String(200) | |
| montant | Float | |
| banque | String(100) | normalisée |
| numero_piece | String(50) | |
| statut | String(50) | colonne DB, **indexé** |
| type_detail | String(50) | `Garantie`, `À encaisser`, `À échéance` |
| entree | String(100) | |
| sortie | String(100) | |
| remarque | Text | |
| cree_par | String(100) | |

**Statuts possibles** : `Encaissé`, `Rejeté`, `Échéance`, `En cours`, `Arrive à échéance`, `Échu`

### CommandeLogistique
| Champ | Type | Notes |
|-------|------|-------|
| id | Integer PK | |
| bon_id | Integer FK → bons_commande, **indexé** | |
| ref_log | String(20) | `LOG-0001` ou numéro BC |
| societe | String(50) | **indexé** |
| annee | String(4) | |
| date_d10 | Date | déclaration douane |
| date_arrivee | Date | **indexé** |
| fournisseur | String(200) | |
| produit | String(200) | |
| emballage | String(100) | |
| quantite | Float | |
| tva | Float | |
| montant_eur | Float | |
| cours | Float | taux EUR/DA (4 déc.) |
| date_facture | Date | |
| code_paiement | String(10) | `T`, `R`, `CAD` |
| nb_jours | Integer | |
| date_echeance | Date | **indexé** |
| date_paiement | Date | |
| date_valeur | Date | |
| remarque | Text | |

**Propriété `statut`** (Python calculé, PAS une colonne DB) :
1. `date_valeur` → `PAYÉ`
2. `date_paiement` → `PAIEMENT EN COURS`
3. `date_echeance < today` → `ÉCHU`
4. `date_echeance ≤ 7j` → `ARRIVE À ÉCHÉANCE`
5. `date_echeance set` → `ÉCHÉANCE`
6. `date_arrivee set` → `ARRIVÉ`
7. `date_d10 set` → `D10`
8. Sinon → `EN COURS`

**Propriété `montant_da`** : `montant_eur × cours`

### BonCommande
| Champ | Type | Notes |
|-------|------|-------|
| id | Integer PK | |
| numero | String(50) | unique, format `BC-YYYY-NNNN` |
| societe | String(50) | **indexé** |
| fournisseur | String(200) | |
| statut | String(30) | DB: `Brouillon`→`En attente`→`Approuvé`→`Envoyé`→`Reçu`, **indexé** |
| date_commande | Date | **indexé** |
| date_livraison_prevue | Date | |
| notes | Text | |
| lignes | rel → LigneCommande | cascade delete |

### LigneCommande
| Champ | Type |
|-------|------|
| id | Integer PK |
| bon_id | FK → bons_commande |
| reference | String(100) |
| designation | String(300) |
| quantite | Float |
| unite | String(50) |
| prix_unitaire | Float |

### Fournisseur
| Champ | Type | Notes |
|-------|------|-------|
| nom | String(200) | |
| societe | String(50) | `SRID` ou `SRID GENETICS` |
| actif | Boolean | |

**Contrainte** : UNIQUE(nom, societe)

### ClientLabel / RemettantLabel
| Champ | Type |
|-------|------|
| nom | String(200), unique |
| actif | Boolean |

### AuditLog
| Champ | Type | Notes |
|-------|------|-------|
| operation_id | Integer | **indexé** |
| action | String(20) | `création`, `modification`, `suppression` |
| utilisateur | String(100) | |
| details | Text | |
| date_action | DateTime | **indexé** |

### Product
| Champ | Type |
|-------|------|
| company | String(100), **indexé** | `SRID` ou `Genetics` |
| reference | String(100) |
| designation | String(300) |

---

## 5. Logique métier

### Finance — Cycle de vie des opérations

**Auto-calcul statut initial** :
- Pas un chèque → `Encaissé`
- Chèque `À échéance` → `Échéance`
- Sinon → `En cours`

**Mise à jour auto** (à chaque consultation) :
- Date échéance passée + statut `Échéance`/`Arrive à échéance` → `Échu`
- Dans les 7 jours + statut `Échéance` → `Arrive à échéance`
- Ne touche PAS `Encaissé`/`Rejeté` (statuts manuels)

### Logistique — Flux

1. Création `BonCommande` (numéro auto `BC-YYYY-NNNN`)
2. Création auto d'une `CommandeLogistique` associée (`bon_id`)
3. Le bon suit : `Brouillon` → `En attente` → `Approuvé` → `Envoyé` → `Reçu`
4. La commande suit le statut calculé (remplissage progressif des dates)

### Référentiels
- **Clients/Remettants** : autocomplete formulaire finance
- **Fournisseurs** : par société, autocomplete logistique (dropdown dynamique selon société sélectionnée)
- **Produits** : par société, autocomplete lignes de bon

### Rôles

| Rôle | Droits |
|------|--------|
| `admin` | TOUT (CRUD, users, suppression, statuts) |
| `saisie` | Création/modification, référentiels. Pas suppression ni gestion users |
| `consultation` | Lecture seule |

---

## 6. Contraintes techniques critiques

### cPanel SCRIPT_ROOT
- L'app tourne à `mondomaine.com/app` (pas à la racine)
- **TOUTE URL JS** doit être préfixée par `SCRIPT_ROOT`
- En Jinja, `url_for()` gère automatiquement
- En JS : `var SCRIPT_ROOT = {{ request.script_root|tojson }};`
- Pour `fetch()`, `form.action`, `<img src>` construits en JS → toujours `SCRIPT_ROOT + '/path'`

### SQLite
- `database.db` est gitignored — chaque environnement a sa propre DB
- `db.create_all()` au démarrage crée tables manquantes mais ne modifie PAS les existantes
- Pour ajouter une colonne : `ALTER TABLE ... ADD COLUMN` en SQL direct
- Pas d'Alembic — migrations manuelles

### HTMX Patterns
- Tables paginées via `hx-get` + `hx-target` + `hx-swap="innerHTML"`
- Formulaires : `hx-post` avec réponse partielle
- Badge notifications : `hx-trigger="load, every 60s"`
- Confirmation suppression : intercepte `htmx:confirm` → modal DaisyUI
- `login_required` retourne 401 JSON pour requêtes HTMX (détecte `HX-Request` header)

### Pagination
- `REF_PER_PAGE = 25`, `LOG_PER_PAGE = 25`, `BON_PER_PAGE = 25`
- Toutes les listes HTMX utilisent `OFFSET/LIMIT`
- **Exception** : filtre statut logistique charge tout en mémoire (statut calculé)

---

## 7. Optimisations en place

- **13 index SQL** sur colonnes fréquemment filtrées/triées
- **Dashboard** : 1 requête GROUP BY au lieu de 216 requêtes individuelles
- **KPIs logistique** : COUNT SQL reproduisant la logique statut (au lieu de charger tous les objets)
- **Rejets** : batch `IN(...)` au lieu de N+1
- **Bons** : produits/fournisseurs JSON chargés seulement si `can_write`

---

## 8. Constantes

```python
CHECK_TYPE_CHOICES = ['Garantie', 'À encaisser', 'À échéance']
STATUS_CHOICES = ['Encaissé', 'Rejeté', 'Échéance', 'En cours', 'Arrive à échéance', 'Échu']
BON_STATUTS = ['Brouillon', 'En attente', 'Approuvé', 'Envoyé', 'Reçu']
LOG_STATUTS = ['EN COURS', 'D10', 'ARRIVÉ', 'ÉCHÉANCE', 'ARRIVE À ÉCHÉANCE', 'ÉCHU', 'PAIEMENT EN COURS', 'PAYÉ']
```

**Sociétés Finance** : `SRID`, `Genetics`
**Sociétés Logistique** : `SRID`, `SRID GENETICS`

---

## 9. Déploiement

### cPanel
1. Upload fichiers via git ou FTP
2. `passenger_wsgi.py` = point d'entrée
3. `pip install -r requirements.txt`
4. DB créée automatiquement au premier lancement
5. Seeds auto au démarrage (utilisateurs, référentiels)

### Migrations (ajouter colonne existante)
```python
import sqlite3
conn = sqlite3.connect('database.db')
conn.execute("ALTER TABLE table_name ADD COLUMN col_name TYPE DEFAULT val")
conn.commit()
conn.close()
```

### Créer index
```python
conn.execute("CREATE INDEX IF NOT EXISTS ix_name ON table(col)")
```

### Reset password
```python
from app import app, db
from models import User
with app.app_context():
    u = User.query.filter_by(username='mehdi').first()
    u.set_password('srid2024boss')
    db.session.commit()
```

### Seed fournisseurs
```bash
python seed_fournisseurs.py
```

---

## 10. Comptes par défaut

| Username | Password | Rôle |
|----------|----------|------|
| mehdi | srid2024boss | admin |
| sabrina | srid2024sab | saisie |
| mobibenz | (dev) | admin |

---

## 11. Quirks à connaître

1. **`date_encaissement`** = date d'échéance (mal nommé historiquement)
2. **Statut logistique** = propriété Python, pas colonne DB → filtrage en mémoire
3. **Société** s'écrit différemment selon la section (`Genetics` vs `SRID GENETICS`)
4. **Thème** : DaisyUI `luxury` (dark) / `corporate` (light), préférence Carbon Blue
5. **Normalisation banques** : `_normalize_bank_name()` gère des dizaines de variantes
6. **PWA** : Service Worker pour installation mobile
