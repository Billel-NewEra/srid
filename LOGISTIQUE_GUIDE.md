# 📦 SRID COM - Module Logistique | Guide Complet

## ✅ État de l'Implémentation

### Fonctionnalités Complétées

| # | Requête Utilisateur | Status | Détails |
|---|-------------------|--------|---------|
| 1 | Retirer "Livraison prévue" du bon | ✅ | Colonne supprimée table + modal |
| 2 | Lister fournisseurs par company | ✅ | Via CommandeLogistique queries |
| 3 | Retirer champ "Unité" | ✅ | Supprimé du formulaire et table |
| 4 | Auto-complétion Référence/Désignation | ✅ | API + datalist + 30 produits test |
| 5 | Ajouter colonne "Total €" par ligne | ✅ | Calcul temps réel: qty × price |
| 6 | Afficher "Montant total" du bon | ✅ | Somme des lignes en footer |
| 7 | Masquer "Notes" de la table | ✅ | Conservé en modal détail |
| 8 | Filtres avec HTMX auto-update | ⚠️ | Conservé design classique (button) |
| 9 | Standardiser colonnes minuscules | ✅ | bons + gestion_commandes |
| 10 | Auto-créer entry dans Logistique | ✅ | À création du bon |
| 11 | Fonction d'impression du bon | ✅ | Template PDF + route + bouton |

---

## 🚀 Comment Tester

### 1. **Lancer l'Application**
```bash
cd c:\WorkSpace_AWS\SRID
python app.py
# Accès: http://localhost:5000/logistique/bons
```

### 2. **Test de la Création de Bon**
1. Clique **"Nouveau bon"**
2. Remplis:
   - **Société**: SRID (ou SRID GENETICS)
   - **Fournisseur**: Commence à taper → suggestions apparaissent
   - **Date**: Auto-remplie aujourd'hui
3. Clique **"Ajouter une ligne"**
4. Dans la ligne, tape une **désignation**:
   - Suggestions de produits apparaissent (datalist)
   - Sélectionne un produit → **Référence auto-remplie**
5. Remplis **Qté** et **Prix**
   - **Total ligne s'affiche** à droite
   - **Montant total** se met à jour en bas
6. Clique **"Créer le bon"** ✅

### 3. **Vérifier Auto-Création de l'Entrée Logistique**
1. Crée un bon avec lignes
2. Va dans **"Gestion de commandes"** (même menu)
3. Cherche le bon créé → **Entrée logistique créée automatiquement**
4. Statut = `EN COURS` (auto-calculé par machine à états)

### 4. **Tester l'Impression**
1. Clique sur l'œil de détail d'un bon
2. Clique **"Imprimer"** dans le modal
3. Nouvelle fenêtre s'ouvre avec mise en page professionnelle
4. **File → Imprimer** (ou Ctrl+P) pour vraiment imprimer

### 5. **Tester les Totaux**
1. Crée un bon avec plusieurs lignes:
   - Ligne 1: 5 × 10.00 = 50.00 €
   - Ligne 2: 2 × 25.50 = 51.00 €
   - Ligne 3: 1 × 100.00 = 100.00 €
2. Vérifier **"Montant total: 201.00 €"** s'affiche bien

---

## 🗄️ Nouvelles Tables & Routes

### Modèle Product
```python
class Product(db.Model):
    company      # 'SRID' ou 'Genetics'
    reference    # REF-001, GEN-015, etc.
    designation  # Description complète du produit
```

**30 produits de test pré-chargés:**
- 15 pour SRID (REF-001 à REF-015)
- 15 pour Genetics (GEN-001 à GEN-015)

### Nouvelles Routes API

| Route | Méthode | Retour |
|-------|---------|--------|
| `/api/products/by-company?company=SRID` | GET | `[{id, company, reference, designation}, ...]` |
| `/api/products/search?company=SRID&q=câble` | GET | Produits filtrés |
| `/logistique/bons/<id>/print` | GET | HTML imprimable |

### Routes Existantes Mises à Jour
- `POST /api/logistique/bons/add` → **Auto-crée CommandeLogistique**

---

## 📋 Changements de Template

### `logistique_bons.html`
✅ Supprimé:
- Colonne "Livraison prévue" (table + modal)
- Champ "Unité" (formulaire)
- Colonne "Notes" (table)

✅ Ajouté:
- Calcul temps réel des totaux (JS: `updateBonTotals()`)
- Datalist pour auto-complétion produits
- Bouton "Imprimer" dans modal détail
- Classes CSS `.line-qty`, `.line-price`, `.line-total-cell`

### `bon_print.html` (Nouveau)
Template d'impression professionnel:
- En-tête SRID COM
- Tableau des articles avec totaux ligne
- Montant total
- Remarques en bloc séparé
- CSS optimisé pour impression

### `logistique_gestion.html`
✅ Colonnes standardisées en minuscules (statut, société, fournisseur, etc.)

---

## 🔄 Machine à États du Statut

Pour `CommandeLogistique.statut` (auto-calculé):

```
PAYÉ
  ↓
PAIEMENT EN COURS
  ↓
ARRIVE À ÉCHÉANCE (si date_echeance - today ≤ 7 jours)
ou
ÉCHU (si date_echeance - today < 0)
ou
ÉCHÉANCE (autre)
  ↓
ARRIVÉ (si date_arrivee est défini)
  ↓
D10 (si date_d10 est défini)
  ↓
EN COURS (statut par défaut à création)
```

---

## 💾 Charger les Vrais Produits

### Étape 1: Convertir Excel en XLSX
1. Ouvre `Liste des produits SRID.xls` avec Excel/LibreOffice
2. **Enregistrer sous** → Format `Excel 2007 (.xlsx)`
3. Répète pour `Produit srid genetics.xls`

### Étape 2: Charger avec le Script
```bash
python populate_products.py
```
Script interactive:
- Affiche colonnes disponibles
- Demande numéro de colonne pour Référence & Désignation
- Charge tous les produits dans DB

### Résultat
Les 30 produits de test seront remplacés par les vrais produits de l'entreprise.

---

## 🛠️ Architecture Détaillée

### Backend (Python/Flask)
```
app.py
  ├─ Route: /logistique/bons → logistique_bons()
  ├─ API: /api/logistique/bons/add → api_bon_add()
  │   └─ Auto-crée CommandeLogistique
  ├─ API: /api/products/by-company → api_products_by_company()
  ├─ API: /api/products/search → api_products_search()
  └─ Route: /logistique/bons/<id>/print → print_bon()

models.py
  ├─ BonCommande (bons de commande)
  ├─ LigneCommande (articles du bon)
  ├─ CommandeLogistique (tracking logistique)
  ├─ Fournisseur (suppliers)
  └─ Product (NEW - product catalog)
```

### Frontend (HTML/JS)
```
logistique_bons.html
  ├─ Form création bon
  │   ├─ Datalist fournisseurs (from DB)
  │   ├─ Tableau lignes dynamique
  │   │   └─ Datalist produits (JS fetch API)
  │   └─ Grand total calculé (updateBonTotals)
  ├─ Modal détail bon
  │   └─ Bouton Imprimer → /logistique/bons/<id>/print
  └─ Scripts
      ├─ addBonLine() - Ajouter ligne
      ├─ removeTableLine() - Supprimer ligne
      ├─ updateBonTotals() - Recalculer totaux
      ├─ bindLineEvents() - Écouter changes input
      └─ openBonDetail() - Afficher détail + print

bon_print.html (NEW)
  ├─ Layout A4 optimisé
  ├─ CSS print media queries
  └─ JS auto-print sur ?autoprint=1
```

---

## 📊 Exemple d'Utilisation Complète

### Scénario: Créer un Bon SRID Genetics

```
1. Clique "Nouveau bon"
   │
2. Remplis:
   │  Société: SRID GENETICS
   │  Fournisseur: [commence à taper "ABD"] → Suggestion "ABD CORP" ✓
   │  Date: 27/06/2026 (auto)
   │
3. Ajoute 2 lignes:
   │
   │  Ligne 1:
   │  ├─ Désignation: [tape "ADN"] → Suggestion "Kit Test ADN Standard"
   │  ├─ Référence: [auto-remplit GEN-001]
   │  ├─ Qté: 5
   │  ├─ Prix: 150.00
   │  └─ Total: 750.00 € ✓
   │
   │  Ligne 2:
   │  ├─ Désignation: [tape "PCR"] → Suggestion "Réactifs PCR Premium"
   │  ├─ Référence: [auto-remplit GEN-002]
   │  ├─ Qté: 2
   │  ├─ Prix: 300.00
   │  └─ Total: 600.00 € ✓
   │
   │  Montant total: 1350.00 € ✓
   │
4. Clique "Créer le bon"
   │
5. En base de données:
   │  ├─ BonCommande créé
   │  ├─ 2 LigneCommande créées
   │  ├─ CommandeLogistique créée avec:
   │  │  └─ ref_log: BC-2026-0002
   │  │  └─ montant_eur: 1350.00
   │  │  └─ statut: EN COURS ✓
   │  │  └─ bon_id: [FK vers BonCommande]
   │
6. Utilisateur voit bon dans "Gestion de commandes"
   │  ├─ Référence: BC-2026-0002
   │  ├─ Montant: 1350.00 €
   │  └─ Statut: EN COURS
   │
7. Clique "Voir détail" → Modal affiche tout
   │  └─ Clique "Imprimer" → Nouvelle fenêtre imprimable
   │     └─ Ctrl+P → Imprime PDF
```

---

## ⚙️ Configuration & Dépendances

### Installed Packages
```
Flask==2.3.3
SQLAlchemy==2.0.20
pandas (for Excel reading when XLSX)
```

### Files Created/Modified
```
NEW:
  ├─ models.py → Product class
  ├─ templates/bon_print.html → Print template
  ├─ populate_products.py → Excel → DB loader
  └─ seed_test_products.py → Test data seeder

MODIFIED:
  ├─ app.py → +3 new routes/APIs
  ├─ templates/logistique_bons.html → JS + UI updates
  └─ templates/logistique_gestion.html → Column normalization
```

---

## 📝 Notes de Maintenance

### Pour Ajouter des Produits Réels
1. Convertir XLS → XLSX (Excel/LibreOffice)
2. `python populate_products.py` (interactive)
3. Les 30 produits test seront gardés (pas de suppression)
4. Script skip les doublons (company + reference unique)

### Pour Changer les Colonnes Excel
1. Ouvre `populate_products.py`
2. Modifie mapping colonnes
3. Re-run le script

### Revert aux Produits Test
```bash
# Supprimer les vrais produits
python -c "from app import *; db.session.execute('DELETE FROM products'); db.session.commit()"

# Re-seed les 30 de test
python seed_test_products.py
```

---

## 🎯 Prochaines Améliorations Possibles

- [ ] HTMX auto-filters (remplacer bouton "Filtrer")
- [ ] Notification "ARRIVE À ÉCHÉANCE" (7j avertissement)
- [ ] Import XML/EDI depuis fournisseurs
- [ ] Validation barcodes produits
- [ ] Multi-warehouse support
- [ ] API REST complète (export JSON/CSV)

---

**Version**: 1.0 - 27/06/2026  
**Status**: ✅ Production Ready  
**Produits**: 30 test (remplaçables par vrais produits)
