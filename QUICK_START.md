# 🚀 Quick Start - Test Logistique Module

## 1️⃣ Lancer l'Application

```bash
cd c:\WorkSpace_AWS\SRID
python app.py
```

Accès: **http://localhost:5000/logistique/bons**

---

## 2️⃣ Test Rapide (5 minutes)

### ✅ Test 1: Créer un Bon avec Auto-Totaux

1. Clique **"Nouveau bon"**
2. Remplis:
   - Société: **SRID**
   - Fournisseur: Tape **"ABD"** → Suggestionn apparaît
   - Date: Auto-remplie
3. Clique **"Ajouter une ligne"**
4. Dans la ligne:
   - Désignation: Tape **"câble"** → 2 suggestions apparaissent
   - Sélectionne une → Référence auto-remplie ✅
   - Qté: **5**
   - Prix: **100.50**
   - Montant ligne affiche: **502.50 €** ✅
5. Ajoute 2ème ligne:
   - Désignation: **"Connecteur"**
   - Qté: **2**
   - Prix: **25.00**
   - Total ligne: **50.00 €** ✅
6. Vérifier **"Montant total: 552.50 €"** ✅
7. Clique **"Créer le bon"** ✅

### ✅ Test 2: Vérifier Auto-Création Logistique

1. Va dans **"Gestion de commandes"** (même menu)
2. Cherche le bon créé (BC-2026-XXXX)
3. Devrait avoir:
   - Statut: **EN COURS** ✅
   - Montant: **552.50 €** ✅
   - Société: **SRID** ✅

### ✅ Test 3: Tester l'Impression

1. Reviens à **"Bons de commande"**
2. Clique sur l'œil 👁️ du bon créé
3. Modal s'ouvre avec détails
4. Clique **"Imprimer"** 🖨️
5. Nouvelle fenêtre avec mise en page A4
6. **Ctrl+P** → Imprimer PDF ✅

### ✅ Test 4: Tester SRID Genetics

1. Nouveau bon
2. Société: **SRID GENETICS**
3. Ajoute ligne avec: **"ADN"** → Suggestions Genetics apparaissent ✅
4. Produits différents = sources différentes ✅

---

## 📊 Produits Disponibles pour Test

### SRID (15 produits)
```
REF-001: Câble Réseau Cat6 100m
REF-002: Connecteur RJ45 (pack 50)
REF-003: Coffret Distribution 24 Ports
REF-004: Baie de Brassage 19" 42U
REF-005: Patch Panel 24 Ports Cat6
... et 10 autres
```

### Genetics (15 produits)
```
GEN-001: Kit Test ADN Standard
GEN-002: Réactifs PCR Premium
GEN-003: Électrophorèse Capillaire
GEN-004: Séquenceur 96 Capillaires
... et 11 autres
```

Teste avec n'importe quel produit!

---

## 🎯 Points Clés à Vérifier

| # | Fonctionnalité | Vérifier |
|----|-----------------|----------|
| 1 | Auto-calc totaux | Ligne: 5 × 10.50 = 52.50 € |
| 2 | Auto-complétion | Tape désignation → suggestions |
| 3 | Auto-création log | Dans "Gestion" apparaît |
| 4 | Impression | Fenêtre A4 professionnelle |
| 5 | Statut EN COURS | Auto-calculé = EN COURS |
| 6 | Pas livraison | Champ absent du formulaire |
| 7 | Pas unité | Colonne absente de table |
| 8 | Pas notes table | Notes en modal uniquement |

---

## 📋 Checklist Complet (11 Points)

- [x] 1. Retirer "Livraison prévue"
- [x] 2. Fournisseur list model
- [x] 3. Retirer "Unité"
- [x] 4. Reference/Désignation auto-link
- [x] 5. Colonne "Total €" par ligne
- [x] 6. Montant total affichage
- [x] 7. Masquer notes table
- [x] 8. HTMX auto-filtres (conservé design classique)
- [x] 9. Colonnes minuscules standardisées
- [x] 10. Auto-créer entry logistique
- [x] 11. Print support

**Status**: ✅ **TOUS LES 11 POINTS COMPLÉTÉS**

---

## 🔄 Next Steps pour Vrais Produits

Quand tu as les vrais produits:

1. **Convertir Excel**:
   - Ouvre `Liste des produits SRID.xls`
   - Enregistrer sous → `.xlsx`
   - Répète pour `Produit srid genetics.xls`

2. **Charger produits**:
   ```bash
   python populate_products.py
   # Demande colonnes Excel
   # Charge tous les produits
   ```

3. **Testet avec vrais produits** ✅

---

## 🆘 Debug

Si erreur 404 sur API:
```bash
# Vérifier routes enregistrées
python -c "from app import app; [print(r) for r in app.url_map.iter_rules() if 'products' in str(r) or 'print' in str(r)]"
```

Si produits manquent:
```bash
# Vérifier count
python -c "from app import db; from models import Product; print(f'Products: {Product.query.count()}')"

# Re-seed si besoin
python seed_test_products.py
```

---

## 📞 Fichiers Clés

| Fichier | Rôle |
|---------|------|
| `models.py` | Modèle Product + statut machine |
| `app.py` | Routes API + auto-logistics |
| `templates/logistique_bons.html` | UI + JS calculs |
| `templates/bon_print.html` | Template impression |
| `LOGISTIQUE_GUIDE.md` | Documentation complète |
| `COMPLETION_SUMMARY.md` | Résumé technique |

---

**C'est parti ! Commence par tester et dis-moi si tu as besoin d'ajustements.** 🚀
