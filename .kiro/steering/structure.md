---
inclusion: always
---

# Structure du projet — SRID COM

## Arborescence
```
SRID/
├── app.py                    # Application Flask (routes, logique métier, filtres, pagination, endpoints HTMX)
├── models.py                 # Modèles SQLAlchemy
├── config.py                 # Configuration Flask
├── passenger_wsgi.py         # Point d'entrée cPanel / Passenger
├── database.db               # SQLite (gitignored)
├── requirements.txt
├── import_excel_cheques.py   # Import opérations depuis "srid finalisé.xlsx" (vide + réimporte)
├── seed_fournisseurs.py      # Seed fournisseurs
├── load_products.py / populate_products.py / seed_test_products.py
├── generate_test_data.py     # Génération de données de test
├── static/
│   ├── manifest.json, sw.js  # PWA
│   └── img/                  # entete-srid.png, entete-srid-genetics.png, icônes
└── templates/
    ├── base.html             # Layout (sidebar + header)
    ├── login.html, dashboard.html, saisie.html, edit.html
    ├── consultation.html     # Liste opérations + filtres
    ├── import.html, referentiels.html, historique.html, users.html
    ├── logistique_gestion.html, logistique_bons.html
    ├── logistique_referentiels.html, logistique_prix_revient.html
    ├── bon_print.html / bon_print
    └── partials/             # Fragments HTMX réutilisables
        ├── operations_table.html, operation_detail.html, edit_form.html
        ├── logistique_gestion_table.html, logistique_bons_table.html
        ├── logistique_frais_table.html, frais_modal.html
        ├── notifications_panel.html, logistique_notifications_panel.html
        ├── ref_clients_list.html, ref_remettants_list.html, ref_fournisseurs_list.html
        └── ...
```

## Modèles (models.py)

- **User** — `username` (unique, lowercase), `password_hash` (werkzeug), `nom_complet`,
  `role` (`admin` | `saisie` | `consultation`).
- **Operation** — `type_operation` (`Chèque`/`Virement`/`Versement`/`Transfer`/`Autre`),
  `societe` (**`SRID`** | **`Genetics`**, NOT NULL), `client` (NOT NULL), `montant` (NOT NULL),
  `date_operation` (NOT NULL), `date_reception`, `date_encaissement`
  ( ⚠️ **stocke la date d'échéance** pour les chèques), `banque` (normalisée), `numero_piece`,
  `statut`, `type_detail` (`Garantie`/`À encaisser`/`À échéance`), `remettant`, `remarque`,
  `cree_par`. Colonnes indexées : `societe`, `date_operation`, `date_encaissement`, `statut`.
- **BonCommande** — `numero` (unique `BC-YYYY-NNNN`), `societe`, `fournisseur`,
  `statut` (`Brouillon`→`En attente`→`Approuvé`→`Envoyé`→`Reçu`), dates ; relation `lignes`
  (cascade delete).
- **LigneCommande** — FK `bon_id`, `reference`, `designation`, `quantite`, `unite`, `prix_unitaire`.
- **CommandeLogistique** — FK `bon_id`, `ref_log`, `societe`, dates (`date_d10`, `date_arrivee`,
  `date_echeance`, `date_paiement`, `date_valeur`), `montant_eur`, `cours`, `code_paiement`.
  `statut` et `montant_da` sont des **propriétés Python calculées** (pas des colonnes).
- **FraisLogistique** — FK `bon_id` ; frais et configuration de prix de revient.
- **Fournisseur** — `nom`, `societe` (`SRID` | `SRID GENETICS`), `actif` ; UNIQUE(nom, societe).
- **ClientLabel / RemettantLabel** — `nom` (unique), `actif`.
- **Product** — `company` (`SRID` | `Genetics`), `reference`, `designation`.
- **AuditLog** — `operation_id`, `action`, `utilisateur`, `details`, `date_action`.

### Clés étrangères vers `bons_commande.id`
`commandes_logistique.bon_id`, `frais_logistique.bon_id`, `lignes_commande.bon_id`.
Ordre de suppression FK-safe : CommandeLogistique → FraisLogistique → LigneCommande →
BonCommande → Operation → AuditLog.

## Logique métier

### Finance — statut des opérations
- Statut initial : non-chèque → `Encaissé` ; chèque `À échéance` → `Échéance` ; sinon → `En cours`.
- Mise à jour auto (à la consultation) : échéance passée → `Échu` ; dans 7 jours →
  `Arrive à échéance`. Ne touche jamais `Encaissé` / `Rejeté` (statuts manuels).

### Logistique — flux
1. Création d'un `BonCommande` (numéro auto `BC-YYYY-NNNN`).
2. Création automatique d'une `CommandeLogistique` associée (`bon_id`).
3. Le bon suit `Brouillon` → `En attente` → `Approuvé` → `Envoyé` → `Reçu`.
4. La commande suit un statut **calculé** par remplissage progressif des dates :
   `date_valeur`→PAYÉ, `date_paiement`→PAIEMENT EN COURS, échéance passée→ÉCHU,
   échéance ≤7j→ARRIVE À ÉCHÉANCE, échéance→ÉCHÉANCE, arrivée→ARRIVÉ, D10→D10, sinon→EN COURS.

## Conventions
- Routes et logique dans `app.py` ; pages dans `templates/`, fragments HTMX dans `templates/partials/`.
- Nommer les endpoints partiels de façon cohérente et réutiliser les helpers de requête.
- Garder desktop et mobile alignés quand une table a les deux rendus.
- Toute URL construite en JS doit passer par `SCRIPT_ROOT` (voir `tech.md`).
