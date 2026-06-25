"""Script pour générer des données de test réalistes"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from app import app, db
from models import Operation
from datetime import date, timedelta
import random

# Données réalistes
clients = [
    "SARL Bâtiment Plus", "SCI Les Oliviers", "ETS Mohamed Ali", "SARL TechnoVert",
    "SPA Hydraulique Nord", "EURL Sécurité Pro", "SARL Transport Express", "SCI Résidence Parc",
    "ETS Boulangerie Centrale", "SARL Climatisation Sud", "SPA Carrelage Luxe",
    "EURL Plomberie Générale", "SARL Électricité Moderne", "SCI Tour Bleue",
    "ETS Menuiserie Fine", "SARL Peinture & Déco", "SPA Import-Export Global",
    "EURL Consulting RH", "SARL Informatique Plus", "SCI Jardin d'Eden",
    "ETS Garage Central", "SARL Pharmacie Santé", "SPA Textile Mode",
    "EURL Restauration Rapide", "SARL Immobilier Prestige",
]

banques = ["BNA", "BEA", "CPA", "BADR", "BDL", "SGA", "AGB", "Gulf Bank", "Trust Bank", "Al Baraka"]

familles = ["Travaux", "Services", "Commerce", "Immobilier", "Transport", "Santé", "Industrie"]

remettants = ["M. Benali", "Mme Kaci", "M. Hamdi", "Mme Zeroual", "M. Boudiaf", "Mme Larbi", "M. Saidi"]

statuts_poids = [
    ("Encaissé", 45),
    ("En attente", 25),
    ("En cours", 15),
    ("Rejeté", 8),
    ("Annulé", 7),
]


def weighted_choice(choices):
    total = sum(w for _, w in choices)
    r = random.uniform(0, total)
    upto = 0
    for choice, weight in choices:
        upto += weight
        if r <= upto:
            return choice
    return choices[0][0]


def generate_data():
    with app.app_context():
        # Vérifier si des données existent déjà
        existing = Operation.query.count()
        if existing > 0:
            print(f"Il y a déjà {existing} opérations. On ajoute quand même les données de test.")

        operations = []
        start_date = date(2025, 1, 1)
        end_date = date(2026, 6, 10)
        total_days = (end_date - start_date).days

        for i in range(150):
            type_op = random.choice(["Chèque", "Chèque", "Chèque", "Virement", "Virement", "Versement"])
            societe = random.choice(["ENT", "ENT", "ENT", "Genetics"])
            jour = start_date + timedelta(days=random.randint(0, total_days))

            # Montants réalistes
            if random.random() < 0.3:
                montant = round(random.uniform(50000, 500000), 2)
            elif random.random() < 0.5:
                montant = round(random.uniform(10000, 50000), 2)
            else:
                montant = round(random.uniform(1000, 15000), 2)

            op = Operation(
                type_operation=type_op,
                societe=societe,
                famille=random.choice(familles) if random.random() > 0.3 else None,
                date_operation=jour,
                date_reception=jour - timedelta(days=random.randint(0, 5)) if type_op == "Chèque" else None,
                date_encaissement=jour + timedelta(days=random.randint(1, 30)) if random.random() > 0.6 else None,
                client=random.choice(clients),
                remettant=random.choice(remettants) if type_op == "Chèque" and random.random() > 0.4 else None,
                montant=montant,
                banque=random.choice(banques),
                numero_piece=str(random.randint(100000, 999999)) if type_op in ["Chèque", "Versement"] else None,
                statut=weighted_choice(statuts_poids),
                type_detail=random.choice(["Courant", "Garantie"]) if type_op == "Chèque" and random.random() > 0.5 else None,
                entree=random.random() > 0.3,
                sortie=random.random() > 0.8,
                remarque=random.choice([None, None, None, "RAS", "À vérifier", "Urgent", "Report demandé par client", "Doublon possible"]),
                cree_par=random.choice(["Agent de saisie", "Directeur"]),
            )
            operations.append(op)

        db.session.add_all(operations)
        db.session.commit()
        print(f"✓ {len(operations)} opérations générées avec succès !")
        print(f"  - Période : {start_date} → {end_date}")
        print(f"  - Total en base : {Operation.query.count()}")


def generate_echeance_test_cases():
    """Génère des cas de test spécifiques pour valider le comportement des notifications d'échéance."""
    today = date.today()

    test_cases = [
        # --- Cas 1 : Échu (date passée depuis 5 jours) — doit passer à Échu automatiquement
        {
            "label": "Échu - date passée il y a 5 jours",
            "date_operation": today - timedelta(days=30),
            "date_reception": today - timedelta(days=32),
            "date_encaissement": today - timedelta(days=5),
            "statut": "Échéance",
            "client": "TEST - SARL Échu Cinq Jours",
            "montant": 75000.00,
        },
        # --- Cas 2 : Échu (date passée hier) — cas limite
        {
            "label": "Échu - date passée hier",
            "date_operation": today - timedelta(days=20),
            "date_reception": today - timedelta(days=22),
            "date_encaissement": today - timedelta(days=1),
            "statut": "Échéance",
            "client": "TEST - ETS Échu Hier",
            "montant": 42000.00,
        },
        # --- Cas 3 : Arrive à échéance - dans 2 jours (alerte urgente)
        {
            "label": "Arrive à échéance - dans 2 jours",
            "date_operation": today - timedelta(days=25),
            "date_reception": today - timedelta(days=27),
            "date_encaissement": today + timedelta(days=2),
            "statut": "Échéance",
            "client": "TEST - SARL Échéance Deux Jours",
            "montant": 120000.00,
        },
        # --- Cas 4 : Arrive à échéance - dans 5 jours
        {
            "label": "Arrive à échéance - dans 5 jours",
            "date_operation": today - timedelta(days=20),
            "date_reception": today - timedelta(days=22),
            "date_encaissement": today + timedelta(days=5),
            "statut": "Échéance",
            "client": "TEST - ETS Échéance Cinq Jours",
            "montant": 88000.00,
        },
        # --- Cas 5 : Arrive à échéance - dans 7 jours (limite de la fenêtre d'alerte)
        {
            "label": "Arrive à échéance - dans 7 jours (limite)",
            "date_operation": today - timedelta(days=15),
            "date_reception": today - timedelta(days=17),
            "date_encaissement": today + timedelta(days=7),
            "statut": "Échéance",
            "client": "TEST - SARL Échéance Sept Jours",
            "montant": 55000.00,
        },
        # --- Cas 6 : Échéance future normale (dans 15 jours — pas d'alerte)
        {
            "label": "Échéance future - dans 15 jours (pas d'alerte)",
            "date_operation": today - timedelta(days=10),
            "date_reception": today - timedelta(days=12),
            "date_encaissement": today + timedelta(days=15),
            "statut": "Échéance",
            "client": "TEST - SPA Échéance Future Quinze Jours",
            "montant": 200000.00,
        },
        # --- Cas 7 : Statut déjà Échu manuellement (ne doit pas être retouché)
        {
            "label": "Échu manuellement (ne doit pas changer)",
            "date_operation": today - timedelta(days=60),
            "date_reception": today - timedelta(days=62),
            "date_encaissement": today + timedelta(days=10),
            "statut": "Échu",
            "client": "TEST - EURL Échu Manuel",
            "montant": 33000.00,
        },
    ]

    with app.app_context():
        created = 0
        for tc in test_cases:
            op = Operation(
                type_operation="Chèque",
                type_detail="À échéance",
                societe="SRID",
                date_operation=tc["date_operation"],
                date_reception=tc["date_reception"],
                date_encaissement=tc["date_encaissement"],
                statut=tc["statut"],
                client=tc["client"],
                remettant="Agent Test",
                montant=tc["montant"],
                banque="BNA",
                numero_piece=str(100000 + created),
                cree_par="Script Test",
                remarque=f"[TEST] {tc['label']}",
            )
            db.session.add(op)
            created += 1
            print(f"  + {tc['label']} → statut initial: {tc['statut']} | échéance: {tc['date_encaissement']}")

        db.session.commit()
        print(f"\n✓ {created} cas de test échéance créés.")
        print("  → Ouvrez /consultation et filtrez par 'TEST' pour les voir.")
        print("  → La route /api/operations appellera _auto_update_echeance_statuts() automatiquement.")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--echeance":
        generate_echeance_test_cases()
    else:
        generate_data()
