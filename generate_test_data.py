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


if __name__ == "__main__":
    generate_data()
