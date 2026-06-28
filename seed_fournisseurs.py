"""Seed fournisseurs table from demo data (run once)."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from app import app, db
from models import Fournisseur

SRID_SUPPLIERS = [
    "AAKO", "AIFAR", "ALTINCO", "BASF", "BELCHIM", "CORTEVA",
    "DE SANGOSSE", "FINE", "GOWAN", "GREEN HAS ITALIA",
    "GREENHAS JORDAN", "ISAGRO", "NANJING AGROCHEMICAL",
    "SYNGENTA", "SYNGENTA CROP", "VCR",
]

GENETICS_SUPPLIERS = [
    "SYNGENTA", "VCR",
]

def seed():
    with app.app_context():
        added = 0
        for nom in SRID_SUPPLIERS:
            exists = Fournisseur.query.filter(
                db.func.lower(Fournisseur.nom) == nom.lower(),
                Fournisseur.societe == 'SRID'
            ).first()
            if not exists:
                db.session.add(Fournisseur(nom=nom, societe='SRID'))
                added += 1

        for nom in GENETICS_SUPPLIERS:
            exists = Fournisseur.query.filter(
                db.func.lower(Fournisseur.nom) == nom.lower(),
                Fournisseur.societe == 'SRID GENETICS'
            ).first()
            if not exists:
                db.session.add(Fournisseur(nom=nom, societe='SRID GENETICS'))
                added += 1

        db.session.commit()
        print(f"Seed terminé: {added} fournisseur(s) ajouté(s).")

if __name__ == '__main__':
    seed()
