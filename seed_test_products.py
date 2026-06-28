#!/usr/bin/env python3
"""
Seed test products for autofill functionality demo.
This is a temporary solution while Excel files are being converted to XLSX.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from app import app, db
from models import Product

# Test product data
PRODUCTS_SRID = [
    ('REF-001', 'Câble Réseau Cat6 100m'),
    ('REF-002', 'Connecteur RJ45 (pack 50)'),
    ('REF-003', 'Coffret Distribution 24 Ports'),
    ('REF-004', 'Baie de Brassage 19" 42U'),
    ('REF-005', 'Patch Panel 24 Ports Cat6'),
    ('REF-006', 'Câble Fibre Optique SM 100m'),
    ('REF-007', 'Splitter Optique 1x32'),
    ('REF-008', 'ODF Optique 48 Cœurs'),
    ('REF-009', 'Cassette de Fusion Optique'),
    ('REF-010', 'Testeur de Câble Réseau'),
    ('REF-011', 'Armoire Électrique 800x600'),
    ('REF-012', 'Disjoncteur 16A Courbe C'),
    ('REF-013', 'Parafoudre Télécom 100A'),
    ('REF-014', 'Climatisation Armoire 1500W'),
    ('REF-015', 'Ventilateur Axial 120mm'),
]

PRODUCTS_GENETICS = [
    ('GEN-001', 'Kit Test ADN Standard'),
    ('GEN-002', 'Réactifs PCR Premium'),
    ('GEN-003', 'Électrophorèse Capillaire'),
    ('GEN-004', 'Séquenceur 96 Capillaires'),
    ('GEN-005', 'Cartouche Polymère POP-7'),
    ('GEN-006', 'Microplaques 384 puits'),
    ('GEN-007', 'Pipettes Multichannel (8-12)'),
    ('GEN-008', 'Tips Stériles 10µL'),
    ('GEN-009', 'Centrifugeuse Réfrigérée'),
    ('GEN-010', 'Thermocycleur PCR en Temps Réel'),
    ('GEN-011', 'Vortex Horizontal Max Vitesse'),
    ('GEN-012', 'Bain-Marie Digital 95°C'),
    ('GEN-013', 'Incubateur Secoué Universel'),
    ('GEN-014', 'Microscope Fluorescence'),
    ('GEN-015', 'Caméra Numérique HD pour Microscope'),
]

def load_test_products():
    """Load test products into database."""
    with app.app_context():
        # Count existing products
        existing = Product.query.count()
        print(f"Current products: {existing}")
        
        if existing > 0:
            print("Products already loaded. Skipping.")
            return
        
        # Load SRID products
        for ref, desig in PRODUCTS_SRID:
            product = Product(company='SRID', reference=ref, designation=desig)
            db.session.add(product)
        
        # Load Genetics products
        for ref, desig in PRODUCTS_GENETICS:
            product = Product(company='Genetics', reference=ref, designation=desig)
            db.session.add(product)
        
        db.session.commit()
        
        total = Product.query.count()
        srid_count = Product.query.filter_by(company='SRID').count()
        genetics_count = Product.query.filter_by(company='Genetics').count()
        
        print(f"✓ Loaded {total} test products:")
        print(f"  - SRID: {srid_count}")
        print(f"  - Genetics: {genetics_count}")

if __name__ == '__main__':
    print("=" * 70)
    print("SEED TEST PRODUCTS")
    print("=" * 70)
    print("\n📦 Loading test product data for autofill demo...")
    load_test_products()
    print("\n✓ Done! Auto-complete is now functional.")
    print("   To load real products later, convert Excel files to XLSX")
    print("   and use: python populate_products.py")
