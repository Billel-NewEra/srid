#!/usr/bin/env python3
"""
Load products from Excel files into database.
Run: python load_products.py

This script reads product data from Excel files and populates the database.
Current status: Need to read Excel files manually or specify columns.
"""

import os
import sys
from pathlib import Path

# Add parent to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from app import app, db
from models import Product

# Product data structure (will be populated from Excel)
PRODUCTS = [
    # Example structure:
    # {
    #     'company': 'SRID',
    #     'reference': 'REF001',
    #     'designation': 'Produit 1'
    # },
    # {
    #     'company': 'Genetics',
    #     'reference': 'GEN001',
    #     'designation': 'Produit Genetics'
    # }
]

def load_products_from_excel():
    """
    Load products from Excel files.
    
    Expected files:
    - Liste des produits SRID.xls (company: 'SRID')
    - Produit srid genetics.xls (company: 'Genetics')
    
    TODO: Specify exact column names in these files.
    User should provide: [company_col, reference_col, designation_col]
    """
    try:
        import pandas as pd
    except ImportError:
        print("pandas not installed. Install with: pip install pandas openpyxl xlrd")
        return False
    
    files_to_load = {
        'Liste des produits SRID.xls': 'SRID',
        'Produit srid genetics.xls': 'Genetics'
    }
    
    for filepath, company_name in files_to_load.items():
        if not os.path.exists(filepath):
            print(f"⚠️  File not found: {filepath}")
            continue
        
        print(f"\n📖 Reading {filepath}...")
        try:
            # TODO: Specify sheet name and column names
            df = pd.read_excel(filepath, sheet_name=0)
            print(f"   Columns found: {list(df.columns)}")
            print(f"   Rows: {len(df)}")
            print(f"   First row:\n{df.head(1).to_dict(orient='records')}")
            
            # TODO: Map columns to [reference, designation]
            # Example:
            # ref_col = 'Reference'  # <-- CHANGE THIS
            # desig_col = 'Designation'  # <-- CHANGE THIS
            # for idx, row in df.iterrows():
            #     product = Product(
            #         company=company_name,
            #         reference=row[ref_col],
            #         designation=row[desig_col]
            #     )
            #     db.session.add(product)
            
        except Exception as e:
            print(f"   ❌ Error reading {filepath}: {e}")
            return False
    
    return True

def load_products_from_dict(products_list):
    """Load products from Python dict list."""
    with app.app_context():
        for prod in products_list:
            existing = Product.query.filter_by(
                company=prod['company'],
                reference=prod['reference']
            ).first()
            
            if not existing:
                product = Product(**prod)
                db.session.add(product)
                print(f"✅ Added: {prod['company']} - {prod['reference']} - {prod['designation']}")
            else:
                print(f"⏭️  Skipped (exists): {prod['company']} - {prod['reference']}")
        
        db.session.commit()
        print(f"\n✓ Loaded {len(products_list)} products")

if __name__ == '__main__':
    print("=" * 70)
    print("PRODUCT LOADER")
    print("=" * 70)
    
    # Try to read Excel
    if load_products_from_excel():
        # If Excel reading works, implement the loading above
        pass
    else:
        # Fallback: user can provide data manually
        print("\n⚠️  Could not read Excel files automatically.")
        print("   Please provide Excel file columns and we'll auto-populate.")
        print("\n   Required info:")
        print("   - Column names in 'Liste des produits SRID.xls': [company, reference, designation]")
        print("   - Column names in 'Produit srid genetics.xls': [company, reference, designation]")
        print("   - Sheet names (if not Sheet1)")

