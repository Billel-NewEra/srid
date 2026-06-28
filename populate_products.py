#!/usr/bin/env python3
"""
Populate Products table from Excel files.

Usage:
  1. Run this script: python populate_products.py
  2. It will detect and display Excel structure
  3. Confirm column names
  4. Load products into database
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# Add parent to path
sys.path.insert(0, str(Path(__file__).parent))

from app import app, db
from models import Product

def detect_excel_structure():
    """Detect and display Excel file structure."""
    try:
        import pandas as pd
    except ImportError:
        print("❌ pandas not installed. Run: pip install pandas")
        return None
    
    files_info = {}
    files_to_read = {
        'Liste des produits SRID.xls': 'SRID',
        'Produit srid genetics.xls': 'Genetics'
    }
    
    for filepath, company_name in files_to_read.items():
        if not os.path.exists(filepath):
            print(f"⚠️  File not found: {filepath}")
            continue
        
        try:
            print(f"\n{'='*70}")
            print(f"FILE: {filepath}")
            print(f"{'='*70}")
            
            # Try to read Excel
            xls = pd.ExcelFile(filepath)
            sheets = xls.sheet_names
            print(f"Sheets available: {sheets}")
            
            # Read first sheet
            sheet_name = sheets[0]
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            
            print(f"\n📊 Sheet: '{sheet_name}'")
            print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
            print(f"\n   Columns: {list(df.columns)}")
            print(f"\n   First 3 rows:")
            for i, row in df.head(3).iterrows():
                print(f"   Row {i+1}: {dict(row)}")
            
            files_info[filepath] = {
                'company': company_name,
                'sheet': sheet_name,
                'columns': list(df.columns),
                'shape': df.shape,
                'dataframe': df
            }
            
        except Exception as e:
            print(f"❌ Error reading {filepath}: {e}")
            return None
    
    return files_info

def ask_columns(files_info):
    """Ask user to specify column names."""
    config = {}
    
    for filepath, info in files_info.items():
        print(f"\n{'='*70}")
        print(f"Configure: {filepath}")
        print(f"{'='*70}")
        print(f"Available columns: {info['columns']}")
        
        # Ask for reference column
        while True:
            ref_col = input(f"\n  Reference column (exact name): ").strip()
            if ref_col in info['columns']:
                break
            print(f"  ❌ '{ref_col}' not found. Try again.")
        
        # Ask for designation column
        while True:
            desig_col = input(f"  Designation column (exact name): ").strip()
            if desig_col in info['columns']:
                break
            print(f"  ❌ '{desig_col}' not found. Try again.")
        
        config[filepath] = {
            'company': info['company'],
            'sheet': info['sheet'],
            'reference_col': ref_col,
            'designation_col': desig_col,
            'dataframe': info['dataframe']
        }
    
    return config

def load_products(config):
    """Load products into database."""
    with app.app_context():
        total_loaded = 0
        total_skipped = 0
        
        for filepath, cfg in config.items():
            print(f"\n{'='*70}")
            print(f"Loading: {filepath}")
            print(f"{'='*70}")
            
            df = cfg['dataframe']
            ref_col = cfg['reference_col']
            desig_col = cfg['designation_col']
            company = cfg['company']
            
            for idx, row in df.iterrows():
                try:
                    ref = str(row[ref_col]).strip() if pd.notna(row[ref_col]) else ''
                    desig = str(row[desig_col]).strip() if pd.notna(row[desig_col]) else ''
                    
                    if not ref or not desig:
                        total_skipped += 1
                        continue
                    
                    # Check if already exists
                    existing = Product.query.filter_by(
                        company=company,
                        reference=ref
                    ).first()
                    
                    if existing:
                        # Update if needed
                        if existing.designation != desig:
                            existing.designation = desig
                            db.session.commit()
                        total_skipped += 1
                    else:
                        product = Product(
                            company=company,
                            reference=ref,
                            designation=desig
                        )
                        db.session.add(product)
                        db.session.commit()
                        total_loaded += 1
                        
                        if total_loaded % 50 == 0:
                            print(f"  ✓ Loaded {total_loaded} products...")
                
                except Exception as e:
                    print(f"  ⚠️  Row {idx+1}: {e}")
                    total_skipped += 1
            
            print(f"  ✓ Loaded: {total_loaded} new products")
            print(f"  ⏭️  Skipped: {total_skipped} (duplicates/empty)")
        
        final_count = Product.query.count()
        print(f"\n{'='*70}")
        print(f"✓ COMPLETE")
        print(f"  Total products in database: {final_count}")
        print(f"{'='*70}")

if __name__ == '__main__':
    try:
        import pandas as pd
    except ImportError:
        print("❌ pandas not installed")
        print("   Install with: pip install pandas openpyxl")
        sys.exit(1)
    
    print("\n" + "="*70)
    print("PRODUCT LOADER - Excel to Database")
    print("="*70)
    
    # Detect Excel structure
    print("\n📖 Scanning Excel files...")
    files_info = detect_excel_structure()
    
    if not files_info:
        print("\n❌ Could not read Excel files")
        sys.exit(1)
    
    # Ask for column configuration
    print("\n🔧 Configure column mapping...")
    config = ask_columns(files_info)
    
    # Load products
    print("\n📦 Loading products into database...")
    load_products(config)
