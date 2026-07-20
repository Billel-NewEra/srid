"""Import des opérations depuis « srid finalisé.xlsx » (feuille CHEQUES).

Vide au préalable : operations, audit_log, bons_commande (+ lignes),
commandes_logistique, frais_logistique. Conserve : users, products,
fournisseurs, client_labels, remettant_labels.

Usage : python import_excel_cheques.py
"""
import sys
from datetime import date, datetime

import openpyxl

from app import app, _normalize_bank_name
from models import (
    db, Operation, AuditLog, BonCommande, LigneCommande,
    CommandeLogistique, FraisLogistique,
)

EXCEL = 'srid finalisé.xlsx'
SHEET = 'CHEQUES'

# Colonnes (0-based) de la feuille CHEQUES
C_TYPE, C_RECEPT, C_REMETTANT, C_BANQUE, C_NUM = 1, 2, 3, 4, 5
C_CLIENT, C_ECHEANCE, C_MONTANT, C_TYPECHEQUE = 6, 7, 8, 9
C_STATUS, C_SOCIETE, C_REMARQUE = 10, 11, 12

TYPE_MAP = {
    'cheque': 'Chèque', 'chèque': 'Chèque',
    'transfert': 'Transfer', 'transfer': 'Transfer',
    'versement': 'Versement',
    'virement': 'Virement',
}
TYPE_CHEQUE_MAP = {
    'à échéance': 'À échéance',
    'a encaisser': 'À encaisser',
    'garantie': 'Garantie',
    'échu': 'À échéance',
    '/': None,
}
STATUT_MAP = {
    'encaissé': 'Encaissé',
    "en cours d'encaissement": 'En cours',
    'rejeté': 'Rejeté',
    'échu': 'Échu',
}
SOCIETE_MAP = {'srid': 'SRID', 'genetics': 'Genetics', '/': 'SRID'}


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _txt(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def main():
    wb = openpyxl.load_workbook(EXCEL, data_only=True, read_only=True)
    ws = wb[SHEET]
    rows = [
        r for r in ws.iter_rows(min_row=5, values_only=True)
        if any(c is not None and str(c).strip() != '' for c in r)
    ]
    print(f'Lignes non vides dans Excel : {len(rows)}')

    with app.app_context():
        # 1) Vidage des tables demandées (ordre = contraintes FK)
        counts_before = {
            'commandes_logistique': CommandeLogistique.query.count(),
            'frais_logistique': FraisLogistique.query.count(),
            'lignes_commande': LigneCommande.query.count(),
            'bons_commande': BonCommande.query.count(),
            'operations': Operation.query.count(),
            'audit_log': AuditLog.query.count(),
        }
        CommandeLogistique.query.delete()
        FraisLogistique.query.delete()
        LigneCommande.query.delete()
        BonCommande.query.delete()
        Operation.query.delete()
        AuditLog.query.delete()
        db.session.commit()
        print('Tables vidées :', counts_before)

        # 2) Insertion des opérations
        now = datetime.utcnow()
        skipped = []
        objets = []
        for i, r in enumerate(rows, start=5):
            raw_type = (str(r[C_TYPE]).strip().lower() if r[C_TYPE] else '')
            type_operation = TYPE_MAP.get(raw_type, 'Autre')

            client = _txt(r[C_CLIENT])
            if not client:
                skipped.append((i, 'client manquant'))
                continue
            client = client.upper()
            try:
                montant = abs(float(r[C_MONTANT]))
            except (TypeError, ValueError):
                # Montant absent ou invalide (ex. « / ») : on importe avec 0.
                montant = 0.0

            date_reception = _to_date(r[C_RECEPT])
            date_echeance = _to_date(r[C_ECHEANCE])

            if type_operation == 'Chèque':
                tc_raw = (str(r[C_TYPECHEQUE]).strip().lower()
                          if r[C_TYPECHEQUE] else '')
                type_detail = TYPE_CHEQUE_MAP.get(tc_raw)
                date_op = date_reception or date_echeance
                date_enc = date_echeance
            else:
                type_detail = None
                date_op = date_reception or date_echeance
                date_enc = None

            if date_op is None:
                skipped.append((i, 'aucune date'))
                continue

            statut_raw = (str(r[C_STATUS]).strip().lower()
                          if r[C_STATUS] else '')
            statut = STATUT_MAP.get(statut_raw, _txt(r[C_STATUS]) or 'Encaissé')

            soc_raw = (str(r[C_SOCIETE]).strip().lower()
                       if r[C_SOCIETE] else '')
            societe = SOCIETE_MAP.get(soc_raw, _txt(r[C_SOCIETE]) or 'SRID')

            objets.append(Operation(
                type_operation=type_operation,
                societe=societe,
                famille=None,
                date_operation=date_op,
                date_reception=date_reception,
                date_encaissement=date_enc,
                date_sortie=None,
                client=client,
                remettant=_txt(r[C_REMETTANT]),
                montant=montant,
                banque=_normalize_bank_name(r[C_BANQUE]),
                numero_piece=_txt(r[C_NUM]),
                statut=statut,
                type_detail=type_detail,
                entree=None,
                sortie=None,
                remarque=_txt(r[C_REMARQUE]),
                cree_par='import-excel',
                date_creation=now,
            ))

        db.session.bulk_save_objects(objets)
        db.session.commit()

        print(f'Opérations importées : {len(objets)}')
        if skipped:
            print(f'Lignes ignorées : {len(skipped)}')
            for row_no, reason in skipped[:20]:
                print(f'  - ligne {row_no}: {reason}')
        print('Total operations en base :', Operation.query.count())


if __name__ == '__main__':
    main()
