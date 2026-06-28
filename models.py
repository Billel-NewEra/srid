from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    nom_complet = db.Column(db.String(150))
    role = db.Column(db.String(20), nullable=False, default='saisie')  # admin, saisie, consultation

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Operation(db.Model):
    __tablename__ = 'operations'

    id = db.Column(db.Integer, primary_key=True)

    # Type et catégorie
    type_operation = db.Column(db.String(20), nullable=False)  # Chèque, Virement, Versement
    societe = db.Column(db.String(100), nullable=False)  # ENT, Genetics, etc.
    famille = db.Column(db.String(100))

    # Dates
    date_operation = db.Column(db.Date, nullable=False)
    date_reception = db.Column(db.Date)
    date_encaissement = db.Column(db.Date)
    date_sortie = db.Column(db.Date)

    # Parties
    client = db.Column(db.String(200), nullable=False)
    remettant = db.Column(db.String(200))

    # Montant et banque
    montant = db.Column(db.Float, nullable=False)
    banque = db.Column(db.String(100))
    numero_piece = db.Column(db.String(50))  # N° chèque ou N° versement

    # Statut et type détail
    statut = db.Column(db.String(50), default='En attente')
    type_detail = db.Column(db.String(50))  # Garantie, Courant, etc.

    # Mouvements
    entree = db.Column(db.String(100))  # Personne qui a reçu
    sortie = db.Column(db.String(100))  # Personne qui a remis

    # Métadonnées
    remarque = db.Column(db.Text)
    cree_par = db.Column(db.String(100))
    date_creation = db.Column(db.DateTime, default=datetime.utcnow)
    date_modification = db.Column(db.DateTime, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'type_operation': self.type_operation,
            'societe': self.societe,
            'famille': self.famille,
            'date_operation': self.date_operation.isoformat() if self.date_operation else None,
            'date_reception': self.date_reception.isoformat() if self.date_reception else None,
            'client': self.client,
            'remettant': self.remettant,
            'montant': self.montant,
            'banque': self.banque,
            'numero_piece': self.numero_piece,
            'statut': self.statut,
            'type_detail': self.type_detail,
            'remarque': self.remarque,
            'cree_par': self.cree_par,
            'date_creation': self.date_creation.isoformat() if self.date_creation else None,
        }


class ClientLabel(db.Model):
    __tablename__ = 'client_labels'

    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(200), unique=True, nullable=False)
    actif = db.Column(db.Boolean, default=True, nullable=False)
    date_creation = db.Column(db.DateTime, default=datetime.utcnow)


class RemettantLabel(db.Model):
    __tablename__ = 'remettant_labels'

    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(200), unique=True, nullable=False)
    actif = db.Column(db.Boolean, default=True, nullable=False)
    date_creation = db.Column(db.DateTime, default=datetime.utcnow)


class Fournisseur(db.Model):
    __tablename__ = 'fournisseurs'

    id            = db.Column(db.Integer, primary_key=True)
    nom           = db.Column(db.String(200), nullable=False)
    societe       = db.Column(db.String(50), nullable=False, default='SRID')
    actif         = db.Column(db.Boolean, default=True, nullable=False)
    date_creation = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (db.UniqueConstraint('nom', 'societe', name='uq_fournisseur_nom_societe'),)


class CommandeLogistique(db.Model):
    __tablename__ = 'commandes_logistique'

    id                = db.Column(db.Integer, primary_key=True)
    bon_id            = db.Column(db.Integer, db.ForeignKey('bons_commande.id'))
    ref_log           = db.Column(db.String(20))          # LOG-0001 etc.
    societe           = db.Column(db.String(50), nullable=False)
    annee             = db.Column(db.String(4))
    date_d10          = db.Column(db.Date)
    date_arrivee      = db.Column(db.Date)
    fournisseur       = db.Column(db.String(200))
    produit           = db.Column(db.String(200))
    emballage         = db.Column(db.String(100))
    quantite          = db.Column(db.Float)
    tva               = db.Column(db.Float)
    montant_eur       = db.Column(db.Float)
    cours             = db.Column(db.Float)               # 4 décimales
    date_facture      = db.Column(db.Date)
    code_paiement     = db.Column(db.String(10))          # T / R / CAD
    nb_jours          = db.Column(db.Integer)
    date_echeance     = db.Column(db.Date)
    date_paiement     = db.Column(db.Date)
    date_valeur       = db.Column(db.Date)
    remarque          = db.Column(db.Text)
    cree_par          = db.Column(db.String(100))
    date_creation     = db.Column(db.DateTime, default=datetime.utcnow)
    date_modification = db.Column(db.DateTime, onupdate=datetime.utcnow)

    @property
    def montant_da(self):
        if self.montant_eur and self.cours:
            return round(self.montant_eur * self.cours, 2)
        return None

    @property
    def statut(self):
        """Machine à états pour le suivi logistique (ordre de priorité strict)."""
        today = date.today()
        if self.date_valeur:
            return 'PAYÉ'
        if self.date_paiement:
            return 'PAIEMENT EN COURS'
        if self.date_echeance:
            if self.date_echeance < today:
                return 'ÉCHU'
            days_left = (self.date_echeance - today).days
            if days_left <= 7:
                return 'ARRIVE À ÉCHÉANCE'
            return 'ÉCHÉANCE'
        if self.date_arrivee:
            return 'ARRIVÉ'
        if self.date_d10:
            return 'D10'
        return 'EN COURS'



class BonCommande(db.Model):
    __tablename__ = 'bons_commande'

    id                    = db.Column(db.Integer, primary_key=True)
    numero                = db.Column(db.String(50), unique=True)
    societe               = db.Column(db.String(50), nullable=False)
    fournisseur           = db.Column(db.String(200))
    statut                = db.Column(db.String(30), default='Brouillon')
    date_commande         = db.Column(db.Date, nullable=False)
    date_livraison_prevue = db.Column(db.Date)
    notes                 = db.Column(db.Text)
    cree_par              = db.Column(db.String(100))
    date_creation         = db.Column(db.DateTime, default=datetime.utcnow)
    lignes                = db.relationship('LigneCommande', backref='bon',
                                            lazy=True, cascade='all, delete-orphan')

    @property
    def total_eur(self):
        return sum((l.quantite or 0) * (l.prix_unitaire or 0) for l in self.lignes)

    def to_dict(self):
        return {
            'id':                    self.id,
            'numero':                self.numero,
            'societe':               self.societe,
            'fournisseur':           self.fournisseur,
            'statut':                self.statut,
            'date_commande':         self.date_commande.strftime('%Y-%m-%d') if self.date_commande else None,
            'date_livraison_prevue': self.date_livraison_prevue.strftime('%Y-%m-%d') if self.date_livraison_prevue else None,
            'notes':                 self.notes,
            'lignes': [
                {
                    'id':            l.id,
                    'reference':     l.reference,
                    'designation':   l.designation,
                    'quantite':      l.quantite,
                    'unite':         l.unite,
                    'prix_unitaire': l.prix_unitaire,
                }
                for l in self.lignes
            ],
        }


class LigneCommande(db.Model):
    __tablename__ = 'lignes_commande'

    id            = db.Column(db.Integer, primary_key=True)
    bon_id        = db.Column(db.Integer, db.ForeignKey('bons_commande.id'), nullable=False)
    reference     = db.Column(db.String(100))
    designation   = db.Column(db.String(300), nullable=False)
    quantite      = db.Column(db.Float, nullable=False, default=1)
    unite         = db.Column(db.String(50))
    prix_unitaire = db.Column(db.Float)


class AuditLog(db.Model):
    __tablename__ = 'audit_log'

    id = db.Column(db.Integer, primary_key=True)
    operation_id = db.Column(db.Integer, nullable=False)
    action = db.Column(db.String(20), nullable=False)  # création, modification, suppression
    utilisateur = db.Column(db.String(100))
    details = db.Column(db.Text)
    date_action = db.Column(db.DateTime, default=datetime.utcnow)


class Product(db.Model):
    __tablename__ = 'products'

    id           = db.Column(db.Integer, primary_key=True)
    company      = db.Column(db.String(100), nullable=False, index=True)  # 'SRID' ou 'Genetics'
    reference    = db.Column(db.String(100), nullable=False)
    designation  = db.Column(db.String(300), nullable=False)
    date_added   = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (db.UniqueConstraint('company', 'reference', name='uq_product_company_ref'),)

    def __repr__(self):
        return f'<Product {self.company}/{self.reference}>'

    def to_dict(self):
        return {
            'id': self.id,
            'company': self.company,
            'reference': self.reference,
            'designation': self.designation,
        }

