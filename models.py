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
    role = db.Column(db.String(20), nullable=False, default='saisisseur')  # boss, saisisseur

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


class AuditLog(db.Model):
    __tablename__ = 'audit_log'

    id = db.Column(db.Integer, primary_key=True)
    operation_id = db.Column(db.Integer, nullable=False)
    action = db.Column(db.String(20), nullable=False)  # création, modification, suppression
    utilisateur = db.Column(db.String(100))
    details = db.Column(db.Text)
    date_action = db.Column(db.DateTime, default=datetime.utcnow)
