from flask import Flask, render_template, request, redirect, url_for, jsonify, session, flash, send_file, send_from_directory, make_response
from config import Config
from models import db, Operation, User, AuditLog, ClientLabel, RemettantLabel, CommandeLogistique, FraisLogistique, BonCommande, LigneCommande, Fournisseur, Product
from datetime import datetime, date, timedelta
from sqlalchemy import or_, func, extract, desc, case
from functools import wraps
import io
import re
import json
import os
import zlib
from openpyxl import Workbook, load_workbook

REF_PER_PAGE = 20
LOG_PER_PAGE = 20
BON_PER_PAGE = 20
PER_PAGE_OPTIONS = (20, 40, 60)
BON_STATUTS  = ['Brouillon', 'En attente', 'Approuvé', 'Envoyé', 'Reçu']


def _get_per_page(args, key, default_value):
    value = args.get(key, default_value, type=int)
    if value not in PER_PAGE_OPTIONS:
        return default_value
    return value


app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

@app.template_filter('currency')
def currency_filter(value, decimals=2):
    """Formate un montant en format français : 1\u2009234\u2009567,89 (ou sans centimes si decimals=0)"""
    try:
        formatted = "{:,.{d}f}".format(float(value), d=decimals)
        # virgule -> espace fine (milliers), point -> virgule (décimale)
        formatted = formatted.replace(",", "\u2009").replace(".", ",")
        return formatted
    except (TypeError, ValueError):
        return value


ROLE_ALIASES = {
    'boss': 'admin',
    'admin': 'admin',
    'saisisseur': 'saisie',
    'saisie': 'saisie',
    'consultation': 'consultation',
    'consultant': 'consultation',
    'viewer': 'consultation',
}

ROLE_LABELS = {
    'admin': 'Admin',
    'saisie': 'Saisie',
    'consultation': 'Consultation',
}

CHECK_TYPE_CHOICES = ['Garantie', 'À encaisser', 'À échéance']
STATUS_CHOICES = ['Encaissé', 'Rejeté', 'Échéance', 'En cours', 'Arrive à échéance', 'Échu']


# --- Décorateurs d'authentification ---


def _normalize_role(role):
    return ROLE_ALIASES.get((role or '').strip().lower(), 'consultation')


def _current_role():
    return _normalize_role(session.get('user_role', ''))


def _forbidden_response():
    if request.path.startswith('/api/') or request.headers.get('HX-Request'):
        return ('Accès refusé.', 403)
    flash('Accès refusé pour ce rôle.', 'error')
    return redirect(url_for('dashboard'))

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest' \
               or request.accept_mimetypes.best == 'application/json' \
               or request.headers.get('HX-Request'):
                return jsonify({'error': 'Non authentifié'}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def role_required(*allowed_roles):
    allowed = {_normalize_role(r) for r in allowed_roles}

    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('login'))
            if _current_role() not in allowed:
                return _forbidden_response()
            return f(*args, **kwargs)

        return decorated_function

    return decorator


# --- Contexte global templates ---

@app.context_processor
def inject_globals():
    from datetime import datetime
    role_key = _current_role()
    return {
        'current_user_id': session.get('user_id'),
        'current_user': session.get('user_nom', ''),
        'current_role': ROLE_LABELS.get(role_key, 'Consultation'),
        'current_role_key': role_key,
        'can_write': role_key in {'admin', 'saisie'},
        'can_delete': role_key == 'admin',
        'is_admin': role_key == 'admin',
        'can_manage_users': role_key == 'admin',
        'is_logged_in': 'user_id' in session,
        'now': datetime.now,
    }


def _normalize_bank_name(raw):
    """Normalise les variantes de banques vers un nom canonique."""
    if raw is None:
        return None
    source = str(raw).strip()
    if not source:
        return None

    cleaned = " ".join(source.split()).upper()
    compact = re.sub(r'[^A-Z0-9]', '', cleaned)
    compact = re.sub(r'\d+$', '', compact)
    if not compact:
        return None

    if compact.startswith('BADR') or compact.startswith('BAXDR'):
        return 'BADR'
    if compact.startswith('HOUSING') or compact in {'HB', 'HH', 'HBTF', 'HTBF'}:
        return 'HOUSING BANK'
    if compact.startswith('BNA'):
        return 'BNA'
    if compact.startswith('BDL'):
        return 'BDL'
    if compact.startswith('CPA'):
        return 'CPA'
    if compact.startswith('BEA'):
        return 'BEA'
    if compact.startswith('AGB') or compact.startswith('GBA'):
        return 'AGB'
    if compact.startswith('SGA') or compact == 'SG':
        return 'SGA'
    if compact.startswith('ALSALAMBANK') or compact.startswith('SALAMBANK'):
        return 'AL SALAM BANK'
    if compact.startswith('ALBARAKA') or compact.startswith('ELBARAKA'):
        return 'AL BARAKA'
    if compact.startswith('TRUST'):
        return 'TRUST BANK'
    if compact.startswith('FRANSABANK') or compact == 'FB':
        return 'FRANSABANK'
    if compact.startswith('ARABBANK'):
        return 'ARAB BANK'
    if compact.startswith('CNEP'):
        return 'CNEP'
    if compact.startswith('CCP'):
        return 'CCP'
    if compact.startswith('CITIBANK'):
        return 'CITIBANK'
    if compact.startswith('BNH'):
        return 'BNH'

    return cleaned


def _get_bank_suggestions():
    """Construit une liste unique de banques normalisées pour les formulaires."""
    rows = db.session.query(Operation.banque).filter(Operation.banque.isnot(None)).all()
    values = {_normalize_bank_name(v) for (v,) in rows}
    values.discard(None)

    defaults = {
        'BADR', 'HOUSING BANK', 'BNA', 'BDL', 'CPA', 'BEA',
        'AGB', 'SGA', 'AL BARAKA', 'AL SALAM BANK', 'TRUST BANK',
        'FRANSABANK', 'ARAB BANK', 'CNEP'
    }
    values.update(defaults)
    return sorted(values)


def _get_client_suggestions():
    # Depuis les opérations existantes
    rows = db.session.query(Operation.client).filter(Operation.client.isnot(None), Operation.client != '').all()
    values = {str(v).strip() for (v,) in rows if v}
    # Depuis le référentiel dédié
    labels = ClientLabel.query.filter_by(actif=True).all()
    values.update(l.nom.strip() for l in labels)
    return sorted(values)


def _get_remettant_suggestions():
    # Depuis les opérations existantes
    rows = db.session.query(Operation.remettant).filter(Operation.remettant.isnot(None), Operation.remettant != '').all()
    values = {str(v).strip() for (v,) in rows if v}
    # Depuis le référentiel dédié
    labels = RemettantLabel.query.filter_by(actif=True).all()
    values.update(l.nom.strip() for l in labels)
    return sorted(values)


def _normalize_type_operation(raw_type):
    s = (raw_type or '').strip().lower()
    if s in {'chèque', 'cheque'}:
        return 'Chèque'
    if s == 'virement':
        return 'Virement'
    if s == 'versement':
        return 'Versement'
    if s == 'transfer':
        return 'Transfer'
    if s == 'autre':
        return 'Autre'
    return 'Autre'


def _compute_statut(type_operation, type_cheque):
    if type_operation != 'Chèque':
        return 'Encaissé'
    if type_cheque == 'À échéance':
        return 'Échéance'
    return 'En cours'


def _normalize_legacy_statut(op):
    raw = (op.statut or '').strip()
    if raw in STATUS_CHOICES:
        return raw

    lowered = raw.lower()
    if 'annul' in lowered:
        return 'Rejeté'
    if 'rejet' in lowered:
        return 'Rejeté'
    if 'ech' in lowered:
        return 'Échéance'

    type_cheque = op.type_detail if op.type_operation == 'Chèque' else None
    return _compute_statut(op.type_operation, type_cheque)


def _auto_update_echeance_statuts():
    """Passe automatiquement les statuts des chèques à échéance en fonction de la date."""
    today = date.today()
    alert_date = today + timedelta(days=7)
    # Échu : date passée uniquement pour les statuts de workflow échéance.
    # Ne pas écraser un statut manuel (ex: Encaissé/Rejeté) choisi par l'admin.
    db.session.query(Operation).filter(
        Operation.type_operation == 'Chèque',
        Operation.type_detail == 'À échéance',
        Operation.date_encaissement < today,
        Operation.statut.in_(['Échéance', 'Arrive à échéance']),
    ).update({'statut': 'Échu'}, synchronize_session=False)
    # Arrive à échéance : dans les 7 prochains jours, statut encore Échéance
    db.session.query(Operation).filter(
        Operation.type_operation == 'Chèque',
        Operation.type_detail == 'À échéance',
        Operation.date_encaissement >= today,
        Operation.date_encaissement <= alert_date,
        Operation.statut == 'Échéance',
    ).update({'statut': 'Arrive à échéance'}, synchronize_session=False)
    db.session.commit()


def _get_echeance_notifications(limit=8):
    """Retourne les alertes d'échéance globales pour affichage sans filtre."""
    today = date.today()
    alert_date = today + timedelta(days=7)
    workflow_statuses = ['Échéance', 'Arrive à échéance', 'Échu']
    base_query = Operation.query.filter(
        Operation.type_operation == 'Chèque',
        Operation.type_detail == 'À échéance',
        Operation.date_encaissement.isnot(None),
        Operation.statut.in_(workflow_statuses),
    )

    overdue_count = base_query.filter(Operation.date_encaissement < today).count()
    upcoming_count = base_query.filter(
        Operation.date_encaissement >= today,
        Operation.date_encaissement <= alert_date,
    ).count()

    critical_operations = base_query.filter(
        or_(
            Operation.date_encaissement < today,
            (
                (Operation.date_encaissement >= today) &
                (Operation.date_encaissement <= alert_date)
            ),
        )
    ).order_by(Operation.date_encaissement.asc()).limit(limit).all()

    return {
        'today': today,
        'overdue_count': overdue_count,
        'upcoming_count': upcoming_count,
        'total_alerts': overdue_count + upcoming_count,
        'critical_operations': critical_operations,
    }


def _get_recent_rejections(limit=8):
    """Retourne les rejets récents (dernières 24h) depuis l'audit log."""
    now = datetime.utcnow()
    yesterday = now - timedelta(hours=24)
    
    # Chercher les audits où les details contiennent "Rejeté" et la date_action est récente
    rejections = AuditLog.query.filter(
        AuditLog.date_action >= yesterday,
        AuditLog.details.ilike('%Rejeté%'),
    ).order_by(desc(AuditLog.date_action)).limit(limit).all()
    
    # Récupérer tous les operation_ids en une seule requête
    op_ids = list({a.operation_id for a in rejections})
    if op_ids:
        ops_map = {op.id: op for op in Operation.query.filter(
            Operation.id.in_(op_ids), Operation.statut == 'Rejeté'
        ).all()}
    else:
        ops_map = {}

    rejection_ops = []
    seen_ops = set()
    for audit in rejections:
        if audit.operation_id not in seen_ops and audit.operation_id in ops_map:
            rejection_ops.append({
                'operation': ops_map[audit.operation_id],
                'rejected_at': audit.date_action,
                'rejected_by': audit.utilisateur,
            })
            seen_ops.add(audit.operation_id)
    
    return {
        'rejections': rejection_ops,
        'total': len(rejection_ops),
    }


def _admin_users_count():
    return User.query.filter_by(role='admin').count()


# --- Routes Auth ---

@app.route('/sw.js')
def service_worker():
    return send_from_directory('static', 'sw.js', mimetype='application/javascript')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            session.permanent = True
            session['user_id'] = user.id
            session['user_nom'] = user.nom_complet or user.username
            session['user_role'] = _normalize_role(user.role)
            flash(f'Bienvenue, {user.nom_complet or user.username} !', 'success')
            return redirect(url_for('dashboard'))
        flash('Identifiants incorrects.', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/utilisateurs', methods=['GET', 'POST'])
@role_required('admin')
def manage_users():
    role_choices = ['admin', 'saisie', 'consultation']

    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        nom_complet = request.form.get('nom_complet', '').strip()
        password = request.form.get('password', '')
        role = _normalize_role(request.form.get('role', 'consultation'))

        if not username or len(username) < 3:
            flash('Nom utilisateur invalide (min 3 caractères).', 'error')
            return redirect(url_for('manage_users'))
        if role not in role_choices:
            flash('Rôle invalide.', 'error')
            return redirect(url_for('manage_users'))
        if len(password) < 6:
            flash('Mot de passe trop court (min 6 caractères).', 'error')
            return redirect(url_for('manage_users'))
        if User.query.filter_by(username=username).first():
            flash('Ce nom utilisateur existe déjà.', 'error')
            return redirect(url_for('manage_users'))

        user = User(
            username=username,
            nom_complet=nom_complet or username,
            role=role,
        )
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        flash(f'Utilisateur {username} créé avec rôle {ROLE_LABELS.get(role, role)}.', 'success')
        return redirect(url_for('manage_users'))

    users = User.query.order_by(User.username.asc()).all()
    return render_template('users.html', users=users, role_labels=ROLE_LABELS)


@app.route('/utilisateurs/<int:user_id>/role', methods=['POST'])
@role_required('admin')
def update_user_role(user_id):
    user = User.query.get_or_404(user_id)
    new_role = _normalize_role(request.form.get('role', 'consultation'))
    if new_role not in {'admin', 'saisie', 'consultation'}:
        flash('Rôle invalide.', 'error')
        return redirect(url_for('manage_users'))

    if user.id == session.get('user_id'):
        flash('Vous ne pouvez pas modifier votre propre rôle.', 'error')
        return redirect(url_for('manage_users'))

    if user.role == 'admin' and new_role != 'admin' and _admin_users_count() <= 1:
        flash('Impossible de retirer le dernier administrateur.', 'error')
        return redirect(url_for('manage_users'))

    user.role = new_role
    db.session.commit()
    flash(
        f'Rôle mis à jour pour {user.username}. Le nouvel accès sera effectif après reconnexion de cet utilisateur.',
        'success'
    )
    return redirect(url_for('manage_users'))


@app.route('/utilisateurs/<int:user_id>/password', methods=['POST'])
@role_required('admin')
def update_user_password(user_id):
    user = User.query.get_or_404(user_id)
    new_password = request.form.get('password', '')
    if len(new_password) < 6:
        flash('Mot de passe trop court (min 6 caractères).', 'error')
        return redirect(url_for('manage_users'))

    user.set_password(new_password)
    db.session.commit()
    flash(
        f'Mot de passe mis à jour pour {user.username}. Le changement sera pris en compte après reconnexion.',
        'success'
    )
    return redirect(url_for('manage_users'))


@app.route('/utilisateurs/<int:user_id>/delete', methods=['POST'])
@role_required('admin')
def delete_user(user_id):
    user = User.query.get_or_404(user_id)

    if user.id == session.get('user_id'):
        flash('Vous ne pouvez pas supprimer votre propre compte.', 'error')
        return redirect(url_for('manage_users'))

    if user.role == 'admin' and _admin_users_count() <= 1:
        flash('Impossible de supprimer le dernier administrateur.', 'error')
        return redirect(url_for('manage_users'))

    username = user.username
    db.session.delete(user)
    db.session.commit()
    flash(f'Utilisateur {username} supprimé.', 'success')
    return redirect(url_for('manage_users'))


# --- Dashboard ---

def _dashboard_available_years():
    years_raw = db.session.query(
        extract('year', Operation.date_operation)
    ).distinct().order_by(extract('year', Operation.date_operation).desc()).all()
    return [int(y[0]) for y in years_raw if y[0]]


def _dashboard_kpi_block(year, month):
    """Calcule les montants (Vue d'ensemble + statuts) pour une année/mois donnés.

    Utilisé par le dashboard initial et par les partiels filtrés indépendamment
    (Vue d'ensemble, carte Encaissé, carte En cours d'encaissement).
    """
    _excluded_types = ['Autre', 'Transfer']
    _excluded_statuts = ['Rejeté']

    base_filter = []
    if year:
        base_filter.append(extract('year', Operation.date_operation) == year)
    if month:
        base_filter.append(extract('month', Operation.date_operation) == month)

    # KPIs globaux (Vue d'ensemble)
    kpi_lookup = {s: float(m) for s, m in db.session.query(
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        ~Operation.type_operation.in_(_excluded_types),
        ~Operation.statut.in_(_excluded_statuts),
        *base_filter
    ).group_by(Operation.societe).all()}
    total_montant = sum(kpi_lookup.values())

    # Statuts avec ventilation par société
    status_labels = STATUS_CHOICES
    statuts_info = {s: {'count': 0, 'montant': 0.0, 'srid_count': 0, 'srid_montant': 0.0,
                        'genetics_count': 0, 'genetics_montant': 0.0} for s in status_labels}
    for statut, societe, count, montant in db.session.query(
        Operation.statut,
        Operation.societe,
        func.count(Operation.id),
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(*base_filter).group_by(Operation.statut, Operation.societe).all():
        if statut in statuts_info:
            statuts_info[statut]['count'] += count
            statuts_info[statut]['montant'] += float(montant)
            if societe == 'SRID':
                statuts_info[statut]['srid_count'] += count
                statuts_info[statut]['srid_montant'] += float(montant)
            elif societe == 'Genetics':
                statuts_info[statut]['genetics_count'] += count
                statuts_info[statut]['genetics_montant'] += float(montant)

    # La carte "En cours d'encaissement" regroupe Échéance, Arrive à échéance et Échu.
    for _s in ['Échéance', 'Arrive à échéance', 'Échu']:
        if _s in statuts_info:
            for k in ('count', 'montant', 'srid_count', 'srid_montant',
                      'genetics_count', 'genetics_montant'):
                statuts_info['En cours'][k] += statuts_info[_s][k]

    return {
        'total_montant': float(total_montant),
        'montant_srid': kpi_lookup.get('SRID', 0.0),
        'montant_genetics': kpi_lookup.get('Genetics', 0.0),
        'statuts_info': statuts_info,
    }


@app.route('/')
@login_required
def dashboard():
    current_year = date.today().year
    previous_year = current_year - 1
    today = date.today()
    status_labels = STATUS_CHOICES

    # Années disponibles
    years_raw = db.session.query(
        extract('year', Operation.date_operation)
    ).distinct().order_by(extract('year', Operation.date_operation).desc()).all()
    available_years = [int(y[0]) for y in years_raw if y[0]]

    # --- KPIs globaux (1 query) ---
    _excluded_types = ['Autre', 'Transfer']
    _excluded_statuts = ['Rejeté']
    kpi_raw = db.session.query(
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        ~Operation.type_operation.in_(_excluded_types),
        ~Operation.statut.in_(_excluded_statuts),
        extract('year', Operation.date_operation) == current_year
    ).group_by(Operation.societe).all()
    kpi_lookup = {s: float(m) for s, m in kpi_raw}
    total_montant = sum(kpi_lookup.values())
    montant_srid = kpi_lookup.get('SRID', 0.0)
    montant_genetics = kpi_lookup.get('Genetics', 0.0)

    # --- Statuts avec montants par société (1 query) ---
    statuts_raw = db.session.query(
        Operation.statut,
        Operation.societe,
        func.count(Operation.id),
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        extract('year', Operation.date_operation) == current_year
    ).group_by(Operation.statut, Operation.societe).all()
    statuts_info = {s: {'count': 0, 'montant': 0.0, 'srid_count': 0, 'srid_montant': 0.0,
                        'genetics_count': 0, 'genetics_montant': 0.0} for s in status_labels}
    for statut, societe, count, montant in statuts_raw:
        if statut in statuts_info:
            statuts_info[statut]['count'] += count
            statuts_info[statut]['montant'] += float(montant)
            if societe == 'SRID':
                statuts_info[statut]['srid_count'] += count
                statuts_info[statut]['srid_montant'] += float(montant)
            elif societe == 'Genetics':
                statuts_info[statut]['genetics_count'] += count
                statuts_info[statut]['genetics_montant'] += float(montant)

    # Point 7 : la carte "En cours" regroupe aussi Échéance, Arrive à échéance et Échu
    _encours_merge = ['Échéance', 'Arrive à échéance', 'Échu']
    for _s in _encours_merge:
        if _s in statuts_info:
            statuts_info['En cours']['count']           += statuts_info[_s]['count']
            statuts_info['En cours']['montant']         += statuts_info[_s]['montant']
            statuts_info['En cours']['srid_count']      += statuts_info[_s]['srid_count']
            statuts_info['En cours']['srid_montant']    += statuts_info[_s]['srid_montant']
            statuts_info['En cours']['genetics_count']  += statuts_info[_s]['genetics_count']
            statuts_info['En cours']['genetics_montant']+= statuts_info[_s]['genetics_montant']

    # --- Types avec montants par société (1 query) ---
    types_raw = db.session.query(
        Operation.type_operation,
        Operation.societe,
        func.count(Operation.id),
        func.coalesce(func.sum(Operation.montant), 0)
    ).group_by(Operation.type_operation, Operation.societe).all()
    type_list = ['Chèque', 'Virement', 'Versement', 'Transfer', 'Autre']
    types_info = {t: {'total': {'count': 0, 'montant': 0.0},
                      'SRID': {'count': 0, 'montant': 0.0},
                      'Genetics': {'count': 0, 'montant': 0.0}} for t in type_list}
    for type_op, societe, count, montant in types_raw:
        if type_op in types_info:
            types_info[type_op]['total']['count'] += count
            types_info[type_op]['total']['montant'] += float(montant)
            if societe in ('SRID', 'Genetics'):
                types_info[type_op][societe]['count'] += count
                types_info[type_op][societe]['montant'] += float(montant)

    # --- Top 5 clients (année en cours) ---
    top_clients = db.session.query(
        Operation.client,
        func.sum(Operation.montant).label('total')
    ).filter(
        extract('year', Operation.date_operation) == current_year,
        Operation.client.isnot(None),
        Operation.client != '',
        ~func.lower(Operation.client).like('%srid%'),
        ~func.lower(Operation.client).like('%genetics%')
    ).group_by(Operation.client).order_by(desc('total')).limit(5).all()

    # --- Dernières opérations ---
    dernieres = Operation.query.order_by(Operation.date_operation.desc()).limit(10).all()

    # --- Données mensuelles année courante + par société (1 query) ---
    monthly_curr_raw = db.session.query(
        extract('month', Operation.date_operation).label('month'),
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        extract('year', Operation.date_operation) == current_year
    ).group_by('month', Operation.societe).all()
    monthly_curr_lookup = {}
    for month, societe, montant in monthly_curr_raw:
        m = int(month)
        if m not in monthly_curr_lookup:
            monthly_curr_lookup[m] = {'total': 0.0, 'SRID': 0.0, 'Genetics': 0.0}
        monthly_curr_lookup[m]['total'] += float(montant)
        if societe in ('SRID', 'Genetics'):
            monthly_curr_lookup[m][societe] += float(montant)
    monthly_data = [monthly_curr_lookup.get(m, {}).get('total', 0.0) for m in range(1, 13)]
    monthly_srid = [monthly_curr_lookup.get(m, {}).get('SRID', 0.0) for m in range(1, 13)]
    monthly_genetics = [monthly_curr_lookup.get(m, {}).get('Genetics', 0.0) for m in range(1, 13)]

    # --- Données mensuelles année précédente (1 query) ---
    monthly_prev_raw = db.session.query(
        extract('month', Operation.date_operation).label('month'),
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        extract('year', Operation.date_operation) == previous_year
    ).group_by('month').all()
    monthly_prev_lookup = {int(m): float(mt) for m, mt in monthly_prev_raw}
    monthly_data_prev = [monthly_prev_lookup.get(m, 0.0) for m in range(1, 13)]

    # --- Statuts par mois année courante (1 query) ---
    monthly_statuts_raw = db.session.query(
        extract('month', Operation.date_operation).label('month'),
        Operation.statut,
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        extract('year', Operation.date_operation) == current_year
    ).group_by('month', Operation.statut, Operation.societe).all()
    monthly_statuts_views = {
        'total': {s: [0.0] * 12 for s in status_labels},
        'srid':  {s: [0.0] * 12 for s in status_labels},
        'genetics': {s: [0.0] * 12 for s in status_labels},
    }
    for month, statut, societe, montant in monthly_statuts_raw:
        m = int(month) - 1
        if statut in status_labels:
            monthly_statuts_views['total'][statut][m] += float(montant)
            if societe == 'SRID':
                monthly_statuts_views['srid'][statut][m] += float(montant)
            elif societe == 'Genetics':
                monthly_statuts_views['genetics'][statut][m] += float(montant)
    monthly_statuts = monthly_statuts_views['total']

    # --- Activité 30 jours (1 query) ---
    thirty_days_ago = today - timedelta(days=29)
    daily_raw = db.session.query(
        func.date(Operation.date_operation).label('day'),
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        func.date(Operation.date_operation) >= thirty_days_ago,
        func.date(Operation.date_operation) <= today
    ).group_by('day', Operation.societe).all()
    daily_lookup = {}
    for day, societe, montant in daily_raw:
        d_str = str(day)
        if d_str not in daily_lookup:
            daily_lookup[d_str] = {'total': 0.0, 'SRID': 0.0, 'Genetics': 0.0}
        daily_lookup[d_str]['total'] += float(montant)
        if societe in ('SRID', 'Genetics'):
            daily_lookup[d_str][societe] += float(montant)
    daily_labels, daily_total, daily_srid, daily_genetics = [], [], [], []
    for i in range(29, -1, -1):
        d = today - timedelta(days=i)
        entry = daily_lookup.get(d.isoformat(), {'total': 0.0, 'SRID': 0.0, 'Genetics': 0.0})
        daily_labels.append(d.strftime('%d/%m'))
        daily_total.append(entry['total'])
        daily_srid.append(entry['SRID'])
        daily_genetics.append(entry['Genetics'])

    return render_template('dashboard.html',
                           total_montant=float(total_montant),
                           montant_srid=float(montant_srid), montant_genetics=float(montant_genetics),
                           statuts_info=statuts_info, types_info=types_info,
                           top_clients=top_clients, dernieres=dernieres,
                           monthly_data=monthly_data, monthly_data_prev=monthly_data_prev,
                           monthly_srid=monthly_srid, monthly_genetics=monthly_genetics,
                           monthly_statuts=monthly_statuts,
                           monthly_statuts_views=monthly_statuts_views,
                           daily_labels=daily_labels, daily_total=daily_total,
                           daily_srid=daily_srid, daily_genetics=daily_genetics,
                           today=date.today().strftime('%d/%m/%Y'),
                           year=current_year, previous_year=previous_year,
                           available_years=available_years,
                           selected_year=current_year, selected_month=0)


# --- API Dashboard (filtres HTMX/JS) ---

@app.route('/api/dashboard/monthly')
@login_required
def api_dashboard_monthly():
    year = request.args.get('year', date.today().year, type=int)
    prev_year = year - 1

    # 1 query pour l'année courante, 1 pour la précédente
    rows = db.session.query(
        extract('month', Operation.date_operation),
        func.sum(Operation.montant)
    ).filter(
        extract('year', Operation.date_operation) == year
    ).group_by(extract('month', Operation.date_operation)).all()
    monthly_map = {int(m): float(t) for m, t in rows}

    rows_prev = db.session.query(
        extract('month', Operation.date_operation),
        func.sum(Operation.montant)
    ).filter(
        extract('year', Operation.date_operation) == prev_year
    ).group_by(extract('month', Operation.date_operation)).all()
    monthly_map_prev = {int(m): float(t) for m, t in rows_prev}

    monthly_data = [monthly_map.get(m, 0) for m in range(1, 13)]
    monthly_data_prev = [monthly_map_prev.get(m, 0) for m in range(1, 13)]
    return jsonify({'year': year, 'prev_year': prev_year, 'data': monthly_data, 'data_prev': monthly_data_prev})


@app.route('/api/dashboard/kpis')
@login_required
def api_dashboard_kpis():
    """HTMX partial: KPIs filtrés par année et/ou mois."""
    year = request.args.get('year', date.today().year, type=int)
    month = request.args.get('month', 0, type=int)

    _excluded_types = ['Autre', 'Transfer']
    _excluded_statuts = ['Rejeté']

    # --- Filtre de base ---
    base_filter = []
    if year:
        base_filter.append(extract('year', Operation.date_operation) == year)
    if month:
        base_filter.append(extract('month', Operation.date_operation) == month)

    # --- KPIs globaux ---
    kpi_q = db.session.query(
        Operation.societe,
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(
        ~Operation.type_operation.in_(_excluded_types),
        ~Operation.statut.in_(_excluded_statuts),
        *base_filter
    ).group_by(Operation.societe)
    kpi_lookup = {s: float(m) for s, m in kpi_q.all()}
    total_montant = sum(kpi_lookup.values())
    montant_srid = kpi_lookup.get('SRID', 0.0)
    montant_genetics = kpi_lookup.get('Genetics', 0.0)

    # --- Statuts avec montants par société ---
    status_labels = STATUS_CHOICES
    statuts_raw = db.session.query(
        Operation.statut,
        Operation.societe,
        func.count(Operation.id),
        func.coalesce(func.sum(Operation.montant), 0)
    ).filter(*base_filter).group_by(Operation.statut, Operation.societe).all()

    statuts_info = {s: {'count': 0, 'montant': 0.0, 'srid_count': 0, 'srid_montant': 0.0,
                        'genetics_count': 0, 'genetics_montant': 0.0} for s in status_labels}
    for statut, societe, count, montant in statuts_raw:
        if statut in statuts_info:
            statuts_info[statut]['count'] += count
            statuts_info[statut]['montant'] += float(montant)
            if societe == 'SRID':
                statuts_info[statut]['srid_count'] += count
                statuts_info[statut]['srid_montant'] += float(montant)
            elif societe == 'Genetics':
                statuts_info[statut]['genetics_count'] += count
                statuts_info[statut]['genetics_montant'] += float(montant)

    _encours_merge = ['Échéance', 'Arrive à échéance', 'Échu']
    for _s in _encours_merge:
        if _s in statuts_info:
            statuts_info['En cours']['count'] += statuts_info[_s]['count']
            statuts_info['En cours']['montant'] += statuts_info[_s]['montant']
            statuts_info['En cours']['srid_count'] += statuts_info[_s]['srid_count']
            statuts_info['En cours']['srid_montant'] += statuts_info[_s]['srid_montant']
            statuts_info['En cours']['genetics_count'] += statuts_info[_s]['genetics_count']
            statuts_info['En cours']['genetics_montant'] += statuts_info[_s]['genetics_montant']

    # Années disponibles pour les filtres
    years_raw = db.session.query(
        extract('year', Operation.date_operation)
    ).distinct().order_by(extract('year', Operation.date_operation).desc()).all()
    available_years = [int(y[0]) for y in years_raw if y[0]]

    return render_template('partials/dashboard_kpis.html',
                           total_montant=float(total_montant),
                           montant_srid=float(montant_srid),
                           montant_genetics=float(montant_genetics),
                           statuts_info=statuts_info,
                           available_years=available_years,
                           selected_year=year,
                           selected_month=month)


@app.route('/api/dashboard/overview')
@login_required
def api_dashboard_overview():
    """HTMX partial: Vue d'ensemble filtrée (année/mois indépendants)."""
    year = request.args.get('year', date.today().year, type=int)
    month = request.args.get('month', 0, type=int)
    block = _dashboard_kpi_block(year, month)
    return render_template('partials/dashboard_overview.html',
                           total_montant=block['total_montant'],
                           montant_srid=block['montant_srid'],
                           montant_genetics=block['montant_genetics'],
                           available_years=_dashboard_available_years(),
                           selected_year=year,
                           selected_month=month)


@app.route('/api/dashboard/encaisse')
@login_required
def api_dashboard_encaisse():
    """HTMX partial: carte Encaissé filtrée (année/mois indépendants)."""
    year = request.args.get('year', date.today().year, type=int)
    month = request.args.get('month', 0, type=int)
    block = _dashboard_kpi_block(year, month)
    return render_template('partials/dashboard_encaisse.html',
                           statuts_info=block['statuts_info'],
                           available_years=_dashboard_available_years(),
                           selected_year=year,
                           selected_month=month)


@app.route('/api/dashboard/encours')
@login_required
def api_dashboard_encours():
    """HTMX partial: carte En cours d'encaissement filtrée (année/mois indépendants)."""
    year = request.args.get('year', date.today().year, type=int)
    month = request.args.get('month', 0, type=int)
    block = _dashboard_kpi_block(year, month)
    return render_template('partials/dashboard_encours.html',
                           statuts_info=block['statuts_info'],
                           available_years=_dashboard_available_years(),
                           selected_year=year,
                           selected_month=month)


@app.route('/api/dashboard/societes')
@login_required
def api_dashboard_societes():
    year = request.args.get('year', date.today().year, type=int)

    rows = db.session.query(
        extract('month', Operation.date_operation),
        Operation.societe,
        func.sum(Operation.montant)
    ).filter(
        extract('year', Operation.date_operation) == year
    ).group_by(extract('month', Operation.date_operation), Operation.societe).all()

    monthly_srid = [0.0] * 12
    monthly_genetics = [0.0] * 12
    for m, soc, total in rows:
        idx = int(m) - 1
        if soc == 'SRID':
            monthly_srid[idx] = float(total)
        elif soc == 'Genetics':
            monthly_genetics[idx] = float(total)
    return jsonify({'year': year, 'srid': monthly_srid, 'genetics': monthly_genetics})


@app.route('/api/dashboard/statuts-monthly')
@login_required
def api_dashboard_statuts_monthly():
    year = request.args.get('year', date.today().year, type=int)
    status_labels = STATUS_CHOICES

    # 1 seule requête groupée au lieu de 216
    rows = db.session.query(
        extract('month', Operation.date_operation),
        Operation.statut,
        Operation.societe,
        func.sum(Operation.montant)
    ).filter(
        extract('year', Operation.date_operation) == year
    ).group_by(
        extract('month', Operation.date_operation), Operation.statut, Operation.societe
    ).all()

    monthly_statuts_views = {
        'total': {s: [0.0] * 12 for s in status_labels},
        'srid': {s: [0.0] * 12 for s in status_labels},
        'genetics': {s: [0.0] * 12 for s in status_labels},
    }

    for m, statut, societe, total in rows:
        idx = int(m) - 1
        amt = float(total or 0)
        if statut in monthly_statuts_views['total']:
            monthly_statuts_views['total'][statut][idx] += amt
            if societe == 'SRID':
                monthly_statuts_views['srid'][statut][idx] += amt
            elif societe == 'Genetics':
                monthly_statuts_views['genetics'][statut][idx] += amt

    return jsonify({'year': year, 'data': monthly_statuts_views['total'], 'views': monthly_statuts_views})


@app.route('/api/dashboard/top-clients')
@login_required
def api_dashboard_top_clients():
    year = request.args.get('year', date.today().year, type=int)
    month = request.args.get('month', 0, type=int)
    filters = [
        extract('year', Operation.date_operation) == year,
        Operation.client.isnot(None),
        Operation.client != '',
        ~func.lower(Operation.client).like('%srid%'),
        ~func.lower(Operation.client).like('%genetics%')
    ]
    if month > 0:
        filters.append(extract('month', Operation.date_operation) == month)
    top_clients = db.session.query(
        Operation.client,
        func.sum(Operation.montant).label('total')
    ).filter(*filters).group_by(Operation.client).order_by(desc('total')).limit(5).all()
    max_total = float(top_clients[0][1]) if top_clients else 1
    html = ''
    for i, (client_name, total) in enumerate(top_clients, 1):
        pct = float(total) / max_total * 100
        html += f'''<div class="flex items-center gap-3">
            <div class="badge badge-sm badge-primary font-bold w-6 h-6">{i}</div>
            <div class="flex-1">
                <div class="flex justify-between items-center mb-1">
                    <span class="text-sm font-medium truncate max-w-[180px]">{client_name or 'N/A'}</span>
                    <span class="text-sm font-bold">{total:,.0f} DA</span>
                </div>
                <progress class="progress progress-primary w-full h-2" value="{pct}" max="100"></progress>
            </div>
        </div>'''
    if not top_clients:
        html = '<p class="text-center opacity-50 py-4">Aucune donnée</p>'
    return html


@app.route('/api/dashboard/types')
@login_required
def api_dashboard_types():
    societe = request.args.get('societe', 'all')
    result = {}
    for type_op in ['Chèque', 'Virement', 'Versement', 'Transfer', 'Autre']:
        q = db.session.query(func.sum(Operation.montant)).filter_by(type_operation=type_op)
        if societe != 'all':
            q = q.filter_by(societe=societe)
        result[type_op] = float(q.scalar() or 0)
    return jsonify(result)


# --- Saisie ---

@app.route('/api/operation/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_operation_add():
    type_operation = _normalize_type_operation(request.form.get('type_operation'))
    type_cheque = (request.form.get('type_cheque') or request.form.get('type_detail')) if type_operation == 'Chèque' else None
    if type_cheque not in CHECK_TYPE_CHOICES:
        type_cheque = None

    date_operation = None
    date_reception = None
    date_sortie = None
    date_echeance = None
    if type_operation == 'Chèque':
        date_reception = _parse_date(request.form.get('date_reception'))
        date_sortie = _parse_date(request.form.get('date_sortie'))
        if type_cheque == 'À échéance':
            date_echeance = _parse_date(request.form.get('date_echeance') or request.form.get('date_encaissement'))
        date_operation = date_sortie
    else:
        date_operation = _parse_date(request.form.get('date_operation'))

    statut_auto = _compute_statut(type_operation, type_cheque)

    op = Operation(
        type_operation=type_operation,
        societe=request.form.get('societe'),
        famille=None,
        date_operation=date_operation,
        date_reception=date_reception,
        date_encaissement=date_echeance,
        date_sortie=date_sortie,
        client=request.form.get('client'),
        remettant=request.form.get('remettant_commercial') or request.form.get('remettant') or None,
        montant=abs(float(request.form.get('montant', 0))),
        banque=_normalize_bank_name(request.form.get('banque')),
        numero_piece=request.form.get('numero_piece') or None,
        statut=statut_auto,
        type_detail=type_cheque,
        entree=request.form.get('entree') or None,
        sortie=request.form.get('sortie') or None,
        remarque=request.form.get('remarque') or None,
        cree_par=session.get('user_nom', ''),
    )
    db.session.add(op)
    db.session.commit()
    _log_audit(op.id, 'création', f"{op.type_operation} - {op.client} - {op.montant}")

    if request.headers.get('HX-Request'):
        return render_template('partials/success_message.html', operation=op)
    flash('Opération enregistrée !', 'success')
    return redirect(url_for('operations'))


# --- Modification ---

@app.route('/edit/<int:op_id>', methods=['GET', 'POST'])
@role_required('admin', 'saisie')
def edit_operation(op_id):
    op = Operation.query.get_or_404(op_id)
    if request.method == 'POST':
        type_operation = _normalize_type_operation(request.form.get('type_operation'))
        type_cheque = (request.form.get('type_cheque') or request.form.get('type_detail')) if type_operation == 'Chèque' else None
        if type_cheque not in CHECK_TYPE_CHOICES:
            type_cheque = None

        date_operation = None
        date_reception = None
        date_sortie = None
        date_echeance = None
        if type_operation == 'Chèque':
            date_reception = _parse_date(request.form.get('date_reception'))
            date_sortie = _parse_date(request.form.get('date_sortie'))
            if type_cheque == 'À échéance':
                date_echeance = _parse_date(request.form.get('date_echeance') or request.form.get('date_encaissement'))
            date_operation = date_sortie
        else:
            date_operation = _parse_date(request.form.get('date_operation'))

        op.type_operation = type_operation
        op.societe = request.form.get('societe')
        op.famille = None
        op.date_operation = date_operation
        op.date_reception = date_reception
        op.date_encaissement = date_echeance
        op.date_sortie = date_sortie
        op.client = request.form.get('client')
        op.remettant = request.form.get('remettant_commercial') or request.form.get('remettant') or None
        op.montant = abs(float(request.form.get('montant', 0)))
        op.banque = _normalize_bank_name(request.form.get('banque'))
        op.numero_piece = request.form.get('numero_piece') or None
        op.statut = _compute_statut(type_operation, type_cheque)
        op.type_detail = type_cheque
        if 'entree' in request.form:
            op.entree = request.form.get('entree') or None
        if 'sortie' in request.form:
            op.sortie = request.form.get('sortie') or None
        op.remarque = request.form.get('remarque') or None
        op.date_modification = datetime.utcnow()
        db.session.commit()
        _log_audit(op.id, 'modification', f"Modifié par {session.get('user_nom', '')}")

        if request.headers.get('HX-Request'):
            return render_template('partials/success_message.html', operation=op, action='modifiée')
        flash('Opération modifiée !', 'success')
        return redirect(url_for('operations'))

    if request.headers.get('HX-Request'):
        return render_template(
            'partials/edit_form.html',
            operation=op,
            bank_options=_get_bank_suggestions(),
            client_options=_get_client_suggestions(),
            remettant_options=_get_remettant_suggestions(),
            check_type_options=CHECK_TYPE_CHOICES,
        )
    return render_template(
        'edit.html',
        operation=op,
        bank_options=_get_bank_suggestions(),
        client_options=_get_client_suggestions(),
        remettant_options=_get_remettant_suggestions(),
        check_type_options=CHECK_TYPE_CHOICES,
    )


# ─── RÉFÉRENTIELS ────────────────────────────────────────────────────────────

def _ref_clients_ctx(is_admin, page=1, search='', per_page=REF_PER_PAGE):
    q = ClientLabel.query.order_by(ClientLabel.nom)
    if search:
        q = q.filter(ClientLabel.nom.ilike(f'%{search}%'))
    total = q.count()
    clients = q.offset((page - 1) * per_page).limit(per_page).all()
    return dict(clients=clients, is_admin=is_admin, page=page, search=search,
                total=total, total_pages=max(1, (total + per_page - 1) // per_page),
                per_page=per_page, per_page_options=PER_PAGE_OPTIONS)


def _ref_remettants_ctx(is_admin, page=1, search='', per_page=REF_PER_PAGE):
    q = RemettantLabel.query.order_by(RemettantLabel.nom)
    if search:
        q = q.filter(RemettantLabel.nom.ilike(f'%{search}%'))
    total = q.count()
    remettants = q.offset((page - 1) * per_page).limit(per_page).all()
    return dict(remettants=remettants, is_admin=is_admin, page=page, search=search,
                total=total, total_pages=max(1, (total + per_page - 1) // per_page),
                per_page=per_page, per_page_options=PER_PAGE_OPTIONS)


@app.route('/referentiels')
@role_required('admin', 'saisie')
def referentiels():
    return render_template('referentiels.html', is_admin=_current_role() == 'admin')


@app.route('/api/referentiels/clients/list')
@role_required('admin', 'saisie')
def api_ref_clients_list():
    page   = request.args.get('page', 1, type=int)
    search = request.args.get('search', '').strip()
    per_page = _get_per_page(request.args, 'per_page_clients', REF_PER_PAGE)
    return render_template('partials/ref_clients_list.html',
                           **_ref_clients_ctx(_current_role() == 'admin', page, search, per_page))


@app.route('/api/referentiels/remettants/list')
@role_required('admin', 'saisie')
def api_ref_remettants_list():
    page   = request.args.get('page', 1, type=int)
    search = request.args.get('search', '').strip()
    per_page = _get_per_page(request.args, 'per_page_remettants', REF_PER_PAGE)
    return render_template('partials/ref_remettants_list.html',
                           **_ref_remettants_ctx(_current_role() == 'admin', page, search, per_page))


@app.route('/api/referentiels/clients/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_client_add():
    nom = (request.form.get('nom') or '').strip()
    if not nom:
        return '<p class="text-error text-sm">Nom requis.</p>', 400
    if ClientLabel.query.filter(db.func.lower(ClientLabel.nom) == nom.lower()).first():
        return '<p class="text-error text-sm">Ce client existe déjà.</p>', 400
    db.session.add(ClientLabel(nom=nom))
    db.session.commit()
    return render_template('partials/ref_clients_list.html',
                           **_ref_clients_ctx(_current_role() == 'admin'))


@app.route('/api/referentiels/clients/<int:item_id>/toggle', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_client_toggle(item_id):
    item = ClientLabel.query.get_or_404(item_id)
    item.actif = not item.actif
    db.session.commit()
    return render_template('partials/ref_clients_list.html',
                           **_ref_clients_ctx(_current_role() == 'admin'))


@app.route('/api/referentiels/clients/<int:item_id>/delete', methods=['DELETE', 'POST'])
@role_required('admin')
def api_ref_client_delete(item_id):
    item = ClientLabel.query.get_or_404(item_id)
    db.session.delete(item)
    db.session.commit()
    return render_template('partials/ref_clients_list.html',
                           **_ref_clients_ctx(_current_role() == 'admin'))


@app.route('/api/referentiels/remettants/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_remettant_add():
    nom = (request.form.get('nom') or '').strip()
    if not nom:
        return '<p class="text-error text-sm">Nom requis.</p>', 400
    if RemettantLabel.query.filter(db.func.lower(RemettantLabel.nom) == nom.lower()).first():
        return '<p class="text-error text-sm">Ce remettant existe déjà.</p>', 400
    db.session.add(RemettantLabel(nom=nom))
    db.session.commit()
    return render_template('partials/ref_remettants_list.html',
                           **_ref_remettants_ctx(_current_role() == 'admin'))


@app.route('/api/referentiels/remettants/<int:item_id>/toggle', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_remettant_toggle(item_id):
    item = RemettantLabel.query.get_or_404(item_id)
    item.actif = not item.actif
    db.session.commit()
    return render_template('partials/ref_remettants_list.html',
                           **_ref_remettants_ctx(_current_role() == 'admin'))


@app.route('/api/referentiels/remettants/<int:item_id>/delete', methods=['DELETE', 'POST'])
@role_required('admin')
def api_ref_remettant_delete(item_id):
    item = RemettantLabel.query.get_or_404(item_id)
    db.session.delete(item)
    db.session.commit()
    return render_template('partials/ref_remettants_list.html',
                           **_ref_remettants_ctx(_current_role() == 'admin'))


# ─── LOGISTIQUE ──────────────────────────────────────────────────────────────

# ── Référentiels Fournisseurs ──

def _ref_fournisseurs_ctx(is_admin, page=1, search='', societe='', per_page=REF_PER_PAGE):
    q = Fournisseur.query.order_by(Fournisseur.nom)
    if societe:
        q = q.filter(Fournisseur.societe == societe)
    if search:
        q = q.filter(Fournisseur.nom.ilike(f'%{search}%'))
    total = q.count()
    fournisseurs = q.offset((page - 1) * per_page).limit(per_page).all()
    return dict(fournisseurs=fournisseurs, is_admin=is_admin, page=page, search=search,
                societe=societe, total=total, total_pages=max(1, (total + per_page - 1) // per_page),
                per_page=per_page, per_page_options=PER_PAGE_OPTIONS)


@app.route('/logistique/referentiels')
@role_required('admin', 'saisie')
def logistique_referentiels():
    return render_template('logistique_referentiels.html', is_admin=_current_role() == 'admin')


@app.route('/api/logistique/referentiels/fournisseurs/list')
@role_required('admin', 'saisie')
def api_ref_fournisseurs_list():
    page    = request.args.get('page', 1, type=int)
    search  = request.args.get('search', '').strip()
    societe = request.args.get('societe', '').strip()
    per_page = _get_per_page(request.args, 'per_page_fournisseurs', REF_PER_PAGE)
    return render_template('partials/ref_fournisseurs_list.html',
                           **_ref_fournisseurs_ctx(_current_role() == 'admin', page, search, societe, per_page))


@app.route('/api/logistique/referentiels/fournisseurs/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_fournisseur_add():
    nom     = (request.form.get('nom') or '').strip()
    societe = (request.form.get('societe') or '').strip()
    if not nom:
        return '<p class="text-error text-sm">Nom requis.</p>', 400
    if not societe:
        return '<p class="text-error text-sm">Société requise.</p>', 400
    if Fournisseur.query.filter(db.func.lower(Fournisseur.nom) == nom.lower(), Fournisseur.societe == societe).first():
        return '<p class="text-error text-sm">Ce fournisseur existe déjà pour cette société.</p>', 400
    db.session.add(Fournisseur(nom=nom, societe=societe))
    db.session.commit()
    return render_template('partials/ref_fournisseurs_list.html',
                           **_ref_fournisseurs_ctx(_current_role() == 'admin', societe=societe))


@app.route('/api/logistique/referentiels/fournisseurs/<int:item_id>/toggle', methods=['POST'])
@role_required('admin', 'saisie')
def api_ref_fournisseur_toggle(item_id):
    item = Fournisseur.query.get_or_404(item_id)
    item.actif = not item.actif
    db.session.commit()
    return render_template('partials/ref_fournisseurs_list.html',
                           **_ref_fournisseurs_ctx(_current_role() == 'admin', societe=item.societe))


@app.route('/api/logistique/referentiels/fournisseurs/<int:item_id>/delete', methods=['DELETE', 'POST'])
@role_required('admin')
def api_ref_fournisseur_delete(item_id):
    item = Fournisseur.query.get_or_404(item_id)
    societe = item.societe
    db.session.delete(item)
    db.session.commit()
    return render_template('partials/ref_fournisseurs_list.html',
                           **_ref_fournisseurs_ctx(_current_role() == 'admin', societe=societe))


def _log_kpis():
    """Calcule les KPIs logistique par statut via SQL (sans charger tous les objets)."""
    dad = CommandeLogistique.query.filter(
        CommandeLogistique.date_arrivee_depot.isnot(None)
    ).count()
    d10 = CommandeLogistique.query.filter(
        CommandeLogistique.date_arrivee_depot.is_(None),
        CommandeLogistique.date_d10.isnot(None)
    ).count()
    dap = CommandeLogistique.query.filter(
        CommandeLogistique.date_arrivee_depot.is_(None),
        CommandeLogistique.date_d10.is_(None),
        CommandeLogistique.date_arrivee.isnot(None)
    ).count()
    etd = CommandeLogistique.query.filter(
        CommandeLogistique.date_arrivee_depot.is_(None),
        CommandeLogistique.date_d10.is_(None),
        CommandeLogistique.date_arrivee.is_(None),
        CommandeLogistique.date_etd.isnot(None)
    ).count()
    arrivage = CommandeLogistique.query.filter(
        CommandeLogistique.date_arrivee_depot.is_(None),
        CommandeLogistique.date_d10.is_(None),
        CommandeLogistique.date_arrivee.is_(None),
        CommandeLogistique.date_etd.is_(None)
    ).count()

    return {
        'DAD': dad, 'D10': d10, 'DAP': dap,
        'ETD': etd, 'ARRIVAGE': arrivage,
    }


def _get_logistique_notifications(limit=8):
    """Retourne les alertes d'échéance pour la gestion logistique."""
    today = date.today()
    alert_date = today + timedelta(days=7)

    # Workflow échéance : uniquement les entrées non payées.
    base_query = CommandeLogistique.query.filter(
        CommandeLogistique.date_echeance.isnot(None),
        CommandeLogistique.date_paiement.is_(None),
        CommandeLogistique.date_valeur.is_(None),
    )

    overdue_count = base_query.filter(CommandeLogistique.date_echeance < today).count()
    upcoming_count = base_query.filter(
        CommandeLogistique.date_echeance >= today,
        CommandeLogistique.date_echeance <= alert_date,
    ).count()

    critical_entries = base_query.filter(
        or_(
            CommandeLogistique.date_echeance < today,
            (
                (CommandeLogistique.date_echeance >= today) &
                (CommandeLogistique.date_echeance <= alert_date)
            ),
        )
    ).order_by(CommandeLogistique.date_echeance.asc(), CommandeLogistique.id.desc()).limit(limit).all()

    return {
        'today': today,
        'overdue_count': overdue_count,
        'upcoming_count': upcoming_count,
        'total_alerts': overdue_count + upcoming_count,
        'critical_entries': critical_entries,
    }


LOG_STATUTS = ['ARRIVAGE', 'ETD', 'DAP', 'D10', 'DAD']


# ── Prix de revient : taux par défaut (modifiables par l'utilisateur) ───────
PR_TAUX_TCS             = 0.03   # TCS 3%
PR_TAUX_TVA             = 0.09   # TVA 9%
PR_TAUX_PRECOMPTE_IBS   = 0.02   # Précompte IBS 2%
PR_TAUX_TAXE_DOM        = 0.003  # Taxe domiciliation 0.3%
PR_TAUX_FRAIS_TRANSFERT = 0.004  # Frais de transfert 0.4%


def _cours_du_bon(bon_id):
    """Récupère le cours (taux de change) déjà saisi dans l'entrée Gestion des commandes du bon."""
    log_entry = CommandeLogistique.query.filter_by(bon_id=bon_id).first()
    return log_entry.cours if log_entry else None


def _attach_devise(items):
    """Attache la devise du bon (item.devise_bon) à chaque entrée logistique.

    La devise est portée par les lignes du bon (LigneCommande.devise). Une seule
    requête groupée est faite pour tous les bons de la page.
    """
    bon_ids = [it.bon_id for it in items if getattr(it, 'bon_id', None)]
    devise_map = {}
    if bon_ids:
        rows = (db.session.query(LigneCommande.bon_id, LigneCommande.devise)
                .filter(LigneCommande.bon_id.in_(bon_ids)).all())
        for bid, dev in rows:
            devise_map.setdefault(bid, dev or 'EUR')
    for it in items:
        it.devise_bon = devise_map.get(getattr(it, 'bon_id', None), 'EUR')
    return items


def _attach_pr(items):
    """Attache l'état des frais et le PR TTC total à chaque entrée logistique.

    - item.frais_saisis : True si des frais/charges ont été enregistrés pour le bon.
    - item.pr_ttc_total : prix de revient TTC total (None si frais non saisis ou cours manquant).
    """
    bon_ids = [it.bon_id for it in items if getattr(it, 'bon_id', None)]
    frais_map = {}
    if bon_ids:
        for f in FraisLogistique.query.filter(FraisLogistique.bon_id.in_(bon_ids)).all():
            frais_map.setdefault(f.bon_id, f)
    for it in items:
        f = frais_map.get(getattr(it, 'bon_id', None))
        it.frais_saisis = bool(f and f.charges_saisies)
        it.pr_ttc_total = None
        if it.frais_saisis and it.bon_id and it.cours:
            bon = BonCommande.query.get(it.bon_id)
            if bon:
                pr = _calculer_pr_bon(bon, f, it.cours)
                if pr:
                    it.pr_ttc_total = sum(r['pr_ttc'] for r in pr)
    return items


def _frais_config_payload(frais, bon, cours):
    """Construit le payload JSON de configuration PR d'un bon (utilisé par la modale)."""
    cfg = _pr_config(frais, bon)
    return {
        'numero': bon.numero if bon else None,
        'cours':  cours,
        'devise': (bon.lignes[0].devise if (bon and bon.lignes) else None) or 'EUR',
        'remarque': frais.remarque or '',
        'rates':  cfg['rates'],
        'charges': cfg['charges'],
        'produits': [dict(id=lid, **vals) for lid, vals in cfg['produits'].items()],
    }


def _pr_config(frais, bon):
    """Construit la configuration de calcul du PR pour un bon.

    Fusionne les valeurs éditables enregistrées (frais.pr_config, JSON) avec les
    valeurs par défaut : taux fixes standards + données produits issues du bon.
    - Prix achat et quantité proviennent du bon (non modifiables).
    - Le taux de droits de douane (D.D.) et le fret sont communs au bon.
    - Seule la TVA peut différer par produit.
    Structure retournée :
      {
        'rates':   {'tcs','tva','precompte','taxe_dom','frais_transfert','douane'},
        'charges': {'rps','fret','echange','honoraires','magasinage','surestaries'},
        'produits': { <ligne_id>: {'nom','prix','qte','tva'} },
      }
    """
    saved = {}
    if frais and frais.pr_config:
        try:
            saved = json.loads(frais.pr_config)
        except (ValueError, TypeError):
            saved = {}

    s_rates    = saved.get('rates', {}) or {}
    s_charges  = saved.get('charges', {}) or {}
    s_produits = saved.get('produits', {}) or {}

    default_douane = frais.taux_douane if (frais and frais.taux_douane is not None) else 0
    rates = {
        'tcs':             s_rates.get('tcs',             PR_TAUX_TCS),
        'tva':             s_rates.get('tva',             PR_TAUX_TVA),
        'precompte':       s_rates.get('precompte',       PR_TAUX_PRECOMPTE_IBS),
        'taxe_dom':        s_rates.get('taxe_dom',        PR_TAUX_TAXE_DOM),
        'frais_transfert': s_rates.get('frais_transfert', PR_TAUX_FRAIS_TRANSFERT),
        'douane':          s_rates.get('douane',          default_douane),
    }
    charges = {
        'rps':         s_charges.get('rps',         frais.rps if frais else 2500),
        'fret':        s_charges.get('fret',        0),
        'echange':     s_charges.get('echange',     frais.echange if frais else 0),
        'honoraires':  s_charges.get('honoraires',  frais.honoraires_transitaire if frais else 0),
        'magasinage':  s_charges.get('magasinage',  frais.magasinage if frais else 0),
        'surestaries': s_charges.get('surestaries', frais.surestaries if frais else 0),
    }

    produits = {}
    if bon:
        for l in bon.lignes:
            sp = s_produits.get(str(l.id), {}) or {}
            produits[str(l.id)] = {
                'nom':  l.designation,
                'prix': l.prix_unitaire or 0,   # depuis le bon (lecture seule)
                'qte':  l.quantite or 0,        # depuis le bon (lecture seule)
                'tva':  sp.get('tva', rates['tva']),
            }
    return {'rates': rates, 'charges': charges, 'produits': produits}


def _calculer_pr_bon(bon, frais, cours):
    """Calcule le prix de revient (PR) détaillé de chaque produit d'un bon.

    Valeur en douane = Montant DA + (fret réparti × cours). Le fret et le RPS
    sont saisis une fois pour le bon puis répartis au prorata de la part de chaque
    produit dans la valeur des marchandises. Le taux de droits de douane est commun
    au bon ; la TVA peut différer par produit. Les taxes proportionnelles sont
    calculées sur la valeur en douane (fret inclus).
    """
    if not bon or not bon.lignes or not cours:
        return []

    cfg = _pr_config(frais, bon)
    rates = cfg['rates']
    charges = cfg['charges']
    fret_total = charges['fret'] or 0
    douane_rate = rates['douane'] or 0

    ordered = [(l, cfg['produits'][str(l.id)]) for l in bon.lignes if str(l.id) in cfg['produits']]

    montants_da = [(p['prix'] or 0) * (p['qte'] or 0) * cours for _, p in ordered]
    total_da = sum(montants_da)

    resultats = []
    for (ligne, p), montant_da in zip(ordered, montants_da):
        part = (montant_da / total_da) if total_da else 0
        qty  = p['qte'] or 0

        fret_part_usd = fret_total * part
        fret_da       = fret_part_usd * cours
        valeur_douane = montant_da + fret_da

        # Taxes douanières (sur la valeur en douane, fret inclus)
        tcs       = valeur_douane * rates['tcs']
        douane    = valeur_douane * douane_rate
        tva       = (valeur_douane + tcs) * (p['tva'] or 0)
        precompte = (valeur_douane + tcs + tva) * rates['precompte']
        rps_l     = (charges['rps'] or 0) * part
        total_taxes_douane = rps_l + tcs + douane + tva + precompte
        total_douanes      = valeur_douane + total_taxes_douane

        # Frais transitaire
        taxe_dom        = valeur_douane * rates['taxe_dom']
        frais_transfert = valeur_douane * rates['frais_transfert']
        echange_l       = (charges['echange'] or 0) * part
        honoraires_l    = (charges['honoraires'] or 0) * part
        magasinage_l    = (charges['magasinage'] or 0) * part
        surestaries_l   = (charges['surestaries'] or 0) * part
        total_transitaire = (taxe_dom + frais_transfert + echange_l
                             + honoraires_l + magasinage_l + surestaries_l)

        pr_ttc = total_douanes + total_transitaire
        pr_ht  = pr_ttc - tva

        resultats.append({
            'ligne':             ligne,
            'nom':               p['nom'],
            'prix':              p['prix'],
            'qte':               qty,
            'prix_usd_total':    (p['prix'] or 0) * qty,
            'fret_usd':          fret_part_usd,
            'fret_da':           fret_da,
            'montant_da':        montant_da,
            'valeur_douane':     valeur_douane,
            'part':              part,
            'rps':               rps_l,
            'tcs':               tcs,
            'douane':            douane,
            'tva':               tva,
            'precompte':         precompte,
            'total_douanes':     total_douanes,
            'taxe_dom':          taxe_dom,
            'frais_transfert':   frais_transfert,
            'echange':           echange_l,
            'honoraires':        honoraires_l,
            'magasinage':        magasinage_l,
            'surestaries':       surestaries_l,
            'total_transitaire': total_transitaire,
            'pr_ttc':            pr_ttc,
            'pr_ht':             pr_ht,
            'pr_unit_ttc':       (pr_ttc / qty) if qty else None,
            'pr_unit_ht':        (pr_ht / qty) if qty else None,
        })
    return resultats


def _build_frais_query(args):
    """Construit la liste/pagination de la table Frais (vue independante)."""
    page = args.get('frais_page', 1, type=int)
    per_page = _get_per_page(args, 'per_page_frais', LOG_PER_PAGE)
    search = args.get('search', '').strip()
    societe = args.get('societe', '').strip()
    sort_col = args.get('frais_sort', '').strip()
    sort_dir = args.get('frais_dir', 'desc').strip()
    if sort_dir not in ('asc', 'desc'):
        sort_dir = 'desc'

    q = FraisLogistique.query
    if search:
        q = q.filter(or_(
            FraisLogistique.ref_log.ilike(f'%{search}%'),
            FraisLogistique.fournisseur.ilike(f'%{search}%'),
            FraisLogistique.annee.ilike(f'%{search}%'),
            FraisLogistique.remarque.ilike(f'%{search}%'),
        ))
    if societe:
        q = q.filter(FraisLogistique.societe == societe)

    frais_sort_columns = {
        'ref_log': FraisLogistique.ref_log,
        'societe': FraisLogistique.societe,
        'annee': FraisLogistique.annee,
        'fournisseur': FraisLogistique.fournisseur,
        'taux_douane': FraisLogistique.taux_douane,
        'rps': FraisLogistique.rps,
        'date_creation': FraisLogistique.date_creation,
    }
    if sort_col in frais_sort_columns:
        col = frais_sort_columns[sort_col]
        order = col.asc().nullslast() if sort_dir == 'asc' else col.desc().nullslast()
        q = q.order_by(order, FraisLogistique.id.desc())
    else:
        q = q.order_by(FraisLogistique.date_creation.desc().nullslast(), FraisLogistique.id.desc())

    total = q.count()
    items = q.offset((page - 1) * per_page).limit(per_page).all()
    total_pages = max(1, (total + per_page - 1) // per_page)

    # Cours (taux de change) et PR agrege par bon, pour affichage dans la table.
    # frais_pr_map : configuration complete par entree pour le calcul PR live dans la modale.
    frais_pr_map = {}
    for item in items:
        item.cours_bon = _cours_du_bon(item.bon_id) if item.bon_id else None
        bon = BonCommande.query.get(item.bon_id) if item.bon_id else None
        frais_pr_map[item.id] = _frais_config_payload(item, bon, item.cours_bon)
        if item.bon_id and item.cours_bon and bon:
            pr_lignes = _calculer_pr_bon(bon, item, item.cours_bon)
            item.pr_ttc_total = sum(r['pr_ttc'] for r in pr_lignes) if pr_lignes else None
            item.pr_ht_total = sum(r['pr_ht'] for r in pr_lignes) if pr_lignes else None
        else:
            item.pr_ttc_total = None
            item.pr_ht_total = None

    return {
        'frais_items': items,
        'frais_pr_map': frais_pr_map,
        'frais_page': page,
        'frais_total': total,
        'frais_total_pages': total_pages,
        'frais_sort_col': sort_col,
        'frais_sort_dir': sort_dir,
        'per_page_frais': per_page,
        'per_page_options': PER_PAGE_OPTIONS,
    }


@app.route('/api/logistique/bons')
@login_required
def api_logistique_bons_list():
    """HTMX partial: retourne la table des bons filtrée."""
    page     = request.args.get('page', 1, type=int)
    search   = request.args.get('search', '').strip()
    societe  = request.args.get('societe', '').strip()
    statut_f = request.args.get('statut', '').strip()
    sort_col = request.args.get('sort', '').strip()
    sort_dir = request.args.get('dir', 'asc').strip()
    per_page = _get_per_page(request.args, 'per_page_bons', BON_PER_PAGE)
    if sort_col != 'statut' and sort_dir not in ('asc', 'desc'):
        sort_dir = 'asc'

    q = BonCommande.query
    if search:
        q = q.filter(or_(
            BonCommande.fournisseur.ilike(f'%{search}%'),
            BonCommande.numero.ilike(f'%{search}%'),
        ))
    if societe:
        q = q.filter(BonCommande.societe == societe)
    if statut_f:
        q = q.filter(BonCommande.statut == statut_f)

    # -- Tri dynamique --
    bon_sort_columns = {
        'fournisseur': BonCommande.fournisseur,
        'date_commande': BonCommande.date_commande,
        'societe': BonCommande.societe,
    }
    if sort_col in bon_sort_columns:
        col = bon_sort_columns[sort_col]
        order = col.asc().nullslast() if sort_dir == 'asc' else col.desc().nullslast()
        q = q.order_by(order, BonCommande.id.desc())
    elif sort_col == 'statut':
        all_bons = q.all()
        present_statuts = [s for s in BON_STATUTS if any(b.statut == s for b in all_bons)]
        if not present_statuts:
            present_statuts = BON_STATUTS
        statut_idx = 0
        try:
            statut_idx = int(sort_dir) % len(present_statuts)
        except (ValueError, ZeroDivisionError):
            pass
        rotated = present_statuts[statut_idx:] + present_statuts[:statut_idx]
        statut_order = {s: i for i, s in enumerate(rotated)}
        all_bons.sort(key=lambda b: (statut_order.get(b.statut, 99), -(b.date_commande.toordinal() if b.date_commande else 0), -b.id))
        total = len(all_bons)
        bons = all_bons[(page - 1) * per_page: page * per_page]
        total_pages = max(1, (total + per_page - 1) // per_page)
        return render_template('partials/logistique_bons_table.html',
                               bons=bons, page=page, total_pages=total_pages, total=total,
                               search=search, societe=societe, statut_f=statut_f,
                               bon_statuts=BON_STATUTS,
                               sort_col=sort_col, sort_dir=sort_dir,
                       per_page_bons=per_page,
                       per_page_options=PER_PAGE_OPTIONS,
                               can_write=_current_role() in ('admin', 'saisie'),
                               is_admin=_current_role() == 'admin')
    else:
        q = q.order_by(BonCommande.date_commande.desc(), BonCommande.id.desc())

    total       = q.count()
    bons        = q.offset((page - 1) * per_page).limit(per_page).all()
    total_pages = max(1, (total + per_page - 1) // per_page)
    return render_template('partials/logistique_bons_table.html',
                           bons=bons, page=page, total_pages=total_pages, total=total,
                           search=search, societe=societe, statut_f=statut_f,
                           bon_statuts=BON_STATUTS,
                           sort_col=sort_col, sort_dir=sort_dir,
                           per_page_bons=per_page,
                           per_page_options=PER_PAGE_OPTIONS,
                           can_write=_current_role() in ('admin', 'saisie'),
                           is_admin=_current_role() == 'admin')


@app.route('/api/logistique/gestion')
@login_required
def api_logistique_gestion_list():
    """HTMX partial: retourne la table de gestion filtrée."""
    page     = request.args.get('page', 1, type=int)
    search   = request.args.get('search', '').strip()
    societe  = request.args.get('societe', '').strip()
    statut_f = request.args.get('statut', '').strip()
    date_filter = request.args.get('date_filter', '').strip()
    date_debut_raw = request.args.get('date_debut', '').strip()
    date_fin_raw = request.args.get('date_fin', '').strip()
    sort_col = request.args.get('sort', '').strip()
    sort_dir = request.args.get('dir', 'asc').strip()
    per_page = _get_per_page(request.args, 'per_page_log', LOG_PER_PAGE)
    if sort_col != 'statut' and sort_dir not in ('asc', 'desc'):
        sort_dir = 'asc'

    def _parse_date(v):
        try:
            return datetime.strptime(v, '%Y-%m-%d').date() if v else None
        except ValueError:
            return None

    date_debut = _parse_date(date_debut_raw)
    date_fin = _parse_date(date_fin_raw)
    date_fields = {
        'date_d10': CommandeLogistique.date_d10,
        'date_arrivee': CommandeLogistique.date_arrivee,
        'date_facture': CommandeLogistique.date_facture,
        'date_echeance': CommandeLogistique.date_echeance,
        'date_paiement': CommandeLogistique.date_paiement,
        'date_valeur': CommandeLogistique.date_valeur,
    }

    q = CommandeLogistique.query
    if search:
        q = q.filter(or_(
            CommandeLogistique.produit.ilike(f'%{search}%'),
            CommandeLogistique.fournisseur.ilike(f'%{search}%'),
            CommandeLogistique.remarque.ilike(f'%{search}%'),
        ))
    if societe:
        q = q.filter(CommandeLogistique.societe == societe)
    if date_filter in date_fields:
        df = date_fields[date_filter]
        if date_debut:
            q = q.filter(df >= date_debut)
        if date_fin:
            q = q.filter(df <= date_fin)

    # -- Tri dynamique --
    sort_columns = {
        'fournisseur': CommandeLogistique.fournisseur,
        'date_arrivee': CommandeLogistique.date_arrivee,
        'date_etd': CommandeLogistique.date_etd,
        'date_d10': CommandeLogistique.date_d10,
        'date_arrivee_depot': CommandeLogistique.date_arrivee_depot,
        'date_echeance': CommandeLogistique.date_echeance,
        'date_paiement': CommandeLogistique.date_paiement,
        'date_valeur': CommandeLogistique.date_valeur,
        'ref_log': CommandeLogistique.ref_log,
    }
    if sort_col in sort_columns:
        col = sort_columns[sort_col]
        order = col.asc().nullslast() if sort_dir == 'asc' else col.desc().nullslast()
        q = q.order_by(order, CommandeLogistique.id.desc())
    elif sort_col == 'statut':
        # Tri par statut géré après fetch (propriété calculée)
        pass
    else:
        q = q.order_by(
            CommandeLogistique.date_creation.desc().nullslast(),
            CommandeLogistique.id.desc(),
        )

    if statut_f:
        all_items  = q.all()
        filtered   = [c for c in all_items if c.statut == statut_f]
        total      = len(filtered)
        if sort_col == 'statut':
            pass  # all same statut, no sort needed
        items      = filtered[(page - 1) * per_page: page * per_page]
    else:
        if sort_col == 'statut':
            all_items = q.all()
            # Only cycle through statuses that actually have entries
            present_statuts = [s for s in LOG_STATUTS if any(c.statut == s for c in all_items)]
            if not present_statuts:
                present_statuts = LOG_STATUTS
            statut_idx = 0
            try:
                statut_idx = int(sort_dir) % len(present_statuts)
            except (ValueError, ZeroDivisionError):
                pass
            # Rotate so that the target statut comes first
            rotated = present_statuts[statut_idx:] + present_statuts[:statut_idx]
            statut_order = {s: i for i, s in enumerate(rotated)}
            all_items.sort(key=lambda c: (statut_order.get(c.statut, 99), -(c.date_creation.timestamp() if c.date_creation else 0), -c.id))
            total = len(all_items)
            items = all_items[(page - 1) * per_page: page * per_page]
        else:
            total = q.count()
            items = q.offset((page - 1) * per_page).limit(per_page).all()

    total_pages = max(1, (total + per_page - 1) // per_page)
    _attach_devise(items)
    _attach_pr(items)
    return render_template('partials/logistique_gestion_table.html',
                           items=items, page=page, total_pages=total_pages, total=total,
                           search=search, societe=societe,
                           statut_f=statut_f, date_filter=date_filter,
                           date_debut=date_debut_raw, date_fin=date_fin_raw,
                           sort_col=sort_col, sort_dir=sort_dir,
                           today=date.today(),
                           per_page_log=per_page,
                           per_page_options=PER_PAGE_OPTIONS,
                           can_write=_current_role() in ('admin', 'saisie'),
                           is_admin=_current_role() == 'admin')


@app.route('/api/logistique/frais')
@login_required
def api_logistique_frais_list():
    data = _build_frais_query(request.args)
    return render_template(
        'partials/logistique_frais_table.html',
        search=request.args.get('search', '').strip(),
        societe=request.args.get('societe', '').strip(),
        can_write=_current_role() in ('admin', 'saisie'),
        **data,
    )


@app.route('/api/logistique/frais/config/<int:bon_id>')
@login_required
def api_frais_config(bon_id):
    """Configuration PR d'un bon, pour la modale de calcul depuis Gestion des commandes."""
    frais = FraisLogistique.query.filter_by(bon_id=bon_id).first()
    if not frais:
        return jsonify({'error': 'Aucun frais associé à ce bon'}), 404
    bon = BonCommande.query.get(bon_id)
    cours = _cours_du_bon(bon_id)
    payload = _frais_config_payload(frais, bon, cours)
    payload['frais_id'] = frais.id
    return jsonify(payload)


@app.route('/api/logistique/frais/<int:item_id>/edit', methods=['POST'])
@role_required('admin', 'saisie')
def api_frais_edit(item_id):
    frais = FraisLogistique.query.get_or_404(item_id)

    def _f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    # Configuration complète (taux éditables + surcharges par produit) au format JSON.
    config = {}
    raw_config = request.form.get('pr_config', '').strip()
    if raw_config:
        try:
            config = json.loads(raw_config)
        except (ValueError, TypeError):
            config = {}

    if config:
        frais.pr_config = json.dumps(config, ensure_ascii=False)
        # Synchronise les colonnes héritées (utilisées dans la table de synthèse).
        rates   = config.get('rates', {}) or {}
        charges = config.get('charges', {}) or {}
        frais.rps         = _f(charges.get('rps'))
        frais.echange     = _f(charges.get('echange'))
        frais.honoraires_transitaire = _f(charges.get('honoraires'))
        frais.magasinage  = _f(charges.get('magasinage'))
        frais.surestaries = _f(charges.get('surestaries'))
        frais.taux_douane = _f(rates.get('douane'))
    else:
        # Repli : ancienne saisie simple par champs de formulaire.
        def ff(k):
            v = request.form.get(k, '').strip()
            try:
                return float(v) if v else None
            except ValueError:
                return None
        taux_douane = ff('taux_douane')
        if taux_douane is not None:
            taux_douane = taux_douane / 100 if taux_douane > 1 else taux_douane
        frais.taux_douane            = taux_douane
        frais.rps                    = ff('rps')
        frais.echange                = ff('echange')
        frais.honoraires_transitaire = ff('honoraires_transitaire')
        frais.magasinage             = ff('magasinage')
        frais.surestaries            = ff('surestaries')

    frais.remarque = request.form.get('remarque', '').strip() or None
    db.session.commit()
    return redirect(url_for('logistique_gestion'))


@app.route('/logistique/prix-revient')
@login_required
def logistique_prix_revient():
    societe = request.args.get('societe', '').strip()

    # On calcule le PR de tous les bons (toutes sociétés) : la liste des produits
    # proposée dans le filtre doit regrouper les produits des 2 sociétés.
    q = FraisLogistique.query.order_by(
        FraisLogistique.date_creation.desc().nullslast(), FraisLogistique.id.desc()
    )

    all_entries = []
    for frais in q.all():
        if not frais.bon_id or not frais.charges_saisies:
            continue
        bon = BonCommande.query.get(frais.bon_id)
        cours = _cours_du_bon(frais.bon_id)
        if not bon or not cours:
            continue
        pr_lignes = _calculer_pr_bon(bon, frais, cours)
        if not pr_lignes:
            continue
        all_entries.append({
            'bon': bon,
            'frais': frais,
            'cours': cours,
            'lignes': pr_lignes,
            'pr_ttc_total': sum(r['pr_ttc'] for r in pr_lignes),
            'pr_ht_total': sum(r['pr_ht'] for r in pr_lignes),
        })

    # Liste complète des produits (toutes sociétés confondues).
    produit_options = sorted(
        {(r['nom'] or (r['ligne'].designation if r['ligne'] else '')).strip()
         for e in all_entries for r in e['lignes']
         if (r['nom'] or (r['ligne'].designation if r['ligne'] else '')).strip()},
        key=lambda s: s.upper()
    )

    # Le tableau lui-même reste filtrable par société.
    if societe:
        bons_pr = [e for e in all_entries if e['frais'].societe == societe]
    else:
        bons_pr = all_entries

    return render_template(
        'logistique_prix_revient.html',
        bons_pr=bons_pr,
        produit_options=produit_options,
        societe=societe,
        can_write=_current_role() in ('admin', 'saisie'),
    )


@app.route('/logistique/gestion')
@login_required
def logistique_gestion():
    page     = request.args.get('page', 1, type=int)
    per_page = _get_per_page(request.args, 'per_page_log', LOG_PER_PAGE)
    search   = request.args.get('search', '').strip()
    societe  = request.args.get('societe', '').strip()
    statut_f = request.args.get('statut', '').strip()
    date_filter = request.args.get('date_filter', '').strip()
    date_debut_raw = request.args.get('date_debut', '').strip()
    date_fin_raw = request.args.get('date_fin', '').strip()

    def _parse_date(v):
        try:
            return datetime.strptime(v, '%Y-%m-%d').date() if v else None
        except ValueError:
            return None

    date_debut = _parse_date(date_debut_raw)
    date_fin = _parse_date(date_fin_raw)
    date_fields = {
        'date_d10': CommandeLogistique.date_d10,
        'date_arrivee': CommandeLogistique.date_arrivee,
        'date_facture': CommandeLogistique.date_facture,
        'date_echeance': CommandeLogistique.date_echeance,
        'date_paiement': CommandeLogistique.date_paiement,
        'date_valeur': CommandeLogistique.date_valeur,
    }

    q = CommandeLogistique.query
    if search:
        q = q.filter(or_(
            CommandeLogistique.produit.ilike(f'%{search}%'),
            CommandeLogistique.fournisseur.ilike(f'%{search}%')
        ))
    if societe:
        q = q.filter(CommandeLogistique.societe == societe)
    if date_filter in date_fields:
        df = date_fields[date_filter]
        if date_debut:
            q = q.filter(df >= date_debut)
        if date_fin:
            q = q.filter(df <= date_fin)

    q = q.order_by(
        CommandeLogistique.date_creation.desc().nullslast(),
        CommandeLogistique.id.desc(),
    )

    if statut_f:
        all_items        = q.all()
        items_filtered   = [c for c in all_items if c.statut == statut_f]
        total            = len(items_filtered)
        items            = items_filtered[(page - 1) * per_page: page * per_page]
    else:
        total = q.count()
        items = q.offset((page - 1) * per_page).limit(per_page).all()

    total_pages  = max(1, (total + per_page - 1) // per_page)
    _attach_devise(items)
    _attach_pr(items)

    return render_template('logistique_gestion.html',
                           items=items, page=page, total_pages=total_pages, total=total,
                           search=search, societe=societe,
                           statut_f=statut_f, date_filter=date_filter,
                           date_debut=date_debut_raw, date_fin=date_fin_raw,
                           kpis=_log_kpis(),
                           notifications=_get_logistique_notifications(),
                           today=date.today(),
                           log_statuts=LOG_STATUTS,
                           per_page_log=per_page,
                           can_write=_current_role() in ('admin', 'saisie'),
                           is_admin=_current_role() == 'admin')


@app.route('/api/logistique/notifications')
@login_required
def api_logistique_notifications():
    return render_template(
        'partials/logistique_notifications_panel.html',
        notifications=_get_logistique_notifications(),
    )


def _log_form_fields(c):
    """Lit les champs du formulaire logistique depuis request.form et les applique à l'objet c."""
    def fd(k):
        v = request.form.get(k, '').strip()
        try:
            return datetime.strptime(v, '%Y-%m-%d').date() if v else None
        except ValueError:
            return None
    def ff(k):
        try:
            v = request.form.get(k, '').strip()
            return float(v) if v else None
        except ValueError:
            return None
    def fi(k):
        try:
            v = request.form.get(k, '').strip()
            return int(float(v)) if v else None
        except ValueError:
            return None

    if 'societe' in request.form:
        c.societe       = request.form.get('societe', '').strip()
    c.annee         = request.form.get('annee', '').strip() or None
    c.date_d10      = fd('date_d10')
    c.date_arrivee  = fd('date_arrivee')
    c.date_arrivee_depot = fd('date_arrivee_depot')
    if 'fournisseur' in request.form:
        c.fournisseur   = request.form.get('fournisseur', '').strip().upper() or None
    c.produit       = request.form.get('produit', '').strip() or None
    c.emballage     = request.form.get('emballage', '').strip() or None
    c.quantite      = ff('quantite')
    c.tva           = ff('tva')
    c.montant_eur   = ff('montant_eur')
    c.cours         = ff('cours')
    c.date_facture  = fd('date_facture')
    c.date_etd      = fd('date_etd')
    c.code_paiement = request.form.get('code_paiement', '').strip() or None
    c.nb_jours      = fi('nb_jours')
    # Point 12 : date d'échéance calculée automatiquement = date facture/BL + délai (jours)
    if c.date_facture is not None and c.nb_jours is not None:
        c.date_echeance = c.date_facture + timedelta(days=c.nb_jours)
    else:
        c.date_echeance = None
    c.date_paiement = fd('date_paiement')
    c.date_valeur   = fd('date_valeur')
    c.remarque      = request.form.get('remarque', '').strip() or None


@app.route('/api/logistique/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_logistique_add():
    c = CommandeLogistique(cree_par=session.get('username', ''))
    _log_form_fields(c)
    db.session.add(c)
    db.session.commit()
    return redirect(url_for('logistique_gestion'))


@app.route('/api/logistique/<int:item_id>/edit', methods=['POST'])
@role_required('admin', 'saisie')
def api_logistique_edit(item_id):
    c = CommandeLogistique.query.get_or_404(item_id)
    _log_form_fields(c)
    db.session.commit()
    return redirect(url_for('logistique_gestion'))


@app.route('/api/logistique/<int:item_id>/delete', methods=['POST', 'DELETE'])
@role_required('admin')
def api_logistique_delete(item_id):
    CommandeLogistique.query.get_or_404(item_id)
    flash('Suppression directe désactivée. Supprimez le bon de commande pour supprimer aussi son entrée de gestion.', 'warning')
    return redirect(url_for('logistique_gestion'))


# ── Products (Referentiels) ──────────────────────────────────────────────────

@app.route('/api/products/by-company')
@login_required
def api_products_by_company():
    """Get products filtered by company for autofill."""
    company = request.args.get('company', '').strip()
    if not company:
        return jsonify([])
    
    products = Product.query.filter_by(company=company).order_by(Product.reference).all()
    return jsonify([p.to_dict() for p in products])


@app.route('/api/products/search')
@login_required
def api_products_search():
    """Search products by reference or designation."""
    q = request.args.get('q', '').strip()
    company = request.args.get('company', '').strip()
    
    query = Product.query
    if company:
        query = query.filter_by(company=company)
    
    if q:
        query = query.filter(or_(
            Product.reference.ilike(f'%{q}%'),
            Product.designation.ilike(f'%{q}%')
        ))
    
    products = query.order_by(Product.reference).limit(50).all()
    return jsonify([p.to_dict() for p in products])


# ── Bons de commande ──────────────────────────────────────────────────────────

@app.route('/logistique/bons')
@login_required
def logistique_bons():
    page     = request.args.get('page', 1, type=int)
    per_page = _get_per_page(request.args, 'per_page_bons', BON_PER_PAGE)
    search   = request.args.get('search', '').strip()
    societe  = request.args.get('societe', '').strip()
    statut_f = request.args.get('statut', '').strip()

    q = BonCommande.query
    if search:
        q = q.filter(or_(
            BonCommande.fournisseur.ilike(f'%{search}%'),
            BonCommande.numero.ilike(f'%{search}%'),
            BonCommande.notes.ilike(f'%{search}%')
        ))
    if societe:
        q = q.filter(BonCommande.societe == societe)
    if statut_f:
        q = q.filter(BonCommande.statut == statut_f)

    q = q.order_by(BonCommande.date_commande.desc(), BonCommande.id.desc())
    total       = q.count()
    bons        = q.offset((page - 1) * per_page).limit(per_page).all()
    total_pages = max(1, (total + per_page - 1) // per_page)

    can_write = _current_role() in ('admin', 'saisie')

    # Ne charger les données du formulaire que si l'utilisateur peut écrire
    import json as _json
    if can_write:
        fournisseurs_srid = [f.nom for f in Fournisseur.query.filter_by(societe='SRID', actif=True).order_by(Fournisseur.nom).all()]
        fournisseurs_genetics = [f.nom for f in Fournisseur.query.filter_by(societe='SRID GENETICS', actif=True).order_by(Fournisseur.nom).all()]
        products = Product.query.order_by(Product.company, Product.designation).all()
        products_json = _json.dumps([p.to_dict() for p in products])
        fournisseurs_json = _json.dumps({'SRID': fournisseurs_srid, 'SRID GENETICS': fournisseurs_genetics})
    else:
        products_json = '[]'
        fournisseurs_json = '{}'

    return render_template('logistique_bons.html',
                           bons=bons, page=page, total_pages=total_pages, total=total,
                           search=search, societe=societe, statut_f=statut_f,
                           bon_statuts=BON_STATUTS, fournisseurs_json=fournisseurs_json,
                           products_json=products_json,
                           per_page_bons=per_page,
                           per_page_options=PER_PAGE_OPTIONS,
                           can_write=can_write,
                           is_admin=_current_role() == 'admin')


@app.route('/api/logistique/bons/add', methods=['POST'])
@role_required('admin', 'saisie')
def api_bon_add():
    def fd(k):
        v = request.form.get(k, '').strip()
        try:
            return datetime.strptime(v, '%Y-%m-%d').date() if v else None
        except ValueError:
            return None

    designations   = request.form.getlist('designation[]')
    quantites      = request.form.getlist('quantite[]')
    unites         = request.form.getlist('unite[]')
    prix_unitaires = request.form.getlist('prix_unitaire[]')
    references     = request.form.getlist('reference[]')
    devises        = request.form.getlist('devise[]')

    parsed_lines = []
    for i, raw_desig in enumerate(designations):
        desig = raw_desig.strip()
        qty_raw = quantites[i].strip() if i < len(quantites) else ''
        prix_raw = prix_unitaires[i].strip() if i < len(prix_unitaires) else ''

        try:
            qty = float(qty_raw) if qty_raw else 0.0
        except ValueError:
            qty = 0.0
        try:
            prix = float(prix_raw) if prix_raw else 0.0
        except ValueError:
            prix = 0.0

        if not desig or qty <= 0 or prix <= 0:
            flash(f'Ligne {i + 1} invalide: désignation requise, quantité > 0 et prix unitaire > 0.', 'error')
            return redirect(url_for('logistique_bons'))

        devise_raw = devises[i].strip().upper() if i < len(devises) and devises[i].strip() else 'EUR'
        devise_norm = 'EUR' if devise_raw in ('EUR', 'EURO') else devise_raw
        if devise_norm not in ('EUR', 'USD'):
            flash(f'Ligne {i + 1} invalide: devise non supportée ({devise_raw}).', 'error')
            return redirect(url_for('logistique_bons'))

        parsed_lines.append({
            'designation': desig,
            'quantite': qty,
            'prix_unitaire': prix,
            'reference': references[i].strip() if i < len(references) else '',
            'unite': unites[i].strip() if i < len(unites) else '',
            'devise': devise_norm,
        })

    if not parsed_lines:
        flash('Ajoutez au moins une ligne de commande valide.', 'error')
        return redirect(url_for('logistique_bons'))

    first_devise = parsed_lines[0]['devise']
    if any(l['devise'] != first_devise for l in parsed_lines):
        flash('Toutes les lignes doivent avoir la même devise (EURO ou USD).', 'error')
        return redirect(url_for('logistique_bons'))

    fournisseur_raw = request.form.get('fournisseur', '').strip().upper()
    if not fournisseur_raw:
        flash('Le fournisseur est obligatoire.', 'error')
        return redirect(url_for('logistique_bons'))

    # Point 17: nomenclature courte par fournisseur, séquence indépendante
    # Format: BC-<CODE>-<NNN> ex: BC-SONA7F-001
    # CODE = 4 premiers alnum + signature fournisseur (2 hexa) pour éviter collisions entre fournisseurs homonymes en préfixe.
    supplier_prefix = re.sub(r'[^A-Z0-9]+', '', fournisseur_raw)[:4] or 'SUPP'
    supplier_sig = format(zlib.crc32(fournisseur_raw.encode('utf-8')) & 0xFF, '02X')
    supplier_code = f'{supplier_prefix}{supplier_sig}'
    supplier_count = BonCommande.query.filter(BonCommande.fournisseur == fournisseur_raw).count()
    seq = supplier_count + 1
    numero = f'BC-{supplier_code}-{seq:03d}'
    while BonCommande.query.filter_by(numero=numero).first() is not None:
        seq += 1
        numero = f'BC-{supplier_code}-{seq:03d}'

    bon = BonCommande(
        numero                = numero,
        societe               = request.form.get('societe', '').strip(),
        fournisseur           = fournisseur_raw,
        statut                = request.form.get('statut', 'Brouillon'),
        date_commande         = fd('date_commande') or date.today(),
        date_livraison_prevue = fd('date_livraison_prevue'),
        notes                 = request.form.get('notes', '').strip() or None,
        cree_par              = session.get('username', ''),
    )
    db.session.add(bon)
    db.session.flush()

    total_montant = 0.0
    for l in parsed_lines:
        db.session.add(LigneCommande(
            bon_id=bon.id,
            reference=l['reference'] or None,
            designation=l['designation'],
            quantite=l['quantite'],
            unite=l['unite'] or None,
            prix_unitaire=l['prix_unitaire'],
            devise=l['devise'],
        ))
        total_montant += l['quantite'] * l['prix_unitaire']

    # Créer automatiquement l'entrée dans CommandeLogistique
    log_entry = CommandeLogistique(
        bon_id            = bon.id,
        ref_log           = numero,  # Même numéro que le bon
        societe           = bon.societe,
        annee             = str(date.today().year),
        fournisseur       = bon.fournisseur,
        montant_eur       = total_montant if total_montant else None,
        cree_par          = session.get('username', ''),
    )
    db.session.add(log_entry)

    # Créer automatiquement l'entrée Frais associée (charges à saisir manuellement ensuite)
    frais_entry = FraisLogistique(
        bon_id      = bon.id,
        ref_log     = numero,
        societe     = bon.societe,
        annee       = str(date.today().year),
        fournisseur = bon.fournisseur,
        cree_par    = session.get('username', ''),
    )
    db.session.add(frais_entry)
    db.session.commit()

    return redirect(url_for('logistique_bons'))


@app.route('/api/logistique/bons/<int:bon_id>/statut', methods=['POST'])
@role_required('admin', 'saisie')
def api_bon_statut(bon_id):
    bon = BonCommande.query.get_or_404(bon_id)
    bon.statut = request.form.get('statut', bon.statut)
    db.session.commit()
    return redirect(url_for('logistique_bons'))


@app.route('/api/logistique/bons/<int:bon_id>/update', methods=['POST'])
@role_required('admin', 'saisie')
def api_bon_update(bon_id):
    bon = BonCommande.query.get_or_404(bon_id)

    def fd(k):
        v = request.form.get(k, '').strip()
        try:
            return datetime.strptime(v, '%Y-%m-%d').date() if v else None
        except ValueError:
            return None

    designations   = request.form.getlist('designation[]')
    quantites      = request.form.getlist('quantite[]')
    prix_unitaires = request.form.getlist('prix_unitaire[]')
    references     = request.form.getlist('reference[]')
    devises        = request.form.getlist('devise[]')

    parsed_lines = []
    for i, raw_desig in enumerate(designations):
        desig = raw_desig.strip()
        qty_raw = quantites[i].strip() if i < len(quantites) else ''
        prix_raw = prix_unitaires[i].strip() if i < len(prix_unitaires) else ''

        try:
            qty = float(qty_raw) if qty_raw else 0.0
        except ValueError:
            qty = 0.0
        try:
            prix = float(prix_raw) if prix_raw else 0.0
        except ValueError:
            prix = 0.0

        if not desig or qty <= 0 or prix <= 0:
            flash(f'Ligne {i + 1} invalide: désignation requise, quantité > 0 et prix unitaire > 0.', 'error')
            return redirect(url_for('logistique_bons'))

        devise_raw = devises[i].strip().upper() if i < len(devises) and devises[i].strip() else 'EUR'
        devise_norm = 'EUR' if devise_raw in ('EUR', 'EURO') else devise_raw
        if devise_norm not in ('EUR', 'USD'):
            flash(f'Ligne {i + 1} invalide: devise non supportée ({devise_raw}).', 'error')
            return redirect(url_for('logistique_bons'))

        parsed_lines.append({
            'designation': desig,
            'quantite': qty,
            'prix_unitaire': prix,
            'reference': references[i].strip() if i < len(references) else '',
            'devise': devise_norm,
        })

    if not parsed_lines:
        flash('Ajoutez au moins une ligne de commande valide.', 'error')
        return redirect(url_for('logistique_bons'))

    first_devise = parsed_lines[0]['devise']
    if any(l['devise'] != first_devise for l in parsed_lines):
        flash('Toutes les lignes doivent avoir la même devise (EURO ou USD).', 'error')
        return redirect(url_for('logistique_bons'))

    bon.societe       = request.form.get('societe', bon.societe).strip()
    bon.fournisseur   = request.form.get('fournisseur', '').strip().upper() or None
    bon.statut        = request.form.get('statut', bon.statut)
    bon.date_commande = fd('date_commande') or bon.date_commande
    bon.notes         = request.form.get('notes', '').strip() or None

    # Remplacer les lignes existantes
    LigneCommande.query.filter_by(bon_id=bon.id).delete()
    db.session.flush()

    total_montant = 0.0
    for l in parsed_lines:
        db.session.add(LigneCommande(
            bon_id=bon.id,
            reference=l['reference'] or None,
            designation=l['designation'],
            quantite=l['quantite'],
            prix_unitaire=l['prix_unitaire'],
            devise=l['devise'],
        ))
        total_montant += l['quantite'] * l['prix_unitaire']

    # Mettre à jour l'entrée CommandeLogistique associée
    log_entry = CommandeLogistique.query.filter_by(bon_id=bon.id).first()
    if log_entry:
        log_entry.societe     = bon.societe
        log_entry.fournisseur = bon.fournisseur
        log_entry.montant_eur = total_montant or None
    else:
        # Garantit la règle : chaque bon doit avoir une entrée de gestion
        db.session.add(CommandeLogistique(
            bon_id      = bon.id,
            ref_log     = bon.numero,
            societe     = bon.societe,
            annee       = str(date.today().year),
            fournisseur = bon.fournisseur,
            montant_eur = total_montant or None,
            cree_par    = session.get('username', ''),
        ))

    # Mettre à jour l'entrée Frais associée (ou la créer si absente)
    frais_entry = FraisLogistique.query.filter_by(bon_id=bon.id).first()
    if frais_entry:
        frais_entry.societe     = bon.societe
        frais_entry.fournisseur = bon.fournisseur
    else:
        db.session.add(FraisLogistique(
            bon_id      = bon.id,
            ref_log     = bon.numero,
            societe     = bon.societe,
            annee       = str(date.today().year),
            fournisseur = bon.fournisseur,
            cree_par    = session.get('username', ''),
        ))

    db.session.commit()
    return redirect(url_for('logistique_bons'))


@app.route('/api/logistique/bons/<int:bon_id>/delete', methods=['POST', 'DELETE'])
@role_required('admin')
def api_bon_delete(bon_id):
    bon = BonCommande.query.get_or_404(bon_id)
    CommandeLogistique.query.filter_by(bon_id=bon.id).delete()
    FraisLogistique.query.filter_by(bon_id=bon.id).delete()
    db.session.delete(bon)
    db.session.commit()
    return redirect(url_for('logistique_bons'))


@app.route('/logistique/bons/<int:bon_id>/print')
@login_required
def print_bon(bon_id):
    """Print bon de commande."""
    bon = BonCommande.query.get_or_404(bon_id)
    from datetime import datetime
    return render_template('bon_print.html', bon=bon, now=datetime.now())


@app.route('/api/logistique/bons/<int:bon_id>/detail')
@login_required
def api_bon_detail(bon_id):
    """Return bon de commande data as JSON (used by gestion page viewer)."""
    try:
        bon = BonCommande.query.get_or_404(bon_id)
        return jsonify(bon.to_dict())
    except Exception as e:
        app.logger.error(f'api_bon_detail error bon_id={bon_id}: {e}')
        return jsonify({'error': str(e)}), 500


# --- Suppression ---

@app.route('/delete/<int:op_id>', methods=['DELETE', 'POST'])
@role_required('admin')
def delete_operation(op_id):
    op = Operation.query.get_or_404(op_id)
    _log_audit(op.id, 'suppression', f"{op.type_operation} - {op.client} - {op.montant}")
    db.session.delete(op)
    db.session.commit()
    if request.headers.get('HX-Request'):
        return '<div class="alert alert-info fade-in"><i class="fas fa-trash mr-2"></i>Opération supprimée.</div>'
    flash('Opération supprimée.', 'info')
    return redirect(url_for('operations'))


# --- Consultation ---

def _build_operations_query(args):
    """Construit et retourne la query + pagination à partir d'un dict de params."""
    query = Operation.query

    search = args.get('search', '').strip()
    if search:
        query = query.filter(or_(
            Operation.client.ilike(f'%{search}%'),
            Operation.remettant.ilike(f'%{search}%'),
            Operation.banque.ilike(f'%{search}%'),
            Operation.numero_piece.ilike(f'%{search}%'),
            Operation.remarque.ilike(f'%{search}%'),
            Operation.societe.ilike(f'%{search}%'),
        ))

    type_op = args.get('type_operation', '').strip()
    if type_op:
        query = query.filter(Operation.type_operation == type_op)

    type_cheque = args.get('type_cheque', '').strip()
    if type_cheque:
        query = query.filter(Operation.type_operation == 'Chèque', Operation.type_detail == type_cheque)

    societe = args.get('societe', '').strip()
    if societe:
        query = query.filter(Operation.societe == societe)

    remettant = args.get('remettant', '').strip()
    if remettant:
        query = query.filter(Operation.remettant.ilike(f'%{remettant}%'))

    statut = args.get('statut', '').strip()
    if statut:
        query = query.filter(Operation.statut == statut)

    date_filter = args.get('date_filter', 'date_operation').strip()
    date_column = Operation.date_operation
    if date_filter == 'date_reception':
        query = query.filter(Operation.type_operation == 'Chèque')
        date_column = Operation.date_reception
    elif date_filter == 'date_echeance':
        query = query.filter(
            Operation.type_operation == 'Chèque',
            Operation.type_detail == 'À échéance',
        )
        date_column = Operation.date_encaissement

    date_debut = args.get('date_debut', '').strip()
    if date_debut:
        query = query.filter(date_column >= date_debut)

    date_fin = args.get('date_fin', '').strip()
    if date_fin:
        query = query.filter(date_column <= date_fin)

    # Tri: les dernières saisies doivent apparaître en haut.
    sort_col = args.get('sort', '').strip()
    sort_dir = args.get('dir', 'asc').strip()
    if sort_col != 'statut' and sort_dir not in ('asc', 'desc'):
        sort_dir = 'asc'

    ops_sort_columns = {
        'client': Operation.client,
        'societe': Operation.societe,
        'banque': Operation.banque,
        'montant': Operation.montant,
        'date_operation': Operation.date_operation,
        'date_reception': Operation.date_reception,
        'date_encaissement': Operation.date_encaissement,
    }
    if sort_col in ops_sort_columns:
        col = ops_sort_columns[sort_col]
        order = col.asc().nullslast() if sort_dir == 'asc' else col.desc().nullslast()
        query = query.order_by(order, Operation.id.desc())
    elif sort_col == 'statut':
        # Will sort in-memory after fetch
        pass
    else:
        query = query.order_by(
            Operation.date_creation.desc().nullslast(),
            Operation.id.desc(),
        )

    total_montant = db.session.query(func.sum(Operation.montant)).filter(
        Operation.id.in_(query.with_entities(Operation.id)),
        ~Operation.type_operation.in_(['Autre', 'Transfer']),
        ~Operation.statut.in_(['Rejeté'])
    ).scalar() or 0
    total_count = query.count()

    page = int(args.get('page', 1) or 1)
    per_page = _get_per_page(args, 'per_page', REF_PER_PAGE)
    offset = (page - 1) * per_page

    if sort_col == 'statut':
        all_ops = query.all()
        present_statuts = [s for s in STATUS_CHOICES if any(o.statut == s for o in all_ops)]
        if not present_statuts:
            present_statuts = STATUS_CHOICES
        statut_idx = 0
        try:
            statut_idx = int(sort_dir) % len(present_statuts)
        except (ValueError, ZeroDivisionError):
            pass
        rotated = present_statuts[statut_idx:] + present_statuts[:statut_idx]
        statut_order = {s: i for i, s in enumerate(rotated)}
        all_ops.sort(key=lambda o: (statut_order.get(o.statut, 99), -(o.date_creation.timestamp() if o.date_creation else 0), -o.id))
        operations = all_ops[offset: offset + per_page]
    else:
        operations = query.limit(per_page).offset(offset).all()

    total_pages = (total_count + per_page - 1) // per_page

    return {
        'operations': operations,
        'total_montant': total_montant,
        'total_count': total_count,
        'page': page,
        'total_pages': total_pages,
        'per_page': per_page,
        'per_page_options': PER_PAGE_OPTIONS,
        'sort_col': sort_col,
        'sort_dir': sort_dir,
    }


@app.route('/operations')
@login_required
def operations():
    _auto_update_echeance_statuts()
    notifications = _get_echeance_notifications()
    role = _current_role()
    rejections = _get_recent_rejections() if role != 'admin' else {'rejections': [], 'total': 0}
    ops_data = _build_operations_query(request.args)
    return render_template(
        'operations.html',
        notifications=notifications,
        rejections=rejections,
        is_admin=(role == 'admin'),
        bank_options=_get_bank_suggestions(),
        client_options=_get_client_suggestions(),
        remettant_options=_get_remettant_suggestions(),
        check_type_options=CHECK_TYPE_CHOICES,
        **ops_data,
    )


@app.route('/api/operations')
@login_required
def api_operations():
    data = _build_operations_query(request.args)
    role = _current_role()
    return render_template(
        'partials/operations_table.html',
        is_admin=role == 'admin',
        can_write=role in ('admin', 'saisie'),
        can_delete=role == 'admin',
        status_choices=STATUS_CHOICES,
        **data,
    )


@app.route('/api/notifications')
@login_required
def api_notifications():
    notifications = _get_echeance_notifications()
    return render_template(
        'partials/notifications_panel.html',
        notifications=notifications,
        is_admin=_current_role() == 'admin',
    )


@app.route('/api/notifications/badge')
@login_required
def api_notifications_badge():
    notifications = _get_echeance_notifications()
    log_notifications = _get_logistique_notifications()
    # Pour les non-admins, ajouter le compte des rejets récents
    rejections_count = 0 if _current_role() == 'admin' else _get_recent_rejections()['total']
    consultation_total = (notifications.get('total_alerts', 0) if notifications else 0) + rejections_count
    logistique_total = log_notifications.get('total_alerts', 0) if log_notifications else 0
    target_url = url_for('operations') if consultation_total > 0 else (
        url_for('logistique_gestion') if logistique_total > 0 else url_for('operations')
    )
    return render_template(
        'partials/global_notifications_badge.html',
        notifications=notifications,
        log_notifications=log_notifications,
        rejections_count=rejections_count,
        target_url=target_url,
    )


@app.route('/api/rejections')
@login_required
def api_rejections():
    """Retourne les rejets récents (sauf pour les admins)."""
    if _current_role() == 'admin':

        rejections_data = {'rejections': [], 'total': 0}  # Admins ne voient pas les rejets
    else:
        rejections_data = _get_recent_rejections()
    
    return render_template(
        'partials/rejections_panel.html',
        rejections=rejections_data,
    )

@app.route('/api/operations/<int:op_id>')
@login_required
def api_operation_detail(op_id):
    op = Operation.query.get_or_404(op_id)
    audits = AuditLog.query.filter_by(operation_id=op_id).order_by(AuditLog.date_action.desc()).all()
    if request.headers.get('HX-Request'):
        return render_template('partials/operation_detail.html', operation=op, audits=audits)
    return jsonify(op.to_dict())


@app.route('/api/operations/<int:op_id>/status', methods=['POST'])
@role_required('admin')
def api_update_operation_status(op_id):
    op = Operation.query.get_or_404(op_id)
    new_status = (request.form.get('statut') or '').strip()

    if new_status not in STATUS_CHOICES:
        return ('Statut invalide.', 400)

    if op.statut == new_status:
        return ('', 204)

    old_status = op.statut
    op.statut = new_status
    op.date_modification = datetime.utcnow()
    db.session.commit()

    # Log audit avec message spécial si c'est un rejet
    if new_status == 'Rejeté':
        audit_message = f"Statut modifié: {old_status} -> Rejeté (Chèque #{op.numero_piece} de {op.client} rejeté)"
    else:
        audit_message = f'Statut modifié: {old_status} -> {new_status}'
    
    _log_audit(op.id, 'modification', audit_message)
    return ('', 204)


# --- Import Excel ---

@app.route('/import', methods=['GET', 'POST'])
@role_required('admin', 'saisie')
def import_excel():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename.endswith(('.xlsx', '.xls')):
            flash('Veuillez sélectionner un fichier Excel (.xlsx)', 'error')
            return redirect(url_for('import_excel'))

        try:
            wb = load_workbook(file, read_only=True, data_only=True)
            imported = 0
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(min_row=2, values_only=True))
                for row in rows:
                    if not row or all(cell is None for cell in row):
                        continue
                    op = _parse_excel_row(row, sheet_name)
                    if op:
                        db.session.add(op)
                        imported += 1
            db.session.commit()
            flash(f'{imported} opérations importées avec succès !', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f'Erreur lors de l\'import : {str(e)}', 'error')

        return redirect(url_for('import_excel'))

    return render_template('import.html')


# --- Historique ---

@app.route('/historique')
@login_required
def historique():
    audits = AuditLog.query.order_by(AuditLog.date_action.desc()).limit(200).all()
    return render_template('historique.html', audits=audits)


# --- Export ---

@app.route('/export')
@login_required
def export_excel():
    query = Operation.query
    search = request.args.get('search', '').strip()
    if search:
        query = query.filter(or_(
            Operation.client.ilike(f'%{search}%'),
            Operation.remettant.ilike(f'%{search}%'),
            Operation.banque.ilike(f'%{search}%'),
            Operation.numero_piece.ilike(f'%{search}%'),
            Operation.remarque.ilike(f'%{search}%'),
            Operation.societe.ilike(f'%{search}%'),
        ))

    type_op = request.args.get('type_operation', '').strip()
    if type_op:
        query = query.filter(Operation.type_operation == type_op)

    type_cheque = request.args.get('type_cheque', '').strip()
    if type_cheque:
        query = query.filter(Operation.type_operation == 'Chèque', Operation.type_detail == type_cheque)

    societe = request.args.get('societe', '').strip()
    if societe:
        query = query.filter(Operation.societe == societe)

    statut = request.args.get('statut', '').strip()
    if statut:
        query = query.filter(Operation.statut == statut)

    date_debut = request.args.get('date_debut', '').strip()
    if date_debut:
        query = query.filter(Operation.date_operation >= date_debut)

    date_fin = request.args.get('date_fin', '').strip()
    if date_fin:
        query = query.filter(Operation.date_operation <= date_fin)

    date_echeance = request.args.get('date_echeance', '').strip()
    if date_echeance:
        query = query.filter(
            Operation.type_operation == 'Chèque',
            Operation.type_detail == 'À échéance',
            Operation.date_encaissement == date_echeance,
        )

    operations = query.order_by(Operation.date_operation.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Opérations"
    headers = ['ID', 'Date', 'Type', 'Société', 'Client', 'Remettant', 'Montant',
               'Banque', 'N° Pièce', 'Statut', 'Famille', 'Remarque', 'Saisi par']
    ws.append(headers)
    for op in operations:
        ws.append([
            op.id,
            op.date_operation.strftime('%d/%m/%Y') if op.date_operation else '',
            op.type_operation, op.societe, op.client, op.remettant or '',
            op.montant, op.banque or '', op.numero_piece or '',
            op.statut, op.famille or '', op.remarque or '', op.cree_par or '',
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f'operations_{date.today().strftime("%Y%m%d")}.xlsx'
    file_bytes = output.getvalue()
    response = make_response(file_bytes)
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return response


# --- Utilitaires ---

def _parse_date(date_str):
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return None


def _log_audit(operation_id, action, details=''):
    log = AuditLog(
        operation_id=operation_id,
        action=action,
        utilisateur=session.get('user_nom', 'Système'),
        details=details,
    )
    db.session.add(log)
    db.session.commit()


def _parse_excel_row(row, sheet_name):
    """Parse une ligne Excel en opération"""
    try:
        sheet_lower = sheet_name.lower()
        if 'chèque' in sheet_lower or 'cheque' in sheet_lower:
            type_op = 'Chèque'
        elif 'virement' in sheet_lower:
            type_op = 'Virement'
        elif 'versement' in sheet_lower:
            type_op = 'Versement'
        else:
            type_op = 'Versement'

        date_val = None
        montant_val = None
        client_val = None
        banque_val = None

        for cell in row:
            if isinstance(cell, datetime):
                date_val = cell.date()
            elif isinstance(cell, date):
                date_val = cell
            elif isinstance(cell, (int, float)) and cell > 10 and montant_val is None:
                montant_val = float(cell)
            elif isinstance(cell, str) and len(cell) > 2:
                if client_val is None:
                    client_val = cell
                elif banque_val is None:
                    banque_val = cell

        if not date_val or not montant_val or not client_val:
            return None

        return Operation(
            type_operation=type_op, societe='ENT',
            date_operation=date_val, client=client_val,
            montant=montant_val, banque=_normalize_bank_name(banque_val),
            statut='Encaissé', cree_par='Import Excel',
        )
    except Exception:
        return None


# --- Initialisation ---
with app.app_context():
    db.create_all()

    # Migration légère : ajoute la colonne date_etd si absente (DB déjà existante).
    with db.engine.connect() as _conn:
        _cols = [row[1] for row in _conn.exec_driver_sql("PRAGMA table_info(commandes_logistique)").fetchall()]
        if 'date_etd' not in _cols:
            _conn.exec_driver_sql("ALTER TABLE commandes_logistique ADD COLUMN date_etd DATE")
            _conn.commit()

        # Migration : refonte de frais_logistique pour le calcul du prix de revient.
        _frais_cols = [row[1] for row in _conn.exec_driver_sql("PRAGMA table_info(frais_logistique)").fetchall()]
        _new_frais_cols = {
            'taux_douane': 'FLOAT',
            'rps': 'FLOAT',
            'echange': 'FLOAT',
            'honoraires_transitaire': 'FLOAT',
            'magasinage': 'FLOAT',
            'surestaries': 'FLOAT',
            'pr_config': 'TEXT',
        }
        for _col, _type in _new_frais_cols.items():
            if _col not in _frais_cols:
                _conn.exec_driver_sql(f"ALTER TABLE frais_logistique ADD COLUMN {_col} {_type}")
        _conn.commit()

    # Normalise legacy roles to the new 3-role model.
    for u in User.query.all():
        normalized = _normalize_role(u.role)
        if u.role != normalized:
            u.role = normalized

    # Rename legacy admin account boss -> mehdi (keeping password hash).
    legacy_boss = User.query.filter_by(username='boss').first()
    mehdi_user = User.query.filter_by(username='mehdi').first()
    if legacy_boss and not mehdi_user:
        legacy_boss.username = 'mehdi'
        if not legacy_boss.nom_complet or legacy_boss.nom_complet.strip().lower() == 'boss':
            legacy_boss.nom_complet = 'Mehdi'
        legacy_boss.role = 'admin'

    if not User.query.filter_by(username='mehdi').first():
        mehdi = User(username='mehdi', nom_complet='Mehdi', role='admin')
        mehdi.set_password('srid2024mehdi')
        db.session.add(mehdi)

    if not User.query.filter_by(username='sabrina').first():
        sabrina = User(username='sabrina', nom_complet='Sabrina', role='saisie')
        sabrina.set_password('srid2024sab')
        db.session.add(sabrina)
    else:
        sabrina = User.query.filter_by(username='sabrina').first()
        sabrina.role = 'saisie'

    # Normalize legacy statuses to the approved status list.
    for op in Operation.query.all():
        normalized_statut = _normalize_legacy_statut(op)
        if op.statut != normalized_statut:
            op.statut = normalized_statut

    # Seed référentiels depuis les opérations existantes (idempotent).
    existing_clients = {c.nom.lower() for c in ClientLabel.query.all()}
    client_vals = db.session.query(Operation.client).filter(
        Operation.client.isnot(None), Operation.client != ''
    ).distinct().all()
    for (nom,) in client_vals:
        nom = (nom or '').strip()
        if nom and nom.lower() not in existing_clients:
            db.session.add(ClientLabel(nom=nom))
            existing_clients.add(nom.lower())

    existing_remettants = {r.nom.lower() for r in RemettantLabel.query.all()}
    remettant_vals = db.session.query(Operation.remettant).filter(
        Operation.remettant.isnot(None), Operation.remettant != ''
    ).distinct().all()
    for (nom,) in remettant_vals:
        nom = (nom or '').strip()
        if nom and nom.lower() not in existing_remettants:
            db.session.add(RemettantLabel(nom=nom))
            existing_remettants.add(nom.lower())

    db.session.commit()

    # Seed logistique depuis bon-md/logistics-data.js (idempotent, une seule fois)
    if CommandeLogistique.query.count() == 0:
        _js_path = os.path.abspath(os.path.join(
            os.path.dirname(os.path.abspath(__file__)), '..', 'bon-md', 'logistics-data.js'))
        if os.path.exists(_js_path):
            with open(_js_path, 'r', encoding='utf-8') as _f:
                _content = _f.read()
            _m = re.search(r'window\.LOGISTICS_SEED\s*=\s*(\[.*?\]);', _content, re.DOTALL)
            if _m:
                _data = json.loads(_m.group(1))
                def _pd(s):
                    if not s or not isinstance(s, str) or not s.strip():
                        return None
                    try:
                        return datetime.strptime(s.strip(), '%Y-%m-%d').date()
                    except ValueError:
                        return None
                def _pf(v):
                    try:
                        return float(v) if v is not None and v != '' else None
                    except (TypeError, ValueError):
                        return None
                def _pi(v):
                    try:
                        return int(float(v)) if v is not None and v != '' else None
                    except (TypeError, ValueError):
                        return None
                for _item in _data:
                    db.session.add(CommandeLogistique(
                        ref_log       = _item.get('id'),
                        societe       = _item.get('company', ''),
                        annee         = str(_item.get('year', '')) if _item.get('year') else None,
                        date_d10      = _pd(_item.get('dateD10')),
                        date_arrivee  = _pd(_item.get('arrivalDate')),
                        fournisseur   = (_item.get('supplier') or '').upper() or None,
                        produit       = _item.get('product') or None,
                        emballage     = _item.get('packaging') or None,
                        quantite      = _pf(_item.get('quantity')),
                        tva           = _pf(_item.get('vat')),
                        montant_eur   = _pf(_item.get('amountEur')),
                        cours         = _pf(_item.get('rate')),
                        date_facture  = _pd(_item.get('invoiceDate')),
                        code_paiement = _item.get('paymentCode') or None,
                        nb_jours      = _pi(_item.get('paymentDays')),
                        date_echeance = _pd(_item.get('dueDate')),
                        date_paiement = _pd(_item.get('paymentDate')),
                        date_valeur   = _pd(_item.get('valueDate')),
                        remarque      = _item.get('remark') or None,
                        cree_par      = 'seed',
                    ))
                db.session.commit()


if __name__ == '__main__':
    app.run(debug=True, port=5000)
