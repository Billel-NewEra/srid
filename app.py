from flask import Flask, render_template, request, redirect, url_for, jsonify, session, flash, send_file, send_from_directory
from config import Config
from models import db, Operation, User, AuditLog
from datetime import datetime, date, timedelta
from sqlalchemy import or_, func, extract, desc
from functools import wraps
import io
import re
from openpyxl import Workbook, load_workbook


app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)


# --- Décorateurs d'authentification ---

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


# --- Contexte global templates ---

@app.context_processor
def inject_globals():
    from datetime import datetime
    return {
        'current_user': session.get('user_nom', ''),
        'current_role': session.get('user_role', ''),
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
            session['user_id'] = user.id
            session['user_nom'] = user.nom_complet or user.username
            session['user_role'] = user.role
            flash(f'Bienvenue, {user.nom_complet or user.username} !', 'success')
            return redirect(url_for('dashboard'))
        flash('Identifiants incorrects.', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# --- Dashboard ---

@app.route('/')
@login_required
def dashboard():
    current_year = date.today().year

    # Années disponibles (toutes les années présentes en base)
    years_raw = db.session.query(
        extract('year', Operation.date_operation)
    ).distinct().order_by(extract('year', Operation.date_operation).desc()).all()
    available_years = [int(y[0]) for y in years_raw if y[0]]

    # --- KPIs par société ---
    total_ops = Operation.query.count()
    total_montant = db.session.query(func.sum(Operation.montant)).scalar() or 0
    ops_srid = Operation.query.filter_by(societe='SRID').count()
    ops_genetics = Operation.query.filter_by(societe='Genetics').count()
    montant_srid = db.session.query(func.sum(Operation.montant)).filter_by(societe='SRID').scalar() or 0
    montant_genetics = db.session.query(func.sum(Operation.montant)).filter_by(societe='Genetics').scalar() or 0

    # --- Statuts avec montants par société ---
    statuts_info = {}
    for statut in ['Encaissé', 'En attente', 'Rejeté', 'Annulé', 'En cours']:
        count = Operation.query.filter_by(statut=statut).count()
        montant = db.session.query(func.sum(Operation.montant)).filter_by(statut=statut).scalar() or 0
        count_srid = Operation.query.filter_by(statut=statut, societe='SRID').count()
        montant_srid_s = db.session.query(func.sum(Operation.montant)).filter_by(statut=statut, societe='SRID').scalar() or 0
        count_gen = Operation.query.filter_by(statut=statut, societe='Genetics').count()
        montant_gen_s = db.session.query(func.sum(Operation.montant)).filter_by(statut=statut, societe='Genetics').scalar() or 0
        statuts_info[statut] = {
            'count': count, 'montant': float(montant),
            'srid_count': count_srid, 'srid_montant': float(montant_srid_s),
            'genetics_count': count_gen, 'genetics_montant': float(montant_gen_s)
        }

    # --- Types avec montants par société ---
    types_info = {}
    for type_op in ['Chèque', 'Virement', 'Versement']:
        total_count = Operation.query.filter_by(type_operation=type_op).count()
        total_mt = db.session.query(func.sum(Operation.montant)).filter_by(type_operation=type_op).scalar() or 0
        srid_count = Operation.query.filter_by(type_operation=type_op, societe='SRID').count()
        srid_mt = db.session.query(func.sum(Operation.montant)).filter_by(type_operation=type_op, societe='SRID').scalar() or 0
        gen_count = Operation.query.filter_by(type_operation=type_op, societe='Genetics').count()
        gen_mt = db.session.query(func.sum(Operation.montant)).filter_by(type_operation=type_op, societe='Genetics').scalar() or 0
        types_info[type_op] = {
            'total': {'count': total_count, 'montant': float(total_mt)},
            'SRID': {'count': srid_count, 'montant': float(srid_mt)},
            'Genetics': {'count': gen_count, 'montant': float(gen_mt)}
        }

    # --- Taux d'encaissement ---
    taux_encaissement = round((statuts_info['Encaissé']['count'] / total_ops * 100), 1) if total_ops > 0 else 0

    # --- Top 5 clients (année en cours, tous mois) ---
    top_clients = db.session.query(
        Operation.client,
        func.sum(Operation.montant).label('total'),
        func.count(Operation.id).label('nb_ops')
    ).filter(
        extract('year', Operation.date_operation) == current_year,
        Operation.client.isnot(None),
        Operation.client != '',
        ~func.lower(Operation.client).like('%srid%'),
        ~func.lower(Operation.client).like('%genetics%')
    ).group_by(Operation.client).order_by(desc('total')).limit(5).all()

    # --- Dernières opérations (par date échéance desc) ---
    dernieres = Operation.query.order_by(Operation.date_operation.desc()).limit(10).all()

    # --- Données mensuelles année courante ---
    monthly_data = []
    for month in range(1, 13):
        total = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == current_year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        monthly_data.append(float(total))

    # --- Données mensuelles année précédente ---
    previous_year = current_year - 1
    monthly_data_prev = []
    for month in range(1, 13):
        total = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == previous_year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        monthly_data_prev.append(float(total))

    # --- Données mensuelles par société ---
    monthly_srid = []
    monthly_genetics = []
    for month in range(1, 13):
        t_srid = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == current_year,
            extract('month', Operation.date_operation) == month,
            Operation.societe == 'SRID'
        ).scalar() or 0
        t_genetics = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == current_year,
            extract('month', Operation.date_operation) == month,
            Operation.societe == 'Genetics'
        ).scalar() or 0
        monthly_srid.append(float(t_srid))
        monthly_genetics.append(float(t_genetics))

    # --- Statuts par mois (année courante) : total + SRID + Genetics ---
    status_labels = ['Encaissé', 'En attente', 'Rejeté', 'En cours']
    monthly_statuts_views = {
        'total': {s: [] for s in status_labels},
        'srid': {s: [] for s in status_labels},
        'genetics': {s: [] for s in status_labels},
    }
    for month in range(1, 13):
        for statut in status_labels:
            total_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == current_year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut
            ).scalar() or 0
            srid_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == current_year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut,
                Operation.societe == 'SRID'
            ).scalar() or 0
            genetics_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == current_year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut,
                Operation.societe == 'Genetics'
            ).scalar() or 0
            monthly_statuts_views['total'][statut].append(float(total_mt))
            monthly_statuts_views['srid'][statut].append(float(srid_mt))
            monthly_statuts_views['genetics'][statut].append(float(genetics_mt))

    monthly_statuts = monthly_statuts_views['total']

    # --- Activité 30 jours par montant et société ---
    today = date.today()
    daily_labels = []
    daily_total = []
    daily_srid = []
    daily_genetics = []
    for i in range(29, -1, -1):
        d = today - timedelta(days=i)
        mt_total = db.session.query(func.sum(Operation.montant)).filter(
            func.date(Operation.date_operation) == d
        ).scalar() or 0
        mt_srid = db.session.query(func.sum(Operation.montant)).filter(
            func.date(Operation.date_operation) == d,
            Operation.societe == 'SRID'
        ).scalar() or 0
        mt_gen = db.session.query(func.sum(Operation.montant)).filter(
            func.date(Operation.date_operation) == d,
            Operation.societe == 'Genetics'
        ).scalar() or 0
        daily_labels.append(d.strftime('%d/%m'))
        daily_total.append(float(mt_total))
        daily_srid.append(float(mt_srid))
        daily_genetics.append(float(mt_gen))

    return render_template('dashboard.html',
                           total_ops=total_ops, total_montant=float(total_montant),
                           ops_srid=ops_srid, ops_genetics=ops_genetics,
                           montant_srid=float(montant_srid), montant_genetics=float(montant_genetics),
                           statuts_info=statuts_info, types_info=types_info,
                           taux_encaissement=taux_encaissement,
                           top_clients=top_clients, dernieres=dernieres,
                           monthly_data=monthly_data, monthly_data_prev=monthly_data_prev,
                           monthly_srid=monthly_srid, monthly_genetics=monthly_genetics,
                           monthly_statuts=monthly_statuts,
                           monthly_statuts_views=monthly_statuts_views,
                           daily_labels=daily_labels, daily_total=daily_total,
                           daily_srid=daily_srid, daily_genetics=daily_genetics,
                           today=date.today().strftime('%d/%m/%Y'),
                           year=current_year, previous_year=previous_year,
                           available_years=available_years)


# --- API Dashboard (filtres HTMX/JS) ---

@app.route('/api/dashboard/monthly')
@login_required
def api_dashboard_monthly():
    year = request.args.get('year', date.today().year, type=int)
    prev_year = year - 1
    monthly_data = []
    monthly_data_prev = []
    for month in range(1, 13):
        total = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        total_prev = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == prev_year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        monthly_data.append(float(total))
        monthly_data_prev.append(float(total_prev))
    return jsonify({'year': year, 'prev_year': prev_year, 'data': monthly_data, 'data_prev': monthly_data_prev})


@app.route('/api/dashboard/societes')
@login_required
def api_dashboard_societes():
    year = request.args.get('year', date.today().year, type=int)
    monthly_srid = []
    monthly_genetics = []
    for month in range(1, 13):
        t_srid = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == year,
            extract('month', Operation.date_operation) == month,
            Operation.societe == 'SRID'
        ).scalar() or 0
        t_genetics = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == year,
            extract('month', Operation.date_operation) == month,
            Operation.societe == 'Genetics'
        ).scalar() or 0
        monthly_srid.append(float(t_srid))
        monthly_genetics.append(float(t_genetics))
    return jsonify({'year': year, 'srid': monthly_srid, 'genetics': monthly_genetics})


@app.route('/api/dashboard/statuts-monthly')
@login_required
def api_dashboard_statuts_monthly():
    year = request.args.get('year', date.today().year, type=int)
    status_labels = ['Encaissé', 'En attente', 'Rejeté', 'En cours']
    monthly_statuts_views = {
        'total': {s: [] for s in status_labels},
        'srid': {s: [] for s in status_labels},
        'genetics': {s: [] for s in status_labels},
    }
    for month in range(1, 13):
        for statut in status_labels:
            total_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut
            ).scalar() or 0
            srid_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut,
                Operation.societe == 'SRID'
            ).scalar() or 0
            genetics_mt = db.session.query(func.sum(Operation.montant)).filter(
                extract('year', Operation.date_operation) == year,
                extract('month', Operation.date_operation) == month,
                Operation.statut == statut,
                Operation.societe == 'Genetics'
            ).scalar() or 0
            monthly_statuts_views['total'][statut].append(float(total_mt))
            monthly_statuts_views['srid'][statut].append(float(srid_mt))
            monthly_statuts_views['genetics'][statut].append(float(genetics_mt))
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
        func.sum(Operation.montant).label('total'),
        func.count(Operation.id).label('nb_ops')
    ).filter(*filters).group_by(Operation.client).order_by(desc('total')).limit(5).all()
    max_total = float(top_clients[0][1]) if top_clients else 1
    html = ''
    for i, (client_name, total, nb_ops) in enumerate(top_clients, 1):
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
            <div class="badge badge-ghost badge-sm">{nb_ops} ops</div>
        </div>'''
    if not top_clients:
        html = '<p class="text-center opacity-50 py-4">Aucune donnée</p>'
    return html


@app.route('/api/dashboard/types')
@login_required
def api_dashboard_types():
    societe = request.args.get('societe', 'all')
    result = {}
    for type_op in ['Chèque', 'Virement', 'Versement']:
        q = db.session.query(func.sum(Operation.montant)).filter_by(type_operation=type_op)
        if societe != 'all':
            q = q.filter_by(societe=societe)
        result[type_op] = float(q.scalar() or 0)
    return jsonify(result)


# --- Saisie ---

@app.route('/saisie', methods=['GET', 'POST'])
@login_required
def saisie():
    if request.method == 'POST':
        op = Operation(
            type_operation=request.form.get('type_operation'),
            societe=request.form.get('societe'),
            famille=request.form.get('famille') or None,
            date_operation=_parse_date(request.form.get('date_operation')),
            date_reception=_parse_date(request.form.get('date_reception')),
            date_encaissement=_parse_date(request.form.get('date_encaissement')),
            date_sortie=_parse_date(request.form.get('date_sortie')),
            client=request.form.get('client'),
            remettant=request.form.get('remettant') or None,
            montant=float(request.form.get('montant', 0)),
            banque=_normalize_bank_name(request.form.get('banque')),
            numero_piece=request.form.get('numero_piece') or None,
            statut=request.form.get('statut', 'En attente'),
            type_detail=request.form.get('type_detail') or None,
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
        return redirect(url_for('saisie'))

    return render_template('saisie.html', bank_options=_get_bank_suggestions())


# --- Modification ---

@app.route('/edit/<int:op_id>', methods=['GET', 'POST'])
@login_required
def edit_operation(op_id):
    op = Operation.query.get_or_404(op_id)
    if request.method == 'POST':
        op.type_operation = request.form.get('type_operation')
        op.societe = request.form.get('societe')
        op.famille = request.form.get('famille') or None
        op.date_operation = _parse_date(request.form.get('date_operation'))
        op.date_reception = _parse_date(request.form.get('date_reception'))
        op.date_encaissement = _parse_date(request.form.get('date_encaissement'))
        op.date_sortie = _parse_date(request.form.get('date_sortie'))
        op.client = request.form.get('client')
        op.remettant = request.form.get('remettant') or None
        op.montant = float(request.form.get('montant', 0))
        op.banque = _normalize_bank_name(request.form.get('banque'))
        op.numero_piece = request.form.get('numero_piece') or None
        op.statut = request.form.get('statut', 'En attente')
        op.type_detail = request.form.get('type_detail') or None
        op.entree = request.form.get('entree') or None
        op.sortie = request.form.get('sortie') or None
        op.remarque = request.form.get('remarque') or None
        op.date_modification = datetime.utcnow()
        db.session.commit()
        _log_audit(op.id, 'modification', f"Modifié par {session.get('user_nom', '')}")

        if request.headers.get('HX-Request'):
            return render_template('partials/success_message.html', operation=op, action='modifiée')
        flash('Opération modifiée !', 'success')
        return redirect(url_for('consultation'))

    if request.headers.get('HX-Request'):
        return render_template('partials/edit_form.html', operation=op, bank_options=_get_bank_suggestions())
    return render_template('edit.html', operation=op, bank_options=_get_bank_suggestions())


# --- Suppression ---

@app.route('/delete/<int:op_id>', methods=['DELETE', 'POST'])
@login_required
def delete_operation(op_id):
    op = Operation.query.get_or_404(op_id)
    _log_audit(op.id, 'suppression', f"{op.type_operation} - {op.client} - {op.montant}")
    db.session.delete(op)
    db.session.commit()
    if request.headers.get('HX-Request'):
        return '<div class="alert alert-info fade-in"><i class="fas fa-trash mr-2"></i>Opération supprimée.</div>'
    flash('Opération supprimée.', 'info')
    return redirect(url_for('consultation'))


# --- Consultation ---

@app.route('/consultation')
@login_required
def consultation():
    return render_template('consultation.html')


@app.route('/api/operations')
@login_required
def api_operations():
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

    # Tri
    sort = request.args.get('sort', 'date_operation')
    order = request.args.get('order', 'desc')
    if hasattr(Operation, sort):
        col = getattr(Operation, sort)
        query = query.order_by(col.desc() if order == 'desc' else col.asc())
    else:
        query = query.order_by(Operation.date_operation.desc())

    # Totaux
    total_montant_filtre = db.session.query(func.sum(Operation.montant)).filter(
        Operation.id.in_(query.with_entities(Operation.id))
    ).scalar() or 0
    total_count = query.count()

    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = 25
    offset = (page - 1) * per_page
    operations = query.limit(per_page).offset(offset).all()
    total_pages = (total_count + per_page - 1) // per_page

    if request.headers.get('HX-Request'):
        return render_template('partials/operations_table.html',
                               operations=operations,
                               total_montant=total_montant_filtre,
                               total_count=total_count,
                               page=page,
                               total_pages=total_pages)

    return jsonify([op.to_dict() for op in operations])


@app.route('/api/operations/<int:op_id>')
@login_required
def api_operation_detail(op_id):
    op = Operation.query.get_or_404(op_id)
    audits = AuditLog.query.filter_by(operation_id=op_id).order_by(AuditLog.date_action.desc()).all()
    if request.headers.get('HX-Request'):
        return render_template('partials/operation_detail.html', operation=op, audits=audits)
    return jsonify(op.to_dict())


# --- Import Excel ---

@app.route('/import', methods=['GET', 'POST'])
@login_required
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
    type_op = request.args.get('type_operation', '').strip()
    if type_op:
        query = query.filter(Operation.type_operation == type_op)
    societe = request.args.get('societe', '').strip()
    if societe:
        query = query.filter(Operation.societe == societe)

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
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


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
            statut='En attente', cree_par='Import Excel',
        )
    except Exception:
        return None


# --- Initialisation ---
with app.app_context():
    db.create_all()
    if not User.query.filter_by(username='boss').first():
        boss = User(username='boss', nom_complet='Boss', role='boss')
        boss.set_password('srid2024boss')
        db.session.add(boss)
    if not User.query.filter_by(username='sabrina').first():
        sabrina = User(username='sabrina', nom_complet='Sabrina', role='saisisseur')
        sabrina.set_password('srid2024sab')
        db.session.add(sabrina)
    db.session.commit()


if __name__ == '__main__':
    app.run(debug=True, port=5000)
