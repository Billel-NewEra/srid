from flask import Flask, render_template, request, redirect, url_for, jsonify, session, flash, send_file
from config import Config
from models import db, Operation, User, AuditLog
from datetime import datetime, date, timedelta
from sqlalchemy import or_, func, extract, desc
from functools import wraps
import io
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


# --- Routes Auth ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
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
    total_ops = Operation.query.count()
    total_montant = db.session.query(func.sum(Operation.montant)).scalar() or 0
    total_cheques = Operation.query.filter_by(type_operation='Chèque').count()
    total_virements = Operation.query.filter_by(type_operation='Virement').count()
    total_versements = Operation.query.filter_by(type_operation='Versement').count()
    en_attente = Operation.query.filter_by(statut='En attente').count()
    encaisses = Operation.query.filter_by(statut='Encaissé').count()
    rejetes = Operation.query.filter_by(statut='Rejeté').count()
    annules = Operation.query.filter_by(statut='Annulé').count()
    en_cours = Operation.query.filter_by(statut='En cours').count()

    # Taux d'encaissement
    taux_encaissement = round((encaisses / total_ops * 100), 1) if total_ops > 0 else 0

    # Montant moyen
    montant_moyen = round(total_montant / total_ops, 2) if total_ops > 0 else 0

    # Données par société
    montant_srid = db.session.query(func.sum(Operation.montant)).filter_by(societe='SRID').scalar() or 0
    montant_genetics = db.session.query(func.sum(Operation.montant)).filter_by(societe='Genetics').scalar() or 0
    ops_srid = Operation.query.filter_by(societe='SRID').count()
    ops_genetics = Operation.query.filter_by(societe='Genetics').count()

    # Top 5 clients par montant (année en cours, exclure les sociétés)
    top_clients = db.session.query(
        Operation.client,
        func.sum(Operation.montant).label('total'),
        func.count(Operation.id).label('nb_ops')
    ).filter(
        extract('year', Operation.date_operation) == current_year,
        Operation.client.isnot(None),
        ~Operation.client.in_(['SRID', 'Genetics', 'srid', 'genetics', ''])
    ).group_by(Operation.client).order_by(desc('total')).limit(5).all()

    # Dernières opérations
    dernieres = Operation.query.order_by(Operation.date_creation.desc()).limit(10).all()

    # Données mensuelles année en cours
    monthly_data = []
    for month in range(1, 13):
        total = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == current_year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        monthly_data.append(float(total))

    # Données mensuelles année précédente (comparaison N vs N-1)
    previous_year = current_year - 1
    monthly_data_prev = []
    for month in range(1, 13):
        total = db.session.query(func.sum(Operation.montant)).filter(
            extract('year', Operation.date_operation) == previous_year,
            extract('month', Operation.date_operation) == month
        ).scalar() or 0
        monthly_data_prev.append(float(total))

    # Données mensuelles par société (année en cours)
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

    # Activité 30 derniers jours (sparkline)
    today = date.today()
    daily_activity = []
    daily_labels = []
    for i in range(29, -1, -1):
        d = today - timedelta(days=i)
        count = Operation.query.filter(
            func.date(Operation.date_operation) == d
        ).count()
        daily_activity.append(count)
        daily_labels.append(d.strftime('%d/%m'))

    type_data = {'Chèque': total_cheques, 'Virement': total_virements, 'Versement': total_versements}
    statut_data = {
        'Encaissé': encaisses,
        'En attente': en_attente,
        'Rejeté': rejetes,
        'Annulé': annules,
        'En cours': en_cours
    }

    return render_template('dashboard.html',
                           total_ops=total_ops, total_montant=total_montant,
                           total_cheques=total_cheques, total_virements=total_virements,
                           total_versements=total_versements, en_attente=en_attente,
                           encaisses=encaisses, rejetes=rejetes, annules=annules,
                           en_cours=en_cours, taux_encaissement=taux_encaissement,
                           montant_moyen=montant_moyen,
                           montant_srid=float(montant_srid), montant_genetics=float(montant_genetics),
                           ops_srid=ops_srid, ops_genetics=ops_genetics,
                           top_clients=top_clients,
                           dernieres=dernieres,
                           monthly_data=monthly_data, monthly_data_prev=monthly_data_prev,
                           monthly_srid=monthly_srid, monthly_genetics=monthly_genetics,
                           type_data=type_data, statut_data=statut_data,
                           daily_activity=daily_activity, daily_labels=daily_labels,
                           today=date.today().strftime('%d/%m/%Y'),
                           year=current_year, previous_year=previous_year)


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
            banque=request.form.get('banque') or None,
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

    return render_template('saisie.html')


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
        op.banque = request.form.get('banque') or None
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
        return render_template('partials/edit_form.html', operation=op)
    return render_template('edit.html', operation=op)


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
            montant=montant_val, banque=banque_val,
            statut='En attente', cree_par='Import Excel',
        )
    except Exception:
        return None


# --- Initialisation ---
with app.app_context():
    db.create_all()
    if not User.query.filter_by(username='boss').first():
        boss = User(username='boss', nom_complet='Directeur', role='boss')
        boss.set_password('boss123')
        db.session.add(boss)
    if not User.query.filter_by(username='saisie').first():
        saisie_user = User(username='saisie', nom_complet='Agent de saisie', role='saisisseur')
        saisie_user.set_password('saisie123')
        db.session.add(saisie_user)
    db.session.commit()


if __name__ == '__main__':
    app.run(debug=True, port=5000)
