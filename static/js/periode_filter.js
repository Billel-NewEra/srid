/* Filtre « Période » réutilisable (Type de date + popover calendrier flatpickr).
 *
 * Initialise un composant identique à celui de la page Opérations sur n'importe
 * quelle table filtrée par HTMX. Le markup correspondant est fourni par le
 * partial `partials/periode_filter.html` (ids préfixés par cfg.prefix).
 *
 * Config attendue :
 *   prefix   : préfixe d'ID unique (ex. 'bons', 'gestion')
 *   url      : endpoint HTMX à recharger
 *   target   : sélecteur de la cible HTMX (ex. '#bons-table')
 *   include  : tableau de sélecteurs des autres champs de filtre à transmettre
 *   onRefresh: callback optionnel exécuté après chaque rafraîchissement
 */
(function () {
    'use strict';
    if (window.initPeriodeFilter) return;

    window.initPeriodeFilter = function (cfg) {
        var p = cfg.prefix;
        var byId = function (suffix) { return document.getElementById(p + suffix); };

        var calEl = byId('-periode-calendar');
        if (!calEl || typeof flatpickr === 'undefined') return;

        var dropdown = byId('-periode-dropdown');
        var btn      = byId('-periode-btn');
        var panel    = byId('-periode-panel');
        var btnLabel = byId('-periode-label');
        var colSel   = byId('-periode-col');
        var preview  = byId('-periode-preview');
        var hidCol   = byId('-date-filter');
        var hidDeb   = byId('-date-debut');
        var hidFin   = byId('-date-fin');

        var pad  = function (n) { return String(n).padStart(2, '0'); };
        var iso  = function (d) { return d ? d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) : ''; };
        var fmt  = function (d) { return d ? pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear() : ''; };
        var fmtS = function (d) { return d ? pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + pad(d.getFullYear() % 100) : ''; };

        var selStart = null, selEnd = null;

        function openPanel()  { panel.classList.remove('hidden'); }
        function closePanel() { panel.classList.add('hidden'); }
        function isOpen()     { return !panel.classList.contains('hidden'); }

        // Voir la note détaillée dans operations.html : la navigation mois de
        // Flatpickr émet un clic synthétique redirigé vers le bouton toggle.
        var calInteracting = false;
        btn.addEventListener('mousedown', function () { calInteracting = false; });
        btn.addEventListener('click', function (e) {
            if (calInteracting) { calInteracting = false; return; }
            e.stopPropagation();
            if (isOpen()) closePanel(); else openPanel();
        });
        panel.addEventListener('mousedown', function (e) {
            e.stopPropagation();
            calInteracting = !!(e.target.closest && e.target.closest('.flatpickr-calendar'));
        });
        document.addEventListener('mousedown', function (e) {
            if (!isOpen()) return;
            if (dropdown.contains(e.target)) return;
            if (e.target.closest && e.target.closest('.flatpickr-calendar')) return;
            closePanel();
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' && isOpen()) closePanel();
        });

        if (hidCol && colSel && hidCol.value) colSel.value = hidCol.value;

        function refresh() {
            var params = new URLSearchParams();
            (cfg.include || []).forEach(function (sel) {
                var f = document.querySelector(sel);
                if (f && f.value) params.set(f.name, f.value);
            });
            if (hidCol.value) params.set('date_filter', hidCol.value);
            if (hidDeb.value) params.set('date_debut', hidDeb.value);
            if (hidFin.value) params.set('date_fin', hidFin.value);
            if (typeof htmx !== 'undefined') {
                htmx.ajax('GET', cfg.url + '?' + params.toString(), { target: cfg.target, swap: 'innerHTML' });
            }
            if (typeof cfg.onRefresh === 'function') cfg.onRefresh();
        }

        var fp = flatpickr(calEl, {
            inline: true, mode: 'range', locale: 'fr', showMonths: 1,
            onChange: function (dates) {
                selStart = dates[0] || null;
                selEnd = dates[1] || dates[0] || null;
                updatePreview();
            }
        });

        function updatePreview() {
            if (selStart && selEnd) {
                preview.textContent = (iso(selStart) === iso(selEnd))
                    ? ('Le ' + fmt(selStart))
                    : ('Du ' + fmt(selStart) + ' au ' + fmt(selEnd));
            } else {
                preview.textContent = 'Aucune date';
            }
        }

        function computeLabel() {
            if (!selStart) return 'Toutes les dates';
            if (iso(selStart) === iso(selEnd)) return fmtS(selStart);
            if (selStart.getMonth() === 0 && selStart.getDate() === 1 &&
                selEnd.getMonth() === 11 && selEnd.getDate() === 31 &&
                selStart.getFullYear() === selEnd.getFullYear()) {
                return String(selStart.getFullYear());
            }
            return fmtS(selStart) + '–' + fmtS(selEnd);
        }

        if (colSel) colSel.addEventListener('change', function () {
            hidCol.value = colSel.value;
            // Sans période active, le type ne filtre rien : pas de rafraîchissement.
            if (!selStart) return;
            refresh();
        });

        function setYear(y) {
            selStart = new Date(y, 0, 1);
            selEnd = new Date(y, 11, 31);
            fp.setDate([selStart, selEnd], false);
            updatePreview();
        }

        function setPreset(pr) {
            var now = new Date(), s = null, e = null;
            if (pr === 'all') { selStart = selEnd = null; fp.clear(); updatePreview(); return; }
            if (pr === 'thisMonth') { s = new Date(now.getFullYear(), now.getMonth(), 1); e = new Date(now.getFullYear(), now.getMonth() + 1, 0); }
            if (pr === 'thisYear')  { s = new Date(now.getFullYear(), 0, 1); e = new Date(now.getFullYear(), 11, 31); }
            selStart = s; selEnd = e;
            fp.setDate([s, e], false);
            updatePreview();
        }

        dropdown.querySelectorAll('[data-preset]').forEach(function (b) {
            b.addEventListener('click', function () { setPreset(b.dataset.preset); });
        });
        dropdown.querySelectorAll('[data-year]').forEach(function (b) {
            b.addEventListener('click', function () { setYear(parseInt(b.dataset.year, 10)); });
        });

        byId('-periode-clear').addEventListener('click', function () {
            selStart = selEnd = null;
            fp.clear();
            updatePreview();
            btnLabel.textContent = 'Toutes les dates';
            hidDeb.value = '';
            hidFin.value = '';
            refresh();
        });

        byId('-periode-apply').addEventListener('click', function () {
            btnLabel.textContent = computeLabel();
            if (colSel) hidCol.value = colSel.value;
            hidDeb.value = selStart ? iso(selStart) : '';
            hidFin.value = selEnd ? iso(selEnd) : '';
            closePanel();
            refresh();
        });

        // Pré-remplissage initial (lien externe, rechargement avec filtres).
        if (hidDeb.value && hidFin.value) {
            selStart = fp.parseDate(hidDeb.value, 'Y-m-d');
            selEnd = fp.parseDate(hidFin.value, 'Y-m-d');
            fp.setDate([selStart, selEnd], false);
            updatePreview();
            btnLabel.textContent = computeLabel();
        }
    };
})();
