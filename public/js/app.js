(function () {
  // Raccourci Ctrl/Cmd + K : recherche globale
  var q = document.getElementById('q');
  document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k' && q) { e.preventDefault(); q.focus(); q.select(); }
  });

  // Onglets (fiche dossier) : mémorisés dans l'adresse (#demandes)
  var tabs = document.querySelectorAll('[data-tab]');
  function showTab(name) {
    var found = false;
    tabs.forEach(function (t) { var on = t.dataset.tab === name; t.classList.toggle('on', on); found = found || on; });
    document.querySelectorAll('.tabpane').forEach(function (p) { p.classList.toggle('on', p.id === 'tab-' + name); });
    return found;
  }
  if (tabs.length) {
    tabs.forEach(function (t) { t.addEventListener('click', function (e) { e.preventDefault(); showTab(t.dataset.tab); history.replaceState(null, '', '#' + t.dataset.tab); }); });
    var h = location.hash.slice(1);
    if (!h || !showTab(h)) showTab(tabs[0].dataset.tab);
  }

  // Assistant de création
  var wiz = document.querySelectorAll('.wiz');
  if (!wiz.length) return;
  var stepsEl = document.querySelectorAll('#wsteps .step'), i = 0;
  var next = document.getElementById('next'), prev = document.getElementById('prev'), submit = document.getElementById('submit');
  function show(n) {
    i = Math.max(0, Math.min(wiz.length - 1, n));
    wiz.forEach(function (w, k) { w.classList.toggle('on', k === i); });
    stepsEl.forEach(function (s, k) {
      s.className = 'step' + (k < i ? ' done' : k === i ? ' cur' : '');
      s.querySelector('i').textContent = k < i ? '✓' : k + 1;
    });
    var last = i === wiz.length - 1;
    next.hidden = last; submit.hidden = !last; prev.hidden = i === 0;
    if (last) summary();
    scrollTo(0, 0);
  }
  function val(name) {
    var el = document.querySelector('[name$="[' + name + ']"]');
    if (!el) return '';
    if (el.type === 'checkbox') return el.checked ? 'Oui' : 'Non';
    if (el.tagName === 'SELECT') return el.selectedIndex > 0 ? el.options[el.selectedIndex].text : '';
    return el.value;
  }
  function summary() {
    var rows = [
      ['Dossier', val('numero')],
      ['Demandeur', (val('prenom') + ' ' + val('nom')).trim()],
      ['Adresse', [val('adresse1'), val('codePostal'), val('ville')].filter(Boolean).join(' ')],
      ['Foyer', val('situationFamiliale') + (val('conjoint') === 'Oui' ? ' · avec ' + (val('conjointPrenom') + ' ' + val('conjointNom')).trim() : '')],
      ['Projet', [val('nombreEnfants') && val('nombreEnfants') + ' enfant(s)', val('nature'), val('fratrie') === 'Oui' ? 'fratrie acceptée' : ''].filter(Boolean).join(', ')],
      ['Date de la demande', val('dateDemande')]
    ];
    var dl = document.getElementById('recap'); dl.innerHTML = '';
    rows.forEach(function (r) {
      var dt = document.createElement('dt'), dd = document.createElement('dd');
      dt.textContent = r[0]; dd.textContent = r[1] || '—'; dl.appendChild(dt); dl.appendChild(dd);
    });
  }
  var conj = document.querySelector('[name$="[conjoint]"]'), conjBox = document.getElementById('conjoint-box');
  function toggleConj() { conjBox.hidden = !conj.checked; }
  if (conj) { conj.addEventListener('change', toggleConj); toggleConj(); }

  next.addEventListener('click', function () { show(i + 1); });
  prev.addEventListener('click', function () { show(i - 1); });
  // En cas d'erreur de validation côté serveur, on revient à l'étape concernée
  var err = -1; wiz.forEach(function (w, k) { if (err < 0 && w.querySelector('ul li')) err = k; });
  show(err >= 0 ? err : 0);
})();
