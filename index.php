<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Missions ACOBI</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700&display=swap" rel="stylesheet">
  <!-- SheetJS pour lire les fichiers Excel -->
  <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --bg:        #1a2340;
      --card-bg:   #212b4c;
      --card-alt:  #263155;
      --border:    rgba(255,255,255,0.08);
      --accent:    #5DD8B8;
      --accent2:   #4ac8ff;
      --text:      #ffffff;
      --text-muted: rgba(255,255,255,0.55);
      --danger:    #ff6b6b;
      --success:   #5DD8B8;
      --warning:   #ffd166;
      --radius:    14px;
    }

    body {
      font-family: 'Nunito', sans-serif;
      background-color: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }

    /* ── HEADER ── */
    header {
      padding: 2rem 3rem 0;
    }
    .header-badge {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 0.4rem;
    }
    header h1 {
      font-size: 2.2rem;
      font-weight: 300;
    }
    header h1 span { color: var(--accent); font-weight: 700; }

    /* ── TABS ── */
    .tabs {
      display: flex;
      gap: 0.5rem;
      padding: 2rem 3rem 0;
      border-bottom: 1px solid var(--border);
    }
    .tab-btn {
      padding: 0.7rem 1.6rem;
      border-radius: 10px 10px 0 0;
      border: none;
      background: transparent;
      color: var(--text-muted);
      font-family: 'Nunito', sans-serif;
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }
    .tab-btn:hover { color: var(--text); }
    .tab-btn.active {
      background: var(--card-bg);
      color: var(--accent);
      border: 1px solid var(--border);
      border-bottom: 1px solid var(--card-bg);
    }

    /* ── CONTENT ── */
    .tab-content { display: none; padding: 2.5rem 3rem; }
    .tab-content.active { display: block; }

    /* ── SECTION TITLES ── */
    .section-title {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 1.2rem;
      margin-top: 2rem;
    }
    .section-title:first-child { margin-top: 0; }

    /* ── CARDS GRID ── */
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 1.2rem;
    }

    /* ── CARD ── */
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.5rem;
      transition: transform 0.2s, border-color 0.2s;
      position: relative;
    }
    .card:hover { transform: translateY(-3px); border-color: rgba(93,216,184,0.3); }

    .card-logo {
      width: 56px;
      height: 56px;
      border-radius: 10px;
      background: var(--card-alt);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.6rem;
      margin-bottom: 1rem;
      overflow: hidden;
    }
    .card-logo img { width: 100%; height: 100%; object-fit: contain; padding: 4px; }

    .card-title {
      font-size: 1rem;
      font-weight: 700;
      margin-bottom: 0.3rem;
    }
    .card-client {
      font-size: 0.85rem;
      color: var(--accent);
      font-weight: 600;
      margin-bottom: 0.5rem;
    }
    .card-collab {
      font-size: 0.82rem;
      color: var(--text-muted);
      margin-bottom: 0.8rem;
    }
    .card-dates {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-bottom: 1rem;
    }
    .card-arrow {
      color: var(--accent);
      font-size: 1.1rem;
      margin-top: 0.5rem;
    }

    /* ── BADGE STATUT ── */
    .badge {
      display: inline-block;
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      font-size: 0.75rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      margin-bottom: 0.8rem;
    }
    .badge-encours  { background: rgba(74,200,255,0.15); color: var(--accent2); }
    .badge-terminee { background: rgba(93,216,184,0.15); color: var(--success); }

    /* ── FILTRES ── */
    .filters {
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
      margin-bottom: 2rem;
      align-items: flex-end;
    }
    .filter-group { display: flex; flex-direction: column; gap: 0.3rem; }
    .filter-group label {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--text-muted);
      font-weight: 600;
    }
    .filter-group select, .filter-group input[type=text] {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text);
      padding: 0.5rem 0.9rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      cursor: pointer;
      min-width: 160px;
    }
    .filter-group select:focus, .filter-group input:focus {
      outline: none;
      border-color: var(--accent);
    }

    /* ── FORMULAIRES ── */
    .form-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.8rem;
      margin-bottom: 1.5rem;
    }
    .form-card h3 {
      font-size: 1rem;
      font-weight: 700;
      margin-bottom: 1.2rem;
      color: var(--accent);
    }
    .form-row {
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
      margin-bottom: 1rem;
    }
    .form-group {
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
      flex: 1;
      min-width: 180px;
    }
    .form-group label {
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.07em;
    }
    .form-group input,
    .form-group select,
    .form-group textarea {
      background: var(--card-alt);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text);
      padding: 0.6rem 0.9rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
    }
    .form-group input:focus,
    .form-group select:focus,
    .form-group textarea:focus {
      outline: none;
      border-color: var(--accent);
    }
    .form-group select[multiple] { height: 100px; }

    /* ── BOUTONS ── */
    .btn {
      padding: 0.6rem 1.4rem;
      border-radius: 8px;
      border: none;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      font-weight: 700;
      cursor: pointer;
      transition: opacity 0.2s, transform 0.1s;
    }
    .btn:hover { opacity: 0.85; transform: translateY(-1px); }
    .btn-primary { background: var(--accent); color: var(--bg); }
    .btn-secondary { background: var(--card-alt); color: var(--text); border: 1px solid var(--border); }
    .btn-danger { background: rgba(255,107,107,0.15); color: var(--danger); border: 1px solid rgba(255,107,107,0.3); }
    .btn-sm { padding: 0.3rem 0.8rem; font-size: 0.8rem; }

    /* ── TABLES ── */
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.9rem;
    }
    .data-table th {
      text-align: left;
      padding: 0.7rem 1rem;
      background: var(--card-alt);
      color: var(--text-muted);
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-weight: 700;
    }
    .data-table th:first-child { border-radius: 8px 0 0 8px; }
    .data-table th:last-child  { border-radius: 0 8px 8px 0; }
    .data-table td {
      padding: 0.7rem 1rem;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }
    .data-table tr:last-child td { border-bottom: none; }
    .data-table .logo-mini {
      width: 32px; height: 32px;
      border-radius: 6px;
      background: var(--card-alt);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 1rem;
      overflow: hidden;
      vertical-align: middle;
      margin-right: 0.5rem;
    }
    .data-table .logo-mini img { width: 100%; height: 100%; object-fit: contain; }

    /* ── IMPORT EXCEL ── */
    .import-zone {
      border: 2px dashed var(--border);
      border-radius: var(--radius);
      padding: 2rem;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s;
      margin-bottom: 1rem;
    }
    .import-zone:hover, .import-zone.drag-over { border-color: var(--accent); }
    .import-zone .icon { font-size: 2rem; margin-bottom: 0.5rem; }
    .import-zone p { color: var(--text-muted); font-size: 0.9rem; }
    .import-zone strong { color: var(--accent); }

    /* ── TOAST ── */
    #toast {
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: var(--accent);
      color: var(--bg);
      padding: 0.8rem 1.5rem;
      border-radius: 10px;
      font-weight: 700;
      font-size: 0.9rem;
      opacity: 0;
      transform: translateY(10px);
      transition: all 0.3s;
      z-index: 1000;
      pointer-events: none;
    }
    #toast.show { opacity: 1; transform: translateY(0); }
    #toast.error { background: var(--danger); color: #fff; }

    /* ── EMPTY STATE ── */
    .empty-state {
      text-align: center;
      padding: 4rem 2rem;
      color: var(--text-muted);
    }
    .empty-state .icon { font-size: 3rem; margin-bottom: 1rem; }

    /* ── MODAL ── */
    .modal-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.6);
      z-index: 500;
      align-items: center;
      justify-content: center;
    }
    .modal-overlay.open { display: flex; }
    .modal {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 2rem;
      width: 90%;
      max-width: 520px;
      max-height: 90vh;
      overflow-y: auto;
    }
    .modal h3 { color: var(--accent); margin-bottom: 1.2rem; font-size: 1.1rem; }
    .modal-actions { display: flex; gap: 0.8rem; justify-content: flex-end; margin-top: 1.5rem; }

    .stat-row {
      display: flex;
      gap: 1.5rem;
      margin-bottom: 2rem;
      flex-wrap: wrap;
    }
    .stat-chip {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 0.8rem 1.4rem;
      display: flex;
      flex-direction: column;
      gap: 0.2rem;
    }
    .stat-chip .val { font-size: 1.6rem; font-weight: 700; color: var(--accent); }
    .stat-chip .lbl { font-size: 0.78rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.07em; }
  </style>
</head>
<body>

<header>
  <div class="header-badge">ACOBI</div>
  <h1>Suivi des <span>missions</span>.</h1>
</header>

<nav class="tabs">
  <button class="tab-btn active" onclick="showTab('visu')">🗂 Visualisation</button>
  <button class="tab-btn" onclick="showTab('param')">⚙️ Paramétrage</button>
</nav>

<!-- ════════════════════════════════════════
     ONGLET VISUALISATION
════════════════════════════════════════ -->
<div id="tab-visu" class="tab-content active">

  <div class="stat-row" id="stats-row">
    <!-- rempli par JS -->
  </div>

  <div class="filters">
    <div class="filter-group">
      <label>Statut</label>
      <select id="f-statut" onchange="renderCards()">
        <option value="">Tous</option>
        <option value="en_cours">En cours</option>
        <option value="terminee">Terminée</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Client</label>
      <select id="f-client" onchange="renderCards()">
        <option value="">Tous</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Collaborateur</label>
      <select id="f-collab" onchange="renderCards()">
        <option value="">Tous</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Recherche</label>
      <input type="text" id="f-search" placeholder="Nom de mission…" oninput="renderCards()">
    </div>
  </div>

  <div class="grid" id="missions-grid">
    <!-- rempli par JS -->
  </div>
</div>

<!-- ════════════════════════════════════════
     ONGLET PARAMETRAGE
════════════════════════════════════════ -->
<div id="tab-param" class="tab-content">

  <!-- COLLABORATEURS -->
  <div class="section-title">👤 Collaborateurs</div>
  <div class="form-card">
    <h3>Ajouter un collaborateur</h3>
    <div class="form-row">
      <div class="form-group">
        <label>Prénom</label>
        <input type="text" id="c-prenom" placeholder="ex. Marie">
      </div>
      <div class="form-group">
        <label>Nom</label>
        <input type="text" id="c-nom" placeholder="ex. Dupont">
      </div>
      <div class="form-group">
        <label>Rôle / Poste</label>
        <input type="text" id="c-role" placeholder="ex. Consultante">
      </div>
    </div>
    <button class="btn btn-primary" onclick="addCollaborateur()">Ajouter</button>
  </div>

  <table class="data-table" id="table-collabs">
    <thead><tr><th>Prénom Nom</th><th>Rôle</th><th>Actions</th></tr></thead>
    <tbody></tbody>
  </table>

  <!-- CLIENTS -->
  <div class="section-title" style="margin-top:2.5rem">🏢 Clients</div>
  <div class="form-card">
    <h3>Ajouter un client</h3>
    <div class="form-row">
      <div class="form-group">
        <label>Nom du client</label>
        <input type="text" id="cl-nom" placeholder="ex. EDF">
      </div>
      <div class="form-group">
        <label>Logo (URL ou emoji)</label>
        <input type="text" id="cl-logo" placeholder="ex. 🏭 ou https://…/logo.png" oninput="document.getElementById('cl-logo-preview').innerHTML = logoHtml(this.value, 36, document.getElementById('cl-nom').value)">
      </div>
      <div class="form-group" style="flex:0;min-width:auto;justify-content:flex-end">
        <label>Ou importer</label>
        <button class="btn btn-secondary" type="button" onclick="document.getElementById('cl-logo-file').click()">📁 Fichier</button>
        <input type="file" id="cl-logo-file" accept="image/*" style="display:none" onchange="loadLogoFile(this, 'cl-logo', 'cl-logo-preview', document.getElementById('cl-nom').value)">
      </div>
      <div class="form-group" style="flex:0;min-width:auto;align-items:center;justify-content:flex-end">
        <label>Aperçu</label>
        <div id="cl-logo-preview" style="width:36px;height:36px;display:flex;align-items:center;justify-content:center"></div>
      </div>
    </div>
    <button class="btn btn-primary" onclick="addClient()">Ajouter</button>
  </div>

  <table class="data-table" id="table-clients">
    <thead><tr><th>Client</th><th>Logo</th><th>Actions</th></tr></thead>
    <tbody></tbody>
  </table>

  <!-- MISSIONS -->
  <div class="section-title" style="margin-top:2.5rem">📋 Missions</div>

  <!-- Import Excel -->
  <div class="form-card">
    <h3>📥 Importer depuis Excel</h3>
    <p style="color:var(--text-muted);font-size:0.85rem;margin-bottom:1rem;">
      Colonnes attendues : <strong>Titre</strong>, <strong>Client</strong>, <strong>Collaborateurs</strong> (séparés par « ; »), <strong>Date début</strong>, <strong>Date fin</strong>, <strong>Statut</strong> (en_cours / terminee)
    </p>
    <div class="import-zone" id="drop-zone" onclick="document.getElementById('excel-input').click()"
         ondragover="event.preventDefault();this.classList.add('drag-over')"
         ondragleave="this.classList.remove('drag-over')"
         ondrop="handleDrop(event)">
      <div class="icon">📊</div>
      <p><strong>Cliquez</strong> ou déposez votre fichier Excel ici</p>
      <p style="margin-top:0.3rem;font-size:0.8rem;">.xlsx ou .xls acceptés</p>
    </div>
    <input type="file" id="excel-input" accept=".xlsx,.xls" style="display:none" onchange="handleExcelFile(this.files[0])">
  </div>

  <!-- Formulaire mission -->
  <div class="form-card">
    <h3>Saisir une mission manuellement</h3>
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label>Titre de la mission</label>
        <input type="text" id="m-titre" placeholder="ex. Transformation digitale">
      </div>
      <div class="form-group">
        <label>Statut</label>
        <select id="m-statut">
          <option value="en_cours">En cours</option>
          <option value="terminee">Terminée</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Client</label>
        <select id="m-client"><option value="">-- Sélectionner --</option></select>
      </div>
      <div class="form-group">
        <label>Collaborateurs</label>
        <select id="m-collabs" multiple title="Maintenez Ctrl/Cmd pour sélectionner plusieurs"></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date de début</label>
        <input type="date" id="m-debut">
      </div>
      <div class="form-group">
        <label>Date de fin</label>
        <input type="date" id="m-fin">
      </div>
    </div>
    <button class="btn btn-primary" onclick="addMission()">Ajouter la mission</button>
  </div>

  <table class="data-table" id="table-missions">
    <thead><tr><th>Mission</th><th>Client</th><th>Collaborateurs</th><th>Période</th><th>Statut</th><th>Actions</th></tr></thead>
    <tbody></tbody>
  </table>

</div>

<!-- TOAST -->
<div id="toast"></div>

<!-- MODAL CONFIRMATION SUPPRESSION -->
<div class="modal-overlay" id="modal-confirm">
  <div class="modal">
    <h3>Confirmer la suppression</h3>
    <p id="modal-msg" style="color:var(--text-muted);line-height:1.6"></p>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeModal()">Annuler</button>
      <button class="btn btn-danger" id="modal-confirm-btn">Supprimer</button>
    </div>
  </div>
</div>

<!-- MODAL MODIFICATION CLIENT -->
<div class="modal-overlay" id="modal-edit-client">
  <div class="modal">
    <h3>✏️ Modifier le client</h3>
    <input type="hidden" id="edit-cl-id">
    <!-- Champ caché pour stocker la valeur du logo -->
    <input type="hidden" id="edit-cl-logo">
    <div class="form-row" style="align-items:center;gap:1.5rem">
      <div class="form-group" style="flex:1">
        <label>Nom du client</label>
        <input type="text" id="edit-cl-nom" placeholder="ex. EDF">
      </div>
      <div class="form-group" style="flex:0;min-width:auto">
        <label>Logo actuel</label>
        <div id="edit-cl-logo-preview" style="width:56px;height:56px;display:flex;align-items:center;justify-content:center;background:var(--card-alt);border-radius:10px;overflow:hidden"></div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:1rem;margin-top:0.5rem">
      <button class="btn btn-secondary" type="button" onclick="document.getElementById('edit-cl-logo-file').click()">📁 Changer le logo</button>
      <input type="file" id="edit-cl-logo-file" accept="image/*" style="display:none" onchange="loadLogoFile(this, 'edit-cl-logo', 'edit-cl-logo-preview')">
      <span id="edit-cl-file-label" style="display:none;color:var(--accent);font-size:0.85rem;font-weight:600">✓ Nouveau logo prêt</span>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeEditClientModal()">Annuler</button>
      <button class="btn btn-primary" onclick="saveEditClient()">Enregistrer</button>
    </div>
  </div>
</div>

<!-- MODAL MODIFICATION MISSION -->
<div class="modal-overlay" id="modal-edit">
  <div class="modal">
    <h3>✏️ Modifier la mission</h3>
    <input type="hidden" id="edit-id">
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label>Titre de la mission</label>
        <input type="text" id="edit-titre" placeholder="ex. Transformation digitale">
      </div>
      <div class="form-group">
        <label>Statut</label>
        <select id="edit-statut">
          <option value="en_cours">En cours</option>
          <option value="terminee">Terminée</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Client</label>
        <select id="edit-client"><option value="">-- Sélectionner --</option></select>
      </div>
      <div class="form-group">
        <label>Collaborateurs</label>
        <select id="edit-collabs" multiple title="Maintenez Ctrl/Cmd pour sélectionner plusieurs"></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date de début</label>
        <input type="date" id="edit-debut">
      </div>
      <div class="form-group">
        <label>Date de fin</label>
        <input type="date" id="edit-fin">
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeEditModal()">Annuler</button>
      <button class="btn btn-primary" onclick="saveEditMission()">Enregistrer</button>
    </div>
  </div>
</div>

<script>
// ══════════════════════════════════════════
//  ÉTAT DE L'APPLICATION
// ══════════════════════════════════════════
let DB = {
  collaborateurs: [],
  clients: [],
  missions: []
};
let deleteCallback = null;

// ── Persistance localStorage ──
function save() {
  localStorage.setItem('missions_acobi', JSON.stringify(DB));
}
function load() {
  const raw = localStorage.getItem('missions_acobi');
  if (raw) {
    try { DB = JSON.parse(raw); } catch(e) {}
  }
}

// ── ID unique ──
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// ══════════════════════════════════════════
//  NAVIGATION ONGLETS
// ══════════════════════════════════════════
function showTab(name) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  event.target.classList.add('active');
  if (name === 'visu') renderAll();
}

// ══════════════════════════════════════════
//  TOAST
// ══════════════════════════════════════════
function toast(msg, type = 'ok') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'show' + (type === 'error' ? ' error' : '');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.className = '', 2800);
}

// ══════════════════════════════════════════
//  MODAL CONFIRMATION
// ══════════════════════════════════════════
function openModal(msg, cb) {
  document.getElementById('modal-msg').textContent = msg;
  document.getElementById('modal-confirm').classList.add('open');
  deleteCallback = cb;
}
function closeModal() {
  document.getElementById('modal-confirm').classList.remove('open');
  deleteCallback = null;
}
document.getElementById('modal-confirm-btn').onclick = () => {
  if (deleteCallback) deleteCallback();
  closeModal();
};

// ══════════════════════════════════════════
//  COLLABORATEURS
// ══════════════════════════════════════════
function addCollaborateur() {
  const prenom = document.getElementById('c-prenom').value.trim();
  const nom    = document.getElementById('c-nom').value.trim();
  const role   = document.getElementById('c-role').value.trim();
  if (!prenom || !nom) { toast('Prénom et nom requis', 'error'); return; }
  DB.collaborateurs.push({ id: uid(), prenom, nom, role });
  save();
  document.getElementById('c-prenom').value = '';
  document.getElementById('c-nom').value = '';
  document.getElementById('c-role').value = '';
  renderCollaborateurs();
  refreshSelects();
  toast('Collaborateur ajouté ✓');
}

function deleteCollaborateur(id) {
  const c = DB.collaborateurs.find(x => x.id === id);
  openModal(`Supprimer ${c.prenom} ${c.nom} ?`, () => {
    DB.collaborateurs = DB.collaborateurs.filter(x => x.id !== id);
    save(); renderCollaborateurs(); refreshSelects();
    toast('Collaborateur supprimé');
  });
}

function renderCollaborateurs() {
  const tbody = document.querySelector('#table-collabs tbody');
  if (!DB.collaborateurs.length) {
    tbody.innerHTML = '<tr><td colspan="3" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun collaborateur</td></tr>';
    return;
  }
  tbody.innerHTML = DB.collaborateurs.map(c => `
    <tr>
      <td>${c.prenom} ${c.nom}</td>
      <td style="color:var(--text-muted)">${c.role || '—'}</td>
      <td><button class="btn btn-danger btn-sm" onclick="deleteCollaborateur('${c.id}')">Supprimer</button></td>
    </tr>
  `).join('');
}

// ══════════════════════════════════════════
//  CLIENTS
// ══════════════════════════════════════════
function isUrl(s) { return s.startsWith('http') || s.startsWith('data:'); }

function logoHtml(logo, size = 40, clientNom = '') {
  if (!logo) return initialesHtml(clientNom, size);
  if (isUrl(logo)) return `<img src="${logo}" style="width:${size}px;height:${size}px;object-fit:contain;border-radius:6px" onerror="this.replaceWith(initialesNode('${clientNom.replace(/'/g,"\\'")}', ${size}))">`;
  return `<span style="font-size:${size * 0.5}px">${logo}</span>`;
}

function loadLogoFile(input, logoFieldId, previewId) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const dataUrl = e.target.result;
    // Stocker dans le champ caché
    document.getElementById(logoFieldId).value = dataUrl;
    // Afficher l'aperçu
    document.getElementById(previewId).innerHTML =
      `<img src="${dataUrl}" style="width:100%;height:100%;object-fit:contain;padding:4px">`;
    // Afficher le label "Nouveau logo prêt"
    const labelId = logoFieldId.replace('logo', 'file-label').replace('cl-', 'edit-cl-').replace('edit-edit-','edit-');
    const lbl = document.getElementById(labelId) || document.getElementById('edit-cl-file-label') || document.getElementById('cl-file-label');
    if (lbl) lbl.style.display = 'inline';
  };
  reader.readAsDataURL(file);
}

function initialesHtml(nom, size = 40) {
  const initiales = (nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  return `<span style="display:inline-flex;align-items:center;justify-content:center;width:${size}px;height:${size}px;border-radius:8px;background:var(--card-alt);color:var(--accent);font-weight:700;font-size:${size * 0.35}px">${initiales}</span>`;
}

function initialesNode(nom, size = 40) {
  const el = document.createElement('span');
  const initiales = (nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  el.style.cssText = `display:inline-flex;align-items:center;justify-content:center;width:${size}px;height:${size}px;border-radius:8px;background:var(--card-alt);color:var(--accent);font-weight:700;font-size:${size * 0.35}px`;
  el.textContent = initiales;
  return el;
}

function addClient() {
  const nom  = document.getElementById('cl-nom').value.trim();
  const logo = document.getElementById('cl-logo').value.trim();
  if (!nom) { toast('Nom du client requis', 'error'); return; }
  DB.clients.push({ id: uid(), nom, logo });
  save();
  document.getElementById('cl-nom').value = '';
  document.getElementById('cl-logo').value = '';
  renderClients();
  refreshSelects();
  toast('Client ajouté ✓');
}

function openEditClient(id) {
  const c = DB.clients.find(x => x.id === id);
  if (!c) return;
  document.getElementById('edit-cl-id').value   = c.id;
  document.getElementById('edit-cl-nom').value  = c.nom;
  document.getElementById('edit-cl-logo').value = c.logo || '';
  document.getElementById('edit-cl-file-label').style.display = 'none';
  document.getElementById('edit-cl-logo-file').value = '';
  // Aperçu logo existant
  const prev = document.getElementById('edit-cl-logo-preview');
  if (c.logo) {
    prev.innerHTML = `<img src="${c.logo}" style="width:100%;height:100%;object-fit:contain;padding:4px" onerror="this.parentElement.innerHTML=initialesHtml('${c.nom.replace(/'/g,"\\'")}',56)">`;
  } else {
    prev.innerHTML = initialesHtml(c.nom, 56);
  }
  document.getElementById('modal-edit-client').classList.add('open');
}

function closeEditClientModal() {
  document.getElementById('modal-edit-client').classList.remove('open');
}

function saveEditClient() {
  const id   = document.getElementById('edit-cl-id').value;
  const nom  = document.getElementById('edit-cl-nom').value.trim();
  const logo = document.getElementById('edit-cl-logo').value.trim();
  if (!nom) { toast('Nom du client requis', 'error'); return; }
  const idx = DB.clients.findIndex(x => x.id === id);
  if (idx === -1) return;
  DB.clients[idx] = { id, nom, logo };
  save();
  closeEditClientModal();
  renderClients();
  refreshSelects();
  renderCards();
  toast('Client modifié ✓');
}

function deleteClient(id) {
  const c = DB.clients.find(x => x.id === id);
  openModal(`Supprimer le client "${c.nom}" ?`, () => {
    DB.clients = DB.clients.filter(x => x.id !== id);
    save(); renderClients(); refreshSelects();
    toast('Client supprimé');
  });
}

function renderClients() {
  const tbody = document.querySelector('#table-clients tbody');
  if (!DB.clients.length) {
    tbody.innerHTML = '<tr><td colspan="3" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun client</td></tr>';
    return;
  }
  tbody.innerHTML = DB.clients.map(c => `
    <tr onclick="openEditClient('${c.id}')" style="cursor:pointer" title="Cliquer pour modifier">
      <td>${c.nom}</td>
      <td><span class="logo-mini">${logoHtml(c.logo, 28, c.nom)}</span></td>
      <td><button class="btn btn-danger btn-sm" onclick="event.stopPropagation();deleteClient('${c.id}')">Supprimer</button></td>
    </tr>
  `).join('');
}

// ══════════════════════════════════════════
//  SELECTS DYNAMIQUES
// ══════════════════════════════════════════
function refreshSelects() {
  // Select client (formulaire mission)
  const mClient = document.getElementById('m-client');
  const prev = mClient.value;
  mClient.innerHTML = '<option value="">-- Sélectionner --</option>' +
    DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
  mClient.value = prev;

  // Select collaborateurs (formulaire mission)
  const mCollabs = document.getElementById('m-collabs');
  const prevSel = Array.from(mCollabs.selectedOptions).map(o => o.value);
  mCollabs.innerHTML = DB.collaborateurs.map(c =>
    `<option value="${c.id}" ${prevSel.includes(c.id) ? 'selected' : ''}>${c.prenom} ${c.nom}</option>`
  ).join('');

  // Filtres visualisation
  const fClient = document.getElementById('f-client');
  const fClient2 = fClient.value;
  fClient.innerHTML = '<option value="">Tous</option>' +
    DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
  fClient.value = fClient2;

  const fCollab = document.getElementById('f-collab');
  const fCollab2 = fCollab.value;
  fCollab.innerHTML = '<option value="">Tous</option>' +
    DB.collaborateurs.map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`).join('');
  fCollab.value = fCollab2;
}

// ══════════════════════════════════════════
//  MISSIONS
// ══════════════════════════════════════════
function addMission(data = null) {
  let titre, clientId, collabIds, debut, fin, statut;
  if (data) {
    ({ titre, clientId, collabIds, debut, fin, statut } = data);
  } else {
    titre    = document.getElementById('m-titre').value.trim();
    clientId = document.getElementById('m-client').value;
    collabIds = Array.from(document.getElementById('m-collabs').selectedOptions).map(o => o.value);
    debut    = document.getElementById('m-debut').value;
    fin      = document.getElementById('m-fin').value;
    statut   = document.getElementById('m-statut').value;
  }
  if (!titre) { toast('Titre de mission requis', 'error'); return; }

  DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin, statut: statut || 'en_cours' });
  save();

  if (!data) {
    document.getElementById('m-titre').value = '';
    document.getElementById('m-client').value = '';
    document.getElementById('m-debut').value = '';
    document.getElementById('m-fin').value = '';
    document.getElementById('m-statut').value = 'en_cours';
    Array.from(document.getElementById('m-collabs').options).forEach(o => o.selected = false);
  }
  renderMissions();
}

function openEditMission(id) {
  const m = DB.missions.find(x => x.id === id);
  if (!m) return;

  document.getElementById('edit-id').value     = m.id;
  document.getElementById('edit-titre').value  = m.titre;
  document.getElementById('edit-statut').value = m.statut;
  document.getElementById('edit-debut').value  = m.debut || '';
  document.getElementById('edit-fin').value    = m.fin   || '';

  // Remplir le select client
  const editClient = document.getElementById('edit-client');
  editClient.innerHTML = '<option value="">-- Sélectionner --</option>' +
    DB.clients.map(c => `<option value="${c.id}" ${c.id === m.clientId ? 'selected' : ''}>${c.nom}</option>`).join('');

  // Remplir le select collaborateurs
  const editCollabs = document.getElementById('edit-collabs');
  editCollabs.innerHTML = DB.collaborateurs.map(c =>
    `<option value="${c.id}" ${(m.collabIds || []).includes(c.id) ? 'selected' : ''}>${c.prenom} ${c.nom}</option>`
  ).join('');

  document.getElementById('modal-edit').classList.add('open');
}

function closeEditModal() {
  document.getElementById('modal-edit').classList.remove('open');
}

function saveEditMission() {
  const id       = document.getElementById('edit-id').value;
  const titre    = document.getElementById('edit-titre').value.trim();
  const statut   = document.getElementById('edit-statut').value;
  const clientId = document.getElementById('edit-client').value;
  const collabIds = Array.from(document.getElementById('edit-collabs').selectedOptions).map(o => o.value);
  const debut    = document.getElementById('edit-debut').value;
  const fin      = document.getElementById('edit-fin').value;

  if (!titre) { toast('Titre de mission requis', 'error'); return; }

  const idx = DB.missions.findIndex(x => x.id === id);
  if (idx === -1) return;
  DB.missions[idx] = { id, titre, statut, clientId, collabIds, debut, fin };
  save();
  closeEditModal();
  renderMissions();
  renderCards();
  renderStats();
  toast('Mission modifiée ✓');
}

function deleteMission(id) {
  const m = DB.missions.find(x => x.id === id);
  openModal(`Supprimer la mission "${m.titre}" ?`, () => {
    DB.missions = DB.missions.filter(x => x.id !== id);
    save(); renderMissions(); renderCards();
    toast('Mission supprimée');
  });
}

function renderMissions() {
  const tbody = document.querySelector('#table-missions tbody');
  if (!DB.missions.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucune mission</td></tr>';
    return;
  }
  tbody.innerHTML = DB.missions.map(m => {
    const client = DB.clients.find(c => c.id === m.clientId);
    const collabs = (m.collabIds || []).map(id => {
      const c = DB.collaborateurs.find(x => x.id === id);
      return c ? `${c.prenom} ${c.nom}` : '';
    }).filter(Boolean).join(', ');
    const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ') || '—';
    const badgeCls = m.statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = m.statut === 'en_cours' ? 'En cours' : 'Terminée';
    return `
      <tr onclick="openEditMission('${m.id}')" style="cursor:pointer" title="Cliquer pour modifier">
        <td>${m.titre}</td>
        <td>${client ? client.nom : '—'}</td>
        <td style="color:var(--text-muted)">${collabs || '—'}</td>
        <td style="color:var(--text-muted);font-size:0.82rem">${periode}</td>
        <td><span class="badge ${badgeCls}">${badgeLbl}</span></td>
        <td><button class="btn btn-danger btn-sm" onclick="event.stopPropagation();deleteMission('${m.id}')">Supprimer</button></td>
      </tr>`;
  }).join('');
}

// ══════════════════════════════════════════
//  VISUALISATION — CARTES
// ══════════════════════════════════════════
function renderCards() {
  const fStatut = document.getElementById('f-statut').value;
  const fClient = document.getElementById('f-client').value;
  const fCollab = document.getElementById('f-collab').value;
  const fSearch = document.getElementById('f-search').value.toLowerCase();

  let missions = DB.missions.filter(m => {
    if (fStatut && m.statut !== fStatut) return false;
    if (fClient && m.clientId !== fClient) return false;
    if (fCollab && !(m.collabIds || []).includes(fCollab)) return false;
    if (fSearch && !m.titre.toLowerCase().includes(fSearch)) return false;
    return true;
  });

  const grid = document.getElementById('missions-grid');

  if (!missions.length) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
      <div class="icon">🔍</div>
      <p>Aucune mission trouvée.<br>Ajoutez des missions dans l'onglet Paramétrage.</p>
    </div>`;
    return;
  }

  grid.innerHTML = missions.map(m => {
    const client = DB.clients.find(c => c.id === m.clientId);
    const collabs = (m.collabIds || []).map(id => {
      const c = DB.collaborateurs.find(x => x.id === id);
      return c ? `${c.prenom} ${c.nom}` : '';
    }).filter(Boolean).join(', ');
    const badgeCls = m.statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = m.statut === 'en_cours' ? '🔄 En cours' : '✅ Terminée';
    const periode  = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ');

    return `
    <div class="card">
      <div class="card-logo">${client ? logoHtml(client.logo, 40, client.nom) : initialesHtml('', 40)}</div>
      <span class="badge ${badgeCls}">${badgeLbl}</span>
      <div class="card-title">${m.titre}</div>
      ${client ? `<div class="card-client">${client.nom}</div>` : ''}
      ${collabs ? `<div class="card-collab">👤 ${collabs}</div>` : ''}
      ${periode ? `<div class="card-dates">📅 ${periode}</div>` : ''}
      <div class="card-arrow">→</div>
    </div>`;
  }).join('');
}

function renderStats() {
  const total    = DB.missions.length;
  const encours  = DB.missions.filter(m => m.statut === 'en_cours').length;
  const termines = DB.missions.filter(m => m.statut === 'terminee').length;
  document.getElementById('stats-row').innerHTML = `
    <div class="stat-chip"><span class="val">${total}</span><span class="lbl">Missions totales</span></div>
    <div class="stat-chip"><span class="val" style="color:var(--accent2)">${encours}</span><span class="lbl">En cours</span></div>
    <div class="stat-chip"><span class="val">${termines}</span><span class="lbl">Terminées</span></div>
    <div class="stat-chip"><span class="val" style="color:var(--warning)">${DB.clients.length}</span><span class="lbl">Clients</span></div>
  `;
}

function renderAll() {
  renderCollaborateurs();
  renderClients();
  renderMissions();
  refreshSelects();
  renderStats();
  renderCards();
}

// ══════════════════════════════════════════
//  IMPORT EXCEL
// ══════════════════════════════════════════
function handleDrop(event) {
  event.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag-over');
  const file = event.dataTransfer.files[0];
  if (file) handleExcelFile(file);
}

function handleExcelFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      let added = 0;
      rows.forEach(row => {
        const titre  = (row['Titre'] || row['titre'] || '').toString().trim();
        if (!titre) return;

        const clientNom = (row['Client'] || row['client'] || '').toString().trim();
        const statut    = (row['Statut'] || row['statut'] || 'en_cours').toString().trim();
        const debut     = excelDateToString(row['Date début'] || row['Date debut'] || row['date_debut'] || '');
        const fin       = excelDateToString(row['Date fin']   || row['Date fin']   || row['date_fin']   || '');
        const collabStr = (row['Collaborateurs'] || row['collaborateurs'] || '').toString().trim();

        // Créer le client s'il n'existe pas
        let clientId = '';
        if (clientNom) {
          let c = DB.clients.find(x => x.nom.toLowerCase() === clientNom.toLowerCase());
          if (!c) { c = { id: uid(), nom: clientNom, logo: '' }; DB.clients.push(c); }
          clientId = c.id;
        }

        // Créer les collaborateurs s'ils n'existent pas
        const collabIds = [];
        if (collabStr) {
          collabStr.split(';').map(s => s.trim()).filter(Boolean).forEach(name => {
            const parts  = name.split(' ');
            const prenom = parts[0] || '';
            const nom    = parts.slice(1).join(' ') || '';
            let c = DB.collaborateurs.find(x =>
              (x.prenom + ' ' + x.nom).toLowerCase() === name.toLowerCase()
            );
            if (!c) { c = { id: uid(), prenom, nom, role: '' }; DB.collaborateurs.push(c); }
            collabIds.push(c.id);
          });
        }

        DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin, statut });
        added++;
      });

      save();
      renderAll();
      toast(`${added} mission(s) importée(s) ✓`);
    } catch(err) {
      toast('Erreur lors de la lecture du fichier', 'error');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function excelDateToString(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    // numéro de série Excel → Date
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().split('T')[0];
  }
  const s = val.toString().trim();
  // Essai formats FR (dd/mm/yyyy) ou ISO
  const fr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) return `${fr[3]}-${fr[2].padStart(2,'0')}-${fr[1].padStart(2,'0')}`;
  return s;
}

function formatDate(d) {
  if (!d) return '';
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y}`;
}

// ══════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════
load();
renderAll();
</script>
</body>
</html>
