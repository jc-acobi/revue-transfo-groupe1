<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Missions ACOBI</title>
  <style>
    @font-face { font-family: 'Nunito'; font-style: normal; font-weight: 300 700; font-display: swap;
      src: url('fonts/nunito-latin-ext.woff2') format('woff2');
      unicode-range: U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF; }
    @font-face { font-family: 'Nunito'; font-style: normal; font-weight: 300 700; font-display: swap;
      src: url('fonts/nunito-latin.woff2') format('woff2');
      unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }
  </style>
  <!-- SheetJS pour lire les fichiers Excel -->
  <script src="xlsx.full.min.js"></script>
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

    /* ── LISTE COLLABORATEUR AVEC FILTRE ── */
    .collab-dropdown {
      display: none;
      position: absolute;
      top: 100%;
      left: 0;
      right: 0;
      z-index: 200;
      background: var(--card-bg);
      border: 1px solid var(--accent);
      border-top: none;
      border-radius: 0 0 8px 8px;
      max-height: 200px;
      overflow-y: auto;
    }
    .collab-dd-item {
      padding: 0.5rem 0.9rem;
      cursor: pointer;
      font-size: 0.9rem;
    }
    .collab-dd-item:hover { background: var(--card-alt); }

    /* ── SOUS-ONGLETS PARAMÉTRAGE ── */
    .subtabs {
      display: flex;
      gap: 0.4rem;
      margin-bottom: 2rem;
      border-bottom: 1px solid var(--border);
      padding-bottom: 0;
    }
    .subtab-btn {
      padding: 0.55rem 1.3rem;
      border-radius: 8px 8px 0 0;
      border: none;
      background: transparent;
      color: var(--text-muted);
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }
    .subtab-btn:hover { color: var(--text); }
    .subtab-btn.active {
      background: var(--card-alt);
      color: var(--accent);
      border: 1px solid var(--border);
      border-bottom: 1px solid var(--card-alt);
    }
    .subtab-content { display: none; }
    .subtab-content.active { display: block; }

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
    .data-table th.sortable { cursor: pointer; user-select: none; }
    .data-table th.sortable:hover { color: var(--accent); }
    .data-table th.sortable::after { content: ' ⇅'; opacity: 0.4; font-size: 0.7rem; }
    .data-table th.sort-asc::after  { content: ' ▲'; opacity: 1; color: var(--accent); }
    .data-table th.sort-desc::after { content: ' ▼'; opacity: 1; color: var(--accent); }
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

  <!-- SOUS-ONGLETS -->
  <div class="subtabs">
    <button class="subtab-btn active" onclick="showSubTab('collab')">👤 Collaborateurs</button>
    <button class="subtab-btn" onclick="showSubTab('clients')">🏢 Clients</button>
    <button class="subtab-btn" onclick="showSubTab('missions')">📋 Missions</button>
  </div>

  <!-- SOUS-ONGLET COLLABORATEURS -->
  <div id="subtab-collab" class="subtab-content active">
  <div class="section-title">👤 Collaborateurs</div>
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3 id="c-form-title">Ajouter un collaborateur</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('c-import-input').click()">📥 Import en masse</button>
      <input type="file" id="c-import-input" accept=".xlsx,.xls" style="display:none" onchange="importCollaborateurs(this)">
    </div>
    <input type="hidden" id="c-edit-id">
    <div class="form-row">
      <div class="form-group">
        <label>Nom</label>
        <input type="text" id="c-nom" placeholder="ex. Dupont">
      </div>
      <div class="form-group">
        <label>Prénom</label>
        <input type="text" id="c-prenom" placeholder="ex. Marie">
      </div>
      <div class="form-group">
        <label>Sexe</label>
        <select id="c-sexe">
          <option value="">-- Sélectionner --</option>
          <option value="Masculin">Masculin</option>
          <option value="Féminin">Féminin</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date d'entrée</label>
        <input type="date" id="c-date-entree">
      </div>
      <div class="form-group">
        <label>Date de sortie</label>
        <input type="date" id="c-date-sortie">
      </div>
      <div class="form-group">
        <label>Co-pilote</label>
        <select id="c-copilote">
          <option value="">-- Aucun --</option>
        </select>
      </div>
    </div>
    <div style="display:flex;gap:0.75rem">
      <button class="btn btn-primary" id="c-submit-btn" onclick="saveCollaborateur()">Ajouter</button>
      <button class="btn" id="c-cancel-btn" style="display:none;background:var(--card-alt);color:var(--text)" onclick="cancelEditCollaborateur()">Annuler</button>
    </div>
  </div>

  <table class="data-table" id="table-collabs" style="width:100%">
    <thead><tr>
      <th class="sortable" onclick="sortTable('collabs','prenom')">Prénom Nom</th>
      <th class="sortable" onclick="sortTable('collabs','sexe')">Sexe</th>
      <th class="sortable" onclick="sortTable('collabs','dateEntree')">Date entrée</th>
      <th class="sortable" onclick="sortTable('collabs','dateSortie')">Date sortie</th>
      <th class="sortable" onclick="sortTable('collabs','copilote')">Co-pilote</th>
      <th>Actions</th>
    </tr></thead>
  </table>
  <div style="max-height:296px;overflow-y:auto;border:1px solid var(--border);border-top:none;border-radius:0 0 8px 8px">
    <table class="data-table" style="width:100%"><tbody id="tbody-collabs"></tbody></table>
  </div>

  </div><!-- /subtab-collab -->

  <!-- SOUS-ONGLET CLIENTS -->
  <div id="subtab-clients" class="subtab-content">
  <div class="section-title">🏢 Clients</div>
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3>Ajouter un client</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('cl-import-input').click()">📥 Import en masse</button>
      <input type="file" id="cl-import-input" accept=".xlsx,.xls" style="display:none" onchange="importClients(this)">
    </div>
    <input type="hidden" id="cl-logo">
    <div class="form-row" style="align-items:center;gap:1.5rem">
      <div class="form-group" style="flex:1">
        <label>Nom du client</label>
        <input type="text" id="cl-nom" placeholder="ex. EDF">
      </div>
      <div class="form-group" style="flex:0;min-width:auto">
        <label>Logo</label>
        <div id="cl-logo-preview" style="width:56px;height:56px;display:flex;align-items:center;justify-content:center;background:var(--card-alt);border-radius:10px;overflow:hidden"></div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:1rem;margin-top:0.5rem;margin-bottom:1rem">
      <button class="btn btn-secondary" type="button" onclick="document.getElementById('cl-logo-file').click()">📁 Charger un logo</button>
      <input type="file" id="cl-logo-file" accept="image/*" style="display:none" onchange="loadLogoFile(this, 'cl-logo', 'cl-logo-preview')">
      <span id="cl-file-label" style="display:none;color:var(--accent);font-size:0.85rem;font-weight:600">✓ Logo prêt</span>
    </div>
    <button class="btn btn-primary" onclick="addClient()">Ajouter</button>
  </div>

  <table class="data-table" id="table-clients">
    <thead><tr>
      <th class="sortable" onclick="sortTable('clients','nom')">Client</th>
      <th>Logo</th>
      <th>Actions</th>
    </tr></thead>
    <tbody></tbody>
  </table>

  </div><!-- /subtab-clients -->

  <!-- SOUS-ONGLET MISSIONS -->
  <div id="subtab-missions" class="subtab-content">
  <div class="section-title">📋 Missions</div>

  <!-- Formulaire mission -->
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3>Ajouter une mission</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('m-import-input').click()">📥 Import en masse</button>
      <input type="file" id="m-import-input" accept=".xlsx,.xls" style="display:none" onchange="importMissions(this)">
    </div>
    <div id="import-missions-summary" style="display:none;margin-bottom:1rem"></div>
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label>Titre de la mission</label>
        <input type="text" id="m-titre" placeholder="ex. Transformation digitale">
      </div>
      <div class="form-group">
        <label>Client *</label>
        <select id="m-client"><option value="">-- Sélectionner --</option></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="position:relative">
        <label>Collaborateur *</label>
        <input type="text" id="m-collab-search" placeholder="Rechercher..." autocomplete="off"
          oninput="filterCollabDropdown('m')" onfocus="showCollabDropdown('m')" onblur="hideCollabDropdown('m')">
        <input type="hidden" id="m-collabs">
        <div id="m-collab-dd" class="collab-dropdown"></div>
      </div>
      <div class="form-group">
        <label>Date de début *</label>
        <input type="date" id="m-debut">
      </div>
      <div class="form-group">
        <label>Date de fin</label>
        <input type="date" id="m-fin">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Détails de la mission</label>
        <textarea id="m-details" rows="3" placeholder="Description, contexte, objectifs…" style="width:100%;background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.6rem 0.9rem;font-family:inherit;font-size:0.95rem;resize:vertical"></textarea>
      </div>
    </div>
    <button class="btn btn-primary" onclick="addMission()">Ajouter la mission</button>
  </div>

  <div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:0.8rem;align-items:flex-end">
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Filtrer par client</label>
      <select id="param-f-client" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:180px">
        <option value="">Tous</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Filtrer par collaborateur</label>
      <select id="param-f-collab" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:200px">
        <option value="">Tous</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Statut</label>
      <select id="param-f-statut" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:150px">
        <option value="">Tous</option>
        <option value="en_cours">En cours</option>
        <option value="terminee">Terminée</option>
      </select>
    </div>
    <button class="btn" style="background:transparent;border:1px solid var(--border);color:var(--text-muted);padding:0.5rem 1rem;font-size:0.85rem" onclick="document.getElementById('param-f-client').value='';document.getElementById('param-f-collab').value='';document.getElementById('param-f-statut').value='';renderMissions()">✕ Effacer</button>
  </div>

  <table class="data-table" id="table-missions">
    <thead><tr>
      <th class="sortable" onclick="sortTable('missions','titre')">Mission</th>
      <th class="sortable" onclick="sortTable('missions','client')">Client</th>
      <th class="sortable" onclick="sortTable('missions','collabs')">Collaborateurs</th>
      <th class="sortable" onclick="sortTable('missions','debut')">Période</th>
      <th class="sortable" onclick="sortTable('missions','statut')">Statut</th>
      <th style="min-width:140px">Actions</th>
    </tr></thead>
    <tbody></tbody>
  </table>

  </div><!-- /subtab-missions -->

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
        <label>Client</label>
        <select id="edit-client"><option value="">-- Sélectionner --</option></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="position:relative">
        <label>Collaborateur</label>
        <input type="text" id="edit-collab-search" placeholder="Rechercher..." autocomplete="off"
          oninput="filterCollabDropdown('edit')" onfocus="showCollabDropdown('edit')" onblur="hideCollabDropdown('edit')">
        <input type="hidden" id="edit-collabs">
        <div id="edit-collab-dd" class="collab-dropdown"></div>
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
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Description</label>
        <textarea id="edit-details" rows="3" placeholder="Description, contexte, objectifs…" style="width:100%;background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.6rem 0.9rem;font-family:inherit;font-size:0.9rem;resize:vertical"></textarea>
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

// ── Persistance base de données ──
function save() {
  fetch('api.php', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(DB)
  }).catch(() => toast('Erreur de sauvegarde', 'error'));
}

async function load() {
  try {
    const res = await fetch('api.php');
    const data = await res.json();
    DB = { collaborateurs: data.collaborateurs || [], missions: data.missions || [], clients: data.clients || [] };
  } catch(e) {
    DB = { collaborateurs: [], missions: [], clients: [] };
  }
}

// ── Tri des tableaux ──
const sortState = {
  collabs:  { col: null, dir: 1 },
  clients:  { col: null, dir: 1 },
  missions: { col: null, dir: 1 }
};
function sortTable(table, col) {
  if (sortState[table].col === col) sortState[table].dir *= -1;
  else { sortState[table].col = col; sortState[table].dir = 1; }
  if (table === 'collabs')  renderCollaborateurs();
  if (table === 'clients')  renderClients();
  if (table === 'missions') renderMissions();
}
function updateSortHeaders(tableId, table) {
  document.querySelectorAll('#' + tableId + ' th.sortable').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.getAttribute('onclick').includes("'" + sortState[table].col + "'")) {
      th.classList.add(sortState[table].dir === 1 ? 'sort-asc' : 'sort-desc');
    }
  });
}
function applySortCollabs(list) {
  const { col, dir } = sortState.collabs;
  if (!col) return list;
  return [...list].sort((a, b) => {
    let va = col === 'prenom' ? (a.prenom + ' ' + a.nom) : (a[col] || '');
    let vb = col === 'prenom' ? (b.prenom + ' ' + b.nom) : (b[col] || '');
    if (col === 'copilote') {
      const ca = DB.collaborateurs.find(x => x.id === a.copilote);
      const cb = DB.collaborateurs.find(x => x.id === b.copilote);
      va = ca ? ca.prenom + ' ' + ca.nom : '';
      vb = cb ? cb.prenom + ' ' + cb.nom : '';
    }
    return va.localeCompare(vb, 'fr') * dir;
  });
}
function applySortClients(list) {
  const { col, dir } = sortState.clients;
  if (!col) return list;
  return [...list].sort((a, b) => (a[col] || '').localeCompare(b[col] || '', 'fr') * dir);
}
function applySortMissions(list) {
  const { col, dir } = sortState.missions;
  if (!col) return list;
  return [...list].sort((a, b) => {
    let va = '', vb = '';
    if (col === 'titre')  { va = a.titre || ''; vb = b.titre || ''; }
    if (col === 'client') {
      const ca = DB.clients.find(x => x.id === a.clientId);
      const cb = DB.clients.find(x => x.id === b.clientId);
      va = ca ? ca.nom : ''; vb = cb ? cb.nom : '';
    }
    if (col === 'collabs') {
      va = (a.collabIds || []).map(id => { const c = DB.collaborateurs.find(x => x.id === id); return c ? c.prenom + ' ' + c.nom : ''; }).join(', ');
      vb = (b.collabIds || []).map(id => { const c = DB.collaborateurs.find(x => x.id === id); return c ? c.prenom + ' ' + c.nom : ''; }).join(', ');
    }
    if (col === 'debut')  { va = a.debut || ''; vb = b.debut || ''; }
    if (col === 'statut') { va = getStatut(a); vb = getStatut(b); }
    return va.localeCompare(vb, 'fr') * dir;
  });
}

// ── ID unique ──
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// ══════════════════════════════════════════
//  NAVIGATION ONGLETS
// ══════════════════════════════════════════
function showSubTab(name) {
  document.querySelectorAll('.subtab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.subtab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('subtab-' + name).classList.add('active');
  document.querySelectorAll('.subtab-btn').forEach(btn => {
    if (btn.getAttribute('onclick').includes("'" + name + "'")) btn.classList.add('active');
  });
}

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
function saveCollaborateur() {
  const editId = document.getElementById('c-edit-id').value;
  const nom         = document.getElementById('c-nom').value.trim();
  const prenom      = document.getElementById('c-prenom').value.trim();
  const sexe        = document.getElementById('c-sexe').value;
  const dateEntree  = document.getElementById('c-date-entree').value;
  const dateSortie  = document.getElementById('c-date-sortie').value;
  const copilote    = document.getElementById('c-copilote').value;
  if (!nom || !prenom) { toast('Nom et prénom requis', 'error'); return; }
  if (editId) {
    const c = DB.collaborateurs.find(x => x.id === editId);
    Object.assign(c, { nom, prenom, sexe, dateEntree, dateSortie, copilote });
    toast('Collaborateur modifié ✓');
  } else {
    DB.collaborateurs.push({ id: uid(), nom, prenom, sexe, dateEntree, dateSortie, copilote });
    toast('Collaborateur ajouté ✓');
  }
  save();
  resetCollaborateurForm();
  renderCollaborateurs();
  refreshSelects();
}

function resetCollaborateurForm() {
  document.getElementById('c-edit-id').value = '';
  document.getElementById('c-nom').value = '';
  document.getElementById('c-prenom').value = '';
  document.getElementById('c-sexe').value = '';
  document.getElementById('c-date-entree').value = '';
  document.getElementById('c-date-sortie').value = '';
  document.getElementById('c-copilote').value = '';
  document.getElementById('c-form-title').textContent = 'Ajouter un collaborateur';
  document.getElementById('c-submit-btn').textContent = 'Ajouter';
  document.getElementById('c-cancel-btn').style.display = 'none';
}

function editCollaborateur(id) {
  const c = DB.collaborateurs.find(x => x.id === id);
  refreshCopiloteSelect(id);
  document.getElementById('c-edit-id').value = id;
  document.getElementById('c-nom').value = c.nom;
  document.getElementById('c-prenom').value = c.prenom;
  document.getElementById('c-sexe').value = c.sexe || '';
  document.getElementById('c-date-entree').value = c.dateEntree || '';
  document.getElementById('c-date-sortie').value = c.dateSortie || '';
  document.getElementById('c-copilote').value = c.copilote || '';
  document.getElementById('c-form-title').textContent = `Modifier ${c.prenom} ${c.nom}`;
  document.getElementById('c-submit-btn').textContent = 'Enregistrer';
  document.getElementById('c-cancel-btn').style.display = 'inline-block';
  document.getElementById('c-form-title').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function cancelEditCollaborateur() {
  resetCollaborateurForm();
  refreshCopiloteSelect(null);
}

function refreshCopiloteSelect(excludeId) {
  const sel = document.getElementById('c-copilote');
  const prev = sel.value;
  sel.innerHTML = '<option value="">-- Aucun --</option>' +
    DB.collaborateurs
      .filter(c => c.id !== excludeId)
      .filter(c => !c.dateSortie || c.dateSortie >= new Date().toISOString().split('T')[0])
      .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'))
      .map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`)
      .join('');
  sel.value = prev;
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
  const tbody = document.getElementById('tbody-collabs');
  updateSortHeaders('table-collabs', 'collabs');
  const list = applySortCollabs(DB.collaborateurs);
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun collaborateur</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(c => {
    const copilote = c.copilote ? DB.collaborateurs.find(x => x.id === c.copilote) : null;
    const copiloteLabel = copilote ? `${copilote.prenom} ${copilote.nom}` : '—';
    const fmt = d => d ? d.split('-').reverse().join('/') : '—';
    return `
    <tr>
      <td>${c.prenom} ${c.nom}</td>
      <td style="color:var(--text-muted)">${c.sexe || '—'}</td>
      <td style="color:var(--text-muted)">${fmt(c.dateEntree)}</td>
      <td style="color:var(--text-muted)">${fmt(c.dateSortie)}</td>
      <td style="color:var(--text-muted)">${copiloteLabel}</td>
      <td style="display:flex;gap:0.5rem">
        <button class="btn btn-primary btn-sm" onclick="editCollaborateur('${c.id}')">Modifier</button>
        <button class="btn btn-danger btn-sm" onclick="deleteCollaborateur('${c.id}')">Supprimer</button>
      </td>
    </tr>`;
  }).join('');
}

function importCollaborateurs(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable dans le fichier', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide ou non reconnu', 'error'); return; }

      // Première passe : créer tous les collaborateurs sans co-pilote
      const newIds = {};
      let added = 0;
      rows.forEach(row => {
        const nom    = (row['Nom']    || '').toString().trim();
        const prenom = (row['Prénom'] || row['Prenom'] || '').toString().trim();
        if (!nom || !prenom) return;
        const sexe       = (row['Sexe'] || '').toString().trim();
        const dateEntree = excelDateToISO(row['Entrée'] || row['Entree'] || row['Date entrée'] || '');
        const dateSortie = excelDateToISO(row['Sortie'] || row['Date sortie'] || '');
        const key = (prenom + ' ' + nom).toLowerCase();
        const existing = DB.collaborateurs.find(c => (c.prenom + ' ' + c.nom).toLowerCase() === key);
        if (existing) {
          Object.assign(existing, { sexe, dateEntree, dateSortie });
          newIds[key] = existing.id;
        } else {
          const id = uid();
          DB.collaborateurs.push({ id, nom, prenom, sexe, dateEntree, dateSortie, copilote: '' });
          newIds[key] = id;
          added++;
        }
      });

      // Deuxième passe : associer les co-pilotes
      rows.forEach(row => {
        const nom    = (row['Nom']    || '').toString().trim();
        const prenom = (row['Prénom'] || row['Prenom'] || '').toString().trim();
        if (!nom || !prenom) return;
        const copiloteStr = (row['Co-pilote'] || row['Copilote'] || '').toString().trim();
        if (!copiloteStr || copiloteStr.toLowerCase() === 'n/a') return;
        const key = (prenom + ' ' + nom).toLowerCase();
        const collab = DB.collaborateurs.find(c => c.id === newIds[key]);
        if (!collab) return;
        const coKey = copiloteStr.toLowerCase();
        const co = DB.collaborateurs.find(c => (c.prenom + ' ' + c.nom).toLowerCase() === coKey);
        if (co) collab.copilote = co.id;
      });

      save();
      renderCollaborateurs();
      refreshSelects();
      toast(`${added} collaborateur(s) importé(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function excelDateToISO(val) {
  if (!val && val !== 0) return '';
  const s = val.toString().trim();
  if (!s) return '';
  // Format DD/MM/YYYY ou DD/MM/YY
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  }
  // Format YYYY-MM-DD déjà correct
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return '';
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
    document.getElementById(logoFieldId).value = dataUrl;
    document.getElementById(previewId).innerHTML =
      `<img src="${dataUrl}" style="width:100%;height:100%;object-fit:contain;padding:4px">`;
    // Afficher le label de confirmation (cl- ou edit-cl-)
    const prefix = logoFieldId.startsWith('edit') ? 'edit-cl' : 'cl';
    const lbl = document.getElementById(prefix + '-file-label');
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
  document.getElementById('cl-logo-file').value = '';
  document.getElementById('cl-file-label').style.display = 'none';
  document.getElementById('cl-logo-preview').innerHTML = '';
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
  updateSortHeaders('table-clients', 'clients');
  const list = applySortClients(DB.clients);
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="3" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun client</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(c => `
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

  // Select co-pilote (formulaire collaborateur)
  const editId = document.getElementById('c-edit-id').value;
  refreshCopiloteSelect(editId || null);

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

function importClients(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide ou non reconnu', 'error'); return; }
      let added = 0;
      rows.forEach(row => {
        const nom = (row['Client'] || '').toString().trim();
        if (!nom) return;
        const exists = DB.clients.find(c => c.nom.toLowerCase() === nom.toLowerCase());
        if (!exists) {
          DB.clients.push({ id: uid(), nom, logo: '' });
          added++;
        }
      });
      save();
      renderClients();
      refreshSelects();
      toast(`${added} client(s) importé(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════
//  MISSIONS
// ══════════════════════════════════════════
function getStatut(m) {
  const today = new Date().toISOString().split('T')[0];
  return (m.fin && m.fin < today) ? 'terminee' : 'en_cours';
}

function addMission(data = null) {
  let titre, clientId, collabIds, debut, fin, details;
  if (data) {
    ({ titre, clientId, collabIds, debut, fin, details } = data);
  } else {
    titre     = document.getElementById('m-titre').value.trim();
    clientId  = document.getElementById('m-client').value;
    const collabId = document.getElementById('m-collabs').value;
    collabIds = collabId ? [collabId] : [];
    debut     = document.getElementById('m-debut').value;
    fin       = document.getElementById('m-fin').value;
    details   = document.getElementById('m-details').value.trim();
    if (!clientId) { toast('Client requis', 'error'); return; }
    if (!collabId) { toast('Collaborateur requis', 'error'); return; }
    if (!debut)    { toast('Date de début requise', 'error'); return; }
  }

  DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin: fin || '', details: details || '' });
  save();

  if (!data) {
    document.getElementById('m-titre').value = '';
    document.getElementById('m-client').value = '';
    document.getElementById('m-collab-search').value = '';
    document.getElementById('m-collabs').value = '';
    document.getElementById('m-debut').value = '';
    document.getElementById('m-fin').value = '';
    document.getElementById('m-details').value = '';
  }
  renderMissions();
  toast('Mission ajoutée ✓');
}

function openEditMission(id) {
  const m = DB.missions.find(x => x.id === id);
  if (!m) return;

  document.getElementById('edit-id').value      = m.id;
  document.getElementById('edit-titre').value   = m.titre;
  document.getElementById('edit-debut').value   = m.debut   || '';
  document.getElementById('edit-fin').value     = m.fin     || '';
  document.getElementById('edit-details').value = m.details || '';

  // Remplir le select client
  const editClient = document.getElementById('edit-client');
  editClient.innerHTML = '<option value="">-- Sélectionner --</option>' +
    DB.clients.map(c => `<option value="${c.id}" ${c.id === m.clientId ? 'selected' : ''}>${c.nom}</option>`).join('');

  // Pré-remplir le collaborateur
  const firstCollabId = (m.collabIds || [])[0] || '';
  const firstCollab = firstCollabId ? DB.collaborateurs.find(x => x.id === firstCollabId) : null;
  document.getElementById('edit-collabs').value = firstCollabId;
  document.getElementById('edit-collab-search').value = firstCollab ? firstCollab.prenom + ' ' + firstCollab.nom : '';

  document.getElementById('modal-edit').classList.add('open');
}

function closeEditModal() {
  document.getElementById('modal-edit').classList.remove('open');
}

function saveEditMission() {
  const id       = document.getElementById('edit-id').value;
  const titre    = document.getElementById('edit-titre').value.trim();
  const clientId = document.getElementById('edit-client').value;
  const collabId = document.getElementById('edit-collabs').value;
  const collabIds = collabId ? [collabId] : [];
  const debut    = document.getElementById('edit-debut').value;
  const fin      = document.getElementById('edit-fin').value;
  const details  = document.getElementById('edit-details').value.trim();

  const idx = DB.missions.findIndex(x => x.id === id);
  if (idx === -1) return;
  DB.missions[idx] = { id, titre, clientId, collabIds, debut, fin: fin || '', details };
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
  updateSortHeaders('table-missions', 'missions');

  // Mettre à jour les filtres
  const pfClient = document.getElementById('param-f-client');
  const pfCollab = document.getElementById('param-f-collab');
  if (pfClient && pfCollab) {
    const prevClient = pfClient.value, prevCollab = pfCollab.value;
    pfClient.innerHTML = '<option value="">Tous</option>' + DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
    pfCollab.innerHTML = '<option value="">Tous</option>' + DB.collaborateurs.map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`).join('');
    pfClient.value = prevClient; pfCollab.value = prevCollab;
  }

  const pfStatut = document.getElementById('param-f-statut');
  const fClient = pfClient ? pfClient.value : '';
  const fCollab = pfCollab ? pfCollab.value : '';
  const fStatut = pfStatut ? pfStatut.value : '';
  let list = DB.missions.filter(m => {
    if (fClient && m.clientId !== fClient) return false;
    if (fCollab && !(m.collabIds || []).includes(fCollab)) return false;
    if (fStatut && getStatut(m) !== fStatut) return false;
    return true;
  });
  list = applySortMissions(list);

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucune mission</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(m => {
    const client = DB.clients.find(c => c.id === m.clientId);
    const collabs = (m.collabIds || []).map(id => {
      const c = DB.collaborateurs.find(x => x.id === id);
      return c ? `${c.prenom} ${c.nom}` : '';
    }).filter(Boolean).join(', ');
    const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ') || '—';
    const statut = getStatut(m);
    const badgeCls = statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = statut === 'en_cours' ? 'En cours' : 'Terminée';
    return `
      <tr>
        <td>${m.titre || '—'}</td>
        <td>${client ? client.nom : '—'}</td>
        <td style="color:var(--text-muted)">${collabs || '—'}</td>
        <td style="color:var(--text-muted);font-size:0.82rem">${periode}</td>
        <td><span class="badge ${badgeCls}">${badgeLbl}</span></td>
        <td style="display:flex;gap:0.5rem">
          <button class="btn btn-primary btn-sm" onclick="openEditMission('${m.id}')">Modifier</button>
          <button class="btn btn-danger btn-sm" onclick="deleteMission('${m.id}')">Supprimer</button>
        </td>
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
    if (fStatut && getStatut(m) !== fStatut) return false;
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
    const statut2  = getStatut(m);
    const badgeCls = statut2 === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = statut2 === 'en_cours' ? '🔄 En cours' : '✅ Terminée';
    const periode  = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ');

    return `
    <div class="card">
      <div class="card-logo">${client ? logoHtml(client.logo, 40, client.nom) : initialesHtml('', 40)}</div>
      <span class="badge ${badgeCls}">${badgeLbl}</span>
      <div class="card-title">${m.titre}</div>
      ${client ? `<div class="card-client">${client.nom}</div>` : ''}
      ${collabs ? `<div class="card-collab">👤 ${collabs}</div>` : ''}
      ${periode ? `<div class="card-dates">📅 ${periode}</div>` : ''}
    </div>`;
  }).join('');
}

function renderStats() {
  const total    = DB.missions.length;
  const encours  = DB.missions.filter(m => getStatut(m) === 'en_cours').length;
  const termines = DB.missions.filter(m => getStatut(m) === 'terminee').length;
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
// ══════════════════════════════════════════
//  DROPDOWN COLLABORATEUR AVEC FILTRE
// ══════════════════════════════════════════
function filterCollabDropdown(prefix) {
  const search = (document.getElementById(prefix + '-collab-search').value || '').toLowerCase();
  const today  = new Date().toISOString().split('T')[0];
  const dd     = document.getElementById(prefix + '-collab-dd');
  const matches = DB.collaborateurs
    .filter(c => !c.dateSortie || c.dateSortie >= today)
    .filter(c => {
      const label = (c.nom + ' ' + c.prenom).toLowerCase();
      return !search || label.includes(search) || c.prenom.toLowerCase().includes(search) || c.nom.toLowerCase().includes(search);
    })
    .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'));
  dd.innerHTML = matches.length
    ? matches.map(c => `<div class="collab-dd-item" onmousedown="selectCollabDD('${prefix}','${c.id}','${c.prenom} ${c.nom}')">${c.prenom} ${c.nom}</div>`).join('')
    : '<div style="padding:0.5rem 0.9rem;color:var(--text-muted);font-size:0.85rem">Aucun résultat</div>';
  dd.style.display = 'block';
}
function showCollabDropdown(prefix) { filterCollabDropdown(prefix); }
function hideCollabDropdown(prefix) {
  setTimeout(() => { const dd = document.getElementById(prefix + '-collab-dd'); if (dd) dd.style.display = 'none'; }, 200);
}
function selectCollabDD(prefix, id, label) {
  document.getElementById(prefix + '-collabs').value = id;
  document.getElementById(prefix + '-collab-search').value = label;
  document.getElementById(prefix + '-collab-dd').style.display = 'none';
}

// ══════════════════════════════════════════
//  IMPORT EN MASSE — MISSIONS
// ══════════════════════════════════════════
function importMissions(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide', 'error'); return; }
      let created = 0, updated = 0;
      const skipped = [];

      rows.forEach((row, i) => {
        const titre      = (row['Titre de la mission'] || '').toString().trim();
        const clientNom  = (row['Client'] || '').toString().trim();
        const collabStr  = (row['Collaborateurs'] || '').toString().trim();
        const debut      = excelDateToISO(row['Date Début Mission'] || row['Date Début'] || row['Date debut'] || '');
        const fin        = excelDateToISO(row['Date Fin Mission']   || row['Date Fin']   || row['Date fin']  || '');
        const details    = (row['Détails Missions'] || row['Détails de la mission'] || '').toString().trim();
        const lineNum    = i + 2;

        if (!clientNom) { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : client manquant`); return; }
        if (!collabStr) { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : collaborateur manquant`); return; }
        if (!debut)     { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : date de début manquante`); return; }

        // Collaborateur : doit exister
        const collab = DB.collaborateurs.find(c => (c.nom + ' ' + c.prenom).toLowerCase() === collabStr.toLowerCase()
          || (c.prenom + ' ' + c.nom).toLowerCase() === collabStr.toLowerCase());
        if (!collab) { skipped.push(`Ligne ${lineNum} (${titre}) : collaborateur "${collabStr}" introuvable`); return; }

        // Client : créer si inexistant
        let client = DB.clients.find(c => c.nom.toLowerCase() === clientNom.toLowerCase());
        if (!client) { client = { id: uid(), nom: clientNom, logo: '' }; DB.clients.push(client); }

        // Mission existante (même collaborateur + client + dates) → mise à jour
        const existing = DB.missions.find(m =>
          m.clientId === client.id &&
          m.debut === debut &&
          m.fin === (fin || '') &&
          (m.collabIds || []).includes(collab.id)
        );
        if (existing) {
          if (!existing.collabIds.includes(collab.id)) existing.collabIds.push(collab.id);
          if (debut) existing.debut = debut;
          if (fin)   existing.fin   = fin;
          if (details) existing.details = details;
          updated++;
        } else {
          DB.missions.push({ id: uid(), titre, clientId: client.id, collabIds: [collab.id], debut, fin: fin || '', details });
          created++;
        }
      });

      save();
      renderAll();

      // Afficher le résumé
      const summaryDiv = document.getElementById('import-missions-summary');
      let html = `<div style="background:var(--card-bg);border:1px solid var(--border);border-radius:var(--radius);padding:1.2rem">
        <div style="font-weight:700;margin-bottom:0.6rem;color:var(--accent)">📊 Résumé de l'import</div>
        <div>✅ ${created} mission(s) créée(s)</div>
        <div>🔄 ${updated} mission(s) mise(s) à jour</div>`;
      if (skipped.length) {
        html += `<div style="margin-top:0.6rem;color:var(--warning);font-weight:600">⚠️ ${skipped.length} ligne(s) ignorée(s) :</div>
          <ul style="margin:0.3rem 0 0 1.2rem;color:var(--text-muted);font-size:0.85rem">`;
        skipped.forEach(s => { html += `<li>${s}</li>`; });
        html += '</ul>';
      }
      html += '</div>';
      summaryDiv.innerHTML = html;
      summaryDiv.style.display = 'block';
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

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
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });

      // Regrouper les lignes par mission (Titre Mission + Client + Date Début + Date Fin)
      const missionsMap = {};
      rows.forEach(row => {
        const titre = (row['Titre Mission'] || '').toString().trim();
        if (!titre) return;
        const clientNom = (row['Client'] || '').toString().trim();
        const debut     = excelDateToISO(row['Date Début'] || row['Date Debut'] || '');
        const fin       = excelDateToISO(row['Date Fin']   || '');
        const details   = (row['Détails Missions'] || row['Details Missions'] || '').toString().trim();
        const nom       = (row['Nom']    || '').toString().trim();
        const prenom    = (row['Prénom'] || row['Prenom'] || '').toString().trim();

        const key = [titre, clientNom, debut, fin].join('|');
        if (!missionsMap[key]) {
          missionsMap[key] = { titre, clientNom, debut, fin, details, collabIds: [] };
        }

        if (nom || prenom) {
          const fullName = (prenom + ' ' + nom).trim().toLowerCase();
          let c = DB.collaborateurs.find(x => (x.prenom + ' ' + x.nom).toLowerCase() === fullName);
          if (!c) { c = { id: uid(), nom, prenom, sexe: '', dateEntree: '', dateSortie: '', copilote: '' }; DB.collaborateurs.push(c); }
          if (!missionsMap[key].collabIds.includes(c.id)) missionsMap[key].collabIds.push(c.id);
        }
      });

      let added = 0;
      Object.values(missionsMap).forEach(({ titre, clientNom, debut, fin, details, collabIds }) => {
        let clientId = '';
        if (clientNom) {
          let c = DB.clients.find(x => x.nom.toLowerCase() === clientNom.toLowerCase());
          if (!c) { c = { id: uid(), nom: clientNom, logo: '' }; DB.clients.push(c); }
          clientId = c.id;
        }
        const statut = (fin && fin < new Date().toISOString().split('T')[0]) ? 'terminee' : 'en_cours';
        DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin, details, statut });
        added++;
      });

      save();
      renderAll();
      toast(`${added} mission(s) importée(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
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
load().then(() => renderAll());
</script>
</body>
</html>
