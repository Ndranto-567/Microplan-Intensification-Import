// ═══════════════════════════════════════════════════════
//  CONFIG – même DEPLOY_ID que index.html
// ═══════════════════════════════════════════════════════
const DEPLOY_ID = 'AKfycbyIsLefLJ0eQLOnrLNOlhU8PbLMW4WOnCtyDDDu1rH8UA-7g41z7eMLxRCNwI1I2eCevQ';
const GS_URL    = `https://script.google.com/macros/s/${DEPLOY_ID}/exec`;

// ═══════════════════════════════════════════════════════
//  DONNÉES RÉGIONS/DISTRICTS (copie depuis index.html)
// ═══════════════════════════════════════════════════════
const REGIONS_DISTRICTS = {
  "ALAOTRA MANGORO":["AMBATONDRAZAKA","AMPARAFARAVOLA","ANDILAMENA","ANOSIBE AN'ALA","MORAMANGA"],
  "AMORON'I MANIA":["AMBATOFINANDRAHANA","AMBOSITRA","FANDRIANA","MANANDRIANA"],
  "ANALAMANGA":["AMBOHIDRATRIMO","ANKAZOBE","ANDRAMASINA","ANTANANARIVO ATSIMONDRANO","ANJOZOROBE","ANTANANARIVO AVARADRANO","ANTANANARIVO RENIVOHITRA","MANJAKANDRIANA"],
  "ANALANJIROFO":["FENERIVE EST","VAVATENINA","SAINTE MARIE","SOANIERANA IVONGO","MANANARA NORD","MAROANTSETRA"],
  "ANDROY":["AMBOVOMBE","ANTANIMORA","BEKILY","BELOHA","TSIHOMBE"],
  "ANOSY":["AMBOASARY SUD","BETROKA","TOLAGNARO"],
  "ATSIMO ATSINANANA":["BEFOTAKA","MIDONGY DU SUD","FARAFANGANA","VANGAINDRANO","VONDROZO"],
  "ATSIMO-ANDREFANA":["AMPANIHY ANDREFANA","BETIOKY-SUD","BENENITRA","SAKARAHA","ANKAZOABO","BEROROHA","MOROMBE","TOLIARA II","TOLIARA I"],
  "ATSINANANA":["ANTANAMBAO MANAMPOTSY","MAHANORO","MAROLAMBO","TOAMASINA I","TOAMASINA II","BRICKAVILLE","VATOMANDRY"],
  "BETSIBOKA":["KANDREHO","MAEVATANANA","TSARATANANA NORD","TSARATANANA SUD"],
  "BOENI":["AMBATO-BOENI","MAROVOAY","MAHAJANGA I","MAHAJANGA II","MITSINJO","SOALALA"],
  "BONGOLAVA":["FENOARIVOBE","TSIROANOMANDIDY"],
  "DIANA":["AMBANJA","AMBILOBE","ANTSIRANANA I","ANTSIRANANA II","NOSY BE"],
  "FITOVINANY":["IKONGO","MANAKARA","VOHIPENO"],
  "HAUTE MATSIATRA":["AMBALAVAO","FII_VOHIBATO","AMBOHIMAHASOA","FIANARANTSOA I","FII_LALANGINA","FII_ISANDRA","IKALAMAVONY"],
  "IHOROMBE":["IAKORA","IHOSY","IVOHIBE"],
  "ITASY":["ARIVONIMAMO","MIARINARIVO","SOAVINANDRIANA"],
  "MELAKY":["AMBATOMAINTY","MORAFENOBE","ANTSALOVA","MAINTIRANO","BESALAMPY"],
  "MENABE":["BELO/TSIRIBIHINA","MORONDAVA","MANJA","MIANDRIVAZO","MAHABO"],
  "S A V A":["ANDAPA","ANTALAHA","SAMBAVA","VOHEMAR"],
  "SOFIA":["ANALALAVA","ANTSOHIHY","BEALALANA","MAMPIKONY","PORT-BERGER","MANDRITSARA","BEFANDRIANA NORD"],
  "VAKINANKARATRA":["AMBATOLAMPY","ANTANIFOTSY","ANTSIRABE I","ANTSIRABE II","FARATSIHO","BETAFO","MANDOTO"],
  "VATOVAVY":["IFANADIANA","MANANJARY","NOSY VARIKA"]
};

// ═══════════════════════════════════════════════════════
//  COLONNES EXCEL ATTENDUES → clés internes
//  Ordre exact de la feuille "Acteurs Communautaires"
//  (la colonne Région et District sont gérées par le formulaire)
// ═══════════════════════════════════════════════════════
const COL_MAP = {
  // Clé normalisée (minuscules, sans accents/espaces) → champ interne
  'region':           'region',
  'district':         'district',
  'csb':              'csb',
  'fokontany':        'fokontany',
  'nomet prenom':     'nom',
  'nom et prenom':    'nom',
  'nom etprenom':     'nom',
  'nom':              'nom',
  'poste':            'poste',
  'fonction':         'poste',
  'postefinction':    'poste',
  'cin':              'cin',
  'nummvola':         'mvola',
  "numm'vola":        'mvola',
  'nummobile money':  'mvola',
  'mobile money':     'mvola',
  'mvola':            'mvola',
  'mvola/airtel':     'mvola',
};

// ═══════════════════════════════════════════════════════
//  ÉTAT
// ═══════════════════════════════════════════════════════
let rows = [];          // { csb, fokontany, nom, poste, cin, mvola }
let filteredRows = [];  // rows après filtre recherche
let rowIdCounter = 0;
let currentPage = 1;
const PAGE_SIZE = 50;
let searchQuery = '';

// ═══════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  // Remplir région
  const sel = document.getElementById('sel-region');
  Object.keys(REGIONS_DISTRICTS).sort().forEach(r => sel.appendChild(new Option(r, r)));

  // Coller depuis clipboard (Ctrl+V dans le tableau)
  document.getElementById('table-wrap').addEventListener('paste', onPaste);
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
      // Uniquement si le focus n'est pas déjà dans un input
      if (document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'SELECT') {
        document.getElementById('table-wrap').focus();
      }
    }
  });
});

// ═══════════════════════════════════════════════════════
//  LOCALISATION
// ═══════════════════════════════════════════════════════
function onRegionChange() {
  const region = document.getElementById('sel-region').value;
  const sd = document.getElementById('sel-district');
  sd.innerHTML = '<option value="">— Sélectionner un district —</option>';
  if (region && REGIONS_DISTRICTS[region]) {
    REGIONS_DISTRICTS[region].forEach(d => sd.appendChild(new Option(d, d)));
    sd.disabled = false;
  } else {
    sd.disabled = true;
  }
  setErr('grp-region', false);
}

function onDistrictChange() {
  setErr('grp-district', false);
}

// ═══════════════════════════════════════════════════════
//  DRAG & DROP
// ═══════════════════════════════════════════════════════
function onDragOver(e) {
  e.preventDefault();
  document.getElementById('import-zone').classList.add('drag-over');
}
function onDragLeave(e) {
  document.getElementById('import-zone').classList.remove('drag-over');
}
function onDrop(e) {
  e.preventDefault();
  document.getElementById('import-zone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function onFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
  e.target.value = ''; // reset pour permettre re-import du même fichier
}

// ═══════════════════════════════════════════════════════
//  TRAITEMENT FICHIER EXCEL / CSV
// ═══════════════════════════════════════════════════════
function processFile(file) {
  showProgress(true, 'Lecture du fichier…', 20);
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      showProgress(true, 'Analyse des colonnes…', 50);
      let importedRows = [];

      if (file.name.endsWith('.csv')) {
        importedRows = parseCSV(e.target.result);
      } else {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        // Chercher la feuille "Acteurs Communautaires" en priorité
        let sheetName = wb.SheetNames.find(n => normalizeKey(n).includes('acteurscommunautaire'))
                     || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });
        importedRows = mapExcelRows(jsonData);
      }

      showProgress(true, `Import de ${importedRows.length} ligne(s)…`, 80);

      // Append ou remplacer ?
      if (rows.length > 0) {
        if (confirm(`Le tableau contient déjà ${rows.length} ligne(s). Voulez-vous AJOUTER les nouvelles lignes (OK) ou REMPLACER tout (Annuler) ?`)) {
          rows = rows.concat(importedRows);
        } else {
          rows = importedRows;
        }
      } else {
        rows = importedRows;
      }

      showProgress(true, 'Rendu du tableau…', 95);
      setTimeout(() => {
        renderTable();
        showProgress(false);
        showToast(`✓ ${importedRows.length} acteur(s) importé(s) depuis "${file.name}"`, 'success');
      }, 100);

    } catch(err) {
      console.error(err);
      showProgress(false);
      showToast('Erreur lors de la lecture du fichier : ' + err.message, 'error');
    }
  };
  reader.onerror = () => { showProgress(false); showToast('Impossible de lire le fichier.', 'error'); };

  if (file.name.endsWith('.csv')) reader.readAsText(file, 'UTF-8');
  else reader.readAsBinaryString(file);
}

function normalizeKey(str) {
  return String(str).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // supprimer accents
    .replace(/[^a-z0-9']/g, ''); // ne garder que lettres, chiffres, apostrophe
}

function mapExcelRows(jsonData) {
  return jsonData.map(row => {
    const mapped = { csb:'', fokontany:'', nom:'', poste:'', cin:'', mvola:'' };
    Object.keys(row).forEach(col => {
      const key = normalizeKey(col);
      const field = COL_MAP[key];
      if (field && field !== 'region' && field !== 'district') {
        mapped[field] = String(row[col] || '').trim();
      }
    });
    mapped._id = ++rowIdCounter;
    return mapped;
  }).filter(r => r.nom || r.csb || r.fokontany); // ignorer lignes vides
}

function parseCSV(text) {
  const lines = text.split('\n').filter(l => l.trim());
  if (lines.length < 2) return [];
  const headers = lines[0].split(',').map(h => h.replace(/"/g,'').trim());
  return lines.slice(1).map(line => {
    const vals = line.split(',').map(v => v.replace(/"/g,'').trim());
    const row = {};
    headers.forEach((h, i) => row[h] = vals[i] || '');
    const mapped = mapExcelRows([row]);
    return mapped[0];
  }).filter(Boolean);
}

// ═══════════════════════════════════════════════════════
//  COLLER DEPUIS PRESSE-PAPIERS (Ctrl+V)
// ═══════════════════════════════════════════════════════
function onPaste(e) {
  const text = (e.clipboardData || window.clipboardData).getData('text');
  if (!text) return;

  const lines = text.split('\n').filter(l => l.trim());
  if (!lines.length) return;

  // Détecter si la première ligne est un en-tête (contient "nom" ou "csb" etc.)
  const firstLineLower = normalizeKey(lines[0]);
  let dataLines = lines;
  let hasHeader = false;

  if (firstLineLower.includes('nom') || firstLineLower.includes('csb') || firstLineLower.includes('region')) {
    hasHeader = true;
    // Parser avec en-tête
    const headers = lines[0].split('\t').map(h => h.trim());
    dataLines = lines.slice(1).map(line => {
      const vals = line.split('\t');
      const row = {};
      headers.forEach((h, i) => row[h] = vals[i] || '');
      return row;
    });
    const mapped = mapExcelRows(dataLines);
    if (mapped.length) {
      rows = rows.concat(mapped);
      renderTable();
      showToast(`✓ ${mapped.length} ligne(s) collée(s)`, 'success');
      e.preventDefault();
      return;
    }
  }

  // Sinon, coller directement dans l'ordre des colonnes : CSB, Fokontany, Nom, Poste, CIN, M'vola
  const newRows = lines.map(line => {
    const cols = line.split('\t');
    const r = {
      _id: ++rowIdCounter,
      csb:        (cols[0] || '').trim(),
      fokontany:  (cols[1] || '').trim(),
      nom:        (cols[2] || '').trim(),
      poste:      (cols[3] || '').trim(),
      cin:        (cols[4] || '').trim(),
      mvola:      (cols[5] || '').trim(),
    };
    return r;
  }).filter(r => r.nom || r.csb);

  if (newRows.length) {
    rows = rows.concat(newRows);
    renderTable();
    showToast(`✓ ${newRows.length} ligne(s) collée(s)`, 'success');
    e.preventDefault();
  }
}

// ═══════════════════════════════════════════════════════
//  RENDU DU TABLEAU
// ═══════════════════════════════════════════════════════
function renderTable() {
  // Appliquer filtre recherche
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    filteredRows = rows.filter(r =>
      (r.nom||'').toLowerCase().includes(q) ||
      (r.csb||'').toLowerCase().includes(q) ||
      (r.fokontany||'').toLowerCase().includes(q) ||
      (r.poste||'').toLowerCase().includes(q) ||
      (r.cin||'').toLowerCase().includes(q)
    );
  } else {
    filteredRows = [...rows];
  }

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;

  const start = (currentPage - 1) * PAGE_SIZE;
  const pageRows = filteredRows.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('ac-tbody');
  tbody.innerHTML = '';

  if (filteredRows.length === 0) {
    tbody.innerHTML = `
      <tr id="ac-empty-row">
        <td colspan="8" class="ac-empty">
          <div class="ac-empty-icon">👥</div>
          <div>${rows.length > 0 ? 'Aucun résultat pour cette recherche.' : 'Aucun acteur. Importez un fichier Excel ou ajoutez des lignes manuellement.'}</div>
        </td>
      </tr>`;
    document.getElementById('stats-bar').style.display = 'none';
    renderPagination(0, 0);
    updateCount();
    return;
  }

  pageRows.forEach((row, pageIdx) => {
    const globalIdx = rows.indexOf(row); // index dans rows (pour édition/suppression)
    const displayNum = start + pageIdx + 1;
    const tr = document.createElement('tr');
    tr.dataset.idx = globalIdx;
    tr.innerHTML = `
      <td class="col-num">${displayNum}</td>
      <td><input type="text" class="f-csb"  value="${esc(row.csb)}"  placeholder="Nom du CSB"        onchange="updateRow(${globalIdx},'csb',this.value)"></td>
      <td><input type="text" class="f-fkt"  value="${esc(row.fokontany)}" placeholder="Fokontany"    onchange="updateRow(${globalIdx},'fokontany',this.value)"></td>
      <td><input type="text" class="f-nom"  value="${esc(row.nom)}"  placeholder="Nom et Prénom"     onchange="updateRow(${globalIdx},'nom',this.value)"></td>
      <td><input type="text" class="f-pos"  value="${esc(row.poste)}" placeholder="Poste/Fonction"   onchange="updateRow(${globalIdx},'poste',this.value)"></td>
      <td><input type="text" class="f-cin"  value="${esc(row.cin)}"  placeholder="XXX XXX XXX XXX"   maxlength="15" onkeydown="onlyDigits(event)" oninput="formatCIN(this)" inputmode="numeric" autocomplete="off" onchange="updateRow(${globalIdx},'cin',this.value)"></td>
      <td><input type="text" class="f-mvola warn" value="${esc(row.mvola)}" placeholder="03X XX XXX XX" maxlength="13" onkeydown="onlyDigits(event)" oninput="formatMobileMoney(this)" inputmode="numeric" autocomplete="off" onchange="updateRow(${globalIdx},'mvola',this.value)"></td>
      <td class="col-act"><button class="btn-del-row" onclick="deleteRow(${globalIdx})">✕</button></td>
    `;
    tbody.appendChild(tr);
  });

  updateStats();
  renderPagination(filteredRows.length, totalPages);
  updateCount();
  document.getElementById('stats-bar').style.display = 'flex';
}

function esc(str) {
  return String(str||'').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function updateRow(idx, field, value) {
  if (rows[idx]) rows[idx][field] = value;
}

function deleteRow(idx) {
  rows.splice(idx, 1);
  renderTable();
}

function addRow() {
  rows.push({ _id: ++rowIdCounter, csb:'', fokontany:'', nom:'', poste:'', cin:'', mvola:'' });
  // Aller à la dernière page
  currentPage = Math.ceil(rows.length / PAGE_SIZE) || 1;
  renderTable();
  // Focus sur le premier champ de la dernière ligne
  setTimeout(() => {
    const lastRow = document.querySelector('#ac-tbody tr:last-child input.f-csb');
    if (lastRow) lastRow.focus();
  }, 50);
}

function clearAll() {
  if (rows.length === 0) return;
  if (!confirm(`Supprimer les ${rows.length} acteur(s) du tableau ?`)) return;
  rows = [];
  currentPage = 1;
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  STATS & PAGINATION
// ═══════════════════════════════════════════════════════
function updateStats() {
  document.getElementById('stat-total').textContent  = filteredRows.length;
  document.getElementById('stat-cin').textContent    = filteredRows.filter(r => r.cin && r.cin.trim()).length;
  document.getElementById('stat-mvola').textContent  = filteredRows.filter(r => r.mvola && r.mvola.trim()).length;
}

function updateCount() {
  const n = rows.length;
  document.getElementById('ac-count').textContent = `(${n} acteur${n>1?'s':''})`;
}

function renderPagination(total, totalPages) {
  const pg = document.getElementById('pagination');
  const info = document.getElementById('page-info');
  pg.innerHTML = '';

  if (totalPages <= 1) { info.textContent = total > 0 ? `${total} acteur(s)` : ''; return; }

  const start = (currentPage - 1) * PAGE_SIZE + 1;
  const end   = Math.min(currentPage * PAGE_SIZE, total);
  info.textContent = `${start}–${end} sur ${total}`;

  // Prev
  const prev = document.createElement('button');
  prev.className = 'page-btn'; prev.textContent = '‹';
  prev.disabled = currentPage === 1;
  prev.onclick = () => { currentPage--; renderTable(); };
  pg.appendChild(prev);

  // Pages
  for (let p = 1; p <= totalPages; p++) {
    if (totalPages > 7 && Math.abs(p - currentPage) > 2 && p !== 1 && p !== totalPages) {
      if (p === 2 || p === totalPages - 1) { const s=document.createElement('span'); s.textContent='…'; s.style.color='var(--muted)'; s.style.padding='0 4px'; pg.appendChild(s); }
      continue;
    }
    const btn = document.createElement('button');
    btn.className = 'page-btn' + (p===currentPage?' active':'');
    btn.textContent = p;
    btn.onclick = (pp => () => { currentPage = pp; renderTable(); })(p);
    pg.appendChild(btn);
  }

  // Next
  const next = document.createElement('button');
  next.className = 'page-btn'; next.textContent = '›';
  next.disabled = currentPage === totalPages;
  next.onclick = () => { currentPage++; renderTable(); };
  pg.appendChild(next);
}

// ═══════════════════════════════════════════════════════
//  RECHERCHE
// ═══════════════════════════════════════════════════════
function onSearch() {
  searchQuery = document.getElementById('search-input').value.trim();
  currentPage = 1;
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  VALIDATION
// ═══════════════════════════════════════════════════════
function validate() {
  let ok = true;
  const region   = document.getElementById('sel-region').value;
  const district = document.getElementById('sel-district').value;
  setErr('grp-region',   !region);   if (!region)   ok = false;
  setErr('grp-district', !district); if (!district) ok = false;
  if (rows.length === 0) {
    showToast('Le tableau est vide. Ajoutez au moins un acteur.', 'error');
    ok = false;
  }
  // Vérifier que chaque ligne a au moins un nom
  const missing = rows.filter(r => !r.nom.trim()).length;
  if (missing > 0) {
    showToast(`${missing} ligne(s) sans nom. Veuillez les renseigner ou les supprimer.`, 'warn');
    ok = false;
  }
  return ok;
}

// ═══════════════════════════════════════════════════════
//  SOUMETTRE À GOOGLE SHEET
// ═══════════════════════════════════════════════════════
async function submitData() {
  if (!validate()) return;

  const region   = document.getElementById('sel-region').value;
  const district = document.getElementById('sel-district').value;

  const acteurs = rows.map(r => ({
    csb:       r.csb.trim(),
    fokontany: r.fokontany.trim(),
    nom:       r.nom.trim(),
    poste:     r.poste.trim(),
    cin:       r.cin.trim(),
    mvola:     r.mvola.trim(),
  })).filter(r => r.nom);

  const payload = {
    action:   'saveActeursCommunautaires',
    region,
    district,
    acteurs
  };

  showToast('Envoi en cours…', '');
  try {
    const res = await fetch(GS_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body:    JSON.stringify(payload)
    });
    const result = await res.json();
    if (result.status === 'ok') {
      showToast(`✓ ${acteurs.length} acteur(s) enregistré(s) dans Google Sheet !`, 'success');
    } else {
      showToast(result.message || 'Erreur Apps Script', 'error');
    }
  } catch (err) {
    console.error(err);
    showToast('Erreur de connexion avec Google Sheet', 'error');
  }
}

// ═══════════════════════════════════════════════════════
//  CHARGER DEPUIS GOOGLE SHEET
// ═══════════════════════════════════════════════════════
async function loadData() {
  const region   = document.getElementById('sel-region').value;
  const district = document.getElementById('sel-district').value;
  if (!region || !district) { showToast('Sélectionnez une région et un district.', 'error'); return; }

  showToast('Chargement des données…', '');
  try {
    const url = `${GS_URL}?action=getActeursCommunautaires&region=${encodeURIComponent(region)}&district=${encodeURIComponent(district)}`;
    const res  = await fetch(url);
    const data = await res.json();

    if (data.status === 'ok' && data.acteurs && data.acteurs.length > 0) {
      rows = data.acteurs.map(a => ({
        _id:       ++rowIdCounter,
        csb:       a.csb       || '',
        fokontany: a.fokontany || '',
        nom:       a.nom       || '',
        poste:     a.poste     || '',
        cin:       a.cin       || '',
        mvola:     a.mvola     || '',
      }));
      currentPage = 1;
      renderTable();
      showToast(`✓ ${rows.length} acteur(s) chargé(s) depuis Google Sheet`, 'success');
    } else {
      showToast(data.message || 'Aucune donnée trouvée pour ce district.', 'warn');
    }
  } catch (err) {
    console.error(err);
    showToast('Erreur de connexion avec Google Sheet', 'error');
  }
}

// ═══════════════════════════════════════════════════════
//  EXPORTER CSV
// ═══════════════════════════════════════════════════════
function exportCSV() {
  if (rows.length === 0) { showToast('Aucune donnée à exporter.', 'warn'); return; }
  const region   = document.getElementById('sel-region').value   || 'REGION';
  const district = document.getElementById('sel-district').value || 'DISTRICT';

  const headers = ['Région','District','CSB','Fokontany','Nom et Prénom','Poste','CIN',"Num M'vola"];
  const lines = [headers.join(',')];
  rows.forEach(r => {
    lines.push([region, district, r.csb, r.fokontany, r.nom, r.poste, r.cin, r.mvola]
      .map(v => `"${String(v||'').replace(/"/g,'""')}"`)
      .join(','));
  });

  const blob = new Blob(['\uFEFF' + lines.join('\n')], { type:'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `ActeursCommunautaires_${district}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  showToast('✓ Export CSV téléchargé', 'success');
}

// ═══════════════════════════════════════════════════════
//  RÉINITIALISER
// ═══════════════════════════════════════════════════════
function resetAll() {
  if (!confirm('Réinitialiser tout le formulaire ?')) return;
  rows = [];
  rowIdCounter = 0;
  currentPage = 1;
  searchQuery = '';
  document.getElementById('sel-region').value = '';
  document.getElementById('sel-district').innerHTML = '<option value="">— Sélectionner un district —</option>';
  document.getElementById('sel-district').disabled = true;
  document.getElementById('search-input').value = '';
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  PROGRÈS
// ═══════════════════════════════════════════════════════
function showProgress(visible, msg, pct) {
  const el = document.getElementById('upload-progress');
  if (!visible) { el.classList.remove('visible'); return; }
  el.classList.add('visible');
  document.getElementById('upload-msg').textContent = msg || '';
  document.getElementById('progress-bar').style.width = (pct||0) + '%';
}

// ═══════════════════════════════════════════════════════
//  UTILITAIRES PARTAGÉS
// ═══════════════════════════════════════════════════════
function setErr(groupId, hasError) {
  const g = document.getElementById(groupId);
  if (g) g.classList.toggle('has-error', hasError);
}

function showToast(msg, type) {
  const t = document.getElementById('toast');
  document.getElementById('toast-msg').textContent = msg;
  t.className = type ? `show ${type}` : 'show';
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.className = type || ''; }, 4000);
}

function onlyDigits(event) {
  if (!/[0-9]/.test(event.key) &&
      event.key !== 'Backspace' &&
      event.key !== 'Delete' &&
      event.key !== 'ArrowLeft' &&
      event.key !== 'ArrowRight') {
    event.preventDefault();
  }
}

function formatCIN(input) {
  let v = input.value.replace(/\D/g,'').substring(0, 12);
  input.value = v.replace(/(\d{0,3})(\d{0,3})(\d{0,3})(\d{0,3})/,
    (_, a, b, c, d) => [a,b,c,d].filter(Boolean).join(' '));
}

function formatMobileMoney(input) {
  let v = input.value.replace(/\D/g,'').substring(0, 10);
  let res = '';
  if (v.length > 0) res += v.substring(0,3);
  if (v.length > 3) res += ' ' + v.substring(3,5);
  if (v.length > 5) res += ' ' + v.substring(5,8);
  if (v.length > 8) res += ' ' + v.substring(8,10);
  input.value = res;
}
