const BU_CONFIG = {
  'ESG':                  { colore: '#7ebd4b', label: 'ESG' },
  'Business Performance': { colore: '#FF6633', label: 'Business Performance' },
  'PCO':                  { colore: '#EE6784', label: 'PCO' },
  'ALL BU':               { colore: '#002c49', label: 'ALL BU' },
};

window.clientiData = [];
window.businessUnit = null; // 'Energy' o 'Business Performance'

// ── UPLOAD ──────────────────────────────────────
const fileInput = document.getElementById('file-input');
const dropZone  = document.getElementById('drop-zone');
const btnGenera = document.getElementById('btn-genera');
const uploadStatus = document.getElementById('upload-status');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});
btnGenera.addEventListener('click', renderAll);

function setupCambiaVista() {
  const btn = document.querySelector('.btn-cambia-vista');
  if (btn) {
    btn.onclick = null;
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      e.preventDefault();
      cambiaBU();
    }, true); // usa capture phase
  }
}

function parseXLSAsHTML(arrayBuffer) {
  const decoder = new TextDecoder('iso-8859-1');
  const text = decoder.decode(arrayBuffer);
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, 'text/html');
  const rows = doc.querySelectorAll('tr');

  const headers = [];
  rows[0].querySelectorAll('th').forEach(th => headers.push(th.textContent.trim()));

  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const cells = rows[i].querySelectorAll('td');
    if (cells.length === 0) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = cells[idx] ? cells[idx].textContent.trim() : '';
    });
    result.push(obj);
  }
  return result;
}

async function handleFile(file) {
  const reader = new FileReader();
  reader.onload = async function(e) {
    const ATECO_MAPPING = await (await fetch('assets/data/mapping.json')).json();
    const mappingPerCodice = {};
    for (const [key, val] of Object.entries(ATECO_MAPPING)) {
      const m = key.match(/^(\d{2})/);
      if (m) mappingPerCodice[m[1]] = val;
    }

    // Prova prima con parseXLSAsHTML, poi fallback a SheetJS
    let rows;
    try {
      rows = parseXLSAsHTML(e.target.result);
      if (!rows[0]['BU']) throw new Error('not HTML XLS');
    } catch {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const sheetName = wb.SheetNames[0];
      rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    }

    window.clientiDataRaw = rows; // tutte le righe, senza deduplicazione

    const rows_mapped = rows;

    // Mapping BU
    const BU_MAP = {
      'cir': 'Business Performance',
      'national tax': 'Business Performance',
      'dvlt': 'Business Performance',
      'perf env': 'Business Performance',
      'lsu': 'Business Performance',
      'grants': 'Business Performance',
      'nrj it': 'ESG',
      'hr performance': 'PCO',
      'hr perf': 'PCO',
    };

    // Per ogni riga mappa il settore da Sub-industry e la BU
    const rowsProcessed = rows_mapped.map(r => {
      const subIndustry = String(r['Sub-industry'] || '').trim();
      const codice2 = (subIndustry.match(/^(\d{2})/) || [])[1] || null;
      const buRaw = String(r['BU'] || '').trim().toLowerCase();
      return {
        ...r,
        'Mapping': (codice2 ? mappingPerCodice[codice2] : null) || null,
        'Business Unit': BU_MAP[buRaw] || null,
      };
    });

    // Filtra righe senza BU mappata
    const rowsValide = rowsProcessed.filter(r => r['Business Unit'] !== null);

    // Calcola somma RDTC per account
    const rdtcPerAccount = {};
    rowsValide.forEach(r => {
      const account = String(r['Account Name'] || '').trim();
      const rdtc = parseFloat(String(r['RDTC Amount'] || '').replace(',', '.')) || 0;
      if (!rdtcPerAccount[account]) rdtcPerAccount[account] = 0;
      rdtcPerAccount[account] += rdtc;
    });

    // Deduplicazione: per ogni coppia Account Name + Business Unit, tieni solo la prima riga
    const seen = new Set();
    const deduplicati = rowsValide.filter(r => {
      const key = String(r['Account Name'] || '').trim() + '|||' + r['Business Unit'];
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    // Aggiungi somma RDTC a ogni riga deduplicata
    window.clientiData = deduplicati.map(r => ({
      ...r,
      'RDTC Totale': rdtcPerAccount[String(r['Account Name'] || '').trim()] || 0
    }));

    // Passa allo step 2
    setStep(2);
    document.getElementById('step2-back-wrap')?.remove();
    const back2Wrap = document.createElement('div');
    back2Wrap.id = 'step2-back-wrap';
    back2Wrap.innerHTML = `<button class="step3-back" onclick="setStep(1)">←</button>`;
    document.getElementById('upload-section').appendChild(back2Wrap);

    const counts = {};
    Object.keys(BU_CONFIG).forEach(bu => {
      counts[bu] = window.clientiData.filter(r => r['Business Unit'] === bu).length;
    });
    document.getElementById('count-esg').innerHTML = `<span class="bu-count" style="color:#7ebd4b">${counts['ESG'] || 0} aziende ESG</span>`;
    document.getElementById('count-bp').innerHTML = `<span class="bu-count" style="color:#FF6633">${counts['Business Performance'] || 0} aziende BP</span>`;
    document.getElementById('count-pco').innerHTML = `<span class="bu-count" style="color:#EE6784">${counts['PCO'] || 0} aziende PCO</span>`;
    document.getElementById('count-allbu').innerHTML = `<span class="bu-count" style="color:#002c49">${window.clientiData.length} aziende (tutte le BU)</span>`;
  };
  reader.readAsArrayBuffer(file);
}

function selezionaBU(bu) {
  window.businessUnit = bu;
  // Filtra i dati per Business Unit
  if (bu === 'ALL BU') {
    window.clientiFiltrati = [...window.clientiData];
  } else {
    window.clientiFiltrati = window.clientiData.filter(r =>
      String(r['Business Unit'] || '').trim() === bu
    );
  }
  // Passa allo step 3
  setStep(3);
  const colore = BU_CONFIG[bu]?.colore || '#7ebd4b';
  document.documentElement.style.setProperty('--accent', colore);
  // Aggiungi freccia in alto a destra nel contenitore upload-section
  document.getElementById('step2-back-wrap')?.remove();
  document.getElementById('step3-back-wrap')?.remove();
  const backWrap = document.createElement('div');
  backWrap.id = 'step3-back-wrap';
  backWrap.innerHTML = `<button class="step3-back" onclick="setStep(2)">←</button>`;
  document.getElementById('upload-section').appendChild(backWrap);

  document.getElementById('step3-info').innerHTML = `
    <div class="step3-bu-name" style="color:${colore}">${bu}</div>
    <div class="step3-count">${window.clientiFiltrati.length} <span>aziende pronte</span></div>`;
}

function cambiaBU() {
  window.businessUnit = null;
  window.clientiFiltrati = [];
  document.getElementById('app').style.display = 'none';
  document.getElementById('upload-section').style.display = 'block';
  document.documentElement.style.setProperty('--accent', '#7ebd4b');
  setStep(2); // torna alla selezione BU senza ricaricare il file
}

function setStep(n) {
  // Rimuovi eventuali frecce precedenti
  document.querySelectorAll('.step3-back').forEach(el => el.remove());
  document.getElementById('step3-back-wrap')?.remove();
  // Aggiorna contenuti
  document.querySelectorAll('.step-content').forEach(el => el.style.display = 'none');
  document.getElementById(`step-content-${n}`).style.display = 'flex';

  // Aggiorna circles
  document.querySelectorAll('.step').forEach((el, i) => {
    el.classList.remove('active', 'done');
    if (i + 1 < n) el.classList.add('done');
    if (i + 1 === n) el.classList.add('active');
  });
}

// ── RENDER ALL ──────────────────────────────────
function renderAll() {
  document.getElementById('upload-section').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  renderDimensioni();
  renderSettori();
  renderDistribuzione();
  setupCambiaVista();
  aggiornaLogo();
  setupRicerca();
}

function aggiornaLogo() {
  const logoWrap = document.querySelector('.leyton-logo');
  if (!logoWrap) return;
  let src = 'assets/logo_leyton.svg';
  let alt = 'Leyton';
  if (window.businessUnit === 'ESG') {
    src = 'assets/leyton_esg_3.svg';
    alt = 'Leyton ESG';
  } else if (window.businessUnit === 'PCO') {
    src = 'assets/logo_pco.svg';
    alt = 'Leyton PCO';
  } else if (window.businessUnit === 'ALL BU') {
    src = 'assets/logo_leyton.svg';
    alt = 'Leyton';
  }
  logoWrap.innerHTML = `<img src="${src}" alt="${alt}" style="height:32px; width:auto; display:block;">`;
}

// ── DIMENSIONI ──────────────────────────────────
function renderDimensioni() {
  const dati = window.clientiFiltrati;
  const total = dati.length;
  let grandi = 0, medie = 0, piccole = 0, micro = 0;
  dati.forEach(r => {
    const emp = parseFloat(r['Employees']);
    if (isNaN(emp)) return;
    if (emp > 250)      grandi++;
    else if (emp >= 50) medie++;
    else if (emp >= 10) piccole++;
    else                micro++;
  });

  const W = 420, H = 320;
  const cx = 200, cy = 160, R = 100, strokeW = 22;
  const circ = 2 * Math.PI * R;
  const gap = 3;

  const segmenti = [
    { val: grandi,  color: '#002c49', label: 'grandi imprese',  sub: '(>250 dipendenti)',  angle: -90  },
    { val: micro,   color: '#cdd5dc', label: 'microimprese',    sub: '(<10 dipendenti)',   angle: -20  },
    { val: piccole, color: '#9cabb6', label: 'piccole imprese', sub: '(<50 dipendenti)',   angle: 40   },
    { val: medie,   color: '#526d81', label: 'medie imprese',   sub: '(<250 dipendenti)', angle: 160  },
  ];

  // Calcola archi
  const totalValid = grandi + medie + piccole + micro;
  let archi = '';
  let currentAngle = -90; // parte dall'alto

  // Cerchio sfondo
  archi += `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="#e5e8eb" stroke-width="${strokeW}"/>`;

  const segmentiOrdinati = [
    { val: grandi,  color: '#002c49', key: 'grandi',  label: 'Grandi Imprese'  },
    { val: medie,   color: '#526d81', key: 'medie',   label: 'Medie Imprese'   },
    { val: piccole, color: '#9cabb6', key: 'piccole', label: 'Piccole Imprese' },
    { val: micro,   color: '#cdd5dc', key: 'micro',   label: 'Micro Imprese'   },
  ];

  segmentiOrdinati.forEach(s => {
    if (s.val === 0) return;
    const dash = (s.val / totalValid) * (circ - gap * 4);
    const offset = circ / 4 - currentAngle * (circ / 360);
    archi += `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none"
      stroke="${s.color}" stroke-width="${strokeW}"
      stroke-dasharray="${dash.toFixed(2)} ${(circ - dash).toFixed(2)}"
      stroke-dashoffset="${(circ / 4).toFixed(2)}"
      transform="rotate(${currentAngle + 90} ${cx} ${cy})"/>`;
    currentAngle += (s.val / totalValid) * 360 + (gap * 360 / circ);
  });

  // Calcola angolo medio di ogni segmento per posizionare le etichette
  const labelData = [];
  let angleAcc = -90;
  const segWithLabel = [
    { val: grandi,  color: '#002c49', label: 'grandi imprese',  sub: '(>250 dipendenti)'  },
    { val: medie,   color: '#526d81', label: 'medie imprese',   sub: '(<250 dipendenti)'  },
    { val: piccole, color: '#9cabb6', label: 'piccole imprese', sub: '(<50 dipendenti)'   },
    { val: micro,   color: '#cdd5dc', label: 'microimprese',    sub: '(<10 dipendenti)'   },
  ];

  segWithLabel.forEach(s => {
    const sweep = (s.val / totalValid) * 360;
    const midAngle = angleAcc + sweep / 2;
    const rad = midAngle * Math.PI / 180;
    const lineR1 = R + strokeW / 2 + 8;  // punto sul bordo esterno
    const lineR2 = R + strokeW / 2 + 24; // punto fine linea
    const x1 = cx + lineR1 * Math.cos(rad);
    const y1 = cy + lineR1 * Math.sin(rad);
    const x2 = cx + lineR2 * Math.cos(rad);
    const y2 = cy + lineR2 * Math.sin(rad);
    const isRight = x2 > cx;
    const textX = isRight ? x2 + 8 : x2 - 8;
    const textAnchor = isRight ? 'start' : 'end';
    labelData.push({ val: s.val, label: s.label, sub: s.sub, x1, y1, x2, y2, textX, y2, textAnchor, color: s.color });
    angleAcc += sweep + (gap * 360 / circ);
  });

  const accentColor = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim();

  // Genera linee e testi
  let linee = '';
  labelData.forEach(l => {
    linee += `
      <line x1="${l.x1.toFixed(1)}" y1="${l.y1.toFixed(1)}"
            x2="${l.x2.toFixed(1)}" y2="${l.y2.toFixed(1)}"
            stroke="#9cabb6" stroke-width="1"/>
      <text x="${l.textX.toFixed(1)}" y="${(l.y2 - 10).toFixed(1)}"
        text-anchor="${l.textAnchor}"
        font-family="Barlow Condensed, sans-serif"
        font-size="22" font-weight="900" fill="#002c49">${l.val}</text>
      <text x="${l.textX.toFixed(1)}" y="${(l.y2 + 4).toFixed(1)}"
        text-anchor="${l.textAnchor}"
        font-family="Barlow, sans-serif"
        font-size="9" fill="#002c49">${l.label}</text>
      <text x="${l.textX.toFixed(1)}" y="${(l.y2 + 14).toFixed(1)}"
        text-anchor="${l.textAnchor}"
        font-family="Barlow, sans-serif"
        font-size="7.5" fill="#6a8093">${l.sub}</text>`;
  });

  document.getElementById('dimensioni').innerHTML = `
    <div class="dim-inner">
      <svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
        ${archi}
        ${linee}
        <text x="${cx}" y="${cy - 8}" text-anchor="middle" dominant-baseline="middle"
          font-family="Barlow Condensed, sans-serif"
          font-size="42" font-weight="900" fill="${accentColor}">${total}</text>
        <text x="${cx}" y="${cy + 20}" text-anchor="middle" dominant-baseline="middle"
          font-family="Barlow, sans-serif"
          font-size="13" fill="#002c49">aziende</text>
      </svg>
    </div>`;

  // Aggancia click ai segmenti donut dopo il render
  setTimeout(() => {
    const svgEl = document.querySelector('#dimensioni svg');
    if (!svgEl) return;

    // I circle colorati sono dal 2° in poi (il primo è il cerchio grigio di sfondo)
    const circles = svgEl.querySelectorAll('circle');
    const mapping = [
      { key: 'grandi',  label: 'Grandi Imprese'  },
      { key: 'medie',   label: 'Medie Imprese'   },
      { key: 'piccole', label: 'Piccole Imprese'  },
      { key: 'micro',   label: 'Micro Imprese'    },
    ];

    // Salta il primo circle (sfondo grigio), prendi i successivi 4
    let idx = 0;
    circles.forEach((c, i) => {
      if (i === 0) return; // skip sfondo
      if (idx >= mapping.length) return;
      const m = mapping[idx];
      c.style.cursor = 'pointer';
      c.setAttribute('pointer-events', 'stroke');
      c.addEventListener('click', () => mostraPopupDimensione(m.key, m.label));
      c.addEventListener('mouseenter', function() { this.style.opacity = '0.75'; });
      c.addEventListener('mouseleave', function() { this.style.opacity = '1'; });
      idx++;
    });
  }, 300);
}

function mostraPopupDimensione(key, label) {
  // Filtra aziende per dimensione
  const aziende = window.clientiFiltrati.filter(r => {
    const emp = parseFloat(r['Employees']);
    if (isNaN(emp)) return false;
    if (key === 'grandi')  return emp > 250;
    if (key === 'medie')   return emp >= 50 && emp <= 250;
    if (key === 'piccole') return emp >= 10 && emp < 50;
    if (key === 'micro')   return emp < 10;
    return false;
  });

  const beneficioTotaleDim = aziende.reduce((sum, r) => sum + (parseFloat(r['RDTC Totale']) || 0), 0);
  const beneficioFormattaDim = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(beneficioTotaleDim);

  // Raggruppa per regione → provincia (stesso schema modale settori)
  const perRegione = {};
  aziende.forEach(r => {
    const regione = String(r['Billing State/Province (text only)'] || 'Sconosciuta').trim().toUpperCase();
    const provincia = String(r['Province'] || 'Altre').trim();
    const nome = String(r['Account Name'] || '');
    if (!perRegione[regione]) perRegione[regione] = {};
    if (!perRegione[regione][provincia]) perRegione[regione][provincia] = [];
    perRegione[regione][provincia].push(nome);
  });

  // Genera HTML regioni ordinate alfabeticamente
  const bodyHtml = Object.keys(perRegione).sort().map(regione => {
    const province = Object.keys(perRegione[regione]).sort().map(prov => {
      const aziList = perRegione[regione][prov].sort().map(n =>
        `<div class="popup-azienda popup-azienda-link" data-account="${n.replace(/"/g, '&quot;')}"><span class="popup-slash">/</span> ${n.toUpperCase()}</div>`
      ).join('');
      return `<div class="popup-provincia">
        <div class="popup-provincia-nome">${prov}</div>
        ${aziList}
      </div>`;
    }).join('');
    return `<div class="popup-regione">
      <div class="popup-regione-nome">${regione}</div>
      ${province}
    </div>`;
  }).join('');

  // Crea overlay e popup
  document.getElementById('popup-overlay')?.remove();
  document.getElementById('regione-popup')?.remove();

  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';

  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">${label}</div>
        <div class="popup-count">${aziende.length} <span>aziende</span></div>
      </div>
      <div class="popup-header-right">
        <div class="popup-beneficio-label">Beneficio</div>
        <div class="popup-beneficio">${beneficioFormattaDim}</div>
      </div>
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body" style="column-count:3">
      ${bodyHtml}
    </div>`;

  const closeDim = () => {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  };
  overlay.addEventListener('click', closeDim);
  popup.querySelector('.popup-close').addEventListener('click', closeDim);

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
  popup.querySelectorAll('.popup-azienda-link').forEach(el => {
    el.addEventListener('click', () => apriDettaglioAzienda(el.dataset.account));
  });
}

// ── SETTORI ─────────────────────────────────────
function renderSettori() {
  const SETTORI = [
    { key:'ALIMENTARE',       label:'Alimentare',     icon:'menu_alimentari.png'              },
    { key:'AUTOMOTIVE',       label:'Automotive',     icon:'menu_automotive.png'              },
    { key:'CHIMICO',          label:'Chimico',        icon:'menu_chimico.png'                 },
    { key:'COMMERCIO',        label:'Commercio',      icon:'menu_commercio.png'               },
    { key:'CONSULENZA',       label:'Consulenza',     icon:'menu_consulenza.png'              },
    { key:'COSTRUZIONI',      label:'Costruzioni',    icon:'menu_costruzioni.png'             },
    { key:'FARMACEUTICO',     label:'Pharma',         icon:'menu_pharma.png'                  },
    { key:'HORECA',           label:'Ho.re.ca.',      icon:'menu_horeca.png'                  },
    { key:'IT',               label:'IT',             icon:'menu_it.png'                      },
    { key:'MANIFATTURIERO',   label:'Manifatturiero', icon:'menu_manifatturiero.png'          },
    { key:'MODA',             label:'Moda',           icon:'menu_moda.png'                    },
    { key:'RICERCA',          label:'Ricerca',        icon:'menu_ricerca.png'                 },
    { key:'SALUTE',           label:'Salute',         icon:'menu_salute.png'                  },
    { key:'SERVIZI',          label:'Servizi',        icon:'menu_servizi.png'                 },
    { key:'TELECOMUNICAZIONI',label:'Telecom.',       icon:'menu_telecomunicazioni.png'       },
    { key:'TRASPORTI',        label:'Trasporti',      icon:'menu_trasporti.png'               },
  ];

  const conteggi = {};
  window.clientiFiltrati.forEach(r => {
    const m = String(r['Mapping'] || '').trim().toUpperCase();
    if (m) conteggi[m] = (conteggi[m] || 0) + 1;
  });

  const tiles = SETTORI.map(s => {
    const cnt = conteggi[s.key] || 0;
    return `<div class="settore-tile${cnt === 0 ? ' zero' : ''}">
      <div class="settore-icon">
        <img src="assets/icons/${s.icon}" alt="${s.label}">
      </div>
      <div class="settore-name">${s.label}</div>
      <div class="settore-count">${cnt > 0 ? cnt : '—'}</div>
    </div>`;
  }).join('');

  document.getElementById('settore').innerHTML =
    `<div class="settore-grid">${tiles}</div>`;

  const tileNodes = document.querySelectorAll('.settore-tile');
  tileNodes.forEach((tile, index) => {
    const s = SETTORI[index];
    tile.addEventListener('click', () => {
      mostraPopupSettore(s.key, s.label);
    });
  });
}

// ── DISTRIBUZIONE ───────────────────────────────
function renderDistribuzione() {
  const NORD = new Set(['lombardia','veneto','piemonte','emilia-romagna',
    'emilia romagna','liguria','trentino-alto adige','trento','bolzano',
    'friuli venezia giulia',"valle d'aosta",'aosta']);
  const CENTRO = new Set(['toscana','lazio','umbria','marche']);
  const SUD = new Set(['campania','puglia','calabria','sicilia','sardegna',
    'basilicata','molise','abruzzo']);

  let nord = 0, centro = 0, sud = 0, intl = 0;
  const intlNomi = [];

  window.clientiFiltrati.forEach(r => {
    const s = String(r['Billing State/Province (text only)'] || '').trim().toLowerCase();
    if (!s) return;
    if (NORD.has(s))        nord++;
    else if (CENTRO.has(s)) centro++;
    else if (SUD.has(s))    sud++;
    else {
      intl++;
      if (intlNomi.length < 5) intlNomi.push(String(r['Account Name'] || '').toUpperCase());
    }
  });

  const nomiHtml = intlNomi.map(n =>
    `<div class="distrib-intl-co">— ${n}</div>`).join('');

  document.getElementById('distribuzione').innerHTML = `
  <div class="distrib-wrapper">
    <div class="distrib-labels-vertical">
      <div class="distrib-item-v">
        <div class="distrib-num">${nord}</div>
        <div class="distrib-word">aziende</div>
        <div class="distrib-place">Nord Italia</div>
        <div class="distrib-line"></div>
      </div>
      <div class="distrib-item-v">
        <div class="distrib-num">${centro}</div>
        <div class="distrib-word">aziende</div>
        <div class="distrib-place">Centro Italia</div>
        <div class="distrib-line"></div>
      </div>
      <div class="distrib-item-v">
        <div class="distrib-num">${sud}</div>
        <div class="distrib-word">aziende</div>
        <div class="distrib-place">Sud Italia e Isole</div>
        <div class="distrib-line"></div>
      </div>
      <div class="distrib-item-v">
        <div class="distrib-num">${intl}</div>
        <div class="distrib-word">aziende</div>
        <div class="distrib-place">Internazionali</div>
      </div>
      <div class="mondo-wrap">
        <img src="assets/mondo.svg" id="icona-mondo" alt="Internazionali"
          class="icona-mondo" onclick="mostraPopupInternazionali()">
      </div>
    </div>
    <div class="distrib-map" id="mappa-container"></div>
  </div>`;
  buildRegioni();
}

async function buildRegioni() {
  const container = document.getElementById('mappa-container');
  container.innerHTML = '';
  container.style.position = 'relative';

  const geojson = await fetch('assets/italy-regions.json').then(r => r.json());

  const width = 260;
  const height = 440;

  const projection = d3.geoMercator().fitSize([width, height], geojson);
  const path = d3.geoPath().projection(projection);

  const svg = d3.create('svg')
    .attr('viewBox', `0 0 ${width} ${height}`)
    .attr('width', width)
    .attr('height', height)
    .style('display', 'block');

  svg.append('defs').html(`
    <filter id="mapShadow">
      <feDropShadow dx="3" dy="3" stdDeviation="2"
        flood-color="#6a8093" flood-opacity="0.5"/>
    </filter>
  `);

  const accent = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim();

  let regioneAttiva = null;

  svg.append('g')
    .attr('filter', 'url(#mapShadow)')
    .selectAll('path')
    .data(geojson.features)
    .join('path')
    .attr('d', path)
    .attr('fill', '#cdd5dc')
    .attr('stroke', '#b8c6cf')
    .attr('stroke-width', 0.5)
    .style('cursor', 'pointer')
    .attr('data-regione', d => {
      const nome = d.properties.NAME_1 || d.properties.reg_name || d.properties.name || '';
      return nome.toLowerCase();
    })
    .on('mouseover', function() {
      if (this !== regioneAttiva) {
        d3.select(this).attr('fill', accent);
      }
    })
    .on('mouseout', function() {
      if (this !== regioneAttiva) {
        d3.select(this).attr('fill', '#cdd5dc');
      }
    })
    .on('click', function(event, d) {
      const nome = d3.select(this).attr('data-regione');
      // Reset tutte le regioni
      svg.selectAll('path').attr('fill', '#cdd5dc');
      // Colora quella cliccata
      d3.select(this).attr('fill', accent);
      regioneAttiva = this;
      mostraPopupRegione(nome);
    });

  container.appendChild(svg.node());
}

function mostraPopupRegione(nomeRegione) {
  const MAPPING_REGIONI = {
    'abruzzo': 'abruzzo',
    'basilicata': 'basilicata',
    'calabria': 'calabria',
    'campania': 'campania',
    'emilia-romagna': 'emilia-romagna',
    'friuli-venezia giulia': 'friuli venezia giulia',
    'lazio': 'lazio',
    'liguria': 'liguria',
    'lombardia': 'lombardia',
    'marche': 'marche',
    'molise': 'molise',
    'piemonte': 'piemonte',
    'puglia': 'puglia',
    'sardegna': 'sardegna',
    'sicilia': 'sicilia',
    'toscana': 'toscana',
    'trentino-alto adige': 'trento',
    'umbria': 'umbria',
    "valle d'aosta": "valle d'aosta",
    'veneto': 'veneto',
  };

  const mappedName = MAPPING_REGIONI[nomeRegione] || nomeRegione;
  const aziendeRegione = window.clientiFiltrati.filter(c => 
    c['Billing State/Province (text only)'] && 
    c['Billing State/Province (text only)'].toLowerCase() === mappedName
  );

  const byProvincia = {};
  aziendeRegione.forEach(c => {
    let prov = c['Province'] ? String(c['Province']).trim() : 'Sconosciuta';
    if (!byProvincia[prov]) byProvincia[prov] = [];
    byProvincia[prov].push({
      nome: c['Account Name'] || 'Azienda Senza Nome',
      settore: c['Mapping'] || ''
    });
  });

  const provinceOrdinate = Object.keys(byProvincia).sort((a,b) => a.localeCompare(b));

  const beneficioTotaleRegione = aziendeRegione.reduce((sum, r) => sum + (parseFloat(r['RDTC Totale']) || 0), 0);
  const beneficioFormattaRegione = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(beneficioTotaleRegione);

  const ICONE_SETTORE = {
    'ALIMENTARE':      'Elenco/elenco_alimentare.svg',
    'AUTOMOTIVE':      'Elenco/elenco_automotive.svg',
    'CHIMICO':         'Elenco/elenco_chimico.svg',
    'COMMERCIO':       'Elenco/elenco_commercio.svg',
    'CONSULENZA':      'Elenco/elenco_consulenza.svg',
    'COSTRUZIONI':     'Elenco/elenco_costruzioni.svg',
    'FARMACEUTICO':    'Elenco/elenco_pharma.svg',
    'HORECA':          'Elenco/elenco_horeca.svg',
    'IT':              'Elenco/elenco_it.svg',
    'MANIFATTURIERO':  'Elenco/elenco_manifatturiero.svg',
    'MODA':            'Elenco/elenco_moda.svg',
    'RICERCA':         'Elenco/elenco_ricerca.svg',
    'SALUTE':          'Elenco/elenco_salute.svg',
    'SERVIZI':         'Elenco/elenco_servizi.svg',
    'TELECOMUNICAZIONI':'Elenco/elenco_telecomunicazioni.svg',
    'TRASPORTI':       'Elenco/elenco_trasporti.svg',
  };

  let htmlBody = '';
  provinceOrdinate.forEach(prov => {
    htmlBody += `<div class="popup-provincia">
      <div class="popup-provincia-nome">${prov.toUpperCase()}</div>`;
    
    const aziende = byProvincia[prov].sort((a,b) => a.nome.localeCompare(b.nome));
    aziende.forEach(a => {
      const iconaPath = ICONE_SETTORE[a.settore.toUpperCase()];
      const iconaHtml = iconaPath ? `<img src="assets/icons/${iconaPath}" class="popup-azienda-icon">` : `<div></div>`;
      htmlBody += `<div class="popup-azienda popup-azienda-link" data-account="${a.nome.replace(/"/g, '&quot;')}">
        ${iconaHtml}
        ${a.nome}
      </div>`;
    });
    htmlBody += `</div>`;
  });

  function chiudiPopup() {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
    // Reset colore regione attiva
    d3.selectAll('.distrib-map path').attr('fill', '#cdd5dc');
  }

  // Se presenti dal precedente click, assicura la rimozione
  chiudiPopup();

  // Crea overlay
  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';
  overlay.addEventListener('click', chiudiPopup);

  // Crea modale
  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">${nomeRegione.toUpperCase()}</div>
        <div class="popup-count">${aziendeRegione.length} <span>aziende</span></div>
      </div>
      <div class="popup-header-right">
        <div class="popup-beneficio-label">Beneficio</div>
        <div class="popup-beneficio">${beneficioFormattaRegione}</div>
      </div>
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body">
      ${htmlBody}
    </div>
  `;

  popup.querySelector('.popup-close').addEventListener('click', chiudiPopup);

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
  popup.querySelectorAll('.popup-azienda-link').forEach(el => {
    el.addEventListener('click', () => apriDettaglioAzienda(el.dataset.account));
  });
}

function mostraPopupSettore(settore, label) {
  const aziendeSettore = window.clientiFiltrati.filter(c => {
    const m = String(c['Mapping'] || '').trim().toUpperCase();
    return m === settore.toUpperCase();
  });

  const byRegione = {};
  aziendeSettore.forEach(c => {
    let regStr = String(c['Billing State/Province (text only)'] || '').trim();
    let reg = regStr !== '' ? regStr : 'Sconosciuta';
    let provStr = String(c['Province'] || '').trim();
    let prov = provStr !== '' ? provStr : 'Altre';
    
    if (!byRegione[reg]) byRegione[reg] = {};
    if (!byRegione[reg][prov]) byRegione[reg][prov] = [];
    byRegione[reg][prov].push({
      nome: c['Account Name'] || 'Azienda Senza Nome'
    });
  });

  const regioniOrdinate = Object.keys(byRegione).sort((a,b) => {
    if (a.toLowerCase() === 'sconosciuta') return 1;
    if (b.toLowerCase() === 'sconosciuta') return -1;
    return a.localeCompare(b);
  });

  let htmlBody = '';
  regioniOrdinate.forEach(reg => {
    htmlBody += `<div class="popup-regione">
      <div class="popup-regione-nome">${reg.toUpperCase()}</div>`;
    
    const byProvincia = byRegione[reg];
    const provinceOrdinate = Object.keys(byProvincia).sort((a,b) => {
      if (a.toLowerCase() === 'altre') return 1;
      if (b.toLowerCase() === 'altre') return -1;
      return a.localeCompare(b);
    });

    provinceOrdinate.forEach(prov => {
      htmlBody += `
      <div class="popup-provincia">
        <div class="popup-provincia-nome">${prov}</div>`;
      
      const aziende = byProvincia[prov].sort((a,b) => a.nome.localeCompare(b.nome));
      aziende.forEach(a => {
        htmlBody += `
        <div class="popup-azienda popup-azienda-link" data-account="${a.nome.replace(/"/g, '&quot;')}"><span class="popup-slash">/</span> ${a.nome}</div>`;
      });
      htmlBody += `
      </div>`;
    });
    htmlBody += `
    </div>`;
  });

  function chiudiPopup() {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
    // Reset colore regione attiva
    d3.selectAll('.distrib-map path').attr('fill', '#cdd5dc');
  }

  chiudiPopup();

  // Crea overlay
  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';
  overlay.addEventListener('click', chiudiPopup);

  const beneficioTotale = aziendeSettore.reduce((sum, r) => sum + (parseFloat(r['RDTC Totale']) || 0), 0);
  const beneficioFormattato = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(beneficioTotale);

  // Crea modale
  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">${label.toUpperCase()}</div>
        <div class="popup-count">${aziendeSettore.length} <span>aziende</span></div>
      </div>
      <div class="popup-header-right">
        <div class="popup-beneficio-label">Beneficio</div>
        <div class="popup-beneficio">${beneficioFormattato}</div>
      </div>
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body">
      ${htmlBody}
    </div>
  `;

  popup.querySelector('.popup-close').addEventListener('click', chiudiPopup);

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
  popup.querySelectorAll('.popup-azienda-link').forEach(el => {
    el.addEventListener('click', () => apriDettaglioAzienda(el.dataset.account));
  });
}

function mostraPopupInternazionali() {
  const aziende = window.clientiFiltrati
    .filter(r => {
      const s = String(r['Billing State/Province (text only)'] || '').trim().toLowerCase();
      const NORD = new Set(['lombardia','veneto','piemonte','emilia-romagna','emilia romagna','liguria','trentino-alto adige','trento','bolzano','friuli venezia giulia',"valle d'aosta",'aosta']);
      const CENTRO = new Set(['toscana','lazio','umbria','marche']);
      const SUD = new Set(['campania','puglia','calabria','sicilia','sardegna','basilicata','molise','abruzzo']);
      return s && !NORD.has(s) && !CENTRO.has(s) && !SUD.has(s);
    })
    .sort((a, b) => String(a['Account Name']).localeCompare(String(b['Account Name'])));

  const ICONE_SETTORE = {
    'ALIMENTARE': 'Elenco/elenco_alimentare.svg',
    'AUTOMOTIVE': 'Elenco/elenco_automotive.svg',
    'CHIMICO': 'Elenco/elenco_chimico.svg',
    'COMMERCIO': 'Elenco/elenco_commercio.svg',
    'CONSULENZA': 'Elenco/elenco_consulenza.svg',
    'COSTRUZIONI': 'Elenco/elenco_costruzioni.svg',
    'FARMACEUTICO': 'Elenco/elenco_pharma.svg',
    'HORECA': 'Elenco/elenco_horeca.svg',
    'IT': 'Elenco/elenco_it.svg',
    'MANIFATTURIERO': 'Elenco/elenco_manifatturiero.svg',
    'MODA': 'Elenco/elenco_moda.svg',
    'RICERCA': 'Elenco/elenco_ricerca.svg',
    'SALUTE': 'Elenco/elenco_salute.svg',
    'SERVIZI': 'Elenco/elenco_servizi.svg',
    'TELECOMUNICAZIONI': 'Elenco/elenco_telecomunicazioni.svg',
    'TRASPORTI': 'Elenco/elenco_trasporti.svg',
  };

  const listaHtml = aziende.map(r => {
    const settore = String(r['Mapping'] || '').trim().toUpperCase();
    const icona = ICONE_SETTORE[settore];
    const iconaHtml = icona
      ? `<img src="assets/icons/${icona}" class="popup-azienda-icon">`
      : `<span style="width:14px;display:inline-block"></span>`;
    const accountName = String(r['Account Name'] || '');
    return `<div class="popup-azienda popup-azienda-link" data-account="${accountName.replace(/"/g, '&quot;')}">${iconaHtml} ${accountName.toUpperCase()}</div>`;
  }).join('');

  function chiudiPopup() {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  }

  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';
  overlay.addEventListener('click', chiudiPopup);

  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">INTERNAZIONALI</div>
        <div class="popup-count">${aziende.length} <span>aziende</span></div>
      </div>
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body" style="column-count:3">
      ${listaHtml}
    </div>`;

  popup.querySelector('.popup-close').addEventListener('click', chiudiPopup);

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
  popup.querySelectorAll('.popup-azienda-link').forEach(el => {
    el.addEventListener('click', () => apriDettaglioAzienda(el.dataset.account));
  });
}

function cercaAzienda() {
  const query = document.getElementById('search-input').value.trim();
  if (!query) return;

  const risultato = window.clientiFiltrati.find(r =>
    String(r['Registered Number'] || '').trim() === query
  );

  if (!risultato) {
    mostraPopupRicerca(null, query);
    return;
  }
  mostraPopupRicerca(risultato, query);
}

function mostraPopupRicerca(azienda, query) {
  document.getElementById('popup-overlay')?.remove();
  document.getElementById('regione-popup')?.remove();

  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';
  overlay.addEventListener('click', () => {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  });

  const popup = document.createElement('div');
  popup.id = 'regione-popup';

  if (!azienda) {
    popup.innerHTML = `
      <div class="popup-header">
        <div class="popup-header-left">
          <div class="popup-titolo">RICERCA</div>
        </div>
        <button class="popup-close">✕</button>
      </div>
      <div class="popup-body" style="column-count:1; text-align:center; padding: 40px;">
        <div style="font-size:14px; color:#6a8093;">Partita IVA <strong>${query}</strong> non presente</div>
      </div>`;
  } else {
    const accent = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim();
    const _anRicerca = String(azienda['Account Name'] || '').trim();
    const buListRicerca = [...new Set((window.clientiData || []).filter(r => String(r['Account Name'] || '').trim() === _anRicerca).map(r => String(r['Business Unit'] || '').trim()).filter(Boolean))].join(', ') || '—';
    popup.innerHTML = `
      <div class="popup-header">
        <div class="popup-header-left">
          <div class="popup-titolo popup-azienda-link" data-account="${String(azienda['Account Name'] || '').replace(/"/g, '&quot;')}" style="cursor:pointer">${String(azienda['Account Name'] || '').toUpperCase()}</div>
        </div>
        <button class="popup-close">✕</button>
      </div>
      <div class="popup-body" style="column-count:1;">
        <div class="ricerca-field">
          <span class="ricerca-label">Partita IVA</span>
          <span class="ricerca-value">${azienda['Registered Number'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Settore</span>
          <span class="ricerca-value">${azienda['Mapping'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Dipendenti</span>
          <span class="ricerca-value">${azienda['Employees'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Regione</span>
          <span class="ricerca-value">${azienda['Billing State/Province (text only)'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Provincia</span>
          <span class="ricerca-value">${azienda['Province'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Business Unit</span>
          <span class="ricerca-value" style="color:${accent}; font-weight:700;">${buListRicerca}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Fatturato</span>
          <span class="ricerca-value">
            ${new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(parseFloat(azienda['Turnover']) || 0)}
          </span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Beneficio</span>
          <span class="ricerca-value" style="color:var(--accent)">
            ${new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(azienda['RDTC Totale'] || 0)}
          </span>
        </div>
      </div>
      ${(() => {
        const accountName = String(azienda['Account Name'] || '').trim();
        const opportunita = (window.clientiDataRaw || []).filter(r =>
          String(r['Account Name'] || '').trim() === accountName
        );
        const righe = opportunita.map(r => `
          <div class="dettaglio-row">
            <span class="dettaglio-opp-nome">${r['Opportunity Name'] || '—'}</span>
            <span class="dettaglio-opp-owner">${r['Opportunity Owner: Full Name'] || '—'}</span>
          </div>`).join('');
        return `<div class="popup-body" style="column-count:1; padding-top:0; border-top:1px solid #e0e5e9;">
          <div class="dettaglio-opp-title" style="margin-bottom:8px">Opportunità (${opportunita.length})</div>
          <div class="dettaglio-opp-list">${opportunita.length === 0 ? '<div class="dettaglio-empty">Nessuna opportunità trovata</div>' : righe}</div>
        </div>`;
      })()}`;
  }

  popup.querySelector('.popup-close').addEventListener('click', () => {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  });

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
  const linkEl = popup.querySelector('.popup-azienda-link[data-account]');
  if (linkEl) linkEl.addEventListener('click', () => apriDettaglioAzienda(linkEl.dataset.account));
}

function setupRicerca() {
  document.getElementById('search-btn').addEventListener('click', cercaAzienda);
  document.getElementById('search-input').addEventListener('keydown', e => {
    if (e.key === 'Enter') cercaAzienda();
  });
}

async function mostraPopupAteco() {
  const mapping = await fetch('assets/data/mapping.json').then(r => r.json());

  document.getElementById('popup-overlay')?.remove();
  document.getElementById('regione-popup')?.remove();

  const overlay = document.createElement('div');
  overlay.id = 'popup-overlay';
  overlay.addEventListener('click', chiudiPopup);

  const righe = Object.entries(mapping).map(([ateco, settore]) => `
    <tr>
      <td class="ateco-desc">${ateco}</td>
      <td class="ateco-settore">${settore}</td>
    </tr>`).join('');

  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">MAPPING ATECO</div>
      </div>
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body" style="column-count:1; padding:0">
      <table class="ateco-table">
        <thead>
          <tr>
            <th>Descrizione ATECO</th>
            <th>Settore</th>
          </tr>
        </thead>
        <tbody>${righe}</tbody>
      </table>
    </div>`;

  function chiudiPopup() {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  }

  popup.querySelector('.popup-close').addEventListener('click', chiudiPopup);
  document.body.appendChild(overlay);
  document.body.appendChild(popup);
}

function apriDettaglioAzienda(accountName) {
  document.getElementById('dettaglio-overlay')?.remove();

  const opportunita = window.clientiDataRaw.filter(r =>
    String(r['Account Name'] || '').trim() === accountName
  );

  const infoAzienda = window.clientiData.find(r =>
    String(r['Account Name'] || '').trim() === accountName
  ) || window.clientiDataRaw.find(r =>
    String(r['Account Name'] || '').trim() === accountName
  ) || {};
  const info = infoAzienda;

  const buList = [...new Set(
    window.clientiData
      .filter(r => String(r['Account Name'] || '').trim() === accountName)
      .map(r => String(r['Business Unit'] || '').trim())
      .filter(Boolean)
  )].join(', ') || '—';

  const fmt = v => new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(parseFloat(v) || 0);

  const anagrafica = `
    <div class="dettaglio-grid">
      <span class="dettaglio-label">Partita IVA</span>
      <span class="dettaglio-val">${info['Registered Number'] || '—'}</span>
      <span class="dettaglio-label">Settore</span>
      <span class="dettaglio-val">${info['Mapping'] || '—'}</span>
      <span class="dettaglio-label">Dipendenti</span>
      <span class="dettaglio-val">${info['Employees'] || '—'}</span>
      <span class="dettaglio-label">Regione</span>
      <span class="dettaglio-val">${info['Billing State/Province (text only)'] || '—'}</span>
      <span class="dettaglio-label">Provincia</span>
      <span class="dettaglio-val">${info['Province'] || '—'}</span>
      <span class="dettaglio-label">Business Unit</span>
      <span class="dettaglio-val">${buList}</span>
      <span class="dettaglio-label">Fatturato</span>
      <span class="dettaglio-val">${fmt(info['Turnover'])}</span>
      <span class="dettaglio-label">Beneficio RDTC</span>
      <span class="dettaglio-val dettaglio-accent">${fmt(info['RDTC Totale'] || info['RDTC Amount'])}</span>
    </div>`;

  const righe = opportunita.map(r => `
    <div class="dettaglio-row">
      <span class="dettaglio-opp-nome">${r['Opportunity Name'] || '—'}</span>
      <span class="dettaglio-opp-owner">${r['Opportunity Owner: Full Name'] || '—'}</span>
    </div>`).join('');

  const overlay = document.createElement('div');
  overlay.id = 'dettaglio-overlay';

  const popup = document.createElement('div');
  popup.id = 'dettaglio-popup';
  popup.innerHTML = `
    <div class="dettaglio-header">
      <div class="dettaglio-nome">${accountName.toUpperCase()}</div>
      <button class="popup-close dettaglio-close-btn">✕</button>
    </div>
    <div class="dettaglio-body">
      ${anagrafica}
      <hr class="dettaglio-sep">
      <div class="dettaglio-opp-title">Opportunità (${opportunita.length})</div>
      <div class="dettaglio-opp-list">
        ${opportunita.length === 0
          ? '<div class="dettaglio-empty">Nessuna opportunità trovata</div>'
          : righe}
      </div>
    </div>`;

  const chiudiDettaglio = () => document.getElementById('dettaglio-overlay')?.remove();
  overlay.addEventListener('click', chiudiDettaglio);
  popup.addEventListener('click', e => e.stopPropagation());
  popup.querySelector('.dettaglio-close-btn').addEventListener('click', chiudiDettaglio);

  overlay.appendChild(popup);
  document.body.appendChild(overlay);
  const oppList = popup.querySelector('.dettaglio-opp-list');
  if (oppList) oppList.scrollTop = 0;
}
