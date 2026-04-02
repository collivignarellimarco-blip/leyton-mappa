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

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const sheetName = wb.SheetNames.includes('Sheet1') ? 'Sheet1' : wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
    window.clientiData = rows.filter(r => {
      const name = String(r['Account Name'] || '');
      return !name.includes('SEGNALATORE') && !name.includes('PARTNERSHIP');
    });
    const nEnergy = window.clientiData.filter(r => String(r['Business Unit'] || '').trim() === 'Energy').length;
    const nBP = window.clientiData.filter(r => String(r['Business Unit'] || '').trim() === 'Business Performance').length;
    document.getElementById('bu-counts').innerHTML = `
      <div class="bu-count-row">
        <span class="bu-count-energy">${nEnergy} aziende Energy</span>
        <span class="bu-count-bp">${nBP} aziende Business Performance</span>
      </div>`;
    uploadStatus.textContent = '';
    btnGenera.style.display = 'none';
    document.getElementById('bu-selection').style.display = 'flex';
  };
  reader.readAsArrayBuffer(file);
}

function selezionaBU(bu) {
  window.businessUnit = bu;
  // Filtra i dati per Business Unit
  window.clientiFiltrati = window.clientiData.filter(r =>
    String(r['Business Unit'] || '').trim() === bu
  );
  document.getElementById('bu-selection').style.display = 'none';
  btnGenera.style.display = 'block';
  // Aggiorna status
  uploadStatus.innerHTML = '';
  // Imposta colore tema
  const colore = bu === 'Energy' ? '#7ebd4b' : '#FF6633';
  document.documentElement.style.setProperty('--accent', colore);
}

function cambiaBU() {
  window.businessUnit = null;
  window.clientiFiltrati = [];
  // Nascondi app
  document.getElementById('app').style.display = 'none';
  // Mostra upload section ma senza resettare il file
  document.getElementById('upload-section').style.display = 'block';
  // Nascondi genera e mostra selezione BU direttamente
  btnGenera.style.display = 'none';
  document.getElementById('bu-selection').style.display = 'flex';
  // Reset colore tema
  document.documentElement.style.setProperty('--accent', '#7ebd4b');
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
  if (window.businessUnit === 'Energy') {
    logoWrap.innerHTML = `<img src="assets/leyton_esg_3.svg" alt="Leyton ESG" style="height:22px; width:auto; display:block;">`;
  } else {
    logoWrap.innerHTML = `<img src="assets/logo_leyton.svg" alt="Leyton" style="height:22px; width:auto; display:block;">`;
  }
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
    { val: grandi,  color: '#002c49' },
    { val: medie,   color: '#526d81' },
    { val: piccole, color: '#9cabb6' },
    { val: micro,   color: '#cdd5dc' },
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

  const ICONE_SETTORE = {
    'ALIMENTARE':      'elenco/elenco_alimentare.svg',
    'AUTOMOTIVE':      'elenco/elenco_automotive.svg',
    'CHIMICO':         'elenco/elenco_chimico.svg',
    'COMMERCIO':       'elenco/elenco_commercio.svg',
    'CONSULENZA':      'elenco/elenco_consulenza.svg',
    'COSTRUZIONI':     'elenco/elenco_costruzioni.svg',
    'FARMACEUTICO':    'elenco/elenco_pharma.svg',
    'HORECA':          'elenco/elenco_horeca.svg',
    'IT':              'elenco/elenco_it.svg',
    'MANIFATTURIERO':  'elenco/elenco_manifatturiero.svg',
    'MODA':            'elenco/elenco_moda.svg',
    'RICERCA':         'elenco/elenco_ricerca.svg',
    'SALUTE':          'elenco/elenco_salute.svg',
    'SERVIZI':         'elenco/elenco_servizi.svg',
    'TELECOMUNICAZIONI':'elenco/elenco_telecomunicazioni.svg',
    'TRASPORTI':       'elenco/elenco_trasporti.svg',
  };

  let htmlBody = '';
  provinceOrdinate.forEach(prov => {
    htmlBody += `<div class="popup-provincia">
      <div class="popup-provincia-nome">${prov.toUpperCase()}</div>`;
    
    const aziende = byProvincia[prov].sort((a,b) => a.nome.localeCompare(b.nome));
    aziende.forEach(a => {
      const iconaPath = ICONE_SETTORE[a.settore.toUpperCase()];
      const iconaHtml = iconaPath ? `<img src="assets/icons/${iconaPath}" class="popup-azienda-icon">` : `<div></div>`;
      htmlBody += `<div class="popup-azienda">
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
      <button class="popup-close">✕</button>
    </div>
    <div class="popup-body">
      ${htmlBody}
    </div>
  `;

  popup.querySelector('.popup-close').addEventListener('click', chiudiPopup);

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
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
        <div class="popup-azienda"><span class="popup-slash">/</span> ${a.nome}</div>`;
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

  // Crea modale
  const popup = document.createElement('div');
  popup.id = 'regione-popup';
  popup.innerHTML = `
    <div class="popup-header">
      <div class="popup-header-left">
        <div class="popup-titolo">${label.toUpperCase()}</div>
        <div class="popup-count">${aziendeSettore.length} <span>aziende</span></div>
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
    'ALIMENTARE': 'elenco/elenco_alimentare.svg',
    'AUTOMOTIVE': 'elenco/elenco_automotive.svg',
    'CHIMICO': 'elenco/elenco_chimico.svg',
    'COMMERCIO': 'elenco/elenco_commercio.svg',
    'CONSULENZA': 'elenco/elenco_consulenza.svg',
    'COSTRUZIONI': 'elenco/elenco_costruzioni.svg',
    'FARMACEUTICO': 'elenco/elenco_pharma.svg',
    'HORECA': 'elenco/elenco_horeca.svg',
    'IT': 'elenco/elenco_it.svg',
    'MANIFATTURIERO': 'elenco/elenco_manifatturiero.svg',
    'MODA': 'elenco/elenco_moda.svg',
    'RICERCA': 'elenco/elenco_ricerca.svg',
    'SALUTE': 'elenco/elenco_salute.svg',
    'SERVIZI': 'elenco/elenco_servizi.svg',
    'TELECOMUNICAZIONI': 'elenco/elenco_telecomunicazioni.svg',
    'TRASPORTI': 'elenco/elenco_trasporti.svg',
  };

  const listaHtml = aziende.map(r => {
    const settore = String(r['Mapping'] || '').trim().toUpperCase();
    const icona = ICONE_SETTORE[settore];
    const iconaHtml = icona
      ? `<img src="assets/icons/${icona}" class="popup-azienda-icon">`
      : `<span style="width:14px;display:inline-block"></span>`;
    return `<div class="popup-azienda">${iconaHtml} ${String(r['Account Name'] || '').toUpperCase()}</div>`;
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
    popup.innerHTML = `
      <div class="popup-header">
        <div class="popup-header-left">
          <div class="popup-titolo">${String(azienda['Account Name'] || '').toUpperCase()}</div>
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
          <span class="ricerca-value" style="color:${accent}; font-weight:700;">${azienda['Business Unit'] || '—'}</span>
        </div>
        <div class="ricerca-field">
          <span class="ricerca-label">Fatturato</span>
          <span class="ricerca-value">${azienda['Turnover'] || '—'}</span>
        </div>
      </div>`;
  }

  popup.querySelector('.popup-close').addEventListener('click', () => {
    document.getElementById('popup-overlay')?.remove();
    document.getElementById('regione-popup')?.remove();
  });

  document.body.appendChild(overlay);
  document.body.appendChild(popup);
}

function setupRicerca() {
  document.getElementById('search-btn').addEventListener('click', cercaAzienda);
  document.getElementById('search-input').addEventListener('keydown', e => {
    if (e.key === 'Enter') cercaAzienda();
  });
}
