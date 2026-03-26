window.clientiData = [];

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
    uploadStatus.textContent = '✓ Caricati ' + window.clientiData.length + ' clienti da ' + file.name;
    btnGenera.style.display = 'block';
  };
  reader.readAsArrayBuffer(file);
}

// ── RENDER ALL ──────────────────────────────────
function renderAll() {
  document.getElementById('upload-section').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  renderDimensioni();
  renderSettori();
  renderDistribuzione();
}

// ── DIMENSIONI ──────────────────────────────────
function renderDimensioni() {
  const dati = window.clientiData;
  const total = dati.length;
  let grandi = 0, medie = 0, piccole = 0, micro = 0;
  dati.forEach(r => {
    const emp = parseFloat(r['Employees']);
    if (isNaN(emp)) return;
    if (emp > 250)       grandi++;
    else if (emp >= 50)  medie++;
    else if (emp >= 10)  piccole++;
    else                 micro++;
  });

  const segmenti = [
    { val: grandi,  color: '#002c49', label: 'grandi imprese',  sub: '(>250 dipendenti)'   },
    { val: medie,   color: '#526d81', label: 'medie imprese',   sub: '(50-250 dipendenti)' },
    { val: piccole, color: '#9cabb6', label: 'piccole imprese', sub: '(10-49 dipendenti)'  },
    { val: micro,   color: '#cdd5dc', label: 'micro imprese',   sub: '(<10 dipendenti)'    },
  ];

  const circ = 2 * Math.PI * 60;
  const gap = 4;
  let offset = circ / 4;
  let archi = '';
  segmenti.forEach(s => {
    const dash = (s.val / total) * (circ - gap * 4);
    archi += `<circle cx="80" cy="80" r="60" fill="none"
      stroke="${s.color}" stroke-width="18"
      stroke-dasharray="${dash} ${circ - dash}"
      stroke-dashoffset="${offset}"
      transform="rotate(-90 80 80)"/>`;
    offset -= (dash + gap);
  });

  const lista = segmenti.map(s => `
    <div class="dim-row">
      <span class="dim-num">${s.val}</span>
      <div class="dim-desc">
        <span class="dim-label">${s.label}</span>
        <span class="dim-sub">${s.sub}</span>
      </div>
    </div>`).join('');

  document.getElementById('dimensioni').innerHTML = `
    <div class="dim-inner">
      <div class="dim-donut-wrap">
        <img src="assets/Donut.png" alt="Donut chart" style="width:160px; height:160px; object-fit:contain; display:block;">
        <div class="donut-total">${total}<span class="donut-label">aziende</span></div>
      </div>
      <div class="dim-list">${lista}</div>
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
  window.clientiData.forEach(r => {
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

  window.clientiData.forEach(r => {
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
    <div class="distrib-body">
      <div class="distrib-labels">
        <div class="distrib-item">
          <div class="distrib-num">${nord}</div>
          <div class="distrib-word">aziende</div>
          <div class="distrib-place">Nord Italia</div>
        </div>
        <div class="distrib-item">
          <div class="distrib-num">${centro}</div>
          <div class="distrib-word">aziende</div>
          <div class="distrib-place">Centro Italia</div>
        </div>
        <div class="distrib-item">
          <div class="distrib-num">${sud}</div>
          <div class="distrib-word">aziende</div>
          <div class="distrib-place">Sud Italia e Isole</div>
        </div>
        <div class="distrib-item">
          <div class="distrib-num">${intl}</div>
          <div class="distrib-word">aziende</div>
          <div class="distrib-place">Internazionali</div>
        </div>
      </div>
      <div class="distrib-map" id="mappa-container">
        <img src="assets/italy-map.png" alt="Mappa Italia" style="display:block; width:100%; max-width:260px;">
      </div>
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
        d3.select(this).attr('fill', '#7ebd4b');
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
      d3.select(this).attr('fill', '#7ebd4b');
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
  const aziendeRegione = window.clientiData.filter(c => 
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
      <div class="popup-titolo">${nomeRegione.toUpperCase()}</div>
      <div class="popup-count">${aziendeRegione.length} <span>aziende</span></div>
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
