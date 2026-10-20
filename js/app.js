const PICTURE_NOTE = 'Each page is saved as a picture, so text in the PDF can’t be searched or copied.';

const TOOLS = [
  {
    id: 'pdf-to-word', kind: 'a PDF', group: 'from', from: 'PDF', to: 'Word', accept: ['.pdf'], out: '.docx',
    note: 'Turns the PDF’s text into editable paragraphs and headings. Pictures inside the PDF aren’t carried over, and scanned pages come through as pictures.',
    options: [{ key: 'mode', label: 'Output', choices: [['editable', 'Editable text'], ['exact', 'Exact look (pages as pictures)']] }],
  },
  {
    id: 'pdf-to-jpg', kind: 'a PDF', group: 'from', from: 'PDF', to: 'JPG', accept: ['.pdf'], out: '.jpg',
    note: 'Saves every page as a JPG image. When there’s more than one page, the images download together in a ZIP.',
    options: [{ key: 'dpi', label: 'Resolution', choices: [['150', 'Standard (150 dpi)'], ['300', 'High (300 dpi)']] }],
  },
  {
    id: 'pdf-to-excel', kind: 'a PDF', group: 'from', from: 'PDF', to: 'Excel', accept: ['.pdf'], out: '.xlsx',
    note: 'Finds the rows and columns in the PDF’s text and puts them into cells. Works best on simple tables; merged cells may need tidying afterwards.',
    options: [{ key: 'layout', label: 'Sheets', choices: [['pages', 'One sheet per page'], ['single', 'All pages in one sheet']] }],
  },
  {
    id: 'pdf-to-epub', kind: 'a PDF', group: 'from', from: 'PDF', to: 'EPUB', accept: ['.pdf'], out: '.epub',
    note: 'Reflows the PDF’s text into an e-book with chapters, for reading on phones and e-readers. Works best on text-heavy PDFs.',
    options: [],
  },
  {
    id: 'word-to-pdf', kind: 'a Word file (.docx)', group: 'to', from: 'Word', to: 'PDF', accept: ['.docx', '.docm', '.dotx'], out: '.pdf',
    note: `Lays out your .docx page by page. Fonts you don’t have installed are swapped for similar ones, so spacing can shift slightly. ${PICTURE_NOTE}`,
    options: [{ key: 'quality', label: 'Quality', choices: [['sharp', 'Sharp'], ['small', 'Smaller file']] }],
  },
  {
    id: 'pptx-to-pdf', kind: 'a PowerPoint file (.pptx)', group: 'to', from: 'PPTX', to: 'PDF', accept: ['.pptx', '.ppsx', '.potx'], out: '.pdf',
    note: `One page per slide. Charts aren’t drawn yet, and fonts you don’t have installed are swapped for similar ones. ${PICTURE_NOTE}`,
    options: [
      { key: 'quality', label: 'Quality', choices: [['sharp', 'Sharp'], ['small', 'Smaller file']] },
      { key: 'hidden', label: 'Hidden slides', choices: [['skip', 'Leave out'], ['include', 'Include']] },
    ],
  },
  {
    id: 'epub-to-pdf', kind: 'an EPUB e-book', group: 'to', from: 'EPUB', to: 'PDF', accept: ['.epub'], out: '.pdf',
    note: `Turns an e-book into a printable PDF with page numbers. DRM-protected books can’t be converted. ${PICTURE_NOTE}`,
    options: [
      { key: 'pageSize', label: 'Page size', choices: [['a5', 'Book (A5)'], ['a4', 'A4'], ['letter', 'Letter']] },
      { key: 'quality', label: 'Quality', choices: [['sharp', 'Sharp'], ['small', 'Smaller file']] },
    ],
  },
  {
    id: 'image-convert', kind: 'an image (JPG, PNG, WebP, GIF, BMP, AVIF, SVG or HEIC)', group: 'images', from: 'Image', to: 'Image',
    title: 'Convert images', dropTitle: 'Choose images',
    accept: ['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.avif', '.svg', '.heic', '.heif'], out: '.jpg',
    note: 'Converts between JPG, PNG, WebP, BMP and AVIF. Opens GIF, SVG and iPhone HEIC photos too. Transparent areas turn white in JPG and BMP, and animated GIFs keep only their first frame. AVIF files take longer to save.',
    options: [
      { key: 'format', label: 'Convert to', default: 'jpg', choices: [['jpg', 'JPG'], ['png', 'PNG'], ['webp', 'WebP'], ['avif', 'AVIF'], ['bmp', 'BMP']] },
      { key: 'quality', label: 'Quality (JPG, WebP, AVIF)', choices: [['high', 'High'], ['balanced', 'Balanced'], ['small', 'Smaller file']] },
      { key: 'size', label: 'Longest side', choices: [['0', 'Original'], ['3840', '3840 px'], ['1920', '1920 px'], ['1280', '1280 px']] },
    ],
  },
  {
    id: 'audio-convert', kind: 'an audio file (WAV, MP3, OGG, M4A or FLAC)', group: 'audio', from: 'Audio', to: 'MP3 / WAV',
    title: 'Convert audio', dropTitle: 'Choose audio files',
    accept: ['.wav', '.mp3', '.ogg', '.oga', '.opus', '.m4a', '.aac', '.flac'], out: '.mp3',
    note: 'Turns WAV, OGG, M4A, AAC and FLAC files into MP3, or any of them (including MP3) into WAV. Audio is saved at 44.1 kHz. Which formats can be opened depends on your browser; Chrome and Edge handle the most.',
    options: [
      { key: 'format', label: 'Convert to', choices: [['mp3', 'MP3'], ['wav', 'WAV (lossless)']] },
      { key: 'bitrate', label: 'MP3 bitrate', default: '192', choices: [['128', '128 kbps'], ['192', '192 kbps'], ['256', '256 kbps'], ['320', '320 kbps (best)']] },
    ],
  },
];

const $ = (s) => document.querySelector(s);
const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; };

let current = TOOLS[0];
const optionValues = {};
const queue = [];
let running = false;

const toolName = (t) => t.title || `${t.from} to ${t.to}`;

function fmtTag(text) { return el('span', 'fmt', text); }

function buildToolList() {
  for (const t of TOOLS) {
    const b = el('button', 'tool');
    b.type = 'button';
    b.dataset.id = t.id;
    b.setAttribute('aria-pressed', 'false');
    b.setAttribute('aria-label', toolName(t));
    b.append(fmtTag(t.from), el('span', 'arrow', '→'), fmtTag(t.to));
    b.addEventListener('click', () => select(t.id, true));
    document.querySelector(`.tool-list[data-group="${t.group}"]`).appendChild(b);
  }
}

function buildOptions(t) {
  const box = $('#options');
  box.textContent = '';
  optionValues[t.id] = optionValues[t.id] || {};
  for (const o of t.options) {
    const fs = el('fieldset', 'opt');
    fs.appendChild(el('legend', '', o.label));
    const seg = el('div', 'seg');
    const chosen = optionValues[t.id][o.key] || o.default || o.choices[0][0];
    optionValues[t.id][o.key] = chosen;
    for (const [value, label] of o.choices) {
      const l = el('label');
      const input = el('input');
      input.type = 'radio';
      input.name = `${t.id}-${o.key}`;
      input.value = value;
      input.checked = value === chosen;
      input.addEventListener('change', () => { optionValues[t.id][o.key] = value; });
      l.append(input, el('span', '', label));
      seg.appendChild(l);
    }
    fs.appendChild(seg);
    box.appendChild(fs);
  }
}

function select(id, fromUser) {
  const t = TOOLS.find((x) => x.id === id) || TOOLS[0];
  current = t;
  for (const b of document.querySelectorAll('.tool')) b.setAttribute('aria-pressed', String(b.dataset.id === t.id));
  $('#tool-title').textContent = toolName(t);
  $('#tool-note').textContent = t.note;
  $('#drop-title').textContent = t.dropTitle || (t.from === 'PDF' ? 'Choose PDF files' : `Choose ${t.accept[0]} files`);
  $('#drop-hint').textContent = 'or drop them here';
  const input = $('#file-input');
  input.accept = t.accept.join(',');
  buildOptions(t);
  document.title = `${toolName(t)} — Kopya`;
  if (fromUser && location.hash !== `#${t.id}`) history.replaceState(null, '', `#${t.id}`);
}

const ext = (name) => (name.match(/\.[^.]+$/) || [''])[0].toLowerCase();

function toolFor(file) {
  const e = ext(file.name);
  if (e === '.pdf') return null;
  return TOOLS.find((t) => t.accept.includes(e));
}

function sizeText(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function addFiles(files) {
  for (const file of files) {
    const tool = current;
    const job = { file, tool, options: { ...optionValues[tool.id] } };
    const li = el('li', 'job');
    const name = el('p', 'job-name', file.name);
    const status = el('p', 'job-status', 'Waiting');
    const bar = el('div', 'bar');
    bar.setAttribute('role', 'progressbar');
    bar.setAttribute('aria-label', `Converting ${file.name}`);
    bar.setAttribute('aria-valuemin', '0');
    bar.setAttribute('aria-valuemax', '100');
    const fill = el('span');
    bar.appendChild(fill);
    li.append(name, el('span'), bar, status);
    job.ui = { li, status, bar, fill, slot: li.children[1] };
    $('#queue').prepend(li);

    if (!tool.accept.includes(ext(file.name))) {
      const e = ext(file.name);
      const other = toolFor(file);
      let msg = `This tool needs ${tool.kind}.`;
      if (e === '.doc') msg = 'This is the older .doc format. Open it in Word or Google Docs, save it as .docx, then try again.';
      else if (e === '.ppt') msg = 'This is the older .ppt format. Open it in PowerPoint or Google Slides, save it as .pptx, then try again.';
      else if (e === '.pdf') msg += ' For PDFs, pick one of the “From PDF” conversions.';
      else if (other) msg += ` For this file, use ${toolName(other)}.`;
      fail(job, msg, other);
      continue;
    }
    queue.push(job);
  }
  run();
}

function fail(job, message, suggest) {
  const { li, status, bar, slot } = job.ui;
  li.classList.add('is-failed');
  bar.remove();
  status.textContent = message;
  if (suggest) {
    const b = el('button', 'btn btn-quiet', `Switch to ${toolName(suggest)}`);
    b.type = 'button';
    b.addEventListener('click', () => { select(suggest.id, true); li.remove(); addFiles([job.file]); });
    slot.replaceWith(b);
    job.ui.slot = b;
  }
}

async function run() {
  if (running) return;
  running = true;
  while (queue.length) {
    const job = queue.shift();
    const { li, status, bar, fill, slot } = job.ui;
    const report = (fraction, label) => {
      const pct = Math.round(Math.min(1, Math.max(0, fraction)) * 100);
      fill.style.width = `${pct}%`;
      bar.setAttribute('aria-valuenow', String(pct));
      if (label) status.textContent = label;
    };
    report(0.01, 'Starting');
    try {
      const mod = await import(`./converters/${job.tool.id}.js`);
      const result = await mod.convert(job.file, { report, options: job.options });
      const url = URL.createObjectURL(result.blob);
      const a = el('a', 'btn', 'Download');
      a.href = url;
      a.download = result.name;
      a.setAttribute('aria-label', `Download ${result.name}`);
      slot.replaceWith(a);
      bar.remove();
      li.classList.add('is-done');
      status.textContent = `${result.name} · ${sizeText(result.blob.size)}${result.detail ? ` · ${result.detail}` : ''}`;
      for (const w of result.warnings || []) li.appendChild(el('p', 'job-warn', w));
    } catch (e) {
      console.error(e);
      const msg = e && e.message && !/^(bad xml|missing)/.test(e.message) ? e.message : 'Something went wrong while converting this file. It may be damaged or use features Kopya can’t read yet.';
      fail(job, msg);
    }
  }
  running = false;
}

function initDrop() {
  const drop = $('#drop');
  const input = $('#file-input');
  input.addEventListener('change', () => { if (input.files.length) addFiles(Array.from(input.files)); input.value = ''; });
  let depth = 0;
  drop.addEventListener('dragenter', (e) => { e.preventDefault(); depth++; drop.classList.add('is-over'); });
  drop.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) drop.classList.remove('is-over'); });
  drop.addEventListener('dragover', (e) => e.preventDefault());
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    depth = 0;
    drop.classList.remove('is-over');
    const files = Array.from(e.dataTransfer.files || []);
    if (files.length) addFiles(files);
  });
  // A file dropped anywhere else on the page shouldn't open in the tab.
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
}

buildToolList();
initDrop();
select(location.hash.slice(1));
window.addEventListener('hashchange', () => select(location.hash.slice(1)));

// For automated tests: window.kopya.convert(toolId, file, options)
window.kopya = {
  async convert(id, file, options = {}) {
    const mod = await import(`./converters/${id}.js`);
    return mod.convert(file, { report: () => {}, options });
  },
};
