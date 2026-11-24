/*
 * app.js — page wiring for Canvas Course Exporter
 */
(function () {
  'use strict';
  const CQE = window.CQE;
  const $ = (id) => document.getElementById(id);

  const state = {
    pkg: null,
    items: {}, // catId -> items[]
    checked: new Set(), // checked category ids
    selected: {}, // catId -> Set(itemId)
    filter: ''
  };

  /* ---------------- Step 1: upload ---------------- */
  const dz = $('dropzone');
  const input = $('file-input');
  input.addEventListener('change', () => input.files[0] && loadFile(input.files[0]));
  ['dragenter', 'dragover'].forEach((ev) =>
    dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.add('over');
    })
  );
  ['dragleave', 'drop'].forEach((ev) =>
    dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.remove('over');
    })
  );
  dz.addEventListener('drop', (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) loadFile(f);
  });

  function setStatus(el, msg, kind) {
    el.className = 'status' + (kind ? ' ' + kind : '');
    el.textContent = msg || '';
  }

  async function loadFile(file) {
    const st = $('load-status');
    ['step-type', 'step-select', 'step-options'].forEach((id) => $(id).classList.add('hidden'));
    setStatus($('gen-status'), '');
    if (!/\.(imscc|zip)$/i.test(file.name)) {
      setStatus(st, 'Please choose a Canvas course export file ending in .imscc.', 'error');
      return;
    }
    const big = file.size > 500 * 1048576;
    setStatus(st, 'Reading ' + file.name + '…' + (big ? ' This is a large export (' + formatBytes(file.size) + '), so it may take a minute.' : ''), 'busy');
    try {
      const pkg = await CQE.parsePackage(file, (i, n) => setStatus(st, 'Reading quiz files… ' + i + ' of ' + n, 'busy'));
      state.pkg = pkg;
      state.items = {};
      state.selected = {};
      state.checked = new Set();
      for (const cat of CQE.CATEGORIES) {
        const items = cat.items(pkg);
        state.items[cat.id] = items;
        state.selected[cat.id] = new Set(items.filter((i) => !i.empty).map((i) => i.id));
        if (cat.defaultOn && items.length) state.checked.add(cat.id);
      }
      setStatus(st, '✓ Loaded ' + file.name + ' (' + formatBytes(file.size) + ')', 'ok');
      if (pkg.warnings.length) console.warn('Warnings while reading export:', pkg.warnings);
      renderCategories();
      refresh();
      ['step-type', 'step-select', 'step-options'].forEach((id) => $(id).classList.remove('hidden'));
      $('step-type').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      console.error(err);
      setStatus(st, err.message && /imsmanifest/.test(err.message) ? err.message : "Couldn't read this file. Make sure it is a Canvas course export (.imscc). (" + err.message + ')', 'error');
    }
  }

  /* ---------------- Step 2: what to extract ---------------- */
  function renderCategories() {
    $('course-summary').innerHTML = '<strong>' + esc(state.pkg.courseTitle) + '</strong>';
    const host = $('type-grid');
    host.innerHTML = '';
    for (const sec of CQE.SECTIONS) {
      const cats = CQE.CATEGORIES.filter((c) => c.section === sec.id);
      const box = document.createElement('fieldset');
      box.className = 'cat-section';
      box.innerHTML = '<legend><span class="cs-title">' + esc(sec.title) + '</span> <span class="muted">→ ' + esc(sec.sub) + '</span></legend>';
      for (const cat of cats) {
        const n = state.items[cat.id].length;
        const label = document.createElement('label');
        label.className = 'cat-row' + (n ? '' : ' disabled');
        label.innerHTML =
          '<input type="checkbox" data-cat="' + cat.id + '"' + (state.checked.has(cat.id) ? ' checked' : '') + (n ? '' : ' disabled') + '>' +
          '<span class="cat-label">' + esc(cat.label) + '</span><span class="badge' + (n ? '' : ' zero') + '">' + n + '</span>';
        label.querySelector('input').addEventListener('change', (e) => {
          if (e.target.checked) state.checked.add(cat.id);
          else state.checked.delete(cat.id);
          refresh();
        });
        box.appendChild(label);
      }
      host.appendChild(box);
    }
  }

  function refresh() {
    renderList();
    const sections = new Set(CQE.CATEGORIES.filter((c) => state.checked.has(c.id)).map((c) => c.section));
    document.querySelectorAll('#step-options fieldset[data-for]').forEach((fs) => fs.classList.toggle('hidden', !sections.has(fs.dataset.for)));
    $('step-select').classList.toggle('hidden', state.checked.size === 0);
    updateCounts();
  }

  /* ---------------- Step 3: choose items ---------------- */
  function renderList() {
    const host = $('collection-list');
    const openState = {};
    host.querySelectorAll('details.group').forEach((d) => (openState[d.dataset.cat] = d.open));
    host.innerHTML = '';
    const f = state.filter.toLowerCase();

    for (const cat of CQE.CATEGORIES) {
      if (!state.checked.has(cat.id)) continue;
      const all = state.items[cat.id];
      const sel = state.selected[cat.id];
      const items = all.filter((c) => !f || c.title.toLowerCase().includes(f));
      const titleCount = {};
      all.forEach((c) => (titleCount[c.title] = (titleCount[c.title] || 0) + 1));

      const box = document.createElement('details');
      box.className = 'group';
      box.dataset.cat = cat.id;
      box.open = cat.id in openState ? openState[cat.id] : all.length <= 60 || !!f;
      box.innerHTML =
        '<summary><span class="g-title">' + esc(cat.label) + '</span> <span class="badge">' + all.length + '</span>' +
        '<span class="g-sel muted"></span>' +
        '<span class="g-actions"><button type="button" class="link-btn" data-act="all">All</button> · <button type="button" class="link-btn" data-act="none">None</button></span></summary>';
      const ul = document.createElement('ul');
      ul.className = 'clist';
      for (const c of items) {
        const li = document.createElement('li');
        const id = 'c_' + cat.id + '_' + c.id;
        const dup = titleCount[c.title] > 1;
        li.className = c.empty ? 'empty' : '';
        li.innerHTML =
          '<input type="checkbox" id="' + esc(id) + '"' + (sel.has(c.id) ? ' checked' : '') + '>' +
          '<label for="' + esc(id) + '"><span class="c-title">' + esc(c.title) + '</span>' +
          (dup ? ' <span class="muted small" title="Another item has the same name">(id …' + esc(String(c.id).slice(-6)) + ')</span>' : '') +
          (c.meta ? '<span class="c-meta">' + esc(c.meta) + '</span>' : '') + '</label>';
        li.querySelector('input').addEventListener('change', (e) => {
          if (e.target.checked) sel.add(c.id);
          else sel.delete(c.id);
          updateCounts();
        });
        ul.appendChild(li);
      }
      if (!items.length) {
        const li = document.createElement('li');
        li.className = 'muted none';
        li.textContent = 'No matches';
        ul.appendChild(li);
      }
      box.appendChild(ul);
      box.querySelectorAll('[data-act]').forEach((b) =>
        b.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          const on = b.dataset.act === 'all';
          items.forEach((c) => (on ? sel.add(c.id) : sel.delete(c.id)));
          renderList();
          updateCounts();
        })
      );
      host.appendChild(box);
    }
  }

  function selection() {
    const out = {};
    for (const cat of CQE.CATEGORIES) {
      if (!state.checked.has(cat.id)) continue;
      const sel = state.selected[cat.id];
      const items = state.items[cat.id].filter((i) => sel.has(i.id));
      if (items.length) out[cat.id] = items;
    }
    return out;
  }

  function updateCounts() {
    if (!state.pkg) return;
    const sel = selection();
    let total = 0;
    document.querySelectorAll('#collection-list details.group').forEach((box) => {
      const n = (sel[box.dataset.cat] || []).length;
      box.querySelector('.g-sel').textContent = n + ' selected';
    });
    Object.values(sel).forEach((a) => (total += a.length));
    $('selected-count').textContent = total + ' selected';
    $('generate').disabled = total === 0;

    // Estimate the output
    const version = document.querySelector('input[name="version"]:checked').value;
    const combined = document.querySelector('input[name="output"]:checked').value === 'combined';
    const versions = version === 'both' ? 2 : 1;
    let assess = 0;
    let files = 0;
    let docs = 0;
    for (const [cid, items] of Object.entries(sel)) {
      const sec = CQE.CATEGORIES.find((c) => c.id === cid).section;
      if (sec === 'assess') assess += items.length;
      else if (sec === 'files') files += items.length;
      else docs += items.length;
    }
    const quizDocs = assess ? (combined ? 1 : assess) * versions : 0;
    const parts = [];
    if (quizDocs) parts.push(quizDocs + ' quiz/bank document' + (quizDocs === 1 ? '' : 's'));
    if (files) parts.push(files + ' course file' + (files === 1 ? '' : 's'));
    if (docs) parts.push(docs + ' content document' + (docs === 1 ? '' : 's'));
    const n = quizDocs + files + docs;
    $('output-summary').textContent = n ? 'You’ll get ' + parts.join(', ') + (n > 1 ? ' in one .zip file.' : '.') : 'Nothing selected yet.';
  }
  document.querySelectorAll('input[name="version"], input[name="output"]').forEach((r) => r.addEventListener('change', updateCounts));

  $('filter').addEventListener('input', (e) => {
    state.filter = e.target.value;
    renderList();
    updateCounts();
  });
  $('select-all').addEventListener('click', () => {
    const f = state.filter.toLowerCase();
    for (const cid of state.checked) state.items[cid].filter((c) => !f || c.title.toLowerCase().includes(f)).forEach((c) => state.selected[cid].add(c.id));
    renderList();
    updateCounts();
  });
  $('select-none').addEventListener('click', () => {
    for (const cid of state.checked) state.selected[cid].clear();
    renderList();
    updateCounts();
  });

  /* ---------------- Step 4: generate ---------------- */
  $('generate').addEventListener('click', async () => {
    const btn = $('generate');
    const st = $('gen-status');
    const bar = $('progress-bar');
    const prog = $('progress');
    const sel = selection();
    if (!Object.keys(sel).length) return;

    const opts = {
      version: document.querySelector('input[name="version"]:checked').value,
      quizOutput: document.querySelector('input[name="output"]:checked').value,
      includeFeedback: $('opt-feedback').checked,
      includePoints: $('opt-points').checked,
      expandBanks: $('opt-expand').checked,
      answerTable: $('opt-table').checked,
      keepFolders: document.querySelector('input[name="folders"]:checked').value === 'keep',
      includeSettings: $('opt-settings').checked
    };

    btn.disabled = true;
    prog.classList.remove('hidden');
    bar.style.width = '0%';
    setStatus(st, 'Preparing…', 'busy');
    try {
      const result = await CQE.generateExport(state.pkg, sel, opts, (i, n, label) => {
        const pct = Math.round((i / n) * 100);
        bar.style.width = pct + '%';
        prog.setAttribute('aria-valuenow', pct);
        setStatus(st, label === 'Creating zip file…' ? label : 'Working… ' + i + ' of ' + n, 'busy');
      });
      download(result.blob, result.filename);
      setStatus(st, '✓ Done: ' + result.count + ' file' + (result.count === 1 ? '' : 's') + (result.filename.endsWith('.zip') ? ' in ' : ' ') + '“' + result.filename + '” (' + formatBytes(result.blob.size) + ').', 'ok');
    } catch (err) {
      console.error(err);
      setStatus(st, 'Something went wrong while creating the download: ' + err.message, 'error');
    } finally {
      btn.disabled = false;
      setTimeout(() => prog.classList.add('hidden'), 800);
    }
  });

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function formatBytes(n) {
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(0) + ' KB';
    if (n < 1073741824) return (n / 1048576).toFixed(1) + ' MB';
    return (n / 1073741824).toFixed(2) + ' GB';
  }
})();
