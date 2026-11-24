/*
 * extractors.js
 * What can be extracted (categories shown in step 2), image loading,
 * and building the final download.
 *
 * To add a new kind of content later, add an entry to CQE.CATEGORIES with
 * an items(pkg) function and handle its id in buildJobs() below.
 */
(function (global) {
  'use strict';
  const CQE = (global.CQE = global.CQE || {});

  const plural = (n, w, pl) => n + ' ' + (n === 1 ? w : pl || w + 's');
  const fmtSize = (n) => (n < 1024 ? n + ' B' : n < 1048576 ? Math.round(n / 1024) + ' KB' : (n / 1048576).toFixed(1) + ' MB');

  CQE.SECTIONS = [
    { id: 'assess', title: 'Assessments', sub: 'formatted Word documents' },
    { id: 'files', title: 'Course files', sub: 'original files, unchanged' },
    { id: 'content', title: 'Course content', sub: 'Word documents' }
  ];

  function quizMeta(c) {
    const parts = [c.kind === 'new-quiz' ? 'New Quiz' : 'Classic'];
    if (c.questionCount) parts.push(plural(c.questionCount, 'question'));
    if (c.pullsFromBanks) parts.push('pulls from ' + plural(c.pullsFromBanks, 'bank'));
    if (!c.questionCount && !c.pullsFromBanks) parts.push('no questions in export');
    return parts.join(' · ');
  }
  const collItems = (pkg, pred, meta) =>
    pkg.collections.filter(pred).map((c) => ({ id: c.id, title: c.title, meta: meta(c), empty: !(c.questionCount > 0 || c.pullsFromBanks > 0), ref: c }));
  const isQuiz = (c) => c.kind === 'classic-quiz' || c.kind === 'new-quiz';
  const fileItems = (pkg, cat) =>
    pkg.content.files
      .filter((f) => f.category === cat)
      .map((f) => ({ id: f.id, title: f.rel, meta: f.kindLabel + ' · ' + fmtSize(f.size) + (f.hidden ? ' · used in quizzes / hidden in Canvas' : ''), ref: f }));

  CQE.CATEGORIES = [
    { id: 'classic-bank', section: 'assess', label: 'Question banks (Classic Quizzes)', folder: 'Question banks (Classic)', defaultOn: true,
      items: (pkg) => collItems(pkg, (c) => c.kind === 'classic-bank', (c) => plural(c.questionCount, 'question')) },
    { id: 'item-bank', section: 'assess', label: 'Item banks (New Quizzes)', folder: 'Item banks (New Quizzes)', defaultOn: true,
      items: (pkg) => collItems(pkg, (c) => c.kind === 'item-bank', (c) => plural(c.questionCount, 'question')) },
    { id: 'quiz-groups', section: 'assess', label: 'Quizzes with question groups', folder: 'Quizzes with question groups', defaultOn: true,
      items: (pkg) => collItems(pkg, (c) => isQuiz(c) && c.hasGroups, quizMeta) },
    { id: 'quiz-fixed', section: 'assess', label: 'Quizzes (fixed questions)', folder: 'Quizzes', defaultOn: true,
      items: (pkg) => collItems(pkg, (c) => isQuiz(c) && !c.hasGroups, quizMeta) },

    { id: 'files-docs', section: 'files', label: 'Documents (Word, PDF, PowerPoint, Excel…)', folder: 'Course files',
      items: (pkg) => fileItems(pkg, 'files-docs') },
    { id: 'files-images', section: 'files', label: 'Images', folder: 'Course files',
      items: (pkg) => fileItems(pkg, 'files-images') },
    { id: 'files-other', section: 'files', label: 'Audio, video & other files', folder: 'Course files',
      items: (pkg) => fileItems(pkg, 'files-other') },

    { id: 'pages', section: 'content', label: 'Pages', folder: 'Pages',
      items: (pkg) => pkg.content.pages.map((p) => ({ id: p.id, title: p.title, meta: [p.frontPage ? 'front page' : '', p.state].filter(Boolean).join(' · '), ref: p })) },
    { id: 'assignments', section: 'content', label: 'Assignments', folder: 'Assignments',
      items: (pkg) =>
        pkg.content.assignments.map((a) => ({
          id: a.id,
          title: a.title,
          meta: [a.points != null ? +a.points + ' pts' : '', (a.submissionTypes || '').replace(/_/g, ' '), CQE.htmlToText(a.html) ? '' : 'no description'].filter(Boolean).join(' · '),
          ref: a
        })) },
    { id: 'discussions', section: 'content', label: 'Discussions & announcements', folder: 'Discussions',
      items: (pkg) => pkg.content.discussions.map((d) => ({ id: d.id, title: d.title, meta: d.isAnnouncement ? 'announcement' : d.graded ? 'graded discussion' : 'discussion', ref: d })) },
    { id: 'rubrics', section: 'content', label: 'Rubrics', folder: 'Rubrics',
      items: (pkg) => pkg.content.rubrics.map((r) => ({ id: r.id, title: r.title, meta: plural(r.criteria.length, 'criterion', 'criteria') + (r.points != null ? ' · ' + +r.points + ' pts' : ''), ref: r })) },
    { id: 'outline', section: 'content', label: 'Course outline (modules, grading groups, outcomes)', folder: '',
      items: (pkg) => {
        const o = pkg.content.outline;
        if (!o) return [];
        return [{ id: 'outline', title: 'Course outline', meta: [plural(o.modules.length, 'module'), plural(o.groups.length, 'assignment group'), plural(o.outcomes.length, 'outcome')].join(' · '), ref: o }];
      } }
  ];

  /* ------------------------------------------------------------------ *
   * Image loading from the zip (with caching and format conversion)
   * ------------------------------------------------------------------ */
  CQE.makeImageLoader = function (zip) {
    const cache = new Map();
    const lowerIndex = {};
    zip.forEach((p) => (lowerIndex[p.toLowerCase()] = p));

    async function load(src) {
      let bytes = null;
      if (/^data:image\//i.test(src)) {
        const m = src.match(/^data:([^;,]+)(;base64)?,(.*)$/i);
        if (!m) return null;
        if (m[2]) {
          const bin = atob(m[3]);
          bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        } else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
      } else {
        const path = CQE.resolveZipPath(src);
        if (!path) return null;
        const real = zip.file(path) ? path : lowerIndex[path.toLowerCase()];
        if (!real || !zip.file(real)) return null;
        bytes = await zip.file(real).async('uint8array');
      }
      return toDocxImage(bytes);
    }

    return function getImage(src) {
      if (!src) return Promise.resolve(null);
      if (!cache.has(src)) cache.set(src, load(src).catch(() => null));
      return cache.get(src);
    };
  };

  function sniff(b) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8) return 'jpg';
    if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
    const head = new TextDecoder().decode(b.slice(0, 300)).toLowerCase();
    if (head.includes('<svg') || head.includes('<?xml')) return 'svg';
    if (b.length > 12 && head.startsWith('riff') && head.slice(8, 12) === 'webp') return 'webp';
    return 'unknown';
  }

  const MIME = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };

  function loadImgEl(blob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        resolve({ img, url });
      };
      img.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error('image decode failed'));
      };
      img.src = url;
    });
  }

  async function toDocxImage(bytes) {
    const type = sniff(bytes);
    if (type === 'unknown') return null;
    const blob = new Blob([bytes], { type: MIME[type] || 'application/octet-stream' });
    const { img, url } = await loadImgEl(blob);
    let w = img.naturalWidth || 300;
    let h = img.naturalHeight || 200;
    try {
      if (type === 'png' || type === 'jpg' || type === 'gif' || type === 'bmp') {
        return { data: bytes, type, width: w, height: h };
      }
      // SVG / WebP: rasterise to PNG, which Word handles everywhere
      const scale = type === 'svg' ? 2 : 1;
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(w * scale));
      canvas.height = Math.max(1, Math.round(h * scale));
      const g = canvas.getContext('2d');
      g.fillStyle = '#fff';
      g.fillRect(0, 0, canvas.width, canvas.height);
      g.drawImage(img, 0, 0, canvas.width, canvas.height);
      const png = await new Promise((r) => canvas.toBlob(r, 'image/png'));
      if (!png) return null;
      return { data: new Uint8Array(await png.arrayBuffer()), type: 'png', width: w, height: h };
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  /* ------------------------------------------------------------------ *
   * Building the download
   * ------------------------------------------------------------------ */
  function safeName(s) {
    return (
      String(s || 'untitled')
        .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-')
        .replace(/\s+/g, ' ')
        .replace(/^[\s.-]+|[\s.]+$/g, '')
        .slice(0, 90) || 'untitled'
    );
  }
  CQE.safeName = safeName;
  const short = (s) => safeName(safeName(s).slice(0, 80));

  // Already-compressed formats are stored without recompressing (much faster)
  const STORE_EXT = /\.(png|jpe?g|gif|webp|heic|mp4|m4v|mov|webm|avi|mp3|m4a|aac|ogg|wav|zip|gz|7z|rar|docx|xlsx|pptx|dotx|ppsx|pdf|epub|key|numbers|pages)$/i;

  /**
   * selection: { [categoryId]: item[] }   (items from CQE.CATEGORIES[].items)
   * opts: {
   *   version: 'key'|'student'|'both', quizOutput: 'separate'|'combined',
   *   includeFeedback, includePoints, expandBanks, answerTable,
   *   keepFolders, includeSettings
   * }
   * Returns { blob, filename, count }
   */
  CQE.generateExport = async function (pkg, selection, opts, onProgress) {
    const progress = onProgress || function () {};
    const getImage = CQE.makeImageLoader(pkg.zip);
    const jobs = buildJobs(pkg, selection, opts, getImage);
    const files = [];
    const used = new Set();
    const uniq = (name) => {
      const dot = name.lastIndexOf('.');
      const base = dot > name.lastIndexOf('/') ? name.slice(0, dot) : name;
      const ext = dot > name.lastIndexOf('/') ? name.slice(dot) : '';
      let n = name;
      let i = 2;
      while (used.has(n.toLowerCase())) n = base + ' (' + i++ + ')' + ext;
      used.add(n.toLowerCase());
      return n;
    };

    for (let i = 0; i < jobs.length; i++) {
      const job = jobs[i];
      const data = await job.run();
      files.push({ name: uniq(job.name), data, store: STORE_EXT.test(job.name) });
      progress(i + 1, jobs.length, job.label);
      if (i % 5 === 4) await new Promise((r) => setTimeout(r, 0)); // keep the page responsive
    }

    if (!files.length) throw new Error('Nothing was selected.');
    if (files.length === 1) {
      const f = files[0];
      const blob = f.data instanceof Blob ? f.data : new Blob([f.data]);
      return { blob, filename: f.name.replace(/^.*\//, ''), count: 1 };
    }
    progress(jobs.length, jobs.length, 'Creating zip file…');
    const out = new JSZip();
    for (const f of files) out.file(f.name, f.data, { compression: f.store ? 'STORE' : 'DEFLATE' });
    const blob = await out.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    return { blob, filename: safeName((pkg.courseTitle || 'Canvas course') + ' - Canvas export') + '.zip', count: files.length };
  };

  function buildJobs(pkg, selection, opts, getImage) {
    const jobs = [];
    const catById = {};
    CQE.CATEGORIES.forEach((c) => (catById[c.id] = c));
    const pack = (doc) => docx.Packer.toBlob(doc);

    /* ---- assessments ---- */
    const versions = opts.version === 'both' ? ['key', 'student'] : [opts.version];
    const suffix = { key: 'Answer Key', student: 'Student Copy' };
    const quizBuilder = (v) =>
      new CQE.WordBuilder(pkg, { version: v, includeFeedback: opts.includeFeedback, includePoints: opts.includePoints, expandBanks: opts.expandBanks, answerTable: opts.answerTable, courseTitle: pkg.courseTitle }, getImage);
    const assessCats = ['classic-bank', 'item-bank', 'quiz-groups', 'quiz-fixed'];
    const chosen = [];
    assessCats.forEach((cid) => (selection[cid] || []).forEach((it) => chosen.push({ cat: catById[cid], coll: it.ref })));
    for (const v of versions) {
      if (!chosen.length) break;
      if (opts.quizOutput === 'combined') {
        const title = (pkg.courseTitle || 'Canvas') + ' - Quizzes and banks';
        jobs.push({ name: short(title) + ' - ' + suffix[v] + '.docx', label: title, run: async () => pack(await quizBuilder(v).buildDocument(chosen.map((x) => x.coll), title)) });
      } else {
        for (const x of chosen) {
          jobs.push({ name: x.cat.folder + '/' + short(x.coll.title) + ' - ' + suffix[v] + '.docx', label: x.coll.title, run: async () => pack(await quizBuilder(v).buildDocument([x.coll], x.coll.title)) });
        }
      }
    }

    /* ---- course files ---- */
    for (const cid of ['files-docs', 'files-images', 'files-other']) {
      for (const it of selection[cid] || []) {
        const f = it.ref;
        const rel = opts.keepFolders ? f.rel.split('/').map(safeName).join('/') : safeName(f.name);
        jobs.push({ name: 'Course files/' + rel, label: f.name, run: () => pkg.zip.file(f.path).async('uint8array') });
      }
    }

    /* ---- course content ---- */
    const cb = () => new CQE.ContentBuilder(pkg, { courseTitle: pkg.courseTitle, includeSettings: opts.includeSettings }, getImage);
    for (const it of selection.pages || []) jobs.push({ name: 'Pages/' + short(it.ref.title) + '.docx', label: it.ref.title, run: async () => pack(await cb().page(it.ref)) });
    for (const it of selection.assignments || []) jobs.push({ name: 'Assignments/' + short(it.ref.title) + '.docx', label: it.ref.title, run: async () => pack(await cb().assignment(it.ref)) });
    for (const it of selection.discussions || []) {
      const folder = it.ref.isAnnouncement ? 'Announcements/' : 'Discussions/';
      jobs.push({ name: folder + short(it.ref.title) + '.docx', label: it.ref.title, run: async () => pack(await cb().discussion(it.ref)) });
    }
    for (const it of selection.rubrics || []) jobs.push({ name: 'Rubrics/' + short(it.ref.title) + '.docx', label: it.ref.title, run: async () => pack(await cb().rubric(it.ref)) });
    for (const it of selection.outline || []) jobs.push({ name: 'Course outline.docx', label: 'Course outline', run: async () => pack(await cb().outline(it.ref)) });

    return jobs;
  }

})(typeof window !== 'undefined' ? window : globalThis);
