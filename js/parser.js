/*
 * parser.js
 * Reads a Canvas course export (.imscc) and turns its quiz content into a
 * plain JavaScript model that the Word builder can render.
 *
 * Handles:
 *   - Classic Quizzes question banks   (<objectbank> in non_cc_assessments/)
 *   - New Quizzes item banks           (<objectbank canvas_item_bank="true">)
 *   - Classic Quizzes                  (<assessment>, assessment_meta.xml)
 *   - New Quizzes                      (<assessment>, New Quizzes-style meta)
 *   - Generic Common Cartridge QTI 1.2 assessments (cc.* profile types)
 *
 * Everything runs in the browser. Nothing is uploaded anywhere.
 */
(function (global) {
  'use strict';

  const CQE = (global.CQE = global.CQE || {});

  /* ------------------------------------------------------------------ *
   * Small XML helpers (namespace-agnostic, work on localName)
   * ------------------------------------------------------------------ */
  function kids(el, name) {
    if (!el) return [];
    const out = [];
    for (const c of el.children) if (c.localName === name) out.push(c);
    return out;
  }
  function kid(el, name) {
    if (!el) return null;
    for (const c of el.children) if (c.localName === name) return c;
    return null;
  }
  function desc(el, name) {
    if (!el) return [];
    return Array.from(el.getElementsByTagName(name));
  }
  function txt(el) {
    return el ? (el.textContent || '').trim() : '';
  }
  function hasAncestor(el, name, stopAt) {
    let p = el.parentElement;
    while (p && p !== stopAt) {
      if (p.localName === name) return true;
      p = p.parentElement;
    }
    return false;
  }
  function metaFields(container) {
    // Reads <qtimetadata><qtimetadatafield><fieldlabel/><fieldentry/>
    const map = {};
    if (!container) return map;
    const qm = container.localName === 'qtimetadata' ? container : kid(container, 'qtimetadata');
    if (!qm) return map;
    for (const f of kids(qm, 'qtimetadatafield')) {
      const label = txt(kid(f, 'fieldlabel'));
      if (label) map[label] = txt(kid(f, 'fieldentry'));
    }
    return map;
  }
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function parseXml(str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML parse error: ' + txt(err).slice(0, 200));
    return doc;
  }

  /** Material -> HTML string (mattext may be text/html or text/plain). */
  function materialHtml(materialEl) {
    if (!materialEl) return '';
    const parts = [];
    for (const m of desc(materialEl, 'mattext')) {
      const type = (m.getAttribute('texttype') || 'text/plain').toLowerCase();
      const raw = m.textContent || '';
      parts.push(type.includes('html') ? raw : escapeHtml(raw).replace(/\n/g, '<br>'));
    }
    for (const img of desc(materialEl, 'matimage')) {
      const uri = img.getAttribute('uri');
      if (uri) parts.push('<img src="' + escapeHtml(uri) + '">');
    }
    return parts.join('');
  }

  /* ------------------------------------------------------------------ *
   * Question type names
   * ------------------------------------------------------------------ */
  const TYPE_LABELS = {
    multiple_choice_question: 'Multiple Choice',
    true_false_question: 'True/False',
    multiple_answers_question: 'Multiple Answers',
    matching_question: 'Matching',
    short_answer_question: 'Fill in the Blank',
    fill_in_multiple_blanks_question: 'Fill in Multiple Blanks',
    multiple_dropdowns_question: 'Multiple Dropdowns',
    numerical_question: 'Numerical Answer',
    calculated_question: 'Formula',
    essay_question: 'Essay',
    file_upload_question: 'File Upload',
    text_only_question: 'Text (no question)',
    ordering_question: 'Ordering',
    categorization_question: 'Categorization',
    hot_spot_question: 'Hot Spot',
    formula_question: 'Formula',
    unknown: 'Question'
  };

  // Common Cartridge profile names -> Canvas names
  const CC_PROFILE_MAP = {
    'cc.multiple_choice.v0p1': 'multiple_choice_question',
    'cc.multiple_response.v0p1': 'multiple_answers_question',
    'cc.true_false.v0p1': 'true_false_question',
    'cc.fib.v0p1': 'short_answer_question',
    'cc.pattern_match.v0p1': 'short_answer_question',
    'cc.essay.v0p1': 'essay_question'
  };

  CQE.TYPE_LABELS = TYPE_LABELS;

  /* ------------------------------------------------------------------ *
   * Item (question) parsing
   * ------------------------------------------------------------------ */
  function scoreOf(respcondition) {
    // Returns the numeric score set by this condition (or null)
    for (const sv of kids(respcondition, 'setvar')) {
      const v = (sv.getAttribute('varname') || 'SCORE').toUpperCase();
      if (v === 'SCORE' || v === 'QUE_SCORE') {
        const n = parseFloat(txt(sv));
        if (!isNaN(n)) return n;
      }
    }
    return null;
  }

  function positiveVarequals(conditionvar, respident) {
    // varequal elements that are not inside <not>
    return desc(conditionvar, 'varequal').filter(
      (v) =>
        !hasAncestor(v, 'not', conditionvar) &&
        (!respident || v.getAttribute('respident') === respident)
    );
  }

  function guessType(item, responses) {
    const lids = responses.filter((r) => r.localName === 'response_lid');
    const strs = responses.filter((r) => r.localName === 'response_str' || r.localName === 'response_num');
    if (lids.length > 1) return 'matching_question';
    if (lids.length === 1) {
      const card = (lids[0].getAttribute('rcardinality') || '').toLowerCase();
      return card === 'multiple' ? 'multiple_answers_question' : 'multiple_choice_question';
    }
    if (strs.length) {
      const hasCorrect = desc(item, 'varequal').length > 0;
      return hasCorrect ? 'short_answer_question' : 'essay_question';
    }
    return 'text_only_question';
  }

  function parseItem(item) {
    const meta = metaFields(kid(item, 'itemmetadata'));
    const presentation = kid(item, 'presentation');
    const resprocessing = kid(item, 'resprocessing');
    const responses = presentation
      ? desc(presentation, 'response_lid')
          .concat(desc(presentation, 'response_str'))
          .concat(desc(presentation, 'response_num'))
      : [];

    let type = meta.question_type || meta.cc_profile || '';
    if (CC_PROFILE_MAP[type]) type = CC_PROFILE_MAP[type];
    if (!type) type = guessType(item, responses);

    const pts = parseFloat(meta.points_possible);
    const q = {
      id: item.getAttribute('ident') || '',
      title: item.getAttribute('title') || '',
      type,
      typeLabel: TYPE_LABELS[type] || prettifyType(type),
      points: isNaN(pts) ? null : pts,
      stemHtml: '',
      isPassage: meta.passage === 'true' || type === 'text_only_question',
      choices: [], // {id, html, correct, feedbackHtml}
      matches: [], // {promptHtml, answer}
      matchOptions: [], // all right-side options
      blanks: [], // {name, options:[{text, correct}], answers:[text]}
      acceptedAnswers: [], // short answer
      numericAnswers: [], // strings describing numeric answers
      formula: null, // calculated question info
      feedback: { general: '', correct: '', incorrect: '' },
      gradingNotes: meta.grading_notes || '',
      notes: []
    };

    // ---- stem ----------------------------------------------------------
    if (presentation) {
      const stemParts = [];
      for (const child of presentation.children) {
        if (child.localName === 'material') stemParts.push(materialHtml(child));
        else if (child.localName === 'flow') {
          for (const m of kids(child, 'material')) stemParts.push(materialHtml(m));
        }
      }
      q.stemHtml = stemParts.join('');
    }

    // ---- feedback -------------------------------------------------------
    const fbById = {};
    for (const fb of kids(item, 'itemfeedback')) {
      const id = fb.getAttribute('ident') || '';
      const html = materialHtml(fb);
      if (!html.replace(/<[^>]*>|&nbsp;|\s/g, '')) continue;
      fbById[id] = html;
    }
    q.feedback.general = fbById.general_fb || '';
    q.feedback.correct = fbById.correct_fb || '';
    q.feedback.incorrect = fbById.general_incorrect_fb || '';

    // Map answer ident -> feedback via respconditions (in addition to naming convention)
    const conditions = resprocessing ? kids(resprocessing, 'respcondition') : [];
    const answerFeedback = {};
    for (const rc of conditions) {
      const df = kids(rc, 'displayfeedback').map((d) => d.getAttribute('linkrefid'));
      const ve = positiveVarequals(kid(rc, 'conditionvar'));
      if (ve.length === 1) {
        for (const ref of df) {
          if (ref && fbById[ref] && !/^(general_fb|correct_fb|general_incorrect_fb)$/.test(ref)) {
            answerFeedback[txt(ve[0])] = fbById[ref];
          }
        }
      }
    }
    const feedbackFor = (ansId) => answerFeedback[ansId] || fbById[ansId + '_fb'] || '';

    // Correct answer values (by respident)
    const correctByResp = {}; // respident -> Set(values)
    for (const rc of conditions) {
      const s = scoreOf(rc);
      if (s === null || s <= 0) continue;
      for (const v of positiveVarequals(kid(rc, 'conditionvar'))) {
        const rid = v.getAttribute('respident') || '';
        (correctByResp[rid] = correctByResp[rid] || new Set()).add(txt(v));
      }
    }

    const labelsOf = (resp) =>
      desc(resp, 'response_label').map((l) => ({
        id: l.getAttribute('ident') || '',
        html: materialHtml(kid(l, 'material') || l)
      }));

    // ---- by type -----------------------------------------------------------
    switch (type) {
      case 'multiple_choice_question':
      case 'true_false_question':
      case 'multiple_answers_question': {
        const resp = responses.find((r) => r.localName === 'response_lid');
        if (resp) {
          const correct = correctByResp[resp.getAttribute('ident')] || allCorrect(correctByResp);
          q.choices = labelsOf(resp).map((l) => ({
            id: l.id,
            html: l.html,
            correct: correct.has(l.id),
            feedbackHtml: feedbackFor(l.id)
          }));
          // Canvas multiple-answers often scores with one <and> condition
        }
        break;
      }

      case 'matching_question': {
        const lids = responses.filter((r) => r.localName === 'response_lid');
        const optionText = {};
        for (const resp of lids) {
          const rid = resp.getAttribute('ident');
          const labels = labelsOf(resp);
          labels.forEach((l) => (optionText[l.id] = htmlToText(l.html)));
          const correctIds = correctByResp[rid] ? Array.from(correctByResp[rid]) : [];
          q.matches.push({
            promptHtml: materialHtml(kid(resp, 'material')),
            answer: correctIds.map((id) => optionText[id] || id).join(' / ')
          });
        }
        const used = new Set(q.matches.map((m) => m.answer));
        q.matchOptions = Object.values(optionText);
        q.matchDistractors = q.matchOptions.filter((o) => !used.has(o));
        break;
      }

      case 'fill_in_multiple_blanks_question':
      case 'multiple_dropdowns_question': {
        for (const resp of responses.filter((r) => r.localName === 'response_lid')) {
          const rid = resp.getAttribute('ident');
          const name = htmlToText(materialHtml(kid(resp, 'material'))) || rid.replace(/^response_/, '');
          const correct = correctByResp[rid] || new Set();
          const opts = labelsOf(resp).map((l) => ({ text: htmlToText(l.html), correct: correct.has(l.id) }));
          q.blanks.push({
            name,
            options: opts,
            // For fill-in-multiple-blanks every listed answer is acceptable
            answers:
              type === 'fill_in_multiple_blanks_question'
                ? opts.map((o) => o.text)
                : opts.filter((o) => o.correct).map((o) => o.text)
          });
        }
        break;
      }

      case 'short_answer_question': {
        const set = new Set();
        Object.values(correctByResp).forEach((s) => s.forEach((v) => set.add(v)));
        q.acceptedAnswers = Array.from(set);
        break;
      }

      case 'numerical_question': {
        for (const rc of conditions) {
          const s = scoreOf(rc);
          if (s === null || s <= 0) continue;
          const d = describeNumeric(kid(rc, 'conditionvar'));
          if (d) q.numericAnswers.push(d);
        }
        break;
      }

      case 'calculated_question': {
        q.formula = parseCalculated(item);
        break;
      }

      case 'essay_question':
      case 'file_upload_question':
      case 'text_only_question':
        break;

      default: {
        // Unknown / newer types: show whatever choices exist
        for (const resp of responses.filter((r) => r.localName === 'response_lid')) {
          const correct = correctByResp[resp.getAttribute('ident')] || new Set();
          labelsOf(resp).forEach((l) =>
            q.choices.push({ id: l.id, html: l.html, correct: correct.has(l.id), feedbackHtml: feedbackFor(l.id) })
          );
        }
        q.notes.push(
          'This question type (' + q.typeLabel + ") can't be fully reproduced in Word; review it in Canvas."
        );
      }
    }

    return q;
  }

  function allCorrect(map) {
    const s = new Set();
    Object.values(map).forEach((set) => set.forEach((v) => s.add(v)));
    return s;
  }

  function prettifyType(t) {
    return String(t || 'question')
      .replace(/_question$/, '')
      .replace(/_/g, ' ')
      .replace(/\b\w/g, (c) => c.toUpperCase());
  }

  function htmlToText(html) {
    if (!html) return '';
    const d = new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html');
    return (d.body.textContent || '').replace(/\s+/g, ' ').trim();
  }
  CQE.htmlToText = htmlToText;

  function fmtNum(n) {
    const x = parseFloat(n);
    if (isNaN(x)) return String(n);
    return String(+x.toPrecision(12));
  }

  function describeNumeric(cv) {
    if (!cv) return '';
    const eq = desc(cv, 'varequal')[0];
    const gte = desc(cv, 'vargte')[0] || desc(cv, 'vargt')[0];
    const lte = desc(cv, 'varlte')[0] || desc(cv, 'varlt')[0];
    if (eq && gte && lte) {
      const v = parseFloat(txt(eq));
      const margin = Math.abs(parseFloat(txt(lte)) - v);
      return margin > 0 ? fmtNum(v) + ' (± ' + fmtNum(margin) + ')' : fmtNum(v);
    }
    if (eq) return fmtNum(txt(eq));
    if (gte && lte) return 'Between ' + fmtNum(txt(gte)) + ' and ' + fmtNum(txt(lte));
    if (gte) return 'At least ' + fmtNum(txt(gte));
    if (lte) return 'At most ' + fmtNum(txt(lte));
    return '';
  }

  function parseCalculated(item) {
    const calc = desc(item, 'calculated')[0];
    if (!calc) return null;
    const info = {
      tolerance: txt(kid(calc, 'answer_tolerance')),
      decimals: '',
      formulas: desc(calc, 'formula').map(txt).filter(Boolean),
      vars: [],
      examples: []
    };
    const formulasEl = kid(calc, 'formulas');
    if (formulasEl) info.decimals = formulasEl.getAttribute('decimal_places') || '';
    const varsEl = kid(calc, 'vars');
    if (varsEl) {
      for (const v of kids(varsEl, 'var')) {
        info.vars.push({
          name: v.getAttribute('name'),
          min: txt(kid(v, 'min')),
          max: txt(kid(v, 'max')),
          scale: v.getAttribute('scale') || ''
        });
      }
    }
    const sets = kid(calc, 'var_sets');
    if (sets) {
      for (const s of kids(sets, 'var_set').slice(0, 5)) {
        const values = kids(s, 'var').map((v) => v.getAttribute('name') + ' = ' + txt(v));
        info.examples.push({ values: values.join(', '), answer: txt(kid(s, 'answer')) });
      }
    }
    return info;
  }

  /* ------------------------------------------------------------------ *
   * Package parsing
   * ------------------------------------------------------------------ */
  const KIND_LABELS = {
    'classic-bank': 'Classic Quizzes question bank',
    'item-bank': 'New Quizzes item bank',
    'classic-quiz': 'Classic Quiz',
    'new-quiz': 'New Quiz'
  };
  CQE.KIND_LABELS = KIND_LABELS;

  function parseSection(sectionEl) {
    // Returns a section object with items and/or bank references
    const sec = {
      id: sectionEl.getAttribute('ident') || '',
      title: sectionEl.getAttribute('title') || '',
      pick: null,
      pointsPerItem: null,
      bankRef: null,
      items: [],
      sections: []
    };
    const so = kid(sectionEl, 'selection_ordering');
    if (so) {
      const sel = kid(so, 'selection');
      if (sel) {
        const ref = txt(kid(sel, 'sourcebank_ref'));
        sec.bankRef = ref || null;
        const n = parseInt(txt(kid(sel, 'selection_number')), 10);
        sec.pick = isNaN(n) ? null : n;
        const ppi = parseFloat(txt(desc(sel, 'points_per_item')[0]));
        sec.pointsPerItem = isNaN(ppi) ? null : ppi;
        sec.isGroup = true;
      }
    }
    for (const child of sectionEl.children) {
      if (child.localName === 'item') sec.items.push(parseItem(child));
      else if (child.localName === 'section') sec.sections.push(parseSection(child));
    }
    return sec;
  }

  function countItems(sections) {
    let n = 0;
    for (const s of sections) n += s.items.length + countItems(s.sections);
    return n;
  }

  /**
   * Parse an .imscc File/Blob/ArrayBuffer.
   * Returns { courseTitle, collections, warnings, zip, stats }
   */
  CQE.parsePackage = async function parsePackage(fileOrBuffer, onProgress) {
    const progress = onProgress || function () {};
    const zip = await JSZip.loadAsync(fileOrBuffer);
    const warnings = [];

    if (!zip.file('imsmanifest.xml')) {
      throw new Error(
        "This doesn't look like a Canvas course export: imsmanifest.xml is missing. " +
          'Export the course from Canvas (Settings → Export Course Content → Course) and try again.'
      );
    }

    // Course title
    let courseTitle = '';
    const cs = zip.file('course_settings/course_settings.xml');
    if (cs) {
      try {
        courseTitle = txt(parseXml(await cs.async('string')).getElementsByTagName('title')[0]);
      } catch (e) {
        /* ignore */
      }
    }
    const manifestDoc = parseXml(await zip.file('imsmanifest.xml').async('string'));
    if (!courseTitle) {
      const t = manifestDoc.getElementsByTagName('title')[0];
      courseTitle = txt(t) || 'Canvas course';
    }

    // Gather QTI files: Canvas-native (non_cc_assessments) first, then CC assessment files
    const qtiPaths = [];
    const seen = new Set();
    zip.forEach((path, f) => {
      if (!f.dir && /^non_cc_assessments\/.+\.(xml\.qti|qti|xml)$/i.test(path)) {
        qtiPaths.push({ path, canvas: true });
        seen.add(path);
      }
    });
    const nonCcIds = new Set(qtiPaths.map((p) => p.path.replace(/^.*\//, '').replace(/\..*$/, '')));
    for (const res of Array.from(manifestDoc.getElementsByTagName('resource'))) {
      const type = res.getAttribute('type') || '';
      if (!/imsqti/i.test(type)) continue;
      const fileEl = kids(res, 'file')[0];
      const href = (fileEl && fileEl.getAttribute('href')) || res.getAttribute('href');
      if (!href || seen.has(href)) continue;
      const ident = res.getAttribute('identifier');
      if (nonCcIds.has(ident) || nonCcIds.has(href.split('/')[0])) continue; // richer copy already present
      if (zip.file(href)) {
        qtiPaths.push({ path: href, canvas: false });
        seen.add(href);
      }
    }

    const collections = [];
    let done = 0;
    for (const { path } of qtiPaths) {
      done++;
      progress(done, qtiPaths.length);
      let doc;
      try {
        doc = parseXml(await zip.file(path).async('string'));
      } catch (e) {
        warnings.push('Skipped ' + path + ': ' + e.message);
        continue;
      }
      const root = doc.documentElement;

      for (const bank of kids(root, 'objectbank')) {
        const meta = metaFields(bank);
        const isItemBank = bank.getAttribute('canvas_item_bank') === 'true';
        const items = kids(bank, 'item').map(parseItem);
        // Some exports wrap items in sections even in banks
        for (const s of kids(bank, 'section')) items.push(...flattenSection(parseSection(s)));
        collections.push({
          id: bank.getAttribute('ident') || path,
          kind: isItemBank ? 'item-bank' : 'classic-bank',
          title: meta.bank_title || bank.getAttribute('title') || 'Untitled bank',
          sections: [{ id: 'root', title: '', items, sections: [] }],
          questionCount: items.length,
          source: path
        });
      }

      for (const asmt of kids(root, 'assessment')) {
        const ident = asmt.getAttribute('ident') || '';
        const title = asmt.getAttribute('title') || 'Untitled quiz';
        let kind = 'classic-quiz';
        let metaXml = '';
        const metaFile = zip.file(ident + '/assessment_meta.xml');
        if (metaFile) metaXml = await metaFile.async('string');
        if (/<(calculator_type|display_item_feedback|build_on_last_attempt)\b/.test(metaXml)) kind = 'new-quiz';
        else if (!metaXml && desc(asmt, 'fieldlabel').some((f) => txt(f) === 'calculator_type')) kind = 'new-quiz';

        let description = '';
        let quizType = '';
        let pointsPossible = null;
        if (metaXml) {
          try {
            const md = parseXml(metaXml);
            description = txt(md.getElementsByTagName('description')[0]);
            quizType = txt(md.getElementsByTagName('quiz_type')[0]);
            const pp = parseFloat(txt(md.getElementsByTagName('points_possible')[0]));
            if (!isNaN(pp)) pointsPossible = pp;
          } catch (e) {
            /* ignore */
          }
        }

        const sections = kids(asmt, 'section').map(parseSection);
        collections.push({
          id: ident || path,
          kind,
          title,
          description,
          quizType,
          pointsPossible,
          sections,
          questionCount: countItems(sections),
          source: path
        });
      }
    }

    // Link quiz sections to banks
    const byId = {};
    collections.forEach((c) => (byId[c.id] = c));
    for (const c of collections) {
      if (c.kind !== 'classic-quiz' && c.kind !== 'new-quiz') continue;
      c.bankRefs = [];
      walkSections(c.sections, (s) => {
        if (s.isGroup && s.bankRef) {
          const bank = byId[s.bankRef];
          s.bankTitle = bank ? bank.title : null;
          s.bankMissing = !bank;
          c.bankRefs.push(s.bankRef);
        } else if (s.isGroup && !s.bankRef && !s.items.length) {
          s.bankMissing = true;
        }
      });
      c.pullsFromBanks = c.bankRefs.length;
    }

    // Sort: kind, then natural title order
    const order = { 'classic-bank': 0, 'item-bank': 1, 'classic-quiz': 2, 'new-quiz': 3 };
    const coll = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    collections.sort((a, b) => order[a.kind] - order[b.kind] || coll.compare(a.title, b.title));

    const stats = { 'classic-bank': 0, 'item-bank': 0, 'classic-quiz': 0, 'new-quiz': 0, questions: 0 };
    collections.forEach((c) => {
      stats[c.kind]++;
      stats.questions += c.questionCount;
    });

    // Quizzes that use question groups (random pools) vs fixed question lists
    collections.forEach((c) => {
      if (c.kind === 'classic-quiz' || c.kind === 'new-quiz') {
        let groups = false;
        walkSections(c.sections, (s) => {
          if (s.isGroup) groups = true;
        });
        c.hasGroups = groups;
      }
    });

    let content = { files: [], pages: [], assignments: [], discussions: [], rubrics: [], outline: null };
    if (CQE.parseCourseContent) {
      try {
        content = await CQE.parseCourseContent(zip, manifestDoc);
      } catch (e) {
        warnings.push('Could not read some course content: ' + e.message);
      }
    }

    return { courseTitle, collections, byId, warnings, zip, stats, content };
  };

  function flattenSection(sec) {
    let out = sec.items.slice();
    for (const s of sec.sections) out = out.concat(flattenSection(s));
    return out;
  }
  function walkSections(sections, fn) {
    for (const s of sections) {
      fn(s);
      walkSections(s.sections, fn);
    }
  }
  CQE.walkSections = walkSections;

  /** Resolve an <img src> from question HTML to a path inside the zip. */
  CQE.resolveZipPath = function (src) {
    if (!src) return null;
    let s = src.trim();
    s = s.replace(/^%24IMS(-|_)CC(-|_)FILEBASE%24/i, '$IMS-CC-FILEBASE$');
    const m = s.match(/^\$IMS[-_]CC[-_]FILEBASE\$\/?(.*)$/i);
    if (!m) return null;
    let p = m[1].split('?')[0].split('#')[0];
    try {
      p = decodeURIComponent(p);
    } catch (e) {
      /* keep raw */
    }
    return 'web_resources/' + p.replace(/^\/+/, '');
  };
})(typeof window !== 'undefined' ? window : globalThis);
