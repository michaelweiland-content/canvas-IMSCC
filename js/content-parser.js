/*
 * content-parser.js
 * Reads the non-quiz parts of a Canvas course export:
 *   course files, pages, assignments, discussions/announcements,
 *   rubrics, learning outcomes, modules and assignment groups.
 */
(function (global) {
  'use strict';
  const CQE = (global.CQE = global.CQE || {});

  function kids(el, name) {
    if (!el) return [];
    return Array.from(el.children).filter((c) => c.localName === name);
  }
  function kid(el, name) {
    return kids(el, name)[0] || null;
  }
  function txt(el) {
    return el ? (el.textContent || '').trim() : '';
  }
  function childText(el, name) {
    return txt(kid(el, name));
  }
  function parseXml(str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    if (doc.getElementsByTagName('parsererror')[0]) return null;
    return doc;
  }
  async function readXml(zip, path) {
    const f = zip.file(path);
    if (!f) return null;
    try {
      return parseXml(await f.async('string'));
    } catch (e) {
      return null;
    }
  }
  function num(s) {
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }

  /* -------------------- file categories -------------------- */
  const DOC_EXT = new Set(['pdf', 'doc', 'docx', 'dotx', 'rtf', 'odt', 'txt', 'ppt', 'pptx', 'pps', 'ppsx', 'odp', 'key', 'xls', 'xlsx', 'xlsm', 'csv', 'ods', 'numbers', 'pages', 'epub', 'md']);
  const IMG_EXT = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp', 'tif', 'tiff', 'heic', 'ico']);
  const FILE_KIND_LABEL = {
    pdf: 'PDF', doc: 'Word', docx: 'Word', dotx: 'Word', rtf: 'Word', odt: 'Word', txt: 'Text', md: 'Text',
    ppt: 'PowerPoint', pptx: 'PowerPoint', pps: 'PowerPoint', ppsx: 'PowerPoint', odp: 'Slides', key: 'Keynote',
    xls: 'Excel', xlsx: 'Excel', xlsm: 'Excel', csv: 'CSV', ods: 'Spreadsheet', numbers: 'Numbers', pages: 'Pages', epub: 'eBook'
  };
  CQE.fileCategory = function (path) {
    const ext = (path.split('.').pop() || '').toLowerCase();
    if (DOC_EXT.has(ext)) return 'files-docs';
    if (IMG_EXT.has(ext)) return 'files-images';
    return 'files-other';
  };
  CQE.fileKindLabel = function (path) {
    const ext = (path.split('.').pop() || '').toLowerCase();
    return FILE_KIND_LABEL[ext] || (IMG_EXT.has(ext) ? 'Image' : ext ? ext.toUpperCase() : 'File');
  };

  function stripHtmlDoc(html) {
    // Return the <body> contents and <title> of a full HTML document
    const d = new DOMParser().parseFromString(html, 'text/html');
    const meta = {};
    d.querySelectorAll('meta[name]').forEach((m) => (meta[m.getAttribute('name')] = m.getAttribute('content')));
    return { title: (d.title || '').trim(), body: d.body ? d.body.innerHTML : html, meta };
  }

  /**
   * Parse everything that isn't a quiz.
   * Returns { files, pages, assignments, discussions, rubrics, outline }
   */
  CQE.parseCourseContent = async function (zip, manifestDoc) {
    const out = { files: [], pages: [], assignments: [], discussions: [], rubrics: [], outline: null };

    /* ---- course files (web_resources) ---- */
    const hidden = new Set();
    const fm = await readXml(zip, 'course_settings/files_meta.xml');
    if (fm) {
      Array.from(fm.getElementsByTagName('folder')).forEach((f) => {
        if (childText(f, 'hidden') === 'true') hidden.add(f.getAttribute('path'));
      });
    }
    zip.forEach((path, f) => {
      if (f.dir || !/^web_resources\//.test(path)) return;
      const rel = path.slice('web_resources/'.length);
      if (!rel) return;
      const folder = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
      const isHidden = Array.from(hidden).some((h) => folder === h || folder.startsWith(h + '/'));
      out.files.push({
        id: path,
        path,
        rel,
        name: rel.split('/').pop(),
        folder,
        size: (f._data && f._data.uncompressedSize) || 0,
        category: CQE.fileCategory(rel),
        kindLabel: CQE.fileKindLabel(rel),
        hidden: isHidden || /^(assessment_questions|Quiz Files)(\/|$)/.test(folder)
      });
    });

    /* ---- pages (wiki_content) ---- */
    const pagePaths = [];
    zip.forEach((path, f) => {
      if (!f.dir && /^wiki_content\/.+\.html?$/i.test(path)) pagePaths.push(path);
    });
    for (const path of pagePaths) {
      const html = await zip.file(path).async('string');
      const d = stripHtmlDoc(html);
      out.pages.push({
        id: path,
        title: d.title || path.split('/').pop().replace(/\.html?$/i, ''),
        html: d.body,
        state: d.meta.workflow_state || '',
        frontPage: d.meta.front_page === 'true'
      });
    }

    /* ---- assignment groups ---- */
    const groups = {};
    const groupList = [];
    const ag = await readXml(zip, 'course_settings/assignment_groups.xml');
    if (ag) {
      Array.from(ag.getElementsByTagName('assignmentGroup')).forEach((g) => {
        const item = { id: g.getAttribute('identifier'), title: childText(g, 'title'), weight: num(childText(g, 'group_weight')), position: num(childText(g, 'position')) };
        groups[item.id] = item;
        groupList.push(item);
      });
    }

    /* ---- assignments (<id>/assignment_settings.xml) ---- */
    const asgPaths = [];
    zip.forEach((path, f) => {
      if (!f.dir && /^[^/]+\/assignment_settings\.xml$/.test(path)) asgPaths.push(path);
    });
    for (const path of asgPaths) {
      const doc = await readXml(zip, path);
      if (!doc) continue;
      const a = doc.documentElement;
      const dir = path.split('/')[0];
      let html = '';
      const htmlFile = zip.file(new RegExp('^' + dir + '/[^/]+\\.html?$'))[0];
      if (htmlFile) html = stripHtmlDoc(await htmlFile.async('string')).body;
      const g = groups[childText(a, 'assignment_group_identifierref')];
      out.assignments.push({
        id: a.getAttribute('identifier') || dir,
        title: childText(a, 'title') || 'Untitled assignment',
        html,
        points: num(childText(a, 'points_possible')),
        gradingType: childText(a, 'grading_type'),
        submissionTypes: childText(a, 'submission_types'),
        due: childText(a, 'due_at'),
        unlock: childText(a, 'unlock_at'),
        lock: childText(a, 'lock_at'),
        state: childText(a, 'workflow_state'),
        group: g ? g.title : '',
        allowedExtensions: childText(a, 'allowed_extensions'),
        externalUrl: childText(a, 'external_tool_url'),
        peerReviews: childText(a, 'peer_reviews') === 'true',
        rubricRef: childText(a, 'rubric_identifierref'),
        position: num(childText(a, 'position'))
      });
    }

    /* ---- discussions & announcements (imsdt resources) ---- */
    const resById = {};
    Array.from(manifestDoc.getElementsByTagName('resource')).forEach((r) => (resById[r.getAttribute('identifier')] = r));
    for (const r of Object.values(resById)) {
      if (!/imsdt/i.test(r.getAttribute('type') || '')) continue;
      const fileEl = kid(r, 'file');
      const href = fileEl && fileEl.getAttribute('href');
      if (!href) continue;
      const doc = await readXml(zip, href);
      if (!doc) continue;
      const topic = doc.documentElement;
      let meta = null;
      for (const dep of kids(r, 'dependency')) {
        const mr = resById[dep.getAttribute('identifierref')];
        const mf = mr && (mr.getAttribute('href') || (kid(mr, 'file') && kid(mr, 'file').getAttribute('href')));
        if (mf) meta = await readXml(zip, mf);
        if (meta) break;
      }
      const m = meta ? meta.documentElement : null;
      const type = m ? childText(m, 'type') : '';
      const asg = m ? kid(m, 'assignment') : null;
      out.discussions.push({
        id: r.getAttribute('identifier'),
        title: childText(topic, 'title') || (m ? childText(m, 'title') : '') || 'Untitled discussion',
        html: childText(topic, 'text'),
        isAnnouncement: type === 'announcement',
        graded: !!asg,
        points: asg ? num(childText(asg, 'points_possible')) : null,
        due: asg ? childText(asg, 'due_at') : '',
        delayedPost: m ? childText(m, 'delayed_post_at') : '',
        state: m ? childText(m, 'workflow_state') : ''
      });
    }

    /* ---- rubrics ---- */
    const rb = await readXml(zip, 'course_settings/rubrics.xml');
    if (rb) {
      Array.from(rb.getElementsByTagName('rubric')).forEach((r) => {
        const criteria = [];
        const crit = kid(r, 'criteria');
        for (const c of kids(crit, 'criterion')) {
          const ratings = kids(kid(c, 'ratings'), 'rating').map((x) => ({
            title: childText(x, 'description'),
            long: childText(x, 'long_description'),
            points: num(childText(x, 'points'))
          }));
          criteria.push({
            title: childText(c, 'description'),
            long: childText(c, 'long_description'),
            points: num(childText(c, 'points')),
            outcome: childText(c, 'learning_outcome_identifierref'),
            ratings
          });
        }
        out.rubrics.push({
          id: r.getAttribute('identifier'),
          title: childText(r, 'title') || 'Untitled rubric',
          points: num(childText(r, 'points_possible')),
          freeForm: childText(r, 'free_form_criterion_comments') === 'true',
          criteria
        });
      });
    }

    /* ---- outline: modules, assignment groups, outcomes ---- */
    const modules = [];
    const mm = await readXml(zip, 'course_settings/module_meta.xml');
    if (mm) {
      Array.from(mm.getElementsByTagName('module')).forEach((m) => {
        const items = kids(kid(m, 'items'), 'item').map((it) => ({
          type: childText(it, 'content_type'),
          title: childText(it, 'title'),
          indent: parseInt(childText(it, 'indent'), 10) || 0,
          url: childText(it, 'url'),
          state: childText(it, 'workflow_state')
        }));
        modules.push({ title: childText(m, 'title'), state: childText(m, 'workflow_state'), sequential: childText(m, 'require_sequential_progress') === 'true', items });
      });
    }

    const outcomes = [];
    const lo = await readXml(zip, 'course_settings/learning_outcomes.xml');
    function walkOutcomes(el, path) {
      for (const c of Array.from(el.children)) {
        if (c.localName === 'learningOutcomeGroup') {
          walkOutcomes(kid(c, 'learningOutcomes') || c, path.concat(childText(c, 'title')));
        } else if (c.localName === 'learningOutcome') {
          outcomes.push({
            group: path.filter(Boolean).join(' › '),
            title: childText(c, 'title'),
            description: childText(c, 'description'),
            points: num(childText(c, 'points_possible')),
            mastery: num(childText(c, 'mastery_points')),
            method: childText(c, 'calculation_method'),
            ratings: kids(kid(c, 'ratings'), 'rating').map((x) => ({ title: childText(x, 'description'), points: num(childText(x, 'points')) }))
          });
        } else if (c.localName === 'learningOutcomes') {
          walkOutcomes(c, path);
        }
      }
    }
    if (lo) walkOutcomes(lo.documentElement, []);

    if (modules.length || groupList.length || outcomes.length) {
      groupList.sort((a, b) => (a.position || 0) - (b.position || 0));
      out.outline = { id: 'outline', title: 'Course outline', modules, groups: groupList, outcomes };
    }

    const coll = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    out.files.sort((a, b) => coll.compare(a.rel, b.rel));
    out.pages.sort((a, b) => coll.compare(a.title, b.title));
    out.assignments.sort((a, b) => coll.compare(a.title, b.title));
    out.discussions.sort((a, b) => coll.compare(a.title, b.title));
    out.rubrics.sort((a, b) => coll.compare(a.title, b.title));
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);
