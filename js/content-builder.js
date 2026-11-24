/*
 * content-builder.js
 * Word documents for non-quiz content: pages, assignments, discussions,
 * rubrics and the course outline.
 */
(function (global) {
  'use strict';
  const CQE = (global.CQE = global.CQE || {});
  const D = () => global.docx;

  const SUBMISSION_LABELS = {
    online_text_entry: 'Text entry',
    online_url: 'Website URL',
    online_upload: 'File upload',
    media_recording: 'Media recording',
    student_annotation: 'Student annotation',
    external_tool: 'External tool',
    on_paper: 'On paper',
    none: 'No submission',
    online_quiz: 'Quiz',
    discussion_topic: 'Discussion',
    not_graded: 'Not graded',
    wiki_page: 'Page'
  };

  function fmtDate(s) {
    if (!s) return '';
    const d = new Date(s);
    if (isNaN(d)) return s;
    return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
  }
  function fmtPts(p) {
    if (p == null) return '';
    const v = +(+p).toFixed(2);
    return v + (v === 1 ? ' pt' : ' pts');
  }

  class ContentBuilder {
    constructor(pkg, options, imageLoader) {
      this.pkg = pkg;
      this.opt = options || {};
      this.conv = new CQE.HtmlConverter({ getImage: imageLoader });
    }
    T(text, props) {
      return new (D().TextRun)(Object.assign({ text: String(text == null ? '' : text) }, props || {}));
    }
    P(children, props) {
      return new (D().Paragraph)(Object.assign({ children: Array.isArray(children) ? children : [children] }, props || {}));
    }
    title(text, sub) {
      const D_ = D();
      const out = [new D_.Paragraph({ heading: D_.HeadingLevel.HEADING_1, children: [this.T(text)] })];
      const line = [sub, this.opt.courseTitle].filter(Boolean).join('  •  ');
      if (line) out.push(this.P(this.T(line, { color: CQE.COLOR.muted, size: 20 }), { spacing: { after: 120 } }));
      out.push(this.P([], { border: { bottom: { style: D_.BorderStyle.SINGLE, size: 8, color: CQE.COLOR.passageBorder, space: 1 } }, spacing: { after: 200 } }));
      return out;
    }
    /** Two-column settings table: [[label, value], ...] (empty values skipped) */
    settingsTable(rows) {
      const D_ = D();
      rows = rows.filter((r) => r[1] !== '' && r[1] != null);
      if (!rows.length) return [];
      const pad = { top: 40, bottom: 40, left: 100, right: 100 };
      const widths = [2600, 6760];
      return [
        new D_.Table({
          width: { size: 9360, type: D_.WidthType.DXA },
          columnWidths: widths,
          rows: rows.map(
            ([k, v]) =>
              new D_.TableRow({
                children: [
                  new D_.TableCell({ width: { size: widths[0], type: D_.WidthType.DXA }, margins: pad, shading: { type: D_.ShadingType.CLEAR, color: 'auto', fill: 'F2F2F2' }, children: [this.P(this.T(k, { bold: true, size: 20 }), { spacing: { after: 0 } })] }),
                  new D_.TableCell({ width: { size: widths[1], type: D_.WidthType.DXA }, margins: pad, children: [this.P(this.T(v, { size: 20 }), { spacing: { after: 0 } })] })
                ]
              })
          )
        }),
        this.P([], { spacing: { after: 200 } })
      ];
    }
    async body(html, emptyText) {
      const blocks = await this.conv.convert(html || '');
      return blocks.length ? blocks : [this.P(this.T(emptyText || '(No description in the export.)', { italics: true, color: CQE.COLOR.muted }))];
    }
    doc(children, title, landscape) {
      return CQE.makeDocument(children, { title, courseTitle: this.opt.courseTitle, landscape });
    }

    async page(p) {
      const ch = this.title(p.title, 'Page' + (p.frontPage ? ' (front page)' : '') + (p.state && p.state !== 'active' ? ' • ' + p.state : ''));
      ch.push(...(await this.body(p.html, '(This page is empty.)')));
      return this.doc(ch, p.title);
    }

    async assignment(a) {
      const ch = this.title(a.title, 'Assignment');
      if (this.opt.includeSettings) {
        const subs = (a.submissionTypes || '')
          .split(',')
          .map((s) => SUBMISSION_LABELS[s.trim()] || s.trim())
          .filter(Boolean)
          .join(', ');
        ch.push(
          ...this.settingsTable([
            ['Points', a.points != null ? fmtPts(a.points) : ''],
            ['Grading', a.gradingType && a.gradingType !== 'points' ? a.gradingType.replace(/_/g, ' ') : ''],
            ['Assignment group', a.group],
            ['Submission type', subs],
            ['Allowed file types', a.allowedExtensions],
            ['External tool', a.externalUrl],
            ['Available from', fmtDate(a.unlock)],
            ['Due', fmtDate(a.due)],
            ['Until', fmtDate(a.lock)],
            ['Peer reviews', a.peerReviews ? 'Yes' : ''],
            ['Status', a.state && a.state !== 'published' ? a.state : '']
          ])
        );
      }
      ch.push(...(await this.body(a.html)));
      const rubric = a.rubricRef && this.pkg.content.rubrics.find((r) => r.id === a.rubricRef);
      if (rubric) {
        ch.push(new (D().Paragraph)({ heading: D().HeadingLevel.HEADING_2, children: [this.T('Rubric: ' + rubric.title)] }));
        ch.push(...this.rubricTable(rubric, 9360));
      }
      return this.doc(ch, a.title);
    }

    async discussion(d) {
      const kind = d.isAnnouncement ? 'Announcement' : d.graded ? 'Graded discussion' : 'Discussion';
      const ch = this.title(d.title, kind);
      if (this.opt.includeSettings) {
        ch.push(
          ...this.settingsTable([
            ['Points', d.points != null ? fmtPts(d.points) : ''],
            ['Due', fmtDate(d.due)],
            ['Post date', fmtDate(d.delayedPost)],
            ['Status', d.state && !/active|published/.test(d.state) ? d.state : '']
          ])
        );
      }
      ch.push(...(await this.body(d.html)));
      ch.push(this.P(this.T('Student replies are not included in Canvas course exports.', { italics: true, color: CQE.COLOR.muted, size: 18 }), { spacing: { before: 240 } }));
      return this.doc(ch, d.title);
    }

    rubricTable(r, width) {
      const D_ = D();
      const maxRatings = Math.max(1, ...r.criteria.map((c) => c.ratings.length));
      const critW = Math.round(width * 0.24);
      const ptsW = 900;
      const rateW = Math.floor((width - critW - ptsW) / maxRatings);
      const widths = [critW].concat(new Array(maxRatings).fill(rateW)).concat([ptsW]);
      const pad = { top: 60, bottom: 60, left: 80, right: 80 };
      const head = { type: D_.ShadingType.CLEAR, color: 'auto', fill: 'E7ECF4' };
      const cell = (children, w, shade, span) =>
        new D_.TableCell({ width: { size: w, type: D_.WidthType.DXA }, margins: pad, shading: shade, columnSpan: span, children });
      const rows = [];
      rows.push(
        new D_.TableRow({
          tableHeader: true,
          children: [cell([this.P(this.T('Criteria', { bold: true, size: 18 }), { spacing: { after: 0 } })], critW, head), cell([this.P(this.T('Ratings', { bold: true, size: 18 }), { spacing: { after: 0 } })], rateW * maxRatings, head, maxRatings > 1 ? maxRatings : undefined), cell([this.P(this.T('Pts', { bold: true, size: 18 }), { spacing: { after: 0 } })], ptsW, head)]
        })
      );
      for (const c of r.criteria) {
        const critChildren = [this.P(this.T(c.title, { bold: true, size: 18 }), { spacing: { after: 40 } })];
        if (c.long) critChildren.push(this.P(this.T(c.long, { size: 16, color: '595959' }), { spacing: { after: 0 } }));
        const ratingCells = [];
        for (let i = 0; i < maxRatings; i++) {
          const x = c.ratings[i];
          if (!x) {
            ratingCells.push(cell([this.P([])], rateW));
            continue;
          }
          const kids = [this.P([this.T(x.points != null ? +x.points.toFixed(2) + ' pts' : '', { bold: true, size: 16, color: CQE.COLOR.heading })], { spacing: { after: 20 } }), this.P(this.T(x.title, { bold: true, size: 17 }), { spacing: { after: 20 } })];
          if (x.long) kids.push(this.P(this.T(x.long, { size: 16 }), { spacing: { after: 0 } }));
          ratingCells.push(cell(kids, rateW));
        }
        rows.push(new D_.TableRow({ cantSplit: true, children: [cell(critChildren, critW)].concat(ratingCells).concat([cell([this.P(this.T(c.points != null ? +c.points.toFixed(2) : '', { size: 18 }), { spacing: { after: 0 } })], ptsW)]) }));
      }
      rows.push(
        new D_.TableRow({
          children: [cell([this.P(this.T('Total points', { bold: true, size: 18 }), { spacing: { after: 0 }, alignment: D_.AlignmentType.RIGHT })], critW + rateW * maxRatings, head, maxRatings + 1), cell([this.P(this.T(r.points != null ? +r.points.toFixed(2) : '', { bold: true, size: 18 }), { spacing: { after: 0 } })], ptsW, head)]
        })
      );
      return [new D_.Table({ width: { size: width, type: D_.WidthType.DXA }, columnWidths: widths, rows }), this.P([], { spacing: { after: 120 } })];
    }

    async rubric(r) {
      const ch = this.title(r.title, 'Rubric  •  ' + r.criteria.length + ' criteria' + (r.points != null ? '  •  ' + fmtPts(r.points) : ''));
      if (!r.criteria.length) ch.push(this.P(this.T('This rubric has no criteria in the export.', { italics: true })));
      else ch.push(...this.rubricTable(r, 12960));
      return this.doc(ch, r.title, true);
    }

    async outline(o) {
      const D_ = D();
      const ch = this.title('Course outline', 'Modules, assignment groups and learning outcomes');
      const h2 = (t) => new D_.Paragraph({ heading: D_.HeadingLevel.HEADING_2, children: [this.T(t)] });

      ch.push(h2('Modules'));
      if (!o.modules.length) ch.push(this.P(this.T('No modules in the export.', { italics: true, color: CQE.COLOR.muted })));
      o.modules.forEach((m, i) => {
        const note = [m.sequential ? 'sequential' : '', m.state && m.state !== 'active' ? m.state : ''].filter(Boolean).join(', ');
        ch.push(this.P([this.T(i + 1 + '. ' + m.title, { bold: true, color: CQE.COLOR.heading }), this.T(note ? '   (' + note + ')' : '', { color: CQE.COLOR.muted, size: 18 })], { spacing: { before: 160, after: 60 }, keepNext: true }));
        if (!m.items.length) ch.push(this.P(this.T('(empty)', { italics: true, color: CQE.COLOR.muted, size: 20 }), { indent: { left: 360 } }));
        for (const it of m.items) {
          const isHeader = it.type === 'ContextModuleSubHeader';
          ch.push(
            this.P(
              [this.T(isHeader ? it.title : '•  ' + it.title, { bold: isHeader, size: 20 }), this.T(isHeader ? '' : '   ' + typeLabel(it.type), { color: CQE.COLOR.muted, size: 16 })],
              { indent: { left: 360 + it.indent * 360 }, spacing: { after: 30 } }
            )
          );
        }
      });

      ch.push(h2('Assignment groups'));
      if (!o.groups.length) ch.push(this.P(this.T('No assignment groups in the export.', { italics: true, color: CQE.COLOR.muted })));
      else {
        const weighted = o.groups.some((g) => g.weight);
        ch.push(...this.settingsTable(o.groups.map((g) => [g.title, weighted ? (g.weight || 0) + '%' : 'Not weighted'])));
      }

      ch.push(h2('Learning outcomes'));
      if (!o.outcomes.length) ch.push(this.P(this.T('No learning outcomes in the export.', { italics: true, color: CQE.COLOR.muted })));
      let lastGroup = null;
      for (const oc of o.outcomes) {
        if (oc.group && oc.group !== lastGroup) {
          ch.push(new D_.Paragraph({ heading: D_.HeadingLevel.HEADING_3, children: [this.T(oc.group)] }));
          lastGroup = oc.group;
        }
        ch.push(this.P(this.T(oc.title, { bold: true }), { spacing: { before: 100, after: 20 }, keepNext: true }));
        const desc = CQE.htmlToText(oc.description);
        if (desc) ch.push(this.P(this.T(desc, { size: 20 }), { indent: { left: 360 }, spacing: { after: 20 } }));
        const facts = [];
        if (oc.mastery != null) facts.push('Mastery at ' + oc.mastery + (oc.points != null ? ' of ' + oc.points : '') + ' pts');
        if (oc.method) facts.push('calculation: ' + oc.method.replace(/_/g, ' '));
        if (oc.ratings.length) facts.push('scale: ' + oc.ratings.map((r) => r.title + ' (' + r.points + ')').join(', '));
        if (facts.length) ch.push(this.P(this.T(facts.join('  •  '), { color: CQE.COLOR.muted, size: 18 }), { indent: { left: 360 }, spacing: { after: 80 } }));
      }
      return this.doc(ch, 'Course outline');
    }
  }

  function typeLabel(t) {
    return (
      {
        Assignment: 'Assignment',
        WikiPage: 'Page',
        Quizzes__Quiz: 'Quiz',
        'Quizzes::Quiz': 'Quiz',
        DiscussionTopic: 'Discussion',
        Attachment: 'File',
        ExternalUrl: 'Link',
        ContextExternalTool: 'External tool',
        ContextModuleSubHeader: ''
      }[t] || (t || '').replace(/([a-z])([A-Z])/g, '$1 $2')
    );
  }

  CQE.ContentBuilder = ContentBuilder;
})(typeof window !== 'undefined' ? window : globalThis);
