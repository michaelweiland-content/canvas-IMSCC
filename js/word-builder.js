/*
 * word-builder.js
 * Builds formatted Word (.docx) documents from parsed quiz collections.
 */
(function (global) {
  'use strict';

  const CQE = (global.CQE = global.CQE || {});
  const D = () => global.docx;

  const COLOR = {
    heading: '1F3864',
    muted: '666666',
    correct: '1E7B34',
    correctFill: 'E8F5EC',
    feedback: '4A4A4A',
    note: '9C5700',
    passageFill: 'F3F6FA',
    passageBorder: '8EA9C9',
    groupFill: 'EEF1F7'
  };

  // Choice layout: [mark] tab [letter] tab [text]
  const CHOICE_INDENT = 1080; // text starts at 0.75"
  const CHOICE_HANG = 720; // first line starts at 0.25"
  const CHOICE_LETTER_POS = 720;

  /**
   * options = {
   *   version: 'key' | 'student',
   *   includeFeedback: bool,
   *   includePoints: bool,
   *   expandBanks: bool,
   *   answerTable: bool,
   *   courseTitle: string
   * }
   */
  class WordBuilder {
    constructor(pkg, options, imageLoader) {
      this.pkg = pkg;
      this.opt = options;
      this.conv = new CQE.HtmlConverter({ getImage: imageLoader });
      this.key = options.version === 'key';
    }

    T(text, props) {
      return new (D().TextRun)(Object.assign({ text }, props || {}));
    }
    P(children, props) {
      return new (D().Paragraph)(Object.assign({ children: Array.isArray(children) ? children : [children] }, props || {}));
    }

    /** Build a full Document for one or more collections. */
    async buildDocument(collections, docTitle) {
      const D_ = D();
      const children = [];
      for (let i = 0; i < collections.length; i++) {
        const blocks = await this.collectionBlocks(collections[i], i > 0);
        children.push(...blocks);
      }
      return CQE.makeDocument(children, { title: docTitle, courseTitle: this.opt.courseTitle });
    }

    async collectionBlocks(c, pageBreak) {
      const D_ = D();
      const out = [];
      this.counter = 0;
      this.summary = [];

      out.push(new D_.Paragraph({ heading: D_.HeadingLevel.HEADING_1, pageBreakBefore: !!pageBreak, children: [this.T(c.title)] }));

      const info = [CQE.KIND_LABELS[c.kind]];
      if (c.kind === 'classic-bank' || c.kind === 'item-bank') info.push(plural(c.questionCount, 'question'));
      else {
        if (c.questionCount) info.push(plural(c.questionCount, 'question') + ' in quiz');
        if (c.pullsFromBanks) info.push('draws from ' + plural(c.pullsFromBanks, 'bank'));
        if (this.opt.includePoints && c.pointsPossible) info.push(fmtPts(c.pointsPossible) + ' total');
      }
      info.push(this.key ? 'Answer key' : 'Student copy');
      out.push(this.P(this.T(info.join('  •  '), { color: COLOR.muted, size: 20 }), { spacing: { after: 60 } }));
      if (this.opt.courseTitle) out.push(this.P(this.T(this.opt.courseTitle, { color: COLOR.muted, size: 20, italics: true }), { spacing: { after: 200 } }));
      out.push(this.P([], { border: { bottom: { style: D_.BorderStyle.SINGLE, size: 8, color: COLOR.passageBorder, space: 1 } }, spacing: { after: 200 } }));

      if (c.description && CQE.htmlToText(c.description)) {
        out.push(this.P(this.T('Instructions', { bold: true, color: COLOR.heading }), { spacing: { after: 60 } }));
        out.push(...(await this.conv.convert(c.description)));
      }

      const isQuiz = c.kind === 'classic-quiz' || c.kind === 'new-quiz';
      let any = false;
      for (const sec of c.sections) {
        const blocks = await this.sectionBlocks(sec, isQuiz, 0);
        if (blocks.length) any = true;
        out.push(...blocks);
      }
      if (!any) {
        out.push(this.P(this.T('This ' + (isQuiz ? 'quiz' : 'bank') + ' has no questions in the export.', { italics: true, color: COLOR.muted })));
      }

      if (this.opt.answerTable && this.summary.length) out.push(...this.answerTable());
      return out;
    }

    async sectionBlocks(sec, isQuiz, depth, inheritedPoints) {
      const D_ = D();
      const out = [];
      if (sec.isGroup) {
        const head = [];
        head.push(this.T(sec.title ? 'Question group: ' + sec.title : 'Question group', { bold: true, color: COLOR.heading, size: 24 }));
        const detail = [];
        if (sec.pick != null) {
          const pool = sec.items.length || (sec.bankRef && this.pkg.byId[sec.bankRef] ? this.pkg.byId[sec.bankRef].questionCount : null);
          detail.push('Canvas picks ' + sec.pick + (pool ? ' of ' + pool : '') + ' question' + (sec.pick === 1 ? '' : 's') + ' at random');
        }
        if (this.opt.includePoints && sec.pointsPerItem != null) detail.push(fmtPts(sec.pointsPerItem) + ' each');
        if (sec.bankRef && sec.bankTitle) detail.push('from bank “' + sec.bankTitle + '”');
        out.push(
          this.P(head, {
            shading: { type: D_.ShadingType.CLEAR, color: 'auto', fill: COLOR.groupFill },
            spacing: { before: 240, after: 0 },
            keepNext: true,
            indent: { left: 120, right: 120 }
          })
        );
        out.push(
          this.P(this.T(detail.join('  •  ') + (detail.length ? '.' : ''), { color: COLOR.muted, size: 20 }), {
            shading: { type: D_.ShadingType.CLEAR, color: 'auto', fill: COLOR.groupFill },
            spacing: { after: 160 },
            keepNext: true,
            indent: { left: 120, right: 120 }
          })
        );
        if (sec.bankMissing) {
          out.push(this.note('The question bank for this group isn’t included in this export, so its questions can’t be listed.'));
        } else if (sec.bankRef && !sec.items.length) {
          const bank = this.pkg.byId[sec.bankRef];
          if (bank && this.opt.expandBanks) {
            for (const s of bank.sections) out.push(...(await this.sectionBlocks(s, false, depth + 1, sec.pointsPerItem)));
          } else if (bank) {
            out.push(this.note('Questions are listed in the bank’s own document (“' + bank.title + '”). Turn on “Include questions pulled from banks” to list them here.'));
          }
        }
      }
      for (const item of sec.items) {
        out.push(...(await this.questionBlocks(item, sec.isGroup ? sec.pointsPerItem : inheritedPoints != null ? inheritedPoints : null)));
      }
      for (const s of sec.sections) out.push(...(await this.sectionBlocks(s, isQuiz, depth + 1)));
      if (sec.isGroup && (sec.items.length || sec.sections.length || sec.bankRef)) {
        out.push(this.P(this.T('End of question group', { italics: true, color: COLOR.muted, size: 18 }), { spacing: { before: 0, after: 200 }, alignment: D_.AlignmentType.RIGHT }));
      }
      return out;
    }

    note(text) {
      return this.P([this.T('Note: ', { bold: true, color: COLOR.note, size: 20 }), this.T(text, { color: COLOR.note, size: 20 })], { spacing: { after: 160 } });
    }

    async questionBlocks(q, pointsOverride) {
      const D_ = D();
      const out = [];

      // Passages / text-only items are not numbered
      if (q.isPassage && !q.choices.length) {
        const shade = { type: D_.ShadingType.CLEAR, color: 'auto', fill: COLOR.passageFill };
        const border = { left: { style: D_.BorderStyle.SINGLE, size: 18, color: COLOR.passageBorder, space: 8 } };
        out.push(this.P(this.T('Passage / instructions', { bold: true, color: COLOR.heading, size: 20 }), { shading: shade, border, spacing: { before: 200, after: 60 }, indent: { left: 180 }, keepNext: true }));
        const body = await this.conv.convert(q.stemHtml, { indent: 180, paragraph: { shading: shade, border, spacing: { after: 80 } } });
        out.push(...body);
        out.push(this.P([], { spacing: { after: 120 } }));
        return out;
      }

      const n = ++this.counter;
      const head = [this.T('Question ' + n, { bold: true, size: 24, color: COLOR.heading })];
      const meta = [q.typeLabel];
      const pts = pointsOverride != null ? pointsOverride : q.points;
      if (this.opt.includePoints && pts != null) meta.push(fmtPts(pts));
      if (q.title && !/^question(\s*\d+)?$/i.test(q.title.trim())) meta.push(q.title.trim());
      head.push(this.T('   ' + meta.join('  •  '), { color: COLOR.muted, size: 18 }));
      out.push(this.P(head, { spacing: { before: 240, after: 80 }, keepNext: true, border: { top: { style: D_.BorderStyle.SINGLE, size: 4, color: 'D9D9D9', space: 6 } } }));

      const stem = await this.conv.convert(q.stemHtml || '', { paragraph: { keepNext: false } });
      out.push(...stem);

      const summary = { n, type: q.typeLabel, answer: '' };

      switch (q.type) {
        case 'multiple_choice_question':
        case 'true_false_question':
        case 'multiple_answers_question':
          out.push(...(await this.choiceBlocks(q, summary)));
          if (q.type === 'multiple_answers_question' && !this.key) out.push(this.hint('Select all that apply.'));
          break;
        case 'matching_question':
          out.push(...this.matchingBlocks(q, summary));
          break;
        case 'fill_in_multiple_blanks_question':
        case 'multiple_dropdowns_question':
          out.push(...this.blankBlocks(q, summary));
          break;
        case 'short_answer_question':
          if (this.key) {
            out.push(this.answerLine('Accepted answers: ', q.acceptedAnswers.length ? q.acceptedAnswers.join('  |  ') : '(none listed)'));
            summary.answer = q.acceptedAnswers.join(' | ');
          } else out.push(this.writeLine());
          break;
        case 'numerical_question':
          if (this.key) {
            out.push(this.answerLine('Correct answer: ', q.numericAnswers.length ? q.numericAnswers.join('  or  ') : '(none listed)'));
            summary.answer = q.numericAnswers.join(' or ');
          } else out.push(this.writeLine());
          break;
        case 'calculated_question':
          out.push(...this.formulaBlocks(q, summary));
          break;
        case 'essay_question':
          if (this.key) {
            summary.answer = '(essay — graded manually)';
            if (q.gradingNotes) out.push(this.answerLine('Grading notes: ', q.gradingNotes));
          } else out.push(this.responseBox(1800));
          break;
        case 'file_upload_question':
          out.push(this.hint('Students submit a file for this question.'));
          summary.answer = '(file upload)';
          break;
        default:
          if (q.choices.length) out.push(...(await this.choiceBlocks(q, summary)));
          break;
      }

      if (this.key && this.opt.includeFeedback) {
        const fbs = [
          ['General feedback: ', q.feedback.general],
          ['If correct: ', q.feedback.correct],
          ['If incorrect: ', q.feedback.incorrect]
        ];
        for (const [label, html] of fbs) {
          if (!html) continue;
          out.push(...(await this.feedbackBlocks(label, html, 0)));
        }
      }
      for (const nt of q.notes) out.push(this.note(nt));

      this.summary.push(summary);
      return out;
    }

    async choiceBlocks(q, summary) {
      const D_ = D();
      const out = [];
      const tabStops = [
        { type: D_.TabStopType.LEFT, position: CHOICE_LETTER_POS },
        { type: D_.TabStopType.LEFT, position: CHOICE_INDENT }
      ];
      const correctLetters = [];
      for (let i = 0; i < q.choices.length; i++) {
        const ch = q.choices[i];
        const letter = CQE.toAlpha(i + 1);
        const mark = this.key && ch.correct;
        if (mark) correctLetters.push(letter + '. ' + truncate(CQE.htmlToText(ch.html), 80));
        const prefix = [
          this.T(mark ? '✓' : '', { bold: true, color: COLOR.correct }),
          new D_.TextRun({ children: [new D_.Tab()] }),
          this.T(letter + '.', { bold: mark, color: mark ? COLOR.correct : undefined }),
          new D_.TextRun({ children: [new D_.Tab()] })
        ];
        const blocks = await this.conv.convert(ch.html, {
          indent: CHOICE_INDENT,
          hanging: CHOICE_HANG,
          tabStops,
          prefixRuns: prefix,
          emptyText: '(blank)',
          run: mark ? { bold: true, color: COLOR.correct } : {},
          paragraph: { spacing: { after: 40 } }
        });
        out.push(...blocks);
        if (this.key && this.opt.includeFeedback && ch.feedbackHtml) {
          out.push(...(await this.feedbackBlocks('Feedback: ', ch.feedbackHtml, CHOICE_INDENT)));
        }
      }
      if (this.key && q.choices.length && !correctLetters.length && /^(multiple_choice|true_false|multiple_answers)_question$/.test(q.type)) {
        out.push(this.note('No correct answer is marked for this question in the export.'));
      }
      summary.answer = correctLetters.join('; ');
      out.push(this.P([], { spacing: { after: 80 } }));
      return out;
    }

    matchingBlocks(q, summary) {
      const D_ = D();
      const out = [];
      const rows = [];
      const cellPad = { top: 60, bottom: 60, left: 100, right: 100 };
      const headFill = { type: D_.ShadingType.CLEAR, color: 'auto', fill: 'F2F2F2' };
      rows.push(
        new D_.TableRow({
          tableHeader: true,
          children: [
            new D_.TableCell({ width: { size: 4680, type: D_.WidthType.DXA }, margins: cellPad, shading: headFill, children: [this.P(this.T('Prompt', { bold: true }))] }),
            new D_.TableCell({ width: { size: 4680, type: D_.WidthType.DXA }, margins: cellPad, shading: headFill, children: [this.P(this.T(this.key ? 'Correct match' : 'Your answer', { bold: true }))] })
          ]
        })
      );
      for (const m of q.matches) {
        rows.push(
          new D_.TableRow({
            children: [
              new D_.TableCell({ width: { size: 4680, type: D_.WidthType.DXA }, margins: cellPad, children: [this.P(this.T(CQE.htmlToText(m.promptHtml)))] }),
              new D_.TableCell({ width: { size: 4680, type: D_.WidthType.DXA }, margins: cellPad, children: [this.P(this.T(this.key ? m.answer : '', this.key ? { bold: true, color: COLOR.correct } : {}))] })
            ]
          })
        );
      }
      out.push(new D_.Table({ rows, width: { size: 9360, type: D_.WidthType.DXA }, columnWidths: [4680, 4680] }));
      out.push(this.P([], { spacing: { after: 60 } }));
      if (!this.key) {
        out.push(this.answerLine('Choose from: ', q.matchOptions.join('  |  '), true));
      } else if (q.matchDistractors && q.matchDistractors.length) {
        out.push(this.answerLine('Distractors (unused options): ', q.matchDistractors.join('  |  '), true));
      }
      summary.answer = q.matches.map((m) => truncate(CQE.htmlToText(m.promptHtml), 30) + ' → ' + m.answer).join('; ');
      return out;
    }

    blankBlocks(q, summary) {
      const out = [];
      const parts = [];
      for (const b of q.blanks) {
        if (this.key) {
          const ans = b.answers.length ? b.answers.join('  |  ') : '(none listed)';
          out.push(this.answerLine('[' + b.name + ']  ', ans));
          parts.push(b.name + ': ' + b.answers.join(' | '));
        } else if (q.type === 'multiple_dropdowns_question') {
          out.push(this.answerLine('[' + b.name + '] options: ', b.options.map((o) => o.text).join('  |  '), true));
        } else {
          out.push(this.answerLine('[' + b.name + ']  ', '______________________', true));
        }
      }
      summary.answer = parts.join('; ');
      out.push(this.P([], { spacing: { after: 60 } }));
      return out;
    }

    formulaBlocks(q, summary) {
      const out = [];
      const f = q.formula;
      if (!this.key) {
        out.push(this.writeLine());
        return out;
      }
      if (!f) {
        out.push(this.note('Formula details were not found in the export.'));
        return out;
      }
      if (f.formulas.length) out.push(this.answerLine('Formula: ', f.formulas.join('  ;  ')));
      if (f.vars.length) out.push(this.answerLine('Variables: ', f.vars.map((v) => v.name + ' from ' + v.min + ' to ' + v.max + (v.scale ? ' (' + v.scale + ' decimals)' : '')).join('; '), true));
      if (f.tolerance) out.push(this.answerLine('Answer tolerance: ', '± ' + f.tolerance, true));
      if (f.examples.length) {
        out.push(this.answerLine('Example sets: ', '', true));
        for (const ex of f.examples) out.push(this.answerLine('    ', ex.values + '  →  ' + ex.answer, true));
      }
      summary.answer = f.formulas.join(' ; ');
      return out;
    }

    answerLine(label, value, plain) {
      return this.P([this.T(label, { bold: true, color: plain ? '404040' : COLOR.correct }), this.T(value, plain ? { color: '404040' } : { bold: true, color: COLOR.correct })], { indent: { left: 360 }, spacing: { after: 60 } });
    }
    hint(text) {
      return this.P(this.T(text, { italics: true, color: COLOR.muted, size: 20 }), { indent: { left: 360 }, spacing: { after: 80 } });
    }
    writeLine() {
      return this.P(this.T('Answer: ______________________________________________', { color: '404040' }), { indent: { left: 360 }, spacing: { before: 120, after: 120 } });
    }
    responseBox(height) {
      const D_ = D();
      return new D_.Table({
        width: { size: 9000, type: D_.WidthType.DXA },
        columnWidths: [9000],
        indent: { size: 360, type: D_.WidthType.DXA },
        rows: [
          new D_.TableRow({
            height: { value: height, rule: D_.HeightRule.ATLEAST },
            children: [new D_.TableCell({ width: { size: 9000, type: D_.WidthType.DXA }, children: [this.P(this.T('Response:', { color: '9E9E9E', size: 18 }))] })]
          })
        ]
      });
    }

    async feedbackBlocks(label, html, indent) {
      const D_ = D();
      const left = (indent || 0) + 0;
      return this.conv.convert(html, {
        indent: left + 240,
        prefixRuns: [this.T(label, { bold: true, italics: true, color: COLOR.feedback, size: 20 })],
        run: { italics: true, color: COLOR.feedback, size: 20 },
        paragraph: {
          spacing: { after: 60 },
          border: { left: { style: D_.BorderStyle.SINGLE, size: 8, color: 'BFBFBF', space: 6 } }
        }
      });
    }

    answerTable() {
      const D_ = D();
      const out = [];
      out.push(new D_.Paragraph({ heading: D_.HeadingLevel.HEADING_2, pageBreakBefore: true, children: [this.T('Answer key summary')] }));
      const widths = [900, 2200, 6260];
      const pad = { top: 40, bottom: 40, left: 80, right: 80 };
      const fill = { type: D_.ShadingType.CLEAR, color: 'auto', fill: 'F2F2F2' };
      const row = (vals, head) =>
        new D_.TableRow({
          tableHeader: !!head,
          children: vals.map(
            (v, i) =>
              new D_.TableCell({
                width: { size: widths[i], type: D_.WidthType.DXA },
                margins: pad,
                shading: head ? fill : undefined,
                children: [this.P(this.T(String(v || '—'), { bold: !!head, size: 20 }), { spacing: { after: 0 } })]
              })
          )
        });
      const rows = [row(['#', 'Type', 'Correct answer'], true)];
      for (const s of this.summary) rows.push(row([s.n, s.type, s.answer]));
      out.push(new D_.Table({ rows, width: { size: 9360, type: D_.WidthType.DXA }, columnWidths: widths }));
      return out;
    }
  }

  /**
   * Shared Word document shell: fonts, heading styles, page size, footer.
   * opts: { title, courseTitle, landscape }
   */
  CQE.makeDocument = function (children, opts) {
    const D_ = D();
    opts = opts || {};
    const land = !!opts.landscape;
    const textWidth = land ? 12960 : 9360;
    const footerText = (opts.courseTitle ? opts.courseTitle + '  •  ' : '') + (opts.title || '');
    const fr = { size: 16, color: '808080' };
    return new D_.Document({
      creator: 'Canvas Course Exporter',
      title: opts.title || '',
      description: 'Exported from a Canvas course export (.imscc)',
      styles: {
        default: {
          document: { run: { font: 'Calibri', size: 22 } },
          heading1: { run: { font: 'Calibri', size: 32, bold: true, color: COLOR.heading }, paragraph: { spacing: { before: 0, after: 80 } } },
          heading2: { run: { font: 'Calibri', size: 26, bold: true, color: COLOR.heading }, paragraph: { spacing: { before: 240, after: 80 } } },
          heading3: { run: { font: 'Calibri', size: 22, bold: true, color: COLOR.heading }, paragraph: { spacing: { before: 200, after: 60 } } }
        }
      },
      sections: [
        {
          properties: {
            page: {
              // docx.js swaps width/height itself when orientation is landscape
              size: { width: 12240, height: 15840, orientation: land ? D_.PageOrientation.LANDSCAPE : D_.PageOrientation.PORTRAIT },
              margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
            }
          },
          footers: {
            default: new D_.Footer({
              children: [
                new D_.Paragraph({
                  tabStops: [{ type: D_.TabStopType.RIGHT, position: textWidth }],
                  children: [
                    new D_.TextRun(Object.assign({ text: footerText.slice(0, 150) }, fr)),
                    new D_.TextRun(Object.assign({ children: [new D_.Tab(), 'Page '] }, fr)),
                    new D_.TextRun(Object.assign({ children: [D_.PageNumber.CURRENT] }, fr)),
                    new D_.TextRun(Object.assign({ text: ' of ' }, fr)),
                    new D_.TextRun(Object.assign({ children: [D_.PageNumber.TOTAL_PAGES] }, fr))
                  ]
                })
              ]
            })
          },
          children
        }
      ]
    });
  };
  CQE.COLOR = COLOR;

  function plural(n, word) {
    return n + ' ' + word + (n === 1 ? '' : 's');
  }
  function fmtPts(p) {
    const v = +(+p).toFixed(2);
    return v + (v === 1 ? ' pt' : ' pts');
  }
  function truncate(s, n) {
    s = s || '';
    return s.length > n ? s.slice(0, n - 1) + '…' : s;
  }

  CQE.WordBuilder = WordBuilder;
})(typeof window !== 'undefined' ? window : globalThis);
