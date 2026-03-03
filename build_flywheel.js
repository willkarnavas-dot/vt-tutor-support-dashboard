const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
        LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition } = require('docx');
const fs = require('fs');

// ── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  dark_blue:  '1F4E79',
  mid_blue:   '2E75B6',
  light_blue: 'D6E4F0',
  pale_blue:  'EBF3FB',
  orange:     'E07B39',
  green:      '1E8449',
  amber:      'D68910',
  slate:      '475569',
  grey_light: 'F2F5F7',
  grey_mid:   '7F8C8D',
  grey_dark:  '2C3E50',
  white:      'FFFFFF',
};

const noBorders = () => ({
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
});

const thinBorder = (color = 'BDC3C7') => ({
  top:    { style: BorderStyle.SINGLE, size: 4, color },
  bottom: { style: BorderStyle.SINGLE, size: 4, color },
  left:   { style: BorderStyle.SINGLE, size: 4, color },
  right:  { style: BorderStyle.SINGLE, size: 4, color },
});

const leftBar = (color) => ({
  ...noBorders(),
  left: { style: BorderStyle.SINGLE, size: 16, color },
});

const run = (text, opts = {}) => new TextRun({
  text, font: 'Arial',
  bold:    opts.bold    || false,
  italic:  opts.italic  || false,
  size:    opts.size    || 20,
  color:   opts.color   || C.grey_dark,
  break:   opts.break   || 0,
});

const p = (children, opts = {}) => new Paragraph({
  spacing: { before: opts.before || 0, after: opts.after || 0 },
  alignment: opts.align || AlignmentType.LEFT,
  numbering: opts.numbering || undefined,
  children,
});

const spacer = (pt = 120) => p([run('')], { before: 0, after: pt });

const cell = (children, opts = {}) => new TableCell({
  borders:  opts.borders || noBorders(),
  shading:  { fill: opts.fill || C.white, type: ShadingType.CLEAR },
  width:    { size: opts.width || 9360, type: WidthType.DXA },
  margins:  opts.margins || { top: 80, bottom: 80, left: 120, right: 120 },
  columnSpan: opts.span || 1,
  children,
});

// ── SECTION HEADER ───────────────────────────────────────────────────────────
const sectionHeader = (text) => [
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [
      cell([p([run(text, { bold: true, size: 17, color: C.white })])],
        { fill: C.dark_blue, borders: noBorders(), margins: { top: 100, bottom: 100, left: 160, right: 160 } })
    ]})]
  }),
  spacer(80),
];

// ── NODE CARD ─────────────────────────────────────────────────────────────────
const nodeCard = (emoji, title, day, dayColor, bullets, accentColor) => {
  const colWidths = [800, 8560];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      // Header row
      new TableRow({ children: [
        cell([p([run(emoji, { size: 24 })], { align: AlignmentType.CENTER })],
          { fill: C.light_blue, borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 20, color: accentColor } },
            width: 800, margins: { top: 100, bottom: 100, left: 80, right: 80 } }),
        cell([
          p([
            run(title, { bold: true, size: 22, color: C.dark_blue }),
            run('  ', { size: 18 }),
            run(day, { size: 17, color: dayColor, italic: true }),
          ], { after: 0 }),
        ], { fill: C.pale_blue, borders: { ...noBorders(), right: noBorders().right }, width: 8560,
             margins: { top: 100, bottom: 100, left: 160, right: 120 } }),
      ]}),
      // Bullets row
      new TableRow({ children: [
        cell([p([run('')])],
          { fill: C.white, borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 20, color: accentColor } },
            width: 800 }),
        cell(
          bullets.map(b => p([run(b, { size: 18, color: C.grey_dark })],
            { numbering: { reference: 'bullets', level: 0 }, before: 0, after: 40 })),
          { fill: C.white, width: 8560, margins: { top: 80, bottom: 100, left: 160, right: 120 },
            borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5EDF4' } } }
        ),
      ]}),
    ]
  });
};

// ── SEQUENCE ROW ─────────────────────────────────────────────────────────────
const sequenceTable = () => {
  const steps = [
    { day: 'MON', label: 'CSO/MP\nSync',                    color: C.orange  },
    { day: 'TUE', label: 'Internal Agent\nComms Published',  color: C.green   },
    { day: 'WED', label: 'External Tutor\nComms Published',  color: C.amber   },
    { day: 'THU', label: 'Update KB with\nNew Deployments',  color: C.mid_blue },
    { day: 'THU', label: 'Dashboard\nRefresh',               color: C.mid_blue },
  ];
  const w = Math.floor(9360 / steps.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: Array(steps.length).fill(w),
    rows: [
      // Day labels
      new TableRow({ children: steps.map((s, i) =>
        cell([p([run(s.day, { bold: true, size: 16, color: C.white })], { align: AlignmentType.CENTER })],
          { fill: s.color, borders: noBorders(), width: w,
            margins: { top: 80, bottom: 60, left: 40, right: 40 } })
      )}),
      // Step labels
      new TableRow({ children: steps.map((s, i) => {
        const lines = s.label.split('\n');
        return cell(
          lines.map(l => p([run(l, { size: 16, color: s.color, bold: true })], { align: AlignmentType.CENTER, after: 0 })),
          { fill: 'EEF4FB', borders: { ...noBorders(),
              bottom: { style: BorderStyle.SINGLE, size: 6, color: s.color } },
            width: w, margins: { top: 80, bottom: 80, left: 40, right: 40 } }
        );
      })}),
    ]
  });
};

// ── DOCUMENT ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2013',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } },
                 run: { font: 'Arial', size: 18, color: C.grey_mid } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [7000, 2360],
          rows: [new TableRow({ children: [
            cell([
              p([run('Voice of Customer Flywheel', { bold: true, size: 22, color: C.dark_blue })]),
            ], { fill: C.white, borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } }, width: 7000 }),
            cell([
              p([run('March 2026', { bold: true, size: 20, color: C.dark_blue })], { align: AlignmentType.RIGHT, after: 0 }),
              p([run('Tutor Support / Marketplace', { size: 16, color: C.grey_mid })], { align: AlignmentType.RIGHT }),
            ], { fill: C.white, borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } }, width: 2360 }),
          ]})]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          spacing: { before: 60, after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'D0D7DE' } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run('Tutor Support Operations  |  Internal Use Only', { size: 16, color: C.grey_mid }),
            new TextRun({ text: '\t', font: 'Arial', size: 16 }),
            run('Page ', { size: 16, color: C.grey_mid }),
            new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: C.grey_mid }),
            run(' of 2', { size: 16, color: C.grey_mid }),
          ]
        })
      ]})
    },
    children: [

      // ── PAGE 1 ─────────────────────────────────────────────────────────────
      spacer(40),

      // Purpose statement
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [
          cell([
            p([run('Purpose', { bold: true, size: 17, color: C.dark_blue })], { after: 40 }),
            p([run('A shared operating rhythm between Tutor Support and Marketplace that turns daily tutor contact data into faster decisions, better-informed releases, and a smarter bot. When both teams are aligned on what tutors are experiencing, we move faster, communicate more effectively, and build toward a better product together.', { size: 17, color: C.grey_dark })]),
          ], { fill: C.pale_blue, borders: leftBar(C.mid_blue), width: 9360,
               margins: { top: 80, bottom: 80, left: 160, right: 160 } })
        ]})]
      }),

      spacer(100),

      // Weekly sequence
      ...sectionHeader('WEEKLY RHYTHM'),
      sequenceTable(),

      spacer(100),

      // Two-column: What we push up / What comes down
      ...sectionHeader('WEEKLY CSO/MP SYNC  |  MONDAY'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4560, 200, 4600],
        rows: [new TableRow({ children: [
          cell([
            p([run('CSO brings to Marketplace', { bold: true, size: 17, color: C.orange })], { after: 60 }),
            ...['New defects crossing the 20-contact/week threshold',
                'Tutor verbatims with Platform IDs for open tickets',
                'Volume trends that signal urgency or stabilization',
                'Emerging patterns before they become spikes',
            ].map(b => p([run(b, { size: 17 })], { numbering: { reference: 'bullets', level: 0 }, after: 30 })),
          ], { fill: 'FEF6F0', borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 16, color: C.orange } },
               width: 4560, margins: { top: 80, bottom: 90, left: 140, right: 100 } }),
          cell([p([run('')])], { fill: C.white, borders: noBorders(), width: 200 }),
          cell([
            p([run('Marketplace brings to CSO', { bold: true, size: 17, color: C.mid_blue })], { after: 60 }),
            ...['Deployments scheduled for the week — input for KB update Thursday',
                'Workflow changes so agents are briefed Tuesday before comms go out',
                'Engineering progress on open defect tickets',
                'Anything requiring tutor-facing communication Wednesday',
            ].map(b => p([run(b, { size: 17 })], { numbering: { reference: 'bullets', level: 0 }, after: 30 })),
          ], { fill: C.pale_blue, borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 16, color: C.mid_blue } },
               width: 4600, margins: { top: 80, bottom: 90, left: 140, right: 100 } }),
        ]})]
      }),

      spacer(100),

      // Communications flow
      ...sectionHeader('COMMUNICATIONS FLOW'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 200, 3000, 200, 2960],
        rows: [new TableRow({ children: [
          cell([
            p([run('1.  Internal Agent Comms Published', { bold: true, size: 17, color: C.green })], { after: 40 }),
            p([run('Tuesday — before any external comms go out', { size: 15, italic: true, color: C.grey_mid })], { after: 60 }),
            ...['Feature deployments scheduled for the week — what changed and how it works',
                'Bugs confirmed fixed — agents stop escalating, update tutor guidance',
                'Known bugs pending fix — agent scripts for how to handle inbound contacts',
                'Drafted from Monday CSO/MP sync output, owned by Tutor Support',
            ].map(b => p([run(b, { size: 16 })], { numbering: { reference: 'bullets', level: 0 }, after: 30 })),
          ], { fill: 'EAF5EE', borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 16, color: C.green } },
               width: 3000, margins: { top: 80, bottom: 90, left: 140, right: 100 } }),
          cell([p([run('')])], { fill: C.white, borders: noBorders(), width: 200 }),
          cell([
            p([run('2.  External Tutor Comms Published', { bold: true, size: 17, color: C.amber })], { after: 40 }),
            p([run('Wednesday (bi-weekly) — after agents are briefed', { size: 15, italic: true, color: C.grey_mid })], { after: 60 }),
            ...['Feature deployments this week — what changed and what tutors need to know',
                'Bugs confirmed fixed — tutors no longer need to contact support for these',
                'Known bugs pending fix — tutors know their issue is recognized and being worked on',
                'Mirrors Internal Agent Comms, written in plain language for tutors',
            ].map(b => p([run(b, { size: 16 })], { numbering: { reference: 'bullets', level: 0 }, after: 30 })),
          ], { fill: 'FDF6E3', borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 16, color: C.amber } },
               width: 3000, margins: { top: 80, bottom: 90, left: 140, right: 100 } }),
          cell([p([run('')])], { fill: C.white, borders: noBorders(), width: 200 }),
          cell([
            p([run('3.  Update KB with New Deployments', { bold: true, size: 17, color: C.mid_blue })], { after: 40 }),
            p([run('Thursday — alongside dashboard refresh', { size: 15, italic: true, color: C.grey_mid })], { after: 60 }),
            ...['Fix stale instructions (no-show, invoicing, cancel)',
                'Update KB articles for changed workflows',
                'New deployments added same week received',
                'Flag bot team with specific article IDs',
            ].map(b => p([run(b, { size: 16 })], { numbering: { reference: 'bullets', level: 0 }, after: 30 })),
          ], { fill: C.pale_blue, borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 16, color: C.mid_blue } },
               width: 2960, margins: { top: 80, bottom: 90, left: 140, right: 100 } }),
        ]})]
      }),

      // Page break
      new Paragraph({ children: [new PageBreak()] }),

      // ── PAGE 2 ─────────────────────────────────────────────────────────────
      spacer(40),

      // Success metrics
      ...sectionHeader('SHARED SUCCESS METRICS'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2240, 2240, 2440, 2440],
        rows: [new TableRow({ children: [
          cell([
            p([run('Bot Containment Rate', { bold: true, size: 16, color: C.dark_blue })], { after: 40 }),
            p([run('% of tutor contacts resolved by the bot without agent handoff. Every KB update and bug fix moves this number. Target: 45%+. YTD: 41.2%.', { size: 16 })]),
          ], { fill: C.pale_blue, borders: { ...noBorders(), top: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } },
               width: 2240, margins: { top: 80, bottom: 80, left: 100, right: 100 } }),
          cell([
            p([run('Contact Reason Reduction', { bold: true, size: 16, color: C.dark_blue })], { after: 40 }),
            p([run('Week-over-week decline in contacts for a category after a bug fix or KB update. Validates the fix actually landed for tutors.', { size: 16 })]),
          ], { fill: C.pale_blue, borders: { ...noBorders(), top: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } },
               width: 2240, margins: { top: 80, bottom: 80, left: 100, right: 100 } }),
          cell([
            p([run('Agent Comms Aware Coverage', { bold: true, size: 16, color: C.dark_blue })], { after: 40 }),
            p([run('% of agents who received Internal Agent Comms before External Tutor Comms went out. Target: 100% every publish week.', { size: 16 })]),
          ], { fill: C.pale_blue, borders: { ...noBorders(), top: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } },
               width: 2440, margins: { top: 80, bottom: 80, left: 100, right: 100 } }),
          cell([
            p([run('Ticket Completion Time', { bold: true, size: 16, color: C.dark_blue })], { after: 40 }),
            p([run('Time from a defect crossing the 20-contact threshold to the ticket being closed as resolved.', { size: 16 })]),
          ], { fill: C.pale_blue, borders: { ...noBorders(), top: { style: BorderStyle.SINGLE, size: 8, color: C.mid_blue } },
               width: 2440, margins: { top: 80, bottom: 80, left: 100, right: 100 } }),
        ]})]
      }),

      spacer(100),

      // Open Questions
      ...sectionHeader('OPEN QUESTIONS FOR ALIGNMENT'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [440, 8920],
        rows: [
          ...([
            'Who will own Internal Agent Comms?',
            'Who will own External Tutor Comms?',
            'Who will own updating the Knowledge Base?',
            'Should we designate specific deployment day(s) each week?',
            'Should bug fixes have fewer deployment restrictions than new feature releases?',
            'Should the Tutor Support Daily Digest be distributed at a daily or weekly level?',
          ]).map((question, idx) =>
            new TableRow({ children: [
              cell([p([run(String(idx+1), { bold: true, size: 16, color: C.mid_blue })], { align: AlignmentType.CENTER })],
                { fill: idx%2===0 ? C.white : C.grey_light, borders: noBorders(),
                  width: 440, margins: { top: 60, bottom: 60, left: 80, right: 40 } }),
              cell([p([run(question, { size: 16, color: C.grey_dark })])],
                { fill: idx%2===0 ? C.white : C.grey_light,
                  borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E8EDF2' } },
                  width: 8920, margins: { top: 60, bottom: 60, left: 80, right: 120 } }),
            ]})
          ),
        ]
      }),

      spacer(100),

      // FAQ section
      ...sectionHeader('FREQUENTLY ASKED QUESTIONS'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
          ...([
            {
              q: 'How much contact volume does it take for CSO to file a ticket to Marketplace? What about a Watch List?',
              a: '20+ contacts on the same issue in a single week triggers a formal Jira ticket — filed with a definition, tutor verbatims, and Platform IDs, surfaced at the Monday CSO/MP sync. 5-19 contacts per week goes on the Watch List, monitored for two consecutive weeks before escalating. Under 5 contacts per week is logged but no action is taken unless the pattern persists.',
            },
            {
              q: 'How does the KB get updated with new deployments?',
              a: 'Every Thursday, after the deployment schedule is confirmed from Monday\'s sync, KB articles are reviewed and updated to reflect new workflows, removed features, and changed processes.',
            },
          ]).map((faq, idx) => new TableRow({ children: [
            cell([
              p([run(faq.q, { bold: true, size: 17, color: C.dark_blue })], { after: 40 }),
              p([run(faq.a, { size: 16, color: C.grey_dark })]),
            ], {
              fill: idx%2===0 ? C.white : C.pale_blue,
              borders: { ...noBorders(),
                left:   { style: BorderStyle.SINGLE, size: 12, color: idx%2===0 ? C.mid_blue : C.light_blue },
                bottom: { style: BorderStyle.SINGLE, size: 4,  color: 'E8EDF2' } },
              width: 9360,
              margins: { top: 90, bottom: 90, left: 160, right: 160 },
            }),
          ]})),
        ]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/flywheel_2pager.docx', buf);
  console.log('OK');
});
