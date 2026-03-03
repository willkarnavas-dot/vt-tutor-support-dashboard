const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  LevelFormat, PageNumber, Header, Footer, TabStopType, TabStopPosition,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ── DATA (injected by Python) ─────────────────────────────────────────────────
const DATA = JSON.parse(fs.readFileSync('/home/claude/report_data.json', 'utf8'));

// ── COLORS ────────────────────────────────────────────────────────────────────
const C = {
  blue_dark:  '1F4E79',
  blue_mid:   '2E75B6',
  blue_light: 'D6E4F0',
  blue_pale:  'EBF3FB',
  white:      'FFFFFF',
  grey_dark:  '2C3E50',
  grey_mid:   '5D6D7E',
  grey_light: 'F2F5F7',
  red:        'C0392B',
  green:      '1E8449',
  amber:      'D68910',
  border:     'BDC3C7',
};

// ── HELPERS ───────────────────────────────────────────────────────────────────
const border = (color = C.border) => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (color = C.border) => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });
const noBorders = () => ({ top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } });
const cell = (children, opts = {}) => new TableCell({
  borders: opts.borders || borders(),
  width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  margins: { top: 80, bottom: 80, left: 140, right: 140 },
  verticalAlign: opts.vAlign || 'top',
  columnSpan: opts.span,
  children,
});
const p = (runs, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.spaceBefore || 0, after: opts.spaceAfter || 60 },
  border: opts.border,
  children: Array.isArray(runs) ? runs : [runs],
});
const run = (text, opts = {}) => new TextRun({
  text,
  bold: opts.bold,
  italics: opts.italic,
  color: opts.color,
  size: opts.size || 20,
  font: 'Arial',
});
const spacer = (pt = 120) => new Paragraph({ spacing: { before: 0, after: pt }, children: [new TextRun({ text: '' })] });
const divider = (color = C.blue_mid) => new Paragraph({
  spacing: { before: 60, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
  children: [new TextRun({ text: '' })],
});

// Trend arrow + color
const trend = (today, prev) => {
  if (!prev) return { arrow: '', color: C.grey_mid };
  const pct = ((today - prev) / prev) * 100;
  if (Math.abs(pct) < 5) return { arrow: ' (flat)', color: C.grey_mid };
  if (pct > 0) return { arrow: ` (+${Math.round(pct)}% vs yesterday)`, color: C.red };
  return { arrow: ` (${Math.round(pct)}% vs yesterday)`, color: C.green };
};

const spike = (today, avg7) => {
  if (!avg7) return null;
  const pct = ((today - avg7) / avg7) * 100;
  if (pct >= 25) return `SPIKE +${Math.round(pct)}% vs 7d avg`;
  if (pct <= -25) return `DOWN ${Math.round(pct)}% vs 7d avg`;
  return null;
};

// ── SECTION HEADER ────────────────────────────────────────────────────────────
const sectionHeader = (text) => [
  spacer(80),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    shading: { fill: C.blue_dark, type: ShadingType.CLEAR },
    children: [new TextRun({ text: '  ' + text, bold: true, size: 22, color: C.white, font: 'Arial' })],
  }),
];

// ── CALLOUT BOX ───────────────────────────────────────────────────────────────
const calloutBox = (label, value, sub, color = C.blue_mid) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [
      cell([
        p([run(label, { size: 18, color: C.grey_mid })], { spaceAfter: 20 }),
        p([run(value, { bold: true, size: 40, color })], { spaceAfter: 20 }),
        p([run(sub, { size: 18, color: C.grey_mid, italic: true })]),
      ], { fill: C.blue_pale, borders: { top: { style: BorderStyle.SINGLE, size: 12, color }, bottom: border(), left: border(), right: border() }, width: 9360 })
    ]})],
  });

// ── METRIC ROW TABLE ──────────────────────────────────────────────────────────
const metricTable = (metrics) => {
  const colW = Math.floor(9360 / metrics.length);
  const cols = Array(metrics.length).fill(colW);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: cols,
    rows: [new TableRow({ children: metrics.map((m, i) =>
      cell([
        p([run(m.label, { size: 17, color: C.grey_mid })], { spaceAfter: 16 }),
        p([run(m.value, { bold: true, size: 32, color: m.color || C.blue_dark })], { spaceAfter: 4 }),
        ...(m.sub2 ? [p([run(m.sub2, { size: 17, color: C.grey_dark, bold: true })], { spaceAfter: 8 })] : []),
        p([run(m.sub || '', { size: 17, color: m.subColor || C.grey_mid, italic: true })]),
      ], { fill: i % 2 === 0 ? C.blue_pale : C.white, width: colW })
    )})]
  });
};

// ── TICKET TABLE ──────────────────────────────────────────────────────────────
const ticketTableRows = (tickets, colWidths) => {
  const headerRow = new TableRow({
    tableHeader: true,
    children: ['Ticket', 'Issue', 'Today', '14-Day', 'This Week', 'Status'].map((h, i) =>
      cell([p([run(h, { bold: true, size: 17, color: C.white })])],
        { fill: C.blue_dark, width: colWidths[i], borders: borders(C.blue_dark) })
    )
  });
  const dataRows = tickets.map((t, idx) => {
    const tr = trend(t.today, t.yesterday);
    const sp = spike(t.today, t.avg7);
    return new TableRow({ children: [
      cell([p([new TextRun({ text: t.ticket, bold: true, color: C.blue_mid, size: 19, font: 'Arial' })])],
        { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[0] }),
      cell([p([run(t.issue, { size: 18 })])],
        { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[1] }),
      cell([
        p([run(String(t.today), { bold: true, size: 22, color: t.today >= 20 ? C.red : C.grey_dark })]),
        t.yesterday !== undefined ? p([run(tr.arrow, { size: 16, color: tr.color, italic: true })], { spaceBefore: 0 }) : null,
        sp ? p([run(sp, { size: 16, color: C.red, bold: true })]) : null,
      ].filter(Boolean), { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[2] }),
      cell([p([run(String(t.contacts14), { size: 19 })])],
        { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[3] }),
      cell([p([run(String(t.contactsWk), { size: 19 })])],
        { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[4] }),
      cell([
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [new TextRun({
            text: t.status,
            size: 16, color: C.grey_mid, font: 'Arial', italics: true
          })]
        })
      ], { fill: idx % 2 === 0 ? C.white : C.grey_light, width: colWidths[5] }),
    ]});
  });
  return [headerRow, ...dataRows];
};

// ── QUOTE BOX ─────────────────────────────────────────────────────────────────
const quoteBox = (ticket, text, time, email) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({ children: [
    cell([
      p([run(ticket + '  ', { bold: true, size: 17, color: C.blue_mid }), run(time, { size: 16, color: C.grey_mid, italic: true })], { spaceAfter: 12 }),
      p([run('\u201C' + text + '\u201D', { size: 18, italic: true, color: C.grey_dark })], { spaceAfter: 8 }),
      ...(email ? [p([run(email, { size: 15, color: C.grey_mid, italic: true })])] : []),
    ], {
      fill: C.white,
      borders: {
        top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 16, color: C.blue_mid }
      },
      width: 9360
    })
  ]})]
});

// ── WATCH LIST TABLE ──────────────────────────────────────────────────────────
const watchTable = (items) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3200, 1800, 1800, 2560],
  rows: [
    new TableRow({ tableHeader: true, children: ['Issue', '14-Day', 'This Week', 'Threshold'].map((h, i) =>
      cell([p([run(h, { bold: true, size: 17, color: C.white })])],
        { fill: C.grey_mid, width: [3200,1800,1800,2560][i], borders: borders(C.grey_mid) })
    )}),
    ...items.map((item, idx) => new TableRow({ children: [
      cell([p([run(item.issue, { size: 18 })])], { fill: idx%2===0?C.white:C.grey_light, width: 3200 }),
      cell([p([run(String(item.contacts14), { size: 18 })])], { fill: idx%2===0?C.white:C.grey_light, width: 1800 }),
      cell([p([run(String(item.contactsWk), { size: 18 })])], { fill: idx%2===0?C.white:C.grey_light, width: 1800 }),
      cell([p([run(item.threshold, { size: 17, color: C.grey_mid, italic: true })])], { fill: idx%2===0?C.white:C.grey_light, width: 2560 }),
    ]}))
  ]
});

// ── ASSEMBLE DOCUMENT ─────────────────────────────────────────────────────────
const D = DATA;
const ticketColWidths = [720, 3200, 960, 1120, 1120, 1240]; // sum = 9360

const doc = new Document({
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 500, hanging: 280 } }, run: { font: 'Arial', size: 20 } } }] }
  ]},
  styles: {
    default: { document: { run: { font: 'Arial', size: 20, color: C.grey_dark } } },
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
          width: { size: 10080, type: WidthType.DXA },
          columnWidths: [7200, 2880],
          rows: [new TableRow({ children: [
            cell([
              p([run('TUTOR SUPPORT', { bold: true, size: 16, color: C.blue_mid })], { spaceAfter: 0 }),
              p([run('Daily Operations Digest', { size: 20, bold: true, color: C.blue_dark })]),
            ], { fill: C.white, borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 10, color: C.blue_mid } }, width: 7200 }),
            cell([
              p([run(D.date, { bold: true, size: 20, color: C.blue_dark })], { align: AlignmentType.RIGHT }),
              p([run('Wk' + (D.weekLabel || '').replace(/.*Wk/,''), { size: 17, color: C.grey_mid })], { align: AlignmentType.RIGHT }),
            ], { fill: C.white, borders: { ...noBorders(), bottom: { style: BorderStyle.SINGLE, size: 10, color: C.blue_mid } }, width: 2880 }),
          ]})],
        }),
        spacer(60),
      ]})
    },
    footers: {
      default: new Footer({ children: [
        divider(C.border),
        new Paragraph({
          spacing: { before: 60, after: 0 },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run('Varsity Tutors  |  Tutor Support Operations  |  Will Karnavas', { size: 16, color: C.grey_mid }),
            new TextRun({ text: '\tPage ', size: 16, color: C.grey_mid, font: 'Arial' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.grey_mid, font: 'Arial' }),
            new TextRun({ text: ' of ', size: 16, color: C.grey_mid, font: 'Arial' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.grey_mid, font: 'Arial' }),
          ]
        })
      ]})
    },
    children: [




      // ── LEADERSHIP SUMMARY ───────────────────────────────────────────────────
      ...sectionHeader('LEADERSHIP SUMMARY'),
      spacer(80),
      ...D.leadershipBullets.map(b =>
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          spacing: { before: 0, after: 80 },
          children: [run(b.text, { bold: b.bold, color: b.color || C.grey_dark })],
        })
      ),
      spacer(100),

      // ── OVERVIEW METRICS ────────────────────────────────────────────────────
      ...sectionHeader('OVERVIEW'),
      spacer(80),
      metricTable([
        { label: 'Total Contacts', value: D.totalContacts.toLocaleString(),
          sub: D.totalVsAvg, color: C.blue_dark },
        { label: 'Resolved (Contained)', value: D.resolvedPct + '%',
          sub2: String(D.resolvedCount || '') + ' contacts',
          sub: D.resolvedVsPrev, color: parseFloat(D.resolvedPct) >= 38 ? C.green : C.amber },
        { label: 'Handed Off to Agents', value: D.handedOffPct + '%',
          sub2: String(D.handedOffCount || '') + ' contacts',
          sub: D.handedOffVsPrev, color: parseFloat(D.handedOffPct) >= 55 ? C.red : C.amber },
        { label: 'Abandoned', value: D.abandonedPct + '%',
          sub2: String(D.abandonedCount || '') + ' contacts',
          sub: D.abandonedVsPrev, color: parseFloat(D.abandonedPct) >= 11 ? C.red : C.grey_mid },
      ]),
      spacer(100),

      // Top contact drivers
      ...sectionHeader('TOP 10 CONTACT DRIVERS  |  ' + D.date),
      spacer(80),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [600, 4960, 1600, 2200],
        rows: [
          new TableRow({ tableHeader: true, children:
            ['#', 'Category', 'Contacts', '% of Volume'].map((h, i) =>
              cell([p([run(h, { bold: true, size: 17, color: C.white })])],
                { fill: C.blue_dark, width: [600,4960,1600,2200][i], borders: borders(C.blue_dark) })
          )}),
          ...D.topDrivers.map((d, idx) => new TableRow({ children: [
            cell([p([run(String(idx+1), { bold: true, size: 18, color: C.blue_mid })])],
              { fill: idx%2===0?C.white:C.grey_light, width: 600 }),
            cell([p([run(d.name, { size: 18 })])],
              { fill: idx%2===0?C.white:C.grey_light, width: 4960 }),
            cell([p([run(String(d.count), { bold: true, size: 20, color: d.count >= 100 ? C.red : C.grey_dark })])],
              { fill: idx%2===0?C.white:C.grey_light, width: 1600 }),
            cell([p([run(d.pct, { size: 18, color: C.grey_mid })])],
              { fill: idx%2===0?C.white:C.grey_light, width: 2200 }),
          ]}))
        ]
      }),
      spacer(100),


      // ── CONTACT DISPOSITION BY CATEGORY ─────────────────────────────────────
      ...sectionHeader('CONTACT DISPOSITION BY CATEGORY'),
      spacer(80),
      (() => {
        // 5-col layout: Category | Total | Resolved (n / %) | Handed Off (n / %) | Abandoned (n / %)
        // 4200 + 760 + 1500 + 1500 + 1400 = 9360
        const W = [4200, 760, 1500, 1500, 1400];
        const cellM = { top: 70, bottom: 70, left: 120, right: 120 };
        const mkCell = (children, fill, w, align) => new TableCell({
          borders: borders(),
          width: { size: w, type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          margins: cellM,
          children,
        });
        const mkP = (text, opts = {}) => new Paragraph({
          spacing: { before: 0, after: 0 },
          alignment: opts.align || AlignmentType.LEFT,
          children: [new TextRun({ text, bold: opts.bold, color: opts.color || C.grey_dark, size: 17, font: 'Arial' })],
        });
        const headers = ['Category', 'Total', 'Resolved', 'Handed Off', 'Abandoned'];
        const subheads = ['', '', 'n  /  %', 'n  /  %', 'n  /  %'];
        const headerRow = new TableRow({ tableHeader: true, children: headers.map((h,i) =>
          mkCell([
            mkP(h, { bold: true, color: C.white }),
            ...(subheads[i] ? [mkP(subheads[i], { color: 'BDC3C7', size: 14 })] : []),
          ], C.blue_dark, W[i])
        )});
        const dataRows = (D.dispositionByCategory || []).map((d, idx) => {
          const fill     = idx % 2 === 0 ? C.white : C.grey_light;
          const resColor = parseInt(d.resolvedPct) >= 50 ? C.green : parseInt(d.resolvedPct) >= 30 ? C.amber : C.red;
          const hoColor  = parseInt(d.handedOffPct) >= 70 ? C.red : parseInt(d.handedOffPct) >= 50 ? C.amber : C.grey_mid;
          const abColor  = parseInt(d.abandonedPct) >= 15 ? C.red : C.grey_mid;
          return new TableRow({ children: [
            mkCell([mkP(d.category)], fill, W[0]),
            mkCell([mkP(String(d.total), { bold: true, color: C.blue_dark, align: AlignmentType.CENTER })], fill, W[1]),
            mkCell([mkP(d.resolved + '  /  ' + d.resolvedPct, { bold: true, color: resColor, align: AlignmentType.CENTER })], fill, W[2]),
            mkCell([mkP(d.handedOff + '  /  ' + d.handedOffPct, { bold: true, color: hoColor, align: AlignmentType.CENTER })], fill, W[3]),
            mkCell([mkP(d.abandoned + '  /  ' + d.abandonedPct, { color: abColor, align: AlignmentType.CENTER })], fill, W[4]),
          ]});
        });
        return new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: W,
          rows: [headerRow, ...dataRows],
        });
      })(),
      spacer(100),

      // ── ENGINEERING: INVOICING TICKETS ──────────────────────────────────────
      ...sectionHeader('INVOICING BUGS  |  ACTIVE TICKETS'),
      spacer(80),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: ticketColWidths,
        rows: ticketTableRows([...D.invoiceTickets].sort((a,b)=>b.today-a.today), ticketColWidths),
      }),
      spacer(100),

      // ── ENGINEERING: SCHEDULING TICKETS ─────────────────────────────────────
      ...sectionHeader('SCHEDULING PLATFORM  |  ACTIVE TICKETS'),
      spacer(80),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: ticketColWidths,
        rows: ticketTableRows([...D.schedTickets].sort((a,b)=>b.today-a.today), ticketColWidths),
      }),
      spacer(100),


      // ── TOP 3 CATEGORY EXAMPLES ──────────────────────────────────────────────
      ...sectionHeader('TOP CONTACT REASONS  |  5 EXAMPLES EACH'),
      spacer(80),
      ...(D.top3Examples || []).flatMap((cat, catIdx) => {
        const colW = [1200, 8160];
        const rows = [
          // Category header row
          new TableRow({ children: [
            new TableCell({
              columnSpan: 2,
              borders: borders(C.blue_mid),
              shading: { fill: C.blue_light, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 140, right: 140 },
              width: { size: 9360, type: WidthType.DXA },
              children: [new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({ text: String(catIdx + 1) + '.  ', bold: true, size: 20, color: C.blue_mid, font: 'Arial' }),
                  new TextRun({ text: cat.category, bold: true, size: 20, color: C.blue_dark, font: 'Arial' }),
                ]
              })]
            })
          ]}),
          // Example rows
          ...(cat.examples || []).map((ex, i) => {
            const parts = ex.match(/^(\d+:\d+) — "(.+)" \(([^)]*)\)$/) || ex.match(/^(\d+:\d+) — "(.+)"$/) || [];
            const time    = parts[1] || '';
            const quote   = parts[2] || ex;
            const pid     = parts[3] || '';
            return new TableRow({ children: [
              new TableCell({
                borders: { top: border(), bottom: border(), left: border(), right: { style: BorderStyle.NONE } },
                shading: { fill: i % 2 === 0 ? C.white : C.grey_light, type: ShadingType.CLEAR },
                margins: { top: 70, bottom: 70, left: 140, right: 80 },
                width: { size: colW[0], type: WidthType.DXA },
                children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: time, size: 17, color: C.grey_mid, font: 'Arial', italics: true })] })]
              }),
              new TableCell({
                borders: { top: border(), bottom: border(), right: border(), left: { style: BorderStyle.NONE } },
                shading: { fill: i % 2 === 0 ? C.white : C.grey_light, type: ShadingType.CLEAR },
                margins: { top: 70, bottom: 70, left: 80, right: 140 },
                width: { size: colW[1], type: WidthType.DXA },
                children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [
                  new TextRun({ text: '\u201C' + quote + '\u201D', size: 17, color: C.grey_dark, font: 'Arial', italics: true }),
                  ...(pid ? [new TextRun({ text: '  (' + pid + ')', size: 15, color: C.grey_mid, font: 'Arial' })] : []),
                ]})]
              }),
            ]});
          }),
        ];
        return [
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 8160], rows }),
          spacer(catIdx < (D.top3Examples.length - 1) ? 80 : 100),
        ];
      }),

      // ── VERBATIM ESCALATIONS ─────────────────────────────────────────────────
      ...sectionHeader('VERBATIM ESCALATIONS FROM TODAY'),
      spacer(80),
      ...D.quotes.flatMap((q, i) => [quoteBox(q.ticket, q.text, q.time, q.email), spacer(60)]),
      spacer(60),

      // ── WATCH LIST ───────────────────────────────────────────────────────────
      ...sectionHeader('WATCH LIST  |  BELOW TICKET THRESHOLD (< 20/WEEK)'),
      spacer(80),
      watchTable(D.watchList),
      spacer(80),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [
          cell([
            p([run('Ticket filing threshold: 20+ contacts per week for tech/feature/bug issues. Issues below this level are monitored but not escalated to engineering.', { size: 17, italic: true, color: C.grey_mid })]),
          ], { fill: C.grey_light, width: 9360 })
        ]})]
      }),
      spacer(100),

      // Data source note
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [
          cell([
            p([run('Data source: Assembled CSV export  |  ' + D.dataNote, { size: 16, italic: true, color: C.grey_mid })]),
          ], { fill: C.blue_pale, width: 9360, borders: { ...noBorders(), left: { style: BorderStyle.SINGLE, size: 12, color: C.blue_mid } } })
        ]})]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/daily_digest.docx', buf);
  console.log('OK');
}).catch(e => { console.error(e); process.exit(1); });
