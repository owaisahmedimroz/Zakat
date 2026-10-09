// Builds the thesis .docx from simple markup files.
// Usage: node build.js out.docx part1.txt [part2.txt ...]
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak, Header,
  PageNumber, BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType,
  FootnoteReferenceRun, HeadingLevel, TableOfContents, VerticalAlign,
  LineRuleType,
} = require('docx');

const UR = 'Alvi Nastaleeq';
const AR = 'Traditional Arabic';
const BODY = 28;      // 14pt
const FN = 22;        // 11pt
const GRAY = 'EDEDED';

const footnotes = {};
let fnCount = 0;

function font(name) { return { ascii: name, hAnsi: name, cs: name, eastAsia: name }; }

// Inline markup: **bold**, {ar}..{/ar}, {en}..{/en}, [^ footnote ^]
function runs(text, o = {}) {
  const out = [];
  const size = o.size || BODY;
  const re = /(\[\^[\s\S]*?\^\])|(\*\*[\s\S]*?\*\*)|(\{ar\}[\s\S]*?\{\/ar\})|(\{en\}[\s\S]*?\{\/en\})/g;
  let last = 0, m;
  const push = (t, extra = {}) => {
    if (!t) return;
    const isEn = extra.en;
    const f = extra.ar ? AR : (isEn ? 'Times New Roman' : (o.font || UR));
    const sz = extra.ar ? Math.round(size * 1.15) : (isEn ? Math.round(size * 0.8) : size);
    out.push(new TextRun({
      text: t, font: font(f), size: sz, sizeComplexScript: sz,
      bold: o.bold || extra.bold, boldComplexScript: o.bold || extra.bold,
      rightToLeft: !isEn, color: o.color,
    }));
  };
  while ((m = re.exec(text))) {
    push(text.slice(last, m.index));
    const s = m[0];
    if (m[1]) {
      fnCount++;
      footnotes[fnCount] = { children: [fnPara(s.slice(2, -2).trim())] };
      out.push(new FootnoteReferenceRun(fnCount));
    } else if (m[2]) push(s.slice(2, -2), { bold: true });
    else if (m[3]) push(s.slice(4, -5), { ar: true });
    else if (m[4]) push(s.slice(4, -5), { en: true });
    last = m.index + s.length;
  }
  push(text.slice(last));
  return out;
}

function fnPara(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.BOTH,
    spacing: { after: 40, line: 240, lineRule: LineRuleType.AUTO },
    children: runs(text, { size: FN }),
  });
}

function para(text, o = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: o.align || AlignmentType.BOTH,
    spacing: { before: o.before || 0, after: o.after ?? 120, line: o.line || 300, lineRule: LineRuleType.AUTO },
    indent: o.indent,
    heading: o.heading,
    shading: o.shade ? { type: ShadingType.CLEAR, color: 'auto', fill: o.shade } : undefined,
    border: o.box ? boxBorder() : undefined,
    pageBreakBefore: o.pageBreakBefore,
    keepNext: o.keepNext,
    children: runs(text, o),
  });
}

function boxBorder() {
  const b = { style: BorderStyle.SINGLE, size: 6, color: '808080', space: 4 };
  return { top: b, bottom: b, left: b, right: b };
}

function blank(n = 1, size = BODY) {
  return Array.from({ length: n }, () => new Paragraph({ children: [new TextRun({ text: '', size })] }));
}

function cell(text, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [para(text, { align: o.align || AlignmentType.CENTER, bold: o.bold, after: 0, size: o.size })],
  });
}

function table(widths, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    visuallyRightToLeft: true,
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    alignment: AlignmentType.CENTER,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0 && o.header,
      children: r.map((t, j) => cell(t, widths[j], {
        bold: i === 0 && o.header, fill: i === 0 && o.header ? GRAY : undefined,
        align: (o.alignCols && o.alignCols[j]) || AlignmentType.CENTER, size: o.size,
      })),
    })),
  });
}

// ---------- special pages ----------
function titlePage(meta) {
  return [
    para('{ar}بِسْمِ اللّٰهِ الرَّحْمٰنِ الرَّحِيْمِ{/ar}', { align: AlignmentType.CENTER, size: 28, after: 200 }),
    para(meta.title, { align: AlignmentType.CENTER, size: 44, bold: true, after: 60 }),
    para(meta.subtitle, { align: AlignmentType.CENTER, size: 30, after: 300 }),
    para('مقالہ برائے: شہادۃ العالمیہ (سالِ دوم) 2026ء', { align: AlignmentType.CENTER, size: 30, bold: true, after: 300 }),
    table([3000, 4000], [
      ['مقالہ نگار', ''], ['ولدیت', ''], ['رول نمبر', ''], ['رجسٹریشن نمبر', ''],
      ['نگرانِ مقالہ (مشرف)', ''], ['ادارہ / جامعہ', ''], ['الحاق نمبر', ''],
    ], { size: 24 }),
    para('', { after: 300 }),
    para('تنظیم المدارس اہلِ سنت پاکستان', { align: AlignmentType.CENTER, size: 30, bold: true, after: 0 }),
    para('1448ھ / 2026ء', { align: AlignmentType.CENTER, size: 26, after: 0 }),
  ];
}

function resultPage() {
  const rows = [
    ['نمبر شمار', 'عنوان', 'کل نمبر', 'حاصل کردہ نمبر'],
    ['1', 'فہرست', '5', ''], ['2', 'مقدمہ', '15', ''], ['3', 'صلبِ موضوع', '50', ''],
    ['4', 'حواشی و حوالہ جات', '10', ''], ['5', 'خلاصۂ بحث', '5', ''],
    ['6', 'تجاویز و سفارشات', '5', ''], ['7', 'فہرستِ آیات و احادیث', '5', ''],
    ['8', 'فہرستِ مصادر و مراجع', '5', ''], ['', 'میزان', '100', ''],
  ];
  return [
    para('نتیجہ از ممتحنِ گرامی', { align: AlignmentType.CENTER, size: 40, bold: true, after: 300, pageBreakBefore: true }),
    table([1300, 4000, 1500, 2200], rows, { header: true, size: 24 }),
    para('', { after: 300 }),
    para('تبصرۂ ممتحن: ______________________________________________', { after: 200 }),
    para('نام ممتحن: ______________   دستخط: ______________   تاریخ: ____________', { after: 0, size: 24 }),
  ];
}

function calliPage(meta) {
  return [
    new Paragraph({ pageBreakBefore: true, children: [] }),
    ...blank(8),
    para(meta.title, { align: AlignmentType.CENTER, size: 64, bold: true, after: 200 }),
    para(meta.subtitle, { align: AlignmentType.CENTER, size: 40 }),
  ];
}

function tocPage() {
  return [
    para('فہرستِ عنوانات', { align: AlignmentType.CENTER, size: 40, bold: true, after: 300, pageBreakBefore: true }),
    new TableOfContents('فہرستِ عنوانات', { hyperlink: true, headingStyleRange: '1-3' }),
  ];
}

function baabPage(line1, line2) {
  return [
    new Paragraph({ pageBreakBefore: true, children: [] }),
    ...blank(9),
    para(line1, { align: AlignmentType.CENTER, size: 60, bold: true, after: 300, heading: HeadingLevel.HEADING_1 }),
    para(line2, { align: AlignmentType.CENTER, size: 40, bold: true }),
  ];
}

// ---------- markup parser ----------
function parse(src, meta) {
  const out = [];
  const blocks = src.replace(/\r/g, '').split(/\n\s*\n/);
  for (let b of blocks) {
    b = b.trim();
    if (!b || b.startsWith('%%')) continue;
    if (b === '@title-page') out.push(...titlePage(meta));
    else if (b === '@result-page') out.push(...resultPage());
    else if (b === '@calli-page') out.push(...calliPage(meta));
    else if (b === '@toc') out.push(...tocPage());
    else if (b.startsWith('@baab ')) { const [a, c] = b.slice(6).split('|'); out.push(...baabPage(a.trim(), c.trim())); }
    else if (b.startsWith('@center ')) out.push(para(b.slice(8), { align: AlignmentType.CENTER, size: 34, line: 400 }));
    else if (b.startsWith('@sign')) {
      out.push(...blank(2));
      for (const l of b.split('\n').slice(1)) out.push(para(l, { align: AlignmentType.LEFT, after: 60 }));
    }
    else if (b.startsWith('# ')) out.push(para(b.slice(2), { align: AlignmentType.CENTER, size: 40, bold: true, after: 300, pageBreakBefore: true, heading: HeadingLevel.HEADING_1 }));
    else if (b.startsWith('## ')) {
      const [a, c] = b.slice(3).split('|');
      out.push(para(a.trim(), { align: AlignmentType.CENTER, size: 34, bold: true, after: 0, shade: GRAY, box: true, pageBreakBefore: true, keepNext: true }));
      out.push(para(c.trim(), { align: AlignmentType.CENTER, size: 32, bold: true, after: 300, heading: HeadingLevel.HEADING_2 }));
    }
    else if (b.startsWith('### ')) out.push(para(b.slice(4), { size: 32, bold: true, before: 160, after: 80, keepNext: true, heading: HeadingLevel.HEADING_3 }));
    else if (b.startsWith('#### ')) out.push(para(b.slice(5), { size: 30, bold: true, before: 100, after: 60, keepNext: true }));
    else if (b.startsWith('> ')) {
      const m = b.slice(2).replace(/\n/g, ' ').match(/^([\s\S]*?)(\[\^[\s\S]*\^\])?$/);
      out.push(para('{ar}' + m[1].trim() + '{/ar}' + (m[2] || ''), { align: AlignmentType.CENTER, size: 30, line: 360, after: 60, keepNext: true }));
    }
    else if (b.startsWith('T: ')) out.push(para('**ترجمہ:** ' + b.slice(3), { after: 160 }));
    else if (b.startsWith('- ')) for (const l of b.split('\n')) out.push(para('٭ ' + l.replace(/^- /, ''), { indent: { start: 300 }, after: 80 }));
    else if (b.startsWith('@table ')) {
      const lines = b.split('\n');
      const widths = lines[0].slice(7).split(',').map(Number);
      out.push(table(widths, lines.slice(1).map(l => l.split('|').map(s => s.trim())), { header: true, size: 24 }));
      out.push(...blank(1));
    }
    else out.push(para(b.replace(/\n/g, ' '), { indent: { firstLine: 0 } }));
  }
  return out;
}

const [outFile, ...parts] = process.argv.slice(2);
const meta = { title: 'زکوٰۃ: غربت کے خاتمے کا بہترین ذریعہ', subtitle: '(ایک مطالعہ)' };
const children = parts.flatMap(p => parse(fs.readFileSync(p, 'utf8'), meta));

const pageBorder = { style: BorderStyle.SINGLE, size: 6, color: '7F7F7F', space: 18 };
const doc = new Document({
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: font(UR), size: BODY, sizeComplexScript: BODY, rightToLeft: true }, paragraph: { bidirectional: true } } },
    paragraphStyles: [1, 2, 3].map(n => ({
      id: 'Heading' + n, name: 'Heading ' + n, basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: font(UR), color: '000000', bold: true, boldComplexScript: true },
      paragraph: { bidirectional: true, outlineLevel: n - 1 },
    })),
  },
  footnotes,
  sections: [{
    properties: {
      titlePage: true,
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1300, bottom: 1300, left: 1300, right: 1300, header: 700, footer: 600 },
        borders: { pageBorderTop: pageBorder, pageBorderBottom: pageBorder, pageBorderLeft: pageBorder, pageBorderRight: pageBorder },
      },
    },
    headers: {
      first: new Header({ children: [] }),
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 24, font: font('Times New Roman') })] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outFile, buf); console.log('wrote', outFile, 'footnotes:', fnCount); });
