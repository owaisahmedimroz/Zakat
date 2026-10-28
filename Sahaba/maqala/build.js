// Builds the thesis .docx from simple markup files.
// Usage: node build.js out.docx part1.txt [part2.txt ...]
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak, Header,
  PageNumber, BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType,
  FootnoteReferenceRun, HeadingLevel, TableOfContents, VerticalAlign,
  LineRuleType, Bookmark, PageReference, ImageRun,
} = require('docx');

// Tanzeem-ul-Madaris requirements: body 15pt, footnotes 13pt, headings 22/20/18/16, title 24
const UR = 'Jameel Noori Nastaleeq';
const AR = 'Al Mushaf';
const BODY = 30;      // 15pt
const FN = 26;        // 13pt
const QUOTE = 28;     // 14pt (اقتباس، متن سے قدرے چھوٹا)
const H_BAAB = 44, H_FASL = 40, H_MABHAS = 36, H_ZAILI = 32, H_TITLE = 48;
const GRAY = 'EDEDED';

const footnotes = {};
let fnCount = 0;
const ayat = [], ahadees = [];
let bmCount = 0, lastAr = '';

function font(name) { return { ascii: name, hAnsi: name, cs: name, eastAsia: name }; }

// Inline markup: **bold**, {ar}..{/ar}, {en}..{/en}, [^ footnote ^]
function runs(text, o = {}) {
  const out = [];
  const size = o.size || BODY;
  const re = /(\[\^[\s\S]*?\^\])|(\*\*[\s\S]*?\*\*)|(\{ar\}[\s\S]*?\{\/ar\})|(\{en\}[\s\S]*?\{\/en\})|(\{hd:[\s\S]*?\})/g;
  let last = 0, m;
  const push = (t, extra = {}) => {
    if (!t) return;
    const isEn = extra.en;
    const f = extra.ar ? AR : (isEn ? 'Times New Roman' : (o.font || UR));
    const sz = extra.ar ? size : (isEn ? Math.round(size * 0.8) : size);
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
      const ft = s.slice(2, -2).trim();
      if (/^سورۃ/.test(ft)) {
        const id = 'ay' + (++bmCount);
        out.push(new Bookmark({ id, children: [new TextRun('')] }));
        ayat.push({ ref: ft, ar: o.arCtx || lastAr, bm: id });
      }
      fnCount++;
      footnotes[fnCount] = { children: [fnPara(ft)] };
      out.push(new FootnoteReferenceRun(fnCount));
    } else if (m[2]) push(s.slice(2, -2), { bold: true });
    else if (m[3]) { lastAr = s.slice(4, -5); push(lastAr, { ar: true }); }
    else if (m[5]) {
      const id = 'hd' + (++bmCount);
      out.push(new Bookmark({ id, children: [new TextRun('')] }));
      ahadees.push({ label: s.slice(4, -1).trim(), bm: id });
    }
    else if (m[4]) push(s.slice(4, -5), { en: true });
    last = m.index + s.length;
  }
  push(text.slice(last));
  return out;
}

function fnPara(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.BOTH,
    spacing: { after: 0, line: 240, lineRule: LineRuleType.AUTO },
    children: runs(text, { size: FN }),
  });
}

function para(text, o = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: o.align || AlignmentType.BOTH,
    spacing: { before: o.before || 0, after: o.after ?? 80, line: o.line || 240, lineRule: LineRuleType.AUTO },
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
    margins: { top: 30, bottom: 30, left: 100, right: 100 },
    children: [para(text, { align: o.align || AlignmentType.CENTER, bold: o.bold, after: 0, size: o.size })],
  });
}

function table(widths, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = rows.map((r, i) => new TableRow({
    tableHeader: i === 0 && o.header,
    children: r.map((t, j) => cell(t, widths[j], {
      bold: (i === 0 && o.header) || (o.boldCol0 && j === 0), fill: i === 0 && o.header ? GRAY : undefined,
      align: (o.alignCols && o.alignCols[j]) || AlignmentType.CENTER, size: o.size,
    })),
  }));
  if (o.span) trs.push(new TableRow({ children: [new TableCell({ columnSpan: widths.length, width: { size: total, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [para(o.span, { align: AlignmentType.BOTH, bold: true, after: 0, size: o.size })] })] }));
  return new Table({ visuallyRightToLeft: true, width: { size: total, type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER, rows: trs });
}

// ---------- special pages ----------
function titlePage(meta) {
  const logo = new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(__dirname + '/logo.png'), transformation: { width: 105, height: 101 } })] });
  const tp = meta.tp;
  return [
    para('{ar}بِسْمِ اللّٰهِ الرَّحْمٰنِ الرَّحِيْمِ{/ar}', { align: AlignmentType.CENTER, after: 100 }),
    para(meta.title, { align: AlignmentType.CENTER, size: H_TITLE, bold: true, after: 60 }),
    para('(مقالہ برائے شہادۃ العالمیہ)', { align: AlignmentType.CENTER, size: 32, bold: true, after: 100 }),
    logo,
    table([3000, 6000], tp.rows, { size: 26, alignCols: [AlignmentType.BOTH, AlignmentType.BOTH], boldCol0: true, span: tp.footerRow }),
    para('تنظیم المدارس اہلِ سنت پاکستان', { align: AlignmentType.CENTER, size: 32, bold: true, before: 300, after: 0 }),
    para(tp.year, { align: AlignmentType.CENTER, size: 28, after: 0 }),
  ];
}

function tasmiyaPage() {
  return [
    new Paragraph({ pageBreakBefore: true, children: [] }),
    ...blank(6),
    para('{ar}بِسْمِ اللّٰهِ الرَّحْمٰنِ الرَّحِيْمِ{/ar}', { align: AlignmentType.CENTER, size: 44, after: 600 }),
    para('{ar}اَلْحَمْدُ لِلّٰهِ رَبِّ الْعٰلَمِيْنَ وَالصَّلٰوةُ وَالسَّلَامُ عَلٰى سَيِّدِ الْمُرْسَلِيْنَ وَعَلٰى اٰلِهٖ وَاَصْحَابِهٖ اَجْمَعِيْنَ{/ar}', { align: AlignmentType.CENTER, size: 36, line: 360, after: 400 }),
    para('{ar}رَضِيَ اللّٰهُ عَنْهُمْ وَرَضُوْا عَنْهُ{/ar}', { align: AlignmentType.CENTER, size: 36, after: 0 }),
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

function baabPage(line1, line2, fusool) {
  const out = [
    new Paragraph({ pageBreakBefore: true, children: [] }),
    ...blank(5),
    para(line1, { align: AlignmentType.CENTER, size: H_BAAB + 12, bold: true, after: 200, heading: HeadingLevel.HEADING_1 }),
    para(line2, { align: AlignmentType.CENTER, size: H_BAAB, bold: true, after: 500 }),
  ];
  for (const f of fusool) out.push(para(f, { align: AlignmentType.CENTER, size: H_FASL, bold: true, after: 200 }));
  return out;
}

// ---------- indexes (page numbers filled by Word via PAGEREF) ----------
const SURAH = ['الفاتحۃ','البقرۃ','آلِ عمران','النساء','المائدۃ','الانعام','الاعراف','الانفال','التوبۃ','یونس','ہود','یوسف','الرعد','ابراہیم','الحجر','النحل','بنی اسرائیل','الکہف','مریم','طٰہٰ','الانبیاء','الحج','المؤمنون','النور','الفرقان','الشعراء','النمل','القصص','العنکبوت','الروم','لقمان','السجدۃ','الاحزاب','سبا','فاطر','یٰس','الصافات','ص','الزمر','المؤمن','حٰم السجدۃ','الشوریٰ','الزخرف','الدخان','الجاثیۃ','الاحقاف','محمد','الفتح','الحجرات','ق','الذاریات','الطور','النجم','القمر','الرحمٰن','الواقعۃ','الحدید','المجادلۃ','الحشر','الممتحنۃ','الصف','الجمعۃ','المنافقون','التغابن','الطلاق','التحریم','الملک','القلم','الحاقۃ','المعارج','نوح','الجن','المزمل','المدثر','القیامۃ','الدہر','المرسلات','النبا','النازعات','عبس','التکویر','الانفطار','المطففین','الانشقاق','البروج','الطارق','الاعلیٰ','الغاشیۃ','الفجر','البلد','الشمس','اللیل','الضحیٰ','الانشراح','التین','العلق'];
function pageCell(bm, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new PageReference(bm, { font: font('Times New Roman'), size: 22 })] })],
  });
}
function indexTable(widths, header, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    visuallyRightToLeft: true, width: { size: total, type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((t, j) => cell(t, widths[j], { bold: true, fill: GRAY, size: 24 })) }),
      ...rows.map(r => new TableRow({ children: r.map((c, j) => typeof c === 'string' ? cell(c, widths[j], { size: 24 }) : c.bm ? pageCell(c.bm, widths[j]) : cell(c.t, widths[j], { size: c.size, align: c.align })) })),
    ],
  });
}
function fehristAyat() {
  const seen = new Set(), items = [];
  for (const a of ayat) {
    const m = a.ref.match(/^سورۃ\s+(.+?):\s*(\d+)(?:\s*تا\s*(\d+))?/);
    if (!m) continue;
    const key = m[1] + ':' + m[2] + (m[3] ? '-' + m[3] : '');
    if (seen.has(key)) continue; seen.add(key);
    const idx = SURAH.indexOf(m[1].trim());
    const words = (a.ar || '').split(/\s+/).filter(Boolean).slice(0, 5).join(' ');
    items.push({ surah: m[1].trim(), idx: idx < 0 ? 999 : idx, n: +m[2], ayah: m[3] ? m[2] + ' تا ' + m[3] : m[2], words, bm: a.bm });
  }
  items.sort((x, y) => x.idx - y.idx || x.n - y.n);
  return [indexTable([800, 4000, 1700, 1000, 1400], ['نمبر شمار', 'آیتِ مبارکہ (ابتدائی کلمات)', 'سورۃ', 'آیت', 'صفحہ'],
    items.map((it, i) => [String(i + 1), { t: '{ar}' + (it.words ? it.words + ' ...' : '—') + '{/ar}', size: 22 }, it.surah, it.ayah, { bm: it.bm }]))];
}
function fehristAhadees() {
  return [indexTable([800, 6900, 1400], ['نمبر شمار', 'حدیث / اثر (مختصر)', 'صفحہ'],
    ahadees.map((h, i) => [String(i + 1), { t: h.label, size: 24, align: AlignmentType.BOTH }, { bm: h.bm }]))];
}

// ---------- markup parser ----------
function parse(src, meta) {
  const out = [];
  const blocks = src.replace(/\r/g, '').split(/\n\s*\n/);
  for (let bi = 0; bi < blocks.length; bi++) {
    let b = blocks[bi].trim();
    const nextB = (blocks.slice(bi + 1).find(x => x.trim() && !x.trim().startsWith('%%')) || '').trim();
    if (!b || b.startsWith('%%')) continue;
    if (b === '@title-page') out.push(...titlePage(meta));
    else if (b === '@result-page') out.push(...resultPage());
    else if (b === '@calli-page') out.push(...calliPage(meta));
    else if (b === '@toc') out.push(...tocPage());
    else if (b === '@fehrist-ayat') out.push(...fehristAyat());
    else if (b === '@fehrist-ahadees') out.push(...fehristAhadees());
    else if (b.startsWith('@baab ')) {
      const lines = b.split('\n'); const [a, c] = lines[0].slice(6).split('|');
      out.push(...baabPage(a.trim(), c.trim(), lines.slice(1).map(s => s.replace(/^- /, '').trim()).filter(Boolean)));
    }
    else if (b === '@tasmiya') out.push(...tasmiyaPage());
    else if (b === '@pagebreak') out.push(new Paragraph({ pageBreakBefore: true, children: [] }));
    else if (b.startsWith('@center ')) out.push(para(b.slice(8), { align: AlignmentType.CENTER, size: 34, line: 300 }));
    else if (b.startsWith('@sign')) {
      out.push(...blank(2));
      for (const l of b.split('\n').slice(1)) out.push(para(l, { align: AlignmentType.BOTH, after: 60 }));
    }
    else if (b.startsWith('# ')) out.push(para(b.slice(2), { align: AlignmentType.CENTER, size: H_BAAB, bold: true, after: 300, pageBreakBefore: true, heading: HeadingLevel.HEADING_1 }));
    else if (b.startsWith('## ')) {
      const [a, c] = b.slice(3).split('|');
      out.push(para(a.trim(), { align: AlignmentType.CENTER, size: H_FASL, bold: true, after: 0, shade: GRAY, box: true, pageBreakBefore: true, keepNext: true }));
      out.push(para(c.trim(), { align: AlignmentType.CENTER, size: H_FASL, bold: true, after: 300, heading: HeadingLevel.HEADING_2 }));
    }
    else if (b.startsWith('### ')) out.push(para(b.slice(4), { align: AlignmentType.BOTH, size: H_MABHAS, bold: true, before: 160, after: 80, keepNext: true, heading: HeadingLevel.HEADING_3 }));
    else if (b.startsWith('#### ')) out.push(para(b.slice(5), { align: AlignmentType.BOTH, size: H_ZAILI, bold: true, before: 100, after: 60, keepNext: true }));
    else if (b.startsWith('> ')) {
      const m = b.slice(2).replace(/\n/g, ' ').match(/^([\s\S]*?)(\[\^[\s\S]*\^\])?$/);
      out.push(para('{ar}' + m[1].trim() + '{/ar}' + (m[2] || ''), { align: AlignmentType.CENTER, size: BODY + 2, line: 300, after: 60, keepNext: true, arCtx: m[1].trim() }));
    }
    else if (b.startsWith('T: ')) out.push(para('**ترجمہ:** ' + b.slice(3), { after: 160 }));
    else if (b.startsWith('- ')) for (const l of b.split('\n')) out.push(para('٭ ' + l.replace(/^- /, ''), { indent: { start: 300 }, after: 80 }));
    else if (b.startsWith('@table ')) {
      const lines = b.split('\n');
      const widths = lines[0].slice(7).split(',').map(Number);
      out.push(table(widths, lines.slice(1).map(l => l.split('|').map(s => s.trim())), { header: true, size: 24 }));
      // no spacer before a block that starts on a new page (avoids an empty page when the table ends at a page bottom)
      if (!/^(# |## |@pagebreak|@baab |@toc|@fehrist)/.test(nextB)) out.push(...blank(1));
    }
    else if (/^["“]/.test(b)) out.push(para(b.replace(/\n/g, ' '), { size: QUOTE, indent: { start: 400, end: 400 } }));
    else out.push(para(b.replace(/\n/g, ' '), { indent: { firstLine: 0 } }));
  }
  return out;
}

const [outFile, ...parts] = process.argv.slice(2);
const meta = {
  title: 'صحابۂ کرام علیہم الرضوان کی عسکری مہمات کا تفصیلی مطالعہ', subtitle: '',
  tp: {
    rows: [
      ['مقالہ نگار', 'محمد حسین بن سرفراز'],
      ['شہادۃ العالیہ اول رول نمبر', '414089'],
      ['رجسٹریشن نمبر', ''],
      ['مقالہ نگران', 'علامہ مفتی علی اشرف اختر القادری'],
      ['منتظم', 'حبیبیہ اسلامک اکیڈمی'],
      ['الحاق نمبر', '08113'],
      ['درجہ الحاق', 'فوقانی'],
    ],
    footerRow: 'شعبہ امتحانات شہادۃ العالمیہ تنظیم المدارس اہلسنت پاکستان ۱۴۴۹ھ بمطابق ۲۰۲۷ء',
    year: '۱۴۴۹ھ / ۲۰۲۷ء',
  },
};
const children = parts.flatMap(p => parse(fs.readFileSync(p, 'utf8'), meta));


const doc = new Document({
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: font(UR), size: BODY, sizeComplexScript: BODY, rightToLeft: true }, paragraph: { bidirectional: true } } },
    paragraphStyles: [1, 2, 3].map(n => ({
      id: 'Heading' + n, name: 'Heading ' + n, basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: font(UR), color: '000000', bold: true, boldComplexScript: true },
      paragraph: { bidirectional: true, outlineLevel: n - 1 },
    })).concat([1, 2, 3].map(n => ({
      // fehrist-e-unwanaat entries keep the Urdu font and right-to-left layout after "Update fields"
      id: 'TOC' + n, name: 'toc ' + n, basedOn: 'Normal', next: 'Normal', uiPriority: 39,
      run: { font: font(UR), size: BODY - 2, sizeComplexScript: BODY - 2, bold: n === 1, boldComplexScript: n === 1, rightToLeft: true },
      paragraph: { bidirectional: true, alignment: AlignmentType.BOTH, indent: { start: (n - 1) * 400 }, spacing: { after: 60 } },
    }))),
  },
  footnotes,
  sections: [{
    properties: {
      titlePage: true,
      page: {
        size: { width: 11906, height: 16838 },
        // A4; جلد بندی (دائیں) 1.25 انچ، دوسری جانب 0.75 انچ، اوپر 1 انچ، نیچے 0.75 انچ
        margin: { top: 1440, bottom: 1080, right: 1800, left: 1080, header: 600, footer: 500 },
      },
    },
    headers: {
      first: new Header({ children: [] }),
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 26, font: font(UR) })] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outFile, buf); console.log('wrote', outFile, 'footnotes:', fnCount); });
