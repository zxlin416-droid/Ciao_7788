/**
 * 執課表彙整程式
 * 依「說明.md」規則，把「私人課表預約明細拆帳版」與「團體課預約報表」
 * 合併成「(年月)執課表.xlsx」，以授課老師分頁。
 *
 * 用法：
 *   node 彙整執課表.js [報表資料夾=報表] [範本=執課表範本.xlsx] [輸出資料夾=輸出]
 */
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const [DIR = '報表', TEMPLATE = '執課表範本.xlsx', OUT = '輸出'] = process.argv.slice(2);
const WEEK = ['日', '一', '二', '三', '四', '五', '六'];
const HEADERS = ['日期', '星期', '課程教練', '分店', '姓名', '課程內容', '收入金額-單價', '成本/支出金額-單價', '備註'];

const text = (v) => (v == null ? '' : String(v.result ?? v.text ?? v).trim());

// 讀取報表：第 1 列為標題文字、第 2 列為欄位名，之後為資料
async function readReport(prefix) {
  const file = fs.readdirSync(DIR).find((f) => f.startsWith(prefix) && f.endsWith('.xlsx') && !f.startsWith('~$'));
  if (!file) throw new Error(`在「${DIR}」找不到以「${prefix}」開頭的 xlsx`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(DIR, file));
  const ws = wb.worksheets[0];
  const names = ws.getRow(2).values;
  const rows = [];
  for (let r = 3; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values;
    if (!vals.some((v) => v != null && v !== '')) continue;
    const o = {};
    names.forEach((n, i) => { if (n) o[n] = vals[i]; });
    rows.push(o);
  }
  console.log(`讀取 ${file}：${rows.length} 筆`);
  return rows;
}

// "2026/09/30" → Date（以 UTC 存，避免時區位移）
function parseDate(s) {
  const m = text(s).match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (!m) return null;
  return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
}

async function main() {
  const priv = await readReport('私人課表預約明細拆帳版');
  const group = await readReport('團體課預約報表');

  const records = []; // {date,time,teacher,name,content,price}
  const skipped = { privEmpty: 0, groupClosed: 0, groupZero: 0, bad: 0 };

  for (const r of priv) {
    const date = parseDate(r['課程日期']);
    const teacher = text(r['授課老師']);
    if (!date || !teacher) { skipped.bad++; continue; }
    const content = text(r['費用類型']);
    const price = r['費用換算現金(元)'];
    if (!content) { skipped.privEmpty++; continue; } // 費用類型無值則整列不填入
    records.push({
      date, time: text(r['課程時間']), teacher,
      name: text(r['會員姓名']), content, note: text(r['備註']),
      price: price == null || price === '' ? null : Number(price),
    });
  }

  for (const r of group) {
    const date = parseDate(r['課程日期']);
    // 代課老師有填且老師狀態為「代課」時，授課老師改填代課老師、備註填「代課」
    const sub = text(r['代課老師']);
    const isSub = !!sub && text(r['老師狀態']) === '代課';
    const teacher = isSub ? sub : text(r['授課老師']);
    if (!date || !teacher) { skipped.bad++; continue; }
    const count = Number(r['預約總人數'] || 0);
    if (text(r['課程狀態']) === '停課') { skipped.groupClosed++; continue; }
    if (count === 0) { skipped.groupZero++; continue; }
    records.push({
      date, time: text(r['課程時間']), teacher,
      name: '', content: `${text(r['課程名稱'])}${count}`, price: null, note: isSub ? '代課' : '',
    });
  }

  if (!records.length) throw new Error('沒有可彙整的資料');

  // 年月以資料中最早日期判定
  const first = records.reduce((a, b) => (a < b.date ? a : b.date), records[0].date);
  const ym = `${first.getUTCFullYear()}${String(first.getUTCMonth() + 1).padStart(2, '0')}`;

  // 依老師分組、日期＋時間排序
  const byTeacher = new Map();
  for (const rec of records) {
    if (!byTeacher.has(rec.teacher)) byTeacher.set(rec.teacher, []);
    byTeacher.get(rec.teacher).push(rec);
  }

  // 沿用範本樣式（欄寬、字型、框線、對齊）
  const tpl = new ExcelJS.Workbook();
  await tpl.xlsx.readFile(TEMPLATE);
  const tws = tpl.worksheets[0];
  const widths = HEADERS.map((_, i) => tws.getColumn(i + 1).width);
  const headStyle = tws.getRow(1).getCell(1).style;
  const headStyles = HEADERS.map((_, i) => tws.getRow(1).getCell(i + 1).style);

  const wb = new ExcelJS.Workbook();
  for (const teacher of [...byTeacher.keys()].sort((a, b) => a.localeCompare(b, 'zh-Hant'))) {
    const list = byTeacher.get(teacher).sort((a, b) => a.date - b.date || a.time.localeCompare(b.time));
    const ws = wb.addWorksheet(teacher.replace(/[\\/?*[\]:]/g, '_').slice(0, 31));
    ws.columns = HEADERS.map((h, i) => ({ width: widths[i] || 12 }));
    const head = ws.addRow(HEADERS);
    head.eachCell((c, i) => {
      c.style = headStyles[i - 1] || headStyle;
      if (i === 2) c.fill = { type: 'pattern', pattern: 'none' }; // 星期欄：無填滿
    });
    for (const r of list) {
      const row = ws.addRow([
        r.date, WEEK[r.date.getUTCDay()], teacher, '',
        r.name, r.content, r.price, null, r.note || null,
      ]);
      row.eachCell({ includeEmpty: true }, (c, i) => {
        c.style = { ...headStyle };
        if (i === 1) c.numFmt = 'yyyy/mm/dd';
        if (i === 7 || i === 8) c.numFmt = 'General';
      });
    }
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    console.log(`  分頁「${teacher}」：${list.length} 列`);
  }

  fs.mkdirSync(OUT, { recursive: true });
  const outFile = path.join(OUT, `${ym}執課表.xlsx`);
  await wb.xlsx.writeFile(outFile);
  console.log(`\n已輸出：${outFile}`);
  console.log(`略過：私人課無費用類型 ${skipped.privEmpty} 列、團體課停課 ${skipped.groupClosed} 列、團體課預約0人 ${skipped.groupZero} 列、資料不全 ${skipped.bad} 列`);
}

main().catch((e) => { console.error('錯誤：' + e.message); process.exit(1); });
