/**
 * L-TEX — Розрахунок зарплати за огляди  (ДОПОВНЕННЯ до backend-скрипта)
 * ---------------------------------------------------------------------
 * Додайте цей блок У КІНЕЦЬ наявного standalone Apps Script (де є
 * getSpreadsheet_(), doPost(), appendResponse_() та константа RESPONSES_SHEET).
 *
 * Створює/оновлює лист «Зарплата»:
 *
 *   ПІДСУМОК (вгорі)
 *     • Зароблена зарплата — кількість оглянутих лотів × ставка (авто)
 *     • Фактично видана    — сума колонки «Виплачено» + виплати без дати (авто)
 *     • Різниця            — Зароблена − Фактично видана (авто)
 *
 *   ТАБЛИЦЯ ПО ДНЯХ (з рядка 12):
 *     День | Лотів | Зароблено, грн | Виплачено, грн | Залишок боргу
 *      ↑ авто   ↑ авто     ↑ авто        ↑ ВПИСУЄТЕ ВРУЧНУ    ↑ авто (наростаюче)
 *
 *   Навпроти кожного дня у колонці «Виплачено, грн» вписуєте фактично видану
 *   того дня суму — решта рахується формулами.
 *
 *   «Виплати без прив'язки до дня» (A10/B10) — для сум, які не стосуються
 *   конкретного дня (напр. премія). Теж додаються у «Фактично видана».
 *
 * Рахуються лише огляди, у яких «Час відправки» >= SALARY_START.
 *
 * ЗАПУСК: редактор Apps Script → функція buildSalarySheet → «Запустити».
 *   Повторний запуск НЕ втрачає введені суми — вони переносяться назад за датою.
 */

// ===================== НАЛАШТУВАННЯ ЗАРПЛАТИ =====================
const SALARY_SHEET     = 'Зарплата';        // лист із розрахунком (створиться автоматично)
const SALARY_DATE_COL  = 'Час відправки';   // колонка з датою-часом огляду в аркуші «Анкети»
const RATE_PER_LOT     = 170;               // грн за 1 лот (1 рядок-огляд)
const SALARY_START     = new Date(2026, 5, 23, 14, 31, 41); // 23.06.2026 14:31:41 (місяць 0-based: 5 = червень)

const DAY_TITLE_ROW    = 10;  // заголовок денної таблиці
const DAY_HEADER_ROW   = 11;  // шапка денної таблиці
const DAY_FIRST_ROW    = 12;  // з цього рядка йдуть дні
// колонки денної таблиці:  E=День  F=Лотів  G=Зароблено  H=Виплачено  I=Залишок
const COL_DAY = 5, COL_LOTS = 6, COL_EARNED = 7, COL_PAID = 8, COL_BAL = 9;
// ================================================================

function buildSalarySheet() {
  const ss  = getSpreadsheet_();                  // ← ваш наявний helper (відкриває за SPREADSHEET_ID)
  const src = ss.getSheetByName(RESPONSES_SHEET); // 'Анкети'
  if (!src) throw new Error('Не знайдено лист «' + RESPONSES_SHEET + '»');
  const tz = ss.getSpreadsheetTimeZone();

  // 1) Колонка з датою — за заголовком
  const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0].map(String);
  let dateCol = headers.indexOf(SALARY_DATE_COL) + 1;
  if (dateCol === 0) dateCol = 1;
  const c = columnToLetter_(dateCol);
  const dateRange = "'" + RESPONSES_SHEET + "'!" + c + '2:' + c;

  // 2) Надійний підрахунок у коді (звірка / текстові дати)
  const lastRow = src.getLastRow();
  let codeCount = 0;
  if (lastRow >= 2) {
    const vals = src.getRange(2, dateCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const d = toDate_(vals[i][0]);
      if (d && d.getTime() >= SALARY_START.getTime()) codeCount++;
    }
  }

  // 3) ЗБЕРЕГТИ вже введені виплати (щоб не втратити при повторному запуску)
  const paidByDate = {}; // 'yyyy-MM-dd' -> сума
  let extra = 0;         // виплати без прив'язки до дня
  const existing = ss.getSheetByName(SALARY_SHEET);
  if (existing) {
    const lr = existing.getLastRow();
    // 3a) нова розкладка: пари День(E) → Виплачено(H)
    if (lr >= DAY_FIRST_ROW) {
      const rows = existing.getRange(DAY_FIRST_ROW, COL_DAY, lr - DAY_FIRST_ROW + 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const d = toDate_(rows[i][0]);
        const paid = rows[i][3];
        if (d && typeof paid === 'number' && paid !== 0) {
          const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
          paidByDate[k] = (paidByDate[k] || 0) + paid;
        }
      }
    }
    // 3b) стара розкладка: журнал зліва A(дата)/B(сума). Дато­вані → по днях, решта → extra
    if (lr >= 12) {
      const jr = existing.getRange(12, 1, lr - 12 + 1, 2).getValues();
      for (let i = 0; i < jr.length; i++) {
        const amt = jr[i][1];
        if (typeof amt === 'number' && amt !== 0) {
          const d = toDate_(jr[i][0]);
          if (d) {
            const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
            paidByDate[k] = (paidByDate[k] || 0) + amt;
          } else {
            extra += amt;
          }
        }
      }
    }
    // 3c) наявна клітинка «без дати»
    if (existing.getRange('B10').getFormula() === '') {
      const b10 = existing.getRange('B10').getValue();
      if (typeof b10 === 'number') extra += b10;
    }
  }

  // 4) Створити / очистити лист
  let sh = ss.getSheetByName(SALARY_SHEET);
  if (!sh) sh = ss.insertSheet(SALARY_SHEET);
  else sh.clear();
  const maxRows = sh.getMaxRows();

  // ---- ПІДСУМОК ----
  sh.getRange('A1').setValue('Розрахунок зарплати за огляди').setFontWeight('bold').setFontSize(13);
  sh.getRange('A3').setValue('Ставка за лот, грн');
  sh.getRange('B3').setValue(RATE_PER_LOT);
  sh.getRange('A4').setValue('Рахувати огляди від');
  sh.getRange('B4').setValue(SALARY_START).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  sh.getRange('A5').setValue('Оглянуто лотів (рядків)');
  sh.getRange('B5').setFormula('=COUNTIFS(' + dateRange + ',">="&$B$4)');

  sh.getRange('A7:C7')
    .setValues([['Зароблена зарплата', 'Фактично видана', 'Різниця']])
    .setFontWeight('bold').setBackground('#e8eef7');
  sh.getRange('A8').setFormula('=$B$5*$B$3');
  sh.getRange('B8').setFormula('=SUM(H' + DAY_FIRST_ROW + ':H)+B10');
  sh.getRange('C8').setFormula('=A8-B8');
  sh.getRange('A8:C8').setNumberFormat('#,##0.00" грн"');

  // ---- Виплати без прив'язки до дня ----
  sh.getRange('A10').setValue("Виплати без прив'язки до дня, грн");
  sh.getRange('B10').setNumberFormat('#,##0.00" грн"').setBackground('#fff2cc');

  // ---- ТАБЛИЦЯ ПО ДНЯХ ----
  sh.getRange(DAY_TITLE_ROW, COL_DAY).setValue('Заробіток і виплати по днях').setFontWeight('bold');
  sh.getRange(DAY_HEADER_ROW, COL_DAY, 1, 5)
    .setValues([['День', 'Лотів', 'Зароблено, грн', 'Виплачено, грн', 'Залишок боргу']])
    .setFontWeight('bold').setBackground('#e8eef7');
  sh.getRange(DAY_HEADER_ROW, COL_PAID).setBackground('#fff2cc'); // підсвітити колонку вводу

  // День + Лотів (QUERY, спіл у E:F)
  const dtLit = Utilities.formatDate(SALARY_START, tz, 'yyyy-MM-dd HH:mm:ss');
  const dailyQuery =
    '=IFERROR(QUERY(' + dateRange + ', "select toDate(' + c + '), count(' + c + ') ' +
    'where ' + c + " >= datetime '" + dtLit + "' " +
    'group by toDate(' + c + ') order by toDate(' + c + ') ' +
    "label toDate(" + c + ") '', count(" + c + ") ''\", 0), )";
  sh.getRange(DAY_FIRST_ROW, COL_DAY).setFormula(dailyQuery);
  // Зароблено = Лотів × ставка
  sh.getRange(DAY_FIRST_ROW, COL_EARNED).setFormula(
    '=ARRAYFORMULA(IF(LEN(F' + DAY_FIRST_ROW + ':F), F' + DAY_FIRST_ROW + ':F*$B$3, ""))');
  // Залишок боргу наростаюче = (зароблено до цього дня) − (виплачено до цього дня) − виплати без дати
  sh.getRange(DAY_FIRST_ROW, COL_BAL).setFormula(
    '=ARRAYFORMULA(IF(LEN(E' + DAY_FIRST_ROW + ':E), ' +
    'SUMIF(E' + DAY_FIRST_ROW + ':E,"<="&E' + DAY_FIRST_ROW + ':E,G' + DAY_FIRST_ROW + ':G) - ' +
    'SUMIF(E' + DAY_FIRST_ROW + ':E,"<="&E' + DAY_FIRST_ROW + ':E,$H$' + DAY_FIRST_ROW + ':H) - $B$10, ""))');

  // формати колонок денної таблиці
  const len = maxRows - DAY_FIRST_ROW + 1;
  sh.getRange(DAY_FIRST_ROW, COL_DAY, len, 1).setNumberFormat('dd.MM.yyyy');
  sh.getRange(DAY_FIRST_ROW, COL_LOTS, len, 1).setNumberFormat('0');
  sh.getRange(DAY_FIRST_ROW, COL_EARNED, len, 1).setNumberFormat('#,##0.00" грн"');
  sh.getRange(DAY_FIRST_ROW, COL_PAID, len, 1).setNumberFormat('#,##0.00" грн"').setBackground('#fffdf5');
  sh.getRange(DAY_FIRST_ROW, COL_BAL, len, 1).setNumberFormat('#,##0.00" грн"');

  // 5) Повернути збережені виплати навпроти відповідних днів
  SpreadsheetApp.flush();
  const eCol = sh.getRange(DAY_FIRST_ROW, COL_DAY, len, 1).getValues();
  let n = 0;
  while (n < eCol.length && eCol[n][0] !== '') n++;
  if (n > 0) {
    const hOut = [];
    for (let i = 0; i < n; i++) {
      const d = toDate_(eCol[i][0]);
      const k = d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';
      if (k && paidByDate[k] != null) { hOut.push([paidByDate[k]]); delete paidByDate[k]; }
      else hOut.push(['']);
    }
    sh.getRange(DAY_FIRST_ROW, COL_PAID, n, 1).setValues(hOut);
  }
  // виплати за дати без оглядів → у «без прив'язки до дня», щоб нічого не загубити
  for (const k in paidByDate) extra += paidByDate[k];
  sh.getRange('B10').setValue(extra);

  // 6) Запобіжник: COUNTIFS = 0, а лоти є → дати як текст
  SpreadsheetApp.flush();
  const formulaCount = Number(sh.getRange('B5').getValue()) || 0;
  if (formulaCount === 0 && codeCount > 0) {
    sh.getRange('B5').setValue(codeCount);
    sh.getRange('A6')
      .setValue('⚠️ Дати в «' + SALARY_DATE_COL + '» як текст — підсумок статичний, запускайте buildSalarySheet повторно.')
      .setFontColor('#b26a00').setFontSize(9);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 9);

  Logger.log('Лист «Зарплата» готовий. Лотів з ' + SALARY_START + ': ' + codeCount +
             ' · Зароблено: ' + (codeCount * RATE_PER_LOT) + ' грн · Без дати: ' + extra + ' грн');
  return ss.getUrl();
}

/** Номер стовпця → літера A1 (1→A, 27→AA). */
function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/** Значення клітинки (Date / серійний номер / текст) → Date або null. */
function toDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (typeof v === 'string' && v.trim()) {
    const s = v.trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (m) return new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +(m[6] || 0));
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}
