/**
 * L-TEX — Розрахунок зарплати за огляди  (ДОПОВНЕННЯ до backend-скрипта)
 * ---------------------------------------------------------------------
 * Додайте цей блок У КІНЕЦЬ наявного standalone Apps Script (того, що містить
 * getSpreadsheet_(), doPost(), appendResponse_() тощо). Він перевикористовує
 * ваш getSpreadsheet_() і константу RESPONSES_SHEET ('Анкети').
 *
 * Створює/оновлює лист «Зарплата» з трьома показниками:
 *   • Зароблена зарплата — кількість оглянутих лотів (рядків) × ставка за лот
 *   • Фактично видана    — вписуєте вручну (клітинка B8)
 *   • Різниця            — Зароблена − Фактично видана
 *
 * Рахуються лише огляди, у яких «Час відправки» >= START_DATE.
 *
 * ЯК ЗАПУСТИТИ:
 *   Редактор Apps Script → виберіть функцію buildSalarySheet → «Запустити».
 *   Далі: «Зароблена» рахується формулою автоматично; вписуйте лише «Фактично видана».
 */

// ===================== НАЛАШТУВАННЯ ЗАРПЛАТИ =====================
const SALARY_SHEET     = 'Зарплата';        // лист із розрахунком (створиться автоматично)
const SALARY_DATE_COL  = 'Час відправки';   // колонка з датою-часом огляду в аркуші «Анкети»
const RATE_PER_LOT     = 170;               // грн за 1 лот (1 рядок-огляд)
const SALARY_START     = new Date(2026, 5, 23, 14, 31, 41); // 23.06.2026 14:31:41 (місяць 0-based: 5 = червень)
// ================================================================

function buildSalarySheet() {
  const ss  = getSpreadsheet_();                 // ← ваш наявний helper (відкриває за SPREADSHEET_ID)
  const src = ss.getSheetByName(RESPONSES_SHEET); // 'Анкети'
  if (!src) throw new Error('Не знайдено лист «' + RESPONSES_SHEET + '»');

  // 1) Колонка з датою — за заголовком (стійко до перестановки стовпців)
  const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0].map(String);
  let dateCol = headers.indexOf(SALARY_DATE_COL) + 1;
  if (dateCol === 0) dateCol = 1; // запасний варіант — перша колонка
  const colLetter = columnToLetter_(dateCol);
  const dateRange = "'" + RESPONSES_SHEET + "'!" + colLetter + '2:' + colLetter;

  // 2) Надійний підрахунок у коді (для звірки / на випадок текстових дат)
  const lastRow = src.getLastRow();
  let codeCount = 0;
  if (lastRow >= 2) {
    const vals = src.getRange(2, dateCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const d = toDate_(vals[i][0]);
      if (d && d.getTime() >= SALARY_START.getTime()) codeCount++;
    }
  }

  // 3) Створити / очистити лист «Зарплата»
  let sh = ss.getSheetByName(SALARY_SHEET);
  if (!sh) sh = ss.insertSheet(SALARY_SHEET);
  else sh.clear();

  // 4) Параметри
  sh.getRange('A1').setValue('Розрахунок зарплати за огляди')
    .setFontWeight('bold').setFontSize(13);
  sh.getRange('A3').setValue('Ставка за лот, грн');
  sh.getRange('B3').setValue(RATE_PER_LOT);
  sh.getRange('A4').setValue('Рахувати огляди від');
  sh.getRange('B4').setValue(SALARY_START).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  sh.getRange('A5').setValue('Оглянуто лотів (рядків)');
  sh.getRange('B5').setFormula('=COUNTIFS(' + dateRange + ',">="&$B$4)');

  // 5) Три колонки-показники
  sh.getRange('A7:C7')
    .setValues([['Зароблена зарплата', 'Фактично видана', 'Різниця']])
    .setFontWeight('bold').setBackground('#e8eef7');
  sh.getRange('A8').setFormula('=$B$5*$B$3'); // зароблено (авто)
  sh.getRange('B8').setValue(0);              // видано — вписуєте вручну
  sh.getRange('C8').setFormula('=A8-B8');     // різниця (авто)

  sh.getRange('A8:C8').setNumberFormat('#,##0.00" грн"');
  sh.getRange('B8').setBackground('#fff7d6'); // жовта підсвітка = ручний ввід

  // 6) Якщо формула = 0, а лоти є → дати як текст: підставляємо код-підрахунок
  SpreadsheetApp.flush();
  const formulaCount = Number(sh.getRange('B5').getValue()) || 0;
  if (formulaCount === 0 && codeCount > 0) {
    sh.getRange('B5').setValue(codeCount);
    sh.getRange('A6')
      .setValue('⚠️ Дати в «' + SALARY_DATE_COL + '» як текст — лічильник статичний. '
              + 'Запускайте buildSalarySheet повторно для оновлення.')
      .setFontColor('#b26a00').setFontSize(9);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);

  Logger.log('Лист «Зарплата» готовий.');
  Logger.log('Лотів з ' + SALARY_START + ': ' + codeCount);
  Logger.log('Зароблено: ' + (codeCount * RATE_PER_LOT) + ' грн');
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
  if (typeof v === 'number') {                       // серійний номер Google Sheets
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (typeof v === 'string' && v.trim()) {
    const s = v.trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?/); // M/D/YYYY H:mm[:ss]
    if (m) return new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +(m[6] || 0));
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}
