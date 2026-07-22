/**
 * L-TEX — Розрахунок зарплати за огляди  (ДОПОВНЕННЯ до backend-скрипта)
 * ---------------------------------------------------------------------
 * Додайте цей блок У КІНЕЦЬ наявного standalone Apps Script (де є
 * getSpreadsheet_(), doPost(), appendResponse_() та константа RESPONSES_SHEET).
 *
 * Створює/оновлює лист «Зарплата» з трьома частинами:
 *
 *   1) ПІДСУМОК (вгорі)
 *        • Зароблена зарплата — кількість оглянутих лотів × ставка (авто)
 *        • Фактично видана    — СУМА журналу виплат (авто)
 *        • Різниця            — Зароблена − Фактично видана (авто)
 *
 *   2) ЖУРНАЛ ВИПЛАТ (зліва, з рядка 12) — заповнюєте ВРУЧНУ:
 *        Дата виплати | Сума, грн | Коментар
 *        Кожна виплата = окремий рядок; таблиця сама все підсумовує.
 *
 *   3) ЗАРОБІТОК ПО ДНЯХ (справа, з рядка 12) — рахується САМ формулою:
 *        День | Лотів | Зароблено, грн
 *
 * Рахуються лише огляди, у яких «Час відправки» >= SALARY_START.
 *
 * ЗАПУСК: редактор Apps Script → функція buildSalarySheet → «Запустити».
 *   Повторний запуск НЕ втрачає введені виплати — вони переносяться назад.
 */

// ===================== НАЛАШТУВАННЯ ЗАРПЛАТИ =====================
const SALARY_SHEET     = 'Зарплата';        // лист із розрахунком (створиться автоматично)
const SALARY_DATE_COL  = 'Час відправки';   // колонка з датою-часом огляду в аркуші «Анкети»
const RATE_PER_LOT     = 170;               // грн за 1 лот (1 рядок-огляд)
const SALARY_START     = new Date(2026, 5, 23, 14, 31, 41); // 23.06.2026 14:31:41 (місяць 0-based: 5 = червень)

const PAY_TITLE_ROW    = 10;  // рядок із заголовком «Журнал виплат»
const PAY_HEADER_ROW   = 11;  // рядок із шапкою журналу
const PAY_FIRST_ROW    = 12;  // з цього рядка йдуть самі виплати
// ================================================================

function buildSalarySheet() {
  const ss  = getSpreadsheet_();                  // ← ваш наявний helper (відкриває за SPREADSHEET_ID)
  const src = ss.getSheetByName(RESPONSES_SHEET); // 'Анкети'
  if (!src) throw new Error('Не знайдено лист «' + RESPONSES_SHEET + '»');

  // 1) Колонка з датою — за заголовком (стійко до перестановки стовпців)
  const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0].map(String);
  let dateCol = headers.indexOf(SALARY_DATE_COL) + 1;
  if (dateCol === 0) dateCol = 1;
  const c = columnToLetter_(dateCol);
  const dateRange = "'" + RESPONSES_SHEET + "'!" + c + '2:' + c;

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

  // 3) ЗБЕРЕГТИ вже введені виплати (щоб не втратити при повторному запуску)
  let savedPayments = [];
  const existing = ss.getSheetByName(SALARY_SHEET);
  if (existing) {
    const lr = existing.getLastRow();
    if (lr >= PAY_FIRST_ROW) {
      const rows = existing.getRange(PAY_FIRST_ROW, 1, lr - PAY_FIRST_ROW + 1, 3).getValues();
      savedPayments = rows.filter(r => r[0] !== '' || r[1] !== '' || r[2] !== '');
    }
    // Перенести стару разову суму «Фактично видана» (B8) як перший запис журналу
    if (savedPayments.length === 0) {
      const b8formula = existing.getRange('B8').getFormula();
      const b8value   = existing.getRange('B8').getValue();
      if (b8formula === '' && typeof b8value === 'number' && b8value > 0) {
        savedPayments.push(['', b8value, 'перенесено зі старого підсумку — впишіть дату']);
      }
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
  sh.getRange('A8').setFormula('=$B$5*$B$3');            // зароблено (авто)
  sh.getRange('B8').setFormula('=SUM(B' + PAY_FIRST_ROW + ':B)'); // фактично видано = сума журналу
  sh.getRange('C8').setFormula('=A8-B8');               // різниця (авто)
  sh.getRange('A8:C8').setNumberFormat('#,##0.00" грн"');

  // ---- ЖУРНАЛ ВИПЛАТ (ліворуч, ручний ввід) ----
  sh.getRange(PAY_TITLE_ROW, 1).setValue('Журнал виплат — додавайте рядки нижче')
    .setFontWeight('bold');
  sh.getRange(PAY_HEADER_ROW, 1, 1, 3)
    .setValues([['Дата виплати', 'Сума, грн', 'Коментар']])
    .setFontWeight('bold').setBackground('#fff2cc');
  // формати колонок журналу
  const payLen = maxRows - PAY_FIRST_ROW + 1;
  sh.getRange(PAY_FIRST_ROW, 1, payLen, 1).setNumberFormat('dd.MM.yyyy');      // дата
  sh.getRange(PAY_FIRST_ROW, 2, payLen, 1).setNumberFormat('#,##0.00" грн"');  // сума
  sh.getRange(PAY_FIRST_ROW, 1, payLen, 2).setBackground('#fffdf5');           // легка підсвітка поля вводу

  // повернути збережені виплати
  if (savedPayments.length) {
    sh.getRange(PAY_FIRST_ROW, 1, savedPayments.length, 3).setValues(savedPayments);
  }

  // ---- ЗАРОБІТОК ПО ДНЯХ (праворуч, авто) ----
  sh.getRange(PAY_TITLE_ROW, 5).setValue('Заробіток по днях').setFontWeight('bold');
  sh.getRange(PAY_HEADER_ROW, 5, 1, 3)
    .setValues([['День', 'Лотів', 'Зароблено, грн']])
    .setFontWeight('bold').setBackground('#e8eef7');

  const tz = ss.getSpreadsheetTimeZone();
  const dtLit = Utilities.formatDate(SALARY_START, tz, 'yyyy-MM-dd HH:mm:ss');
  // День + кількість лотів за день (спіл у E:F)
  const dailyQuery =
    '=IFERROR(QUERY(' + dateRange + ', "select toDate(' + c + '), count(' + c + ') ' +
    'where ' + c + " >= datetime '" + dtLit + "' " +
    'group by toDate(' + c + ') order by toDate(' + c + ') ' +
    "label toDate(" + c + ") '', count(" + c + ") ''\", 0), )";
  sh.getRange(PAY_FIRST_ROW, 5).setFormula(dailyQuery);
  // Зароблено за день = Лотів × ставка (окрема колонка G)
  sh.getRange(PAY_FIRST_ROW, 7).setFormula(
    '=ARRAYFORMULA(IF(LEN(F' + PAY_FIRST_ROW + ':F), F' + PAY_FIRST_ROW + ':F*$B$3, ""))');
  // формати
  const dayLen = maxRows - PAY_FIRST_ROW + 1;
  sh.getRange(PAY_FIRST_ROW, 5, dayLen, 1).setNumberFormat('dd.MM.yyyy');       // день
  sh.getRange(PAY_FIRST_ROW, 6, dayLen, 1).setNumberFormat('0');                // лотів
  sh.getRange(PAY_FIRST_ROW, 7, dayLen, 1).setNumberFormat('#,##0.00" грн"');   // зароблено

  // 5) Запобіжник: формула COUNTIFS = 0, а лоти є → дати як текст
  SpreadsheetApp.flush();
  const formulaCount = Number(sh.getRange('B5').getValue()) || 0;
  if (formulaCount === 0 && codeCount > 0) {
    sh.getRange('B5').setValue(codeCount);
    sh.getRange('A6')
      .setValue('⚠️ Дати в «' + SALARY_DATE_COL + '» як текст — підсумок статичний, запускайте buildSalarySheet повторно.')
      .setFontColor('#b26a00').setFontSize(9);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 7);

  Logger.log('Лист «Зарплата» готовий. Лотів з ' + SALARY_START + ': ' + codeCount +
             ' · Зароблено: ' + (codeCount * RATE_PER_LOT) + ' грн');
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
