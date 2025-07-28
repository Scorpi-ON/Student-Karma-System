const DATA_HEADER = [
  'Отметка времени',
  'Наименование',
  'Уровень',
  'Дата начала',
  'Длительность',
  'Твоя роль',
  'Файлы',
  'ФИО',
  'Курс',
  'Факультет',
  'Буква группы',
  'Профиль',
  'Профиль'
];
const OUTPUT_DATA_HEADER = [
  'Факультет',
  'Группа',
  'ФИО',
  'Роль',
  'Мероприятие',
  'Уровень',
  'Дата начала',
  'Дата окончания',
  'Баллы',
  'Файлы'
];
const DATE_INDEX = OUTPUT_DATA_HEADER.indexOf('Дата начала');
const RATING_INDEX = OUTPUT_DATA_HEADER.indexOf('Баллы');
const MIN_DATE = new Date(2023, 0, 1);
const REQUEST_SHEET_NAME = 'Заявки';
const ERROR_SHEET_NAME = 'Просроченные заявки'

const MILLENNIUM = 2000;
const SEPTEMBER_INDEX = 8;
const BACHELORS_DEGREE_TERM = 4;
const MASTER_LETTER = 'м';


Date.prototype.myFormat = function() {
  let day = this.getDate();
  let month = this.getMonth() + 1;
  let year = this.getFullYear();

  if (day < 10)
    day = '0' + day;
  if (month < 10)
    month = '0' + month;

  return `${day}.${month}.${year}`;
}


Date.fromMyFormat = function(string) {
  let [day, month, year] = string.split('.');
  --month;
  return new Date(year, month, day);
}


Date.prototype.addDaysToNew = function(days) {
  date = new Date(this);
  date.setDate(this.getDate() + days);
  return date;
}


function getAcademicGroup(profile, course, groupLetter) {
  let masterLetter = '';
  if (course > BACHELORS_DEGREE_TERM) {
    course -= BACHELORS_DEGREE_TERM;
    masterLetter = MASTER_LETTER;
  }

  const todayDate = new Date();
  const currentShortYear = todayDate.getFullYear() - MILLENNIUM;
  const isNewAcademicYear = todayDate.getMonth() - SEPTEMBER_INDEX >= 0;
  const year = currentShortYear - course - Number(isNewAcademicYear);
  
  return `${profile}${masterLetter}-${year}${groupLetter}`;
}


function parseRow(row) {
  const event = row[1];
  const level = row[2];

  const startDate = Date.fromMyFormat(row[3]);
  const duration = parseInt(row[4]) - 1;
  const endDate = duration ? startDate.addDaysToNew(duration) : '';

  const student_role = row[5];
  const files = row[6];
  const student = row[7];
  const course = parseInt(row[8]);
  const faculty = row[9];
  const groupLetter = row[10];
  const profile = row[11] || row[12];

  const group = getAcademicGroup(profile, course, groupLetter);

  return [
    faculty,
    group,
    student,
    student_role,
    event,
    level,
    startDate,
    endDate,
    '',  // Плейсхолдер для баллов
    files
  ];
}


function autoResizeSheetColumns(sheet) {
  const columnCount = sheet.getLastColumn();
  for (let i = 1; i <= columnCount; ++i)
    sheet.autoResizeColumn(i);
}


function getSheetDataRange(sheet) {
  const columnCount = sheet.getLastColumn();
  const dataRowCount = sheet.getLastRow() - 1;
  return sheet.getRange(2, 1, dataRowCount, columnCount);
}


function getSheetHeaderRange(sheet, columnCount = OUTPUT_DATA_HEADER.length) {
  return sheet.getRange(1, 1, 1, columnCount);
}


function setSheetHeader(sheet, header = OUTPUT_DATA_HEADER) {
  getSheetHeaderRange(sheet)
    .setValues([header])
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  autoResizeSheetColumns(sheet);
}


function getOutputSheet(name, clear = false) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let outputSheet = ss.getSheetByName(name);
  if (outputSheet === null) {
    outputSheet = ss.insertSheet(name, ss.getSheets().length);
    setSheetHeader(outputSheet);
  } else if (clear) {
    outputSheet.clear();
    setSheetHeader(outputSheet);
  } else {
    const header = getSheetHeaderRange(outputSheet).getValues()[0];
    if (header.length === 0) {
      setSheetHeader(outputSheet);
    } else if (header.toString() !== OUTPUT_DATA_HEADER.toString()) {
      throw Error(
        `Ошибка парсинга таблицы на листе "${outputSheet.getName()}". `
        + 'Приведите её в соответствие с форматом, '
        + 'с которым работает скрипт, и повторите попытку.'
      );
    }
  }
  //getSheetDataRange(outputSheet).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  return outputSheet;
}


function genReport() {
  const ui = SpreadsheetApp.getUi();

  const answer = ui.alert(
    'Подтверждение',
    'Перед генерацией отчёта убедитесь, что '
    + 'вы оценили все заявки и исправили ошибки. '
    + 'Не оценённые в баллах строки в отчёт не попадут.'
    + 'Если отчёт уже существует, он будет перезаписан. '
    + 'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  if (answer === ui.Button.NO)
    return;

  let requestSheet;
  let errorSheet;
  try {
    requestSheet = getOutputSheet(REQUEST_SHEET_NAME);
    errorSheet = getOutputSheet(ERROR_SHEET_NAME);
  } catch (e) {
    ui.alert(
      `${e.name}`,
      e.message,
      ui.ButtonSet.OK
    );
    return;
  }

  const rows = [
    ...getSheetDataRange(requestSheet).getValues(), 
    ...getSheetDataRange(errorSheet).getValues()
  ];
  if (rows.length === 0) {
    ui.alert(
      'Нечего добавлять в отчёт',
      'Не найдено ни одной заявки на карму. '
      + 'Видимо, форму ещё никто не заполнял.',
      ui.ButtonSet.OK
    );
    return;
  }

  const reportSheet = getOutputSheet(`Отчёт (${new Date().myFormat()})`, true);
  let ratedRows = 0;
  for (let i = 0; i < rows.length; ++i) {
    const row = rows[i];
    if (row[RATING_INDEX]) {
      reportSheet.appendRow(row);
      ++ratedRows;
    }
  }
  autoResizeSheetColumns(reportSheet);
  
  ui.alert(
    'Отчёт сгенерирован',
    `Всего обработано ${rows.length} строк. `
    + `Из них в отчёт попало ${ratedRows} оценённых вами заявок.`,
    ui.ButtonSet.OK
  );
}


function clearRequests() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();

  const answer = ui.alert(
    'Предупреждение',
    'Вы собираетесь очистить все имеющиеся  на данный момент заявки. '
    + 'Если вы это сделаете, они больше не будут попадать в отчёты. '
    + 'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  if (answer === ui.ButtonSet.NO)
    return;

  const sheetsToClear = [
    ss.getSheetByName(REQUEST_SHEET_NAME),
    ss.getSheetByName(ERROR_SHEET_NAME)
  ];
  for (let i = 0; i < sheetsToClear.length; ++i) {
    let sheet = sheetsToClear[i];
    if (sheet !== null) {
      sheet.clear();
      setSheetHeader(sheet);
    }
  }
}


function isRowValid(row) {
  return row[DATE_INDEX] >= MIN_DATE 
      && row[DATE_INDEX] <= new Date();
}


function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Запрограммированные действия')
    .addItem('Сгенерировать отчёт', genReport.name)
    .addItem('Удалить все заявки', clearRequests.name)
    .addToUi();
}


function onAppend(e) {
  const dataRow = parseRow(e.values);
  let sheet = getOutputSheet(
    isRowValid(dataRow) ? REQUEST_SHEET_NAME : ERROR_SHEET_NAME
  );
  sheet.appendRow(dataRow);
  autoResizeSheetColumns(sheet);
}


function onEdit(e) {
  if (e.range.getNote() === '' && e.oldValue !== e.value)
    e.range.setNote(
      e.oldValue ? e.oldValue : 'Пустая ячейка'
    );
}
