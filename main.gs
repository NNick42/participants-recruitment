// Copyright (c) 2017-2019, Sho Ishiguro (c) 2019-2021, Satoru Nishiyama and Sho Ishiguro
// Use of this source code is governed by the BSD 2-Clause License

const TYPE = 3; // 1: 自由回答, 2 or 3: 選択式 どちらかの半角数字を入れてください。

function init() {
  settings.init();
}

///////////////////////////////////////////////////////////////////////////////
// トリガー用の関数
///////////////////////////////////////////////////////////////////////////////

function onOpening() {
  try {
    if (sheets.isConfigured) {
      mail.alertFewMails();
      if (TYPE == 3) {
        SpreadsheetApp.getUi()
          .createMenu('カレンダー')
          .addItem('空き予定とフォームを更新', 'updateFormAndSchedule')
          .addItem('旧形式を複数人共同募集形式へ移行', 'migrateType3MultiResourceSchema')
          .addToUi();
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
  }
}

function onFormSubmission(e) {
  const lock = LockService.getScriptLock();
  let mailJob;
  try {
    lock.waitLock(30000);
    // systemを利用しないなら以降の処理を行わない
    if (settings.config.useFormSystem != 1) {
      return;
    }
    // 実際の回答に続けて値のない回答が送られることがあるので以下のif文で回避
    if (e.values[settings.config.colAddress].length > 0 && !settings.isDefault()) {
      booking.values = e.values;
      booking.setEventType(ScriptApp.EventType.ON_FORM_SUBMIT).validate().allocate(e.range.getRow());
      const { name, address, from: fromWhen, to: toWhen, trigger } = booking;
      mailJob = { name: name, address: address, from: fromWhen, to: toWhen, trigger: trigger };
      console.log('SUCCESS!');
    } else {
      console.log(e.values);
    }
  } catch (err) {
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    MailApp.sendEmail(settings.config.experimenterMailAddress, 'エラーが発生しました', msg);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
  if (mailJob !== undefined) {
    mail
      .create(mailJob.name, mailJob.trigger, mailJob.from, mailJob.to)
      .setBcc('', settings.config.selfBccTentative)
      .send(mailJob.address)
      .alertFewMails();
  }
}

function rangeTouchesHeaders(range, headers, names) {
  const first = range.getColumn() - 1;
  const last = range.getLastColumn() - 1;
  return names.some((name) => {
    const index = getHeaderIndex(headers, name);
    return index >= first && index <= last;
  });
}

function editedSettingKeys(range) {
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();
  if (lastRow < firstRow) {
    return [];
  }
  return range.getSheet().getRange(firstRow, 2, lastRow - firstRow + 1, 1).getValues().map((row) => String(row[0]));
}

function sendBookingMailJobs(jobs) {
  jobs.forEach((job) => {
    mail.create(job.name, job.trigger, job.from, job.to).setBcc(job.assistant, job.selfBcc).send(job.address).alertFewMails();
    if (job.markMailed) {
      sheets.sheets[0].getRange(job.row, settings.config.colMailed + 1).setValue(1);
    }
  });
}

function onSheetEdit(e) {
  const sh = e.range.getSheet();
  const sheetName = sh.getSheetName();
  const answerSheetName = sheets.sheets[0].getSheetName();
  const isAnswerStatusEdit =
    sheetName === answerSheetName && e.range.getColumn() <= settings.config.colStatus + 1 && settings.config.colStatus + 1 <= e.range.getLastColumn();
  let refreshType3 = false;
  let collectSheet;
  let oldConfig;

  if (sheetName === '設定') {
    oldConfig = copy(settings.config);
    collectSheet = sheetName;
    const availabilityKeys = new Set([
      'numberOfRooms',
      'openDate',
      'closeDate',
      'expTimeZone',
      'colExpDate',
      'colStatus',
      'colAssistant',
      'colRoom',
      'colAssignmentStatus',
      'colCalendarEventId',
    ]);
    refreshType3 = TYPE == 3 && editedSettingKeys(e.range).some((key) => availabilityKeys.has(key));
  } else if (sheetName === 'メンバー') {
    collectSheet = sheetName;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    refreshType3 = TYPE == 3 && (e.range.getRow() === 1 || rangeTouchesHeaders(e.range, headers, ['キー', 'カレンダーID', '有効']));
  } else if (sheetName === 'テンプレート') {
    collectSheet = sheetName;
  } else if (sheetName === '空き予定' && TYPE == 3) {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    refreshType3 = e.range.getRow() === 1 || rangeTouchesHeaders(e.range, headers, ['日付', '開始', '終了', '担当候補', '受付可']);
  }

  if (collectSheet !== undefined) {
    settings.collect(collectSheet, true).save();
  }
  const needsLegacySettingsUpdate = sheetName === '設定' && TYPE != 3;
  if (!isAnswerStatusEdit && !refreshType3 && !needsLegacySettingsUpdate) {
    return;
  }

  const lock = LockService.getScriptLock();
  const mailJobs = [];
  const ignoredEventIds = [];
  try {
    lock.waitLock(30000);
    sheets.ss.toast('スクリプトを実行しています。終了までお待ちください。');
    if (isAnswerStatusEdit) {
      const answers = sh.getDataRange().getValues();
      for (let row = e.range.getRow(); row <= e.range.getLastRow(); row++) {
        if (answers[row - 1][settings.config.colStatus] == '') {
          continue;
        }
        booking.values = answers[row - 1];
        if (TYPE == 3) {
          booking.setEventType(ScriptApp.EventType.ON_EDIT).validate().allocate(row);
          if (booking.didRelease && booking.releasedEventId !== '') {
            ignoredEventIds.push(booking.releasedEventId);
          }
          if (booking.shouldSendMail) {
            const { name, address, from: fromWhen, to: toWhen, trigger, assistant } = booking;
            mailJobs.push({
              row: row,
              name: name,
              address: address,
              from: fromWhen,
              to: toWhen,
              trigger: trigger,
              assistant: assistant,
              selfBcc: booking.isFinalization ? settings.config.selfBccFinalize : settings.config.selfBccTentative,
              markMailed: true,
            });
          }
        } else if (booking.values[settings.config.colMailed] !== 1) {
          booking.setEventType(ScriptApp.EventType.ON_EDIT).validate().allocate(row);
          const { name, address, from: fromWhen, to: toWhen, trigger, assistant } = booking;
          mailJobs.push({
            row: row,
            name: name,
            address: address,
            from: fromWhen,
            to: toWhen,
            trigger: trigger,
            assistant: assistant,
            selfBcc: settings.config.selfBccTentative,
            markMailed: false,
          });
        }
      }
      if (TYPE == 3 && ignoredEventIds.length > 0) {
        updateFormAndScheduleUnlocked({ ignoredEventIds: ignoredEventIds });
      }
    } else if (refreshType3) {
      updateFormAndScheduleUnlocked();
    }

    if (sheetName === '設定' && TYPE == 3 && settings.config.expTimeZone != oldConfig.expTimeZone) {
      sheets.ss.setSpreadsheetTimeZone(settings.config.expTimeZone);
      scriptTriggers.updateClockTrigger(settings.config.remindHour, settings.config.expTimeZone);
    }

    if (sheetName === '設定' && TYPE != 3) {
      if (settings.config.remindHour != oldConfig.remindHour) {
        scriptTriggers.updateClockTrigger(settings.config.remindHour, settings.config.expTimeZone);
      } else if (settings.config.expTimeZone != oldConfig.expTimeZone) {
        sheets.ss.setSpreadsheetTimeZone(settings.config.expTimeZone);
        scriptTriggers.updateClockTrigger(settings.config.remindHour, settings.config.expTimeZone);
      } else if (settings.config.workingCalendar != oldConfig.workingCalendar) {
        schedule.calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
        alertInitWithChangeOf('参照するカレンダー');
      } else if (settings.config.experimentLength != oldConfig.experimentLength) {
        alertInitWithChangeOf('実験の所要時間');
      } else if (fmtDate(settings.config.openTime, 'HH:mm') != fmtDate(new Date(oldConfig.openTime), 'HH:mm')) {
        alertInitWithChangeOf('実験の開始時刻');
      } else if (fmtDate(settings.config.closeTime, 'HH:mm') != fmtDate(new Date(oldConfig.closeTime), 'HH:mm')) {
        alertInitWithChangeOf('実験の終了時刻');
      }
    }
    sheets.ss.toast('スクリプトが終了しました。', '', 3);
  } catch (err) {
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
  sendBookingMailJobs(mailJobs);
}

function onClock() {
  try {
    // リマインダーの送信
    const answers = sheets.sheets[0].getDataRange().getValues();
    const timeNow = new Date();
    const tomorrowExps = [];
    for (let row = 0; row < answers.length; row++) {
      let ans = answers[row];
      if (ans[settings.config.colReminded] == '送信準備') {
        const remindDatetime = ans[settings.config.colRemindDate];
        if (is(remindDatetime, 'Date') && remindDatetime <= timeNow) {
          booking.values = ans;
          booking.setEventType(ScriptApp.EventType.CLOCK).allocate(row + 1);
          const { name, address, from: fromWhen, to: toWhen, assistant } = booking;
          mail.create(name, 'リマインダー', fromWhen, toWhen).setBcc(assistant, settings.config.selfBccReminder).send(address);
          tomorrowExps.push({ name: name, from: fromWhen, to: toWhen });
        }
      }
    }
    // 自分にもリマインダーを送る場合
    if (tomorrowExps.length > 0 && settings.config.sendTmrwExps > 0) {
      const tomorrow = new Date();
      tomorrow.setDate(new Date().getDate() + 1);
      tomorrowExps.sort((a, b) => {
        return a.from < b.from ? -1 : 1;
      });
      const time_table = tomorrowExps.map((exp) => {
        return `${fmtDate(exp.from, 'HH:mm')} - ${fmtDate(exp.to, 'HH:mm')} ${exp.name}`;
      });
      const body = time_table.join('\n');
      const title = `明日（${fmtDate(tomorrow, 'MM/dd')}）の実験予定`;
      MailApp.sendEmail(settings.config.experimenterMailAddress, title, body);
    }

    // フォームの修正
    if (settings.config.useFormSystem == 1) {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        if (TYPE == 3) {
          updateFormAndScheduleUnlocked();
        } else {
          form.modify();
        }
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    MailApp.sendEmail(settings.config.experimenterMailAddress, 'エラーが発生しました', msg);
  }
}

// function onCalendarUpdated() {
//   try {
//     // type 3の時だけ動作させる
//     if (TYPE != 3) {
//       return;
//     }
//     schedule.update().allocate(); // スケジュールを更新してシートに反映する
//     form.modify();
//   } catch (err) {
//     //実行に失敗した時に通知
//     const msg = `[${err.name}] ${err.stack}`;
//     console.error(msg);
//     dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
//   }
// }

///////////////////////////////////////////////////////////////////////////////
// Utility functions
///////////////////////////////////////////////////////////////////////////////

// 型判定のための関数https://qiita.com/Layzie/items/465e715dae14e2f601de より
function is(obj, type) {
  const clas = Object.prototype.toString.call(obj).slice(8, -1);
  return obj !== undefined && obj !== null && clas === type;
}

// 全角を半角に変換する関数
function zenToHan(str) {
  if (is(str, 'String')) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      // 全角を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 65248); // 10進数の場合
    });
  } else {
    return str;
  }
}

function numToColumnNotation(num) {
  const alphabet_upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let dgt = Math.floor(num / alphabet_upper.length);
  let remain = num % alphabet_upper.length;
  if (dgt < 1) {
    return alphabet_upper[remain];
  }
  return numToColumnNotation(dgt - 1) + alphabet_upper[remain];
}

function columnNotationToNum(notation) {
  const alphabet_upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  if (notation.length < 1) {
    return -1;
  } else if (notation.length == 1) {
    return alphabet_upper.indexOf(notation);
  }
  return (columnNotationToNum(notation.slice(0, -1)) + 1) * alphabet_upper.length + alphabet_upper.indexOf(notation.slice(-1));
}

function fmtDate(datetime, pattern) {
  if (is(datetime, 'Date')) {
    if (/yobi/.test(pattern)) {
      var yobi = new Array('日', '月', '火', '水', '木', '金', '土')[datetime.getDay()];
      pattern = pattern.replace(/yobi/, yobi);
    }
    return Utilities.formatDate(datetime, settings.config.expTimeZone, pattern);
  }
  return datetime;
}

function copy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function getHeaderIndex(headers, name) {
  return headers.indexOf(name);
}

function parseKeys(value) {
  const matches = String(zenToHan(value || '')).match(/\d+/g);
  return matches === null ? [] : matches;
}

function isEnabled(value) {
  return String(zenToHan(value)).trim() === '1';
}

function dateFromSlotParts(dateValue, timeValue) {
  const date = new Date(dateValue);
  if (is(timeValue, 'Date')) {
    date.setHours(timeValue.getHours(), timeValue.getMinutes(), 0, 0);
    return date.toString() == 'Invalid Date' ? undefined : date;
  }
  const time = String(zenToHan(timeValue || '')).match(/\d+/g);
  if (date.toString() == 'Invalid Date' || time === null || time.length < 2) {
    return undefined;
  }
  date.setHours(Number(time[0]), Number(time[1]), 0, 0);
  return date;
}

function parseType3Slot(value) {
  const numbers = String(zenToHan(value || '')).match(/\d+/g);
  if (numbers === null || numbers.length !== 7) {
    return undefined;
  }
  const from = new Date();
  from.setFullYear(Number(numbers[0]), Number(numbers[1]) - 1, Number(numbers[2]));
  from.setHours(Number(numbers[3]), Number(numbers[4]), 0, 0);
  const to = new Date(from);
  to.setHours(Number(numbers[5]), Number(numbers[6]), 0, 0);
  return { from: from, to: to };
}

function type3SlotLabel(from, to) {
  return `${fmtDate(from, 'yyyy/MM/dd HH:mm')}-${fmtDate(to, 'HH:mm')}`;
}

function intervalsOverlap(fromA, toA, fromB, toB) {
  return fromA < toB && fromB < toA;
}

// https://qiita.com/jz4o/items/d4e978f9085129155ca6 を改変
function isHoliday(time) {
  //土日か判定
  let weekInt = time.getDay();
  if (weekInt <= 0 || 6 <= weekInt) {
    return true;
  }

  //祝日か判定
  const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  if (calendar == null) {
    let msg = '祝日のカレンダーがご自身のgoogleカレンダーに登録されていません。実験日の休日判定を行う場合は祝日カレンダーを登録してください。';
    msg += 'もし休日判定を行わない場合は「テンプレート」シートのB列の数字をすべて0に変更してください。';
    msg += '\n\nなおこのエラーが発生したため参加者にはメールは送られていません。休日判定のための設定したのち';
    msg += '予約ステータスを含む右4列の内容を削除し再度予約ステータスにトリガーを入力してください。';
    throw new Error(msg);
  }
  const todayEvents = calendar.getEventsForDay(time);

  return todayEvents.length > 0;
}

function alertInitWithChangeOf(changed) {
  if (TYPE != 3) {
    return;
  }
  const choice = dlg.alert(`${changed}が変更されました`, '空き予定を初期化しますか？', dlg.ui.ButtonSet.OK_CANCEL);
  if (choice == dlg.ui.Button.OK) {
    schedule.init();
    form.modify();
  }
}

function updateFormAndSchedule() {
  // type 3の時だけ動作させる
  if (TYPE != 3) {
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    updateFormAndScheduleUnlocked();
  } finally {
    lock.releaseLock();
  }
}

function updateFormAndScheduleUnlocked(options) {
  if (TYPE == 3) {
    const snapshot = multiSchedule.buildSnapshot(options);
    multiSchedule.applySnapshot(snapshot);
    return;
  }
  schedule.update().allocate(); // スケジュールを更新してシートに反映する
  form.modify();
}

function migrateType3MultiResourceSchema() {
  if (TYPE != 3) {
    throw new Error('複数実験者・実験室の移行は TYPE = 3 でのみ実行できます。');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    settings.collect('設定', true);
    multiSchedule.assertMigrationIsSafe();
    settings.ensureType3Schema();
    multiSchedule.migrateLegacySlots();
    updateFormAndScheduleUnlocked();
  } finally {
    lock.releaseLock();
  }
}

///////////////////////////////////////////////////////////////////////////////
// Objects
///////////////////////////////////////////////////////////////////////////////

// シートをいい感じに扱いやすくしてくれるはずのオブジェクト
const sheets = (function () {
  const __ss = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet
  let __sheets = __ss.getSheets();
  let __name_idx = new Map();
  __sheets.forEach((sh, idx) => __name_idx.set(sh.getName(), idx));

  let __values = { 設定: undefined, テンプレート: undefined, メンバー: undefined, 空き予定: undefined, Cached: undefined };

  return {
    get length() {
      return __sheets.length;
    },
    get sheets() {
      return __sheets;
    },
    get isConfigured() {
      return Array.from(__name_idx.keys()).includes('設定');
    },
    get ss() {
      return __ss;
    },

    update: function () {
      __sheets = __ss.getSheets();
      __name_idx = new Map();
      __sheets.forEach((sh, idx) => __name_idx.set(sh.getName(), idx));
      return this;
    },

    getSheetByName: function (name) {
      if (__name_idx.get(name) === undefined) {
        new Error(`「${name}」シートがありません。`);
      }
      return __sheets[__name_idx.get(name)];
    },

    getValuesOf: function (name, update = false) {
      if (__values[name] === undefined || update) {
        const sh = this.getSheetByName(name);
        __values[name] = sh.getDataRange().getValues();
      }
      return __values[name];
    },

    getValueAt: function (sheetName, row, col) {
      const values = this.getValuesOf(sheetName);
      return values[row][col];
    },

    getTargetRowID(sheetName, col, target) {
      let sheetValues = this.getValuesOf(sheetName);
      for (let row = 0; row < sheetValues.length; row++) {
        var rowValues = sheetValues[row];
        if (rowValues[col] == target) {
          return row + 1; // getRangeで使うことを想定しているので，+1する
        }
      }
      return undefined;
    },
  };
})();

// 設定をいい感じに扱いやすくしてくれるはずのオブジェクト
const settings = (function () {
  let __settings = {};
  const __name_to_key = new Map([
    ['設定', 'config'],
    ['テンプレート', 'templates'],
    ['メンバー', 'members'],
  ]);

  function __getConfig(update) {
    const table = sheets.getValuesOf('設定', update);
    __settings.config = {};
    for (let row = 1; row < table.length; row++) {
      let key = table[row][1];
      let val = zenToHan(table[row][2]); // 念の為
      if (key.indexOf('col') == 0) {
        const is_alphabet = new RegExp(/^[a-zA-Z]*$/);
        if (!is_alphabet.test(val)) {
          throw new Error(
            `${key}に英字以外が入力されています。列に関する設定には英字を入力してください。一見，英字を入力しているにもかかわらずこのエラーが表示される場合は，入力内容にスペースが含まれているかもしれません。`
          );
        }
        val = columnNotationToNum(val.toUpperCase()); // 列番号に関する設定は，Numberに変更しておく
      }
      __settings.config[key] = val;
    }
    __arrangeExpPeriod();
  }

  function __getMailTemplates(update) {
    const table = sheets.getValuesOf('テンプレート', update);
    __settings.templates = {};
    for (let row = 1; row < table.length; row++) {
      let key = table[row][0];
      let property = {};
      property.changeByDay = table[row][1];
      property.title = table[row][2];
      property.bodywd = table[row][3];
      property.bodywe = table[row][4];
      __settings.templates[key] = property;
    }
  }

  function __getMembers(update) {
    const table = sheets.getValuesOf('メンバー', update);
    __settings.members = {};
    __settings.memberDetails = {};
    const headers = table[0] || [];
    const keyCol = getHeaderIndex(headers, 'キー');
    const nameCol = getHeaderIndex(headers, '名前');
    const addressCol = getHeaderIndex(headers, 'アドレス');
    const calendarCol = getHeaderIndex(headers, 'カレンダーID');
    const enabledCol = getHeaderIndex(headers, '有効');
    for (let row = 1; row < table.length; row++) {
      let key = zenToHan(table[row][keyCol]);
      let address = zenToHan(table[row][addressCol]);
      if (String(key).trim() === '') {
        continue;
      }
      __settings.members[key] = address;
      __settings.memberDetails[key] = {
        key: String(key),
        name: nameCol >= 0 ? table[row][nameCol] : '',
        address: address,
        calendarId: calendarCol >= 0 ? String(zenToHan(table[row][calendarCol] || '')).trim() : '',
        enabled: enabledCol >= 0 && isEnabled(table[row][enabledCol]),
      };
    }
  }

  function __appendAnswerColumns() {
    const sh = sheets.sheets[0];
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const additions = ['実験室', '割当状態', 'CalendarEventId'];
    const missing = additions.filter((header) => !headers.includes(header));
    if (missing.length > 0) {
      sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    }
  }

  function __appendConfigRows() {
    const sh = sheets.getSheetByName('設定');
    const table = sh.getDataRange().getValues();
    const existingRows = new Map();
    table.slice(1).forEach((row, index) => existingRows.set(row[1], index + 2));
    const answerHeaders = sheets.sheets[0].getRange(1, 1, 1, sheets.sheets[0].getLastColumn()).getValues()[0];
    const required = [
      ['実験室数 (1-5)', 'numberOfRooms', 1],
      ['実験室の列', 'colRoom', numToColumnNotation(getHeaderIndex(answerHeaders, '実験室'))],
      ['割当状態の列', 'colAssignmentStatus', numToColumnNotation(getHeaderIndex(answerHeaders, '割当状態'))],
      ['CalendarEventId の列', 'colCalendarEventId', numToColumnNotation(getHeaderIndex(answerHeaders, 'CalendarEventId'))],
    ];
    const missing = required.filter((row) => !existingRows.has(row[1]));
    if (missing.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
    }
    required.slice(1).forEach((row) => {
      const existingRow = existingRows.get(row[1]);
      if (existingRow !== undefined) {
        sh.getRange(existingRow, 3).setValue(row[2]);
      }
    });
  }

  function __appendMemberColumns() {
    const sh = sheets.getSheetByName('メンバー');
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const additions = ['カレンダーID', '有効'].filter((header) => !headers.includes(header));
    if (additions.length > 0) {
      sh.getRange(1, headers.length + 1, 1, additions.length).setValues([additions]);
    }
  }

  function __collect(sheetName, update) {
    switch (sheetName) {
      case '設定':
        __getConfig(update);
        break;
      case 'テンプレート':
        __getMailTemplates(update);
        break;
      case 'メンバー':
        __getMembers(update);
        break;
      default:
        throw new Error(`${sheetName} は「設定」「テンプレート」「メンバー」のいずれにも一致しません`);
    }
  }

  function __retrieve() {
    const sh = sheets.getSheetByName('Cached');
    const cache_json = sh.getRange(1, 1).getValue();
    if (cache_json.length < 10) {
      // cache用JSONが存在しない（適切ではない）場合
      for (let sheetName of __name_to_key.keys()) {
        __collect(sheetName);
      }
    } else {
      __settings = JSON.parse(cache_json);
      __arrangeExpPeriod();
    }
  }

  function __arrangeExpPeriod() {
    // parseしたままだと以下の2つがstringのままで機能しない。
    // シートから直接値を取得した場合は問題ないが，それほど処理速度に影響が出るとも思えないので，処理を分けない
    __settings.config.openDate = new Date(__settings.config.openDate);
    __settings.config.closeDate = new Date(__settings.config.closeDate);

    // openTime, closeTime
    let temp_date = new Date();
    if (is(__settings.config.openTime, 'String')) {
      __settings.config.openTime = new Date(__settings.config.openTime);
    } else if (is(__settings.config.openTime, 'Number')) {
      temp_date.setHours(__settings.config.openTime, 0, 0, 0);
      __settings.config.openTime = new Date(temp_date);
    }
    if (is(__settings.config.closeTime, 'String')) {
      __settings.config.closeTime = new Date(__settings.config.closeTime);
    } else if (is(__settings.config.closeTime, 'Number')) {
      temp_date.setHours(__settings.config.closeTime, 0, 0, 0);
      __settings.config.closeTime = new Date(temp_date);
    }
    if (__settings.config.openTime.toString() == 'Invalid Date') {
      throw new Error('開始時刻の設定が適切ではありません。"10:00" のような時間表記 あるいは "10" のような"時"だけを示す数値を入力してください。');
    } else if (__settings.config.closeTime.toString() == 'Invalid Date') {
      throw new Error('終了時刻の設定が適切ではありません。"10:00" のような時間表記 あるいは "10" のような"時"だけを示す数値を入力してください。');
    }

    // 実験開始日・終了日の調整
    __settings.config.outOfDate = false;
    const now = new Date();
    if (__settings.config.openDate < now) {
      __settings.config.openDate = now;
    }
    if (__settings.config.closeDate < now) {
      __settings.config.outOfDate = true;
    }
    // 実験開始日・終了日の日時の設定
    const openHour = __settings.config.openTime.getHours();
    const openMin = __settings.config.openTime.getMinutes();
    const closeHour = __settings.config.closeTime.getHours();
    const closeMin = __settings.config.closeTime.getMinutes();
    __settings.config.openDate.setHours(openHour, openMin, 0, 0);
    __settings.config.closeDate.setHours(closeHour, closeMin, 0, 0);
  }

  if (sheets.isConfigured) {
    __retrieve();
  }

  return {
    get config() {
      if (__settings.config === undefined) {
        this.collect('設定');
      }
      return __settings.config;
    },
    get templates() {
      if (__settings.templates === undefined) {
        this.collect('テンプレート');
      }
      return __settings.templates;
    },
    get members() {
      if (__settings.members === undefined) {
        this.collect('メンバー');
      }
      return __settings.members;
    },
    get activeMembers() {
      if (__settings.memberDetails === undefined) {
        this.collect('メンバー', true);
      }
      const members = Object.values(__settings.memberDetails).filter((member) => member.enabled && member.calendarId.length > 0);
      if (members.length > 5) {
        throw new Error('有効な実験者は最大5名です。「メンバー」シートの「有効」を見直してください。');
      }
      return members;
    },
    memberByKey: function (key) {
      if (__settings.memberDetails === undefined) {
        this.collect('メンバー', true);
      }
      return __settings.memberDetails[String(key)];
    },

    collect: function (sheetName, update = false) {
      __collect(sheetName, update);
      return this;
    },

    retrieve: function () {
      __retrieve();
      return this;
    },

    save: function () {
      const sh = sheets.getSheetByName('Cached');
      sh.getRange(1, 1).setValue(JSON.stringify(__settings));
    },

    ensureType3Schema: function () {
      if (TYPE != 3) {
        return this;
      }
      __appendAnswerColumns();
      __appendConfigRows();
      __appendMemberColumns();
      sheets.getValuesOf('設定', true);
      sheets.getValuesOf('メンバー', true);
      this.collect('設定', true).collect('メンバー', true).save();
      return this;
    },

    isDefault: function () {
      const is_default = {
        実験者名: this.config.experimenterName == '実験太郎',
        電話番号: this.config.experimenterPhone == 'xxx-xxx-xxx',
        実施場所: this.config.experimentRoom == '実施場所',
      };

      const title = '設定がデフォルトのままです';
      let msg = '以下の重要な設定がデフォルトのままだったので，参加希望者への予約確認メールの送信を中止しました。\n\n';
      let is_any_default = false;
      for (const key in is_default) {
        if (is_default[key]) {
          msg += `${key}\n`;
          is_any_default = true;
        }
      }

      if (is_any_default) {
        msg += '\n変更後，再度参加者応募のテストをして，予約確認のメールが送信されるかどうか，およびその本文が適切かどうかを確認してください。';
        MailApp.sendEmail(this.config.experimenterMailAddress, title, msg);
      }

      return is_any_default;
    },
  };
})();

// メールの内容やらを作成するためのオブジェクト
const mail = (function () {
  let __title;
  let __body;
  let __bcc;
  let __remaining;

  return {
    get remaining() {
      if (__remaining === undefined) {
        __remaining = MailApp.getRemainingDailyQuota();
      }
      const sh = sheets.getSheetByName('設定');
      const row_id = sheets.getTargetRowID('設定', 1, 'remainingMails');

      sh.getRange(row_id, 3).setValue(__remaining);
      return __remaining;
    },

    create: function (name, trigger, from, to) {
      const template = settings.templates[trigger];

      // タイトル
      __title = template.title;

      // 本文
      __body = template.bodywd;
      if (template.changeByDay == 1 && isHoliday(from)) {
        __body = template.bodywe;
      }
      const config = copy(settings.config);
      config.participantName = name;
      config.expDate = fmtDate(from, 'MM/dd（yobi）');
      config.fromWhen = fmtDate(from, 'HH:mm');
      config.toWhen = fmtDate(to, 'HH:mm');
      config.openDate = fmtDate(settings.config.openDate, 'yyyy/MM/dd');
      config.closeDate = fmtDate(settings.config.closeDate, 'yyyy/MM/dd');
      // メールの本文の変数を置換する
      for (const key in config) {
        let regex = new RegExp(key, 'g');
        __body = __body.replace(regex, config[key]);
      }

      return this;
    },

    setBcc: function (assistants, selfBcc) {
      const bcc_array = [];
      if (selfBcc > 0) {
        bcc_array.push(settings.config.experimenterMailAddress);
      }
      assistants = String(zenToHan(assistants));
      // 担当が空欄でなければ
      if (assistants.length > 0) {
        const assistantIDs = assistants.match(/\d+/g);
        assistantIDs.forEach((ast_id) => bcc_array.push(settings.members[ast_id]));
      }
      __bcc = bcc_array.join(','); // 配列が空なら''が返される

      return this;
    },

    send: function (address) {
      if (__bcc.length > 5) {
        MailApp.sendEmail(address, __title, __body, { bcc: __bcc });
      } else {
        MailApp.sendEmail(address, __title, __body);
      }

      return this;
    },

    alertFewMails() {
      const thresholds = [5, 10, 20];
      if (thresholds.includes(this.remaining)) {
        const title = '自動送信メールの残数が' + String(this.remaining) + 'です。';
        const message =
          title +
          'この24時間以内に送信されるかもしれない予約の確認やリマインダーのメール数を考慮して予約を完了させてください。' +
          '自分や分担者にもメールが送信されるようにしている場合は1通あたりに減る数が 2, 3... 大きくなります。';
        dlg.alert(title, message, dlg.ui.ButtonSet.OK);
        // Browser.msgBox(title, message, Browser.Buttons.OK);
      }
    },
  };
})();

// フォームを扱うオブジェクト
const form = (function () {
  let __form;

  function __modifyType2() {
    if (__form === undefined) {
      __form = FormApp.openByUrl(sheets.ss.getFormUrl());
    }
    const items = __form.getItems();
    const itemForDate = items[settings.config.colExpDate - 1]; // -1 なのは，シートで フォームの送信時間が増えているから
    let item;
    if (itemForDate.getType() == 'LIST') {
      item = itemForDate.asListItem();
    } else if (itemForDate.getType() == 'MULTIPLE_CHOICE') {
      item = itemForDate.asMultipleChoiceItem();
    } else {
      return;
    }

    let firstDateOfChoices = new Date(settings.config.openDate);
    firstDateOfChoices.setHours(0, 0, 0, 0);
    settings.config.closeDate.setHours(0, 0, 0, 0);
    // 設定された実験の開始日が関数の動作日時よりも前の場合
    if (firstDateOfChoices < new Date()) {
      firstDateOfChoices.setDate(new Date().getDate() + 1);
    }
    const choices = [];
    for (const choiceDate = firstDateOfChoices; choiceDate <= settings.config.closeDate; choiceDate.setDate(choiceDate.getDate() + 1)) {
      const newChoice = item.createChoice(fmtDate(choiceDate, 'yyyy/MM/dd'));
      choices.push(newChoice);
    }
    item.setChoices(choices);
  }

  function __modifyType3(snapshot) {
    if (__form === undefined) {
      __form = FormApp.openByUrl(sheets.ss.getFormUrl());
    }
    const items = __form.getItems();
    const itemForDate = items[settings.config.colExpDate - 1]; // -1 なのは，シートで フォームの送信時間が増えているから
    let item;
    if (itemForDate.getType() == 'LIST') {
      item = itemForDate.asListItem();
    } else if (itemForDate.getType() == 'MULTIPLE_CHOICE') {
      item = itemForDate.asMultipleChoiceItem();
    } else {
      return;
    }

    if (snapshot === undefined) {
      throw new Error('TYPE 3 のフォーム更新には availability snapshot が必要です。');
    }
    const labels = snapshot.choices;
    if (labels.length === 0) {
      __form.setAcceptingResponses(false);
      return;
    }
    __form.setAcceptingResponses(true);
    const choices = labels.map((label) => item.createChoice(label));
    item.setChoices(choices);
  }

  return {
    modify: function (snapshot) {
      // 実験実施期間を過ぎていたらフォームを閉じる
      if (settings.config.outOfDate) {
        if (__form === undefined) {
          __form = FormApp.openByUrl(sheets.ss.getFormUrl());
        }
        __form.setAcceptingResponses(false);
        return;
      }
      switch (TYPE) {
        case 2:
          return __modifyType2();
        case 3:
          return __modifyType3(snapshot);
        default:
          return;
      }
    },
  };
})();

// 予約情報を扱うオブジェクト
const booking = (function () {
  let __values;
  let __name;
  let __address;
  let __from;
  let __to;
  let __valid = false;
  let __trigger;
  let __status;
  let __event_type;
  let __calendar;
  let __finalizeTriggers;
  let __type3Allocation;
  let __isFinalization = false;
  let __isNoop = false;
  let __shouldSendMail = false;
  let __didRelease = false;
  let __releasedEventId = '';
  if (sheets.isConfigured) {
    __calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
    __finalizeTriggers = String(settings.config.finalizeTrigger).match(/\d+/g);
  }

  function __isValidDatetime() {
    if (__from === undefined || __to === undefined) {
      return false;
    }
    settings.config.openTime.setFullYear(__from.getFullYear(), __from.getMonth(), __from.getDate());
    settings.config.closeTime.setFullYear(__from.getFullYear(), __from.getMonth(), __from.getDate());
    const isValidTime = settings.config.openTime <= __from && __to <= settings.config.closeTime;
    const isValidDate = settings.config.openDate <= __from && __from <= settings.config.closeDate;
    return isValidTime && isValidDate;
  }

  function __validateSubmission() {
    __isFinalization = false;
    __isNoop = false;
    __shouldSendMail = false;
    __didRelease = false;
    __releasedEventId = '';
    if (TYPE == 3) {
      __type3Allocation = multiSchedule.prepareTentative(__values);
      __trigger = __type3Allocation.trigger;
      __valid = __type3Allocation.valid;
      __status = __valid ? ['', '', '', ''] : [__trigger, 1, 'N/A', 'N/A'];
      return;
    }
    if (__from === undefined || __to === undefined) {
      return undefined;
    }
    const events = __calendar.getEvents(__from, __to);
    __trigger = '仮予約';
    __status = ['', '', '', ''];
    __valid = true;
    if (events.length > 0) {
      __trigger = '重複';
      __status = [__trigger, 1, 'N/A', 'N/A'];
      __valid = false;
    } else if (!__isValidDatetime()) {
      __trigger = '時間外';
      __status = [__trigger, 1, 'N/A', 'N/A'];
      __valid = false;
    }
  }

  function __validateEdit() {
    __trigger = String(__values[settings.config.colStatus]);
    const validTriggers = Object.keys(settings.templates);
    __isFinalization = false;
    __isNoop = false;
    __shouldSendMail = false;
    __didRelease = false;
    __releasedEventId = '';

    if (TYPE == 3) {
      const assignmentState = String(__values[settings.config.colAssignmentStatus]);
      const hasMailed = __values[settings.config.colMailed] === 1;
      if (__finalizeTriggers.includes(__trigger)) {
        if (assignmentState === '確定') {
          __isNoop = true;
          __isFinalization = true;
          __shouldSendMail = !hasMailed;
          return;
        }
        if (assignmentState !== '仮割当') {
          throw new Error('111 で最終確認できるのは「仮割当」の予約だけです。新しい pair の再割当は行いません。');
        }
        const remindDate = new Date(__from);
        remindDate.setDate(__from.getDate() - 1);
        const today = new Date();
        today.setHours(19);
        __status = [1, remindDate, remindDate <= today ? '直前のため省略' : '送信準備'];
        __valid = true;
        __isFinalization = true;
        __shouldSendMail = !hasMailed;
        return;
      }
      if (!validTriggers.includes(__trigger)) {
        throw new Error('予約ステータスに入力された文字列（トリガー）が「テンプレート」に存在しないため，メールの送信等の処理は行われませんでした。');
      }
      if (assignmentState === '解放') {
        __isNoop = true;
        __shouldSendMail = !hasMailed;
        return;
      }
      if (assignmentState !== '仮割当' && assignmentState !== '確定') {
        throw new Error('222/333 などで解放できるのは「仮割当」または「確定」の予約だけです。');
      }
      __status = [1, 'N/A', 'N/A'];
      __valid = false;
      __shouldSendMail = true;
      return;
    }

    if (__finalizeTriggers.includes(__trigger)) {
      // 予約確定のトリガーなら
      // リマインダーの設定
      const remindDate = new Date(__from);
      remindDate.setDate(__from.getDate() - 1);
      const today = new Date();
      today.setHours(19);
      __status = [1, remindDate, '送信準備'];
      if (remindDate <= today) {
        // リマインド日が，予約確定させた日の19時よりも前の場合
        __status[2] = '直前のため省略';
      }
      __valid = true;
    } else if (validTriggers.includes(__trigger)) {
      // 予約確定トリガーではないが，有効なトリガーの場合
      __status = [1, 'N/A', 'N/A'];
      __valid = false; // トリガーはvalidだが，実験の応募はvalidではない
    } else {
      // 登録されたトリガーではない場合
      throw new Error('予約ステータスに入力された文字列（トリガー）が「テンプレート」に存在しないため，メールの送信等の処理は行われませんでした。');
    }
  }

  function __allocateOnSubmission(numRow) {
    if (TYPE == 3) {
      multiSchedule.commitTentative(numRow, __values, __type3Allocation);
      return;
    }
    sheets.sheets[0].getRange(numRow, settings.config.colStatus + 1, 1, __status.length).setValues([__status]);
    // カレンダーの編集
    if (__valid) {
      const eventTitle = '仮予約: ' + __name;
      __calendar.createEvent(eventTitle, __from, __to);
    }
  }

  function __allocateOnEdit(numRow) {
    if (TYPE == 3) {
      if (__isNoop) {
        return;
      }
      if (__isFinalization) {
        multiSchedule.finalize(numRow, __values);
      } else {
        __releasedEventId = multiSchedule.release(numRow, __values);
        __didRelease = true;
      }
      if (!__isNoop) {
        sheets.sheets[0].getRange(numRow, settings.config.colRemindDate + 1, 1, 2).setValues([[__status[1], __status[2]]]);
      }
      return;
    }
    sheets.sheets[0].getRange(numRow, settings.config.colMailed + 1, 1, __status.length).setValues([__status]);
    // カレンダーの編集
    // まず予約イベントを削除する
    const events = __calendar.getEvents(__from, __to);
    events.forEach((e) => {
      if (e.getTitle().includes(__name)) e.deleteEvent();
    });
    if (__valid) {
      //予約確定情報をカレンダーに追加
      let newEventName = '予約完了:' + __name;
      if (settings.config.colParNameKana > 0) {
        newEventName = newEventName + '(' + __values[settings.config.colParNameKana] + ')';
      }
      __calendar.createEvent(newEventName, __from, __to);
    }
  }

  function __allocateOnTime(numRow) {
    sheets.sheets[0].getRange(numRow, settings.config.colReminded + 1).setValue('送信済み'); // シートの修正
  }

  function __fmtExpDateTimeType1() {
    const date = __values[settings.config.colExpDate];
    __from = new Date(date);
    __to = new Date(__from);
    __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
  }

  function __fmtExpDateTimeType2() {
    const date = zenToHan(__values[settings.config.colExpDate]);
    const time = zenToHan(__values[settings.config.colExpTime]);
    // 日付の処理
    __from = new Date();
    const date_info = date.match(/\d+/g); // 数字の部分だけを取り出す
    if (date_info.length == 3) {
      const [year, month, day] = date_info;
      __from.setFullYear(year, month - 1, day);
    } else if (date_info.length == 2) {
      const [month, day] = date_info;
      __from.setMonth(month - 1, day);
    } else if (date_info.length == 1) {
      const [day] = date_info;
      __from.setDate(day);
    }

    __to = new Date(__from);

    // 時間の処理
    const from_to = time.match(/\d+/g); // 数字の部分だけを取り出す
    if (from_to.length == 4) {
      // timeが hh:mm-hh:mm 形式なら
      const [fromHour, fromMin, toHour, toMin] = from_to;
      __from.setHours(fromHour, fromMin);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 2) {
      // timeが hh:mm 形式なら
      const [fromHour, fromMin] = from_to;
      __from.setHours(fromHour, fromMin);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    }
  }
  /*
    yyyy-MM-dd HH:mm -> 5
    yyyy/MM/dd HH:mm
    yyyy/MM/dd HH時mm分
    yyyy年MM月dd日 HH時mm分
    yyyy年MM月dd日HH時mm分

    MM/dd HH:mm -> 4
    MM月dd日HH時mm分 -> 4

    yyyy/MM/dd HH:mm-HH:mm -> 7
    yyyy年MM月dd日HH時mm分-HH時mm分
    
    HH:mm-HH:mm -> 4
  */
  function __fmtExpDateTimeType3() {
    const datetime = zenToHan(__values[settings.config.colExpDate]);
    __from = new Date();
    const from_to = datetime.match(/\d+/g); // 数字の部分だけを取り出す
    if (from_to.length == 7) {
      // timeが yyyy/MM/dd HH:mm-HH:mm 形式なら
      const [year, month, day, fromHour, fromMin, toHour, toMin] = from_to;
      __from.setFullYear(year, month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 6) {
      // timeが MM/dd HH:mm-HH:mm 形式なら
      const [month, day, fromHour, fromMin, toHour, toMin] = from_to;
      __from.setFullYear(month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 5) {
      // timeが yyyy-MM-dd HH:mm 形式なら
      const [year, month, day, fromHour, fromMin] = from_to;
      __from.setFullYear(year, month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    } else if (from_to.length == 4) {
      // timeが MM-dd HH:mm 形式なら
      const [month, day, fromHour, fromMin] = from_to;
      __from.setFullYear(month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    }
  }

  return {
    get name() {
      return __name;
    },
    get address() {
      return __address;
    },
    get from() {
      return __from;
    },
    get to() {
      return __to;
    },
    set values(val) {
      __values = val;
      __name = __values[settings.config.colParName];
      __address = __values[settings.config.colAddress];
      if (TYPE == 1) {
        __fmtExpDateTimeType1();
      } else if (TYPE == 2) {
        __fmtExpDateTimeType2();
      } else if (TYPE == 3) {
        const slot = parseType3Slot(__values[settings.config.colExpDate]);
        if (slot === undefined) {
          __from = undefined;
          __to = undefined;
        } else {
          __from = slot.from;
          __to = slot.to;
        }
      }
      if (__from !== undefined) {
        __from.setSeconds(0, 0);
        __to.setSeconds(0, 0);
      }
    },
    get values() {
      return __values;
    },
    get trigger() {
      return __trigger;
    },
    get status() {
      return __status;
    },
    get assistant() {
      return __values[settings.config.colAssistant];
    },
    get room() {
      return TYPE == 3 ? __values[settings.config.colRoom] : undefined;
    },
    get isFinalization() {
      return __isFinalization;
    },
    get shouldSendMail() {
      return __shouldSendMail;
    },
    get didRelease() {
      return __didRelease;
    },
    get releasedEventId() {
      return __releasedEventId;
    },

    setEventType: function (eventType) {
      __event_type = eventType;
      return this;
    },

    validate: function () {
      if (__event_type == ScriptApp.EventType.ON_FORM_SUBMIT) {
        __validateSubmission();
      } else if (__event_type == ScriptApp.EventType.ON_EDIT) {
        __validateEdit();
      }

      return this;
    },

    allocate: function (numRow) {
      if (__event_type == ScriptApp.EventType.ON_FORM_SUBMIT) {
        __allocateOnSubmission(numRow);
      } else if (__event_type == ScriptApp.EventType.ON_EDIT) {
        __allocateOnEdit(numRow);
      } else if (__event_type == ScriptApp.EventType.CLOCK) {
        __allocateOnTime(numRow);
      }

      return this;
    },

    markMailed: function (numRow) {
      sheets.sheets[0].getRange(numRow, settings.config.colMailed + 1).setValue(1);
      return this;
    },
  };
})();

// スクリプトトリガーをいじるオブジェクト
const scriptTriggers = (function () {
  let __triggers;

  return {
    get triggers() {
      if (__triggers === undefined) {
        __triggers = ScriptApp.getProjectTriggers();
      }
      return __triggers;
    },

    init: function () {
      this.triggers.forEach((tr) => ScriptApp.deleteTrigger(tr)); // 削除する

      // 新しく設定する
      ScriptApp.newTrigger('onOpening').forSpreadsheet(sheets.ss).onOpen().create();
      ScriptApp.newTrigger('onFormSubmission').forSpreadsheet(sheets.ss).onFormSubmit().create();
      ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(sheets.ss).onEdit().create();
      ScriptApp.newTrigger('onClock').timeBased().atHour(19).nearMinute(30).everyDays(1).inTimezone('Asia/Tokyo').create();
      // ScriptApp.newTrigger('onCalendarUpdated').forUserCalendar(Session.getActiveUser().getEmail()).onEventUpdated().create();
    },

    updateClockTrigger: function (newHour, timeZone) {
      this.triggers.forEach((tr) => {
        if (tr.getEventType() == ScriptApp.EventType.CLOCK) {
          ScriptApp.deleteTrigger(tr);
          ScriptApp.newTrigger('onClock').timeBased().atHour(newHour).nearMinute(30).everyDays(1).inTimezone(timeZone).create();
        }
      });
    },

    // updateCalendarTrigger: function (calendar_id) {
    //   this.triggers.forEach((tr) => {
    //     if (tr.getEventType() == ScriptApp.EventType.ON_EVENT_UPDATED) {
    //       ScriptApp.deleteTrigger(tr);
    //       ScriptApp.newTrigger('onCalendarUpdated').forUserCalendar(calendar_id).onEventUpdated().create();
    //     }
    //   });
    // },
  };
})();

const dlg = (function () {
  let __ui;

  return {
    get ui() {
      if (__ui === undefined) {
        __ui = SpreadsheetApp.getUi();
      }
      return __ui;
    },

    alert: function (title, prompt, buttons) {
      return this.ui.alert(title, prompt, buttons);
    },
  };
})();

const schedule = (function () {
  let __calendar;
  let __available;

  function __getAvailable(available_array) {
    __available = new Map();
    for (let row = 0; row < available_array.length; row++) {
      let available_date = available_array[row][0];
      // A列に日付が適切に入力されていないなら処理をスキップする
      if (!is(available_date, 'Date')) {
        continue;
      }
      let key = fmtDate(available_date, 'yyyy/MM/dd');
      let datetimes = [];
      for (let col = 1; col < available_array[row].length; col++) {
        let available_time = available_array[row][col];
        if (is(available_time, 'Date')) {
          available_time.setFullYear(available_date.getFullYear(), available_date.getMonth(), available_date.getDate());
          // 空き予定でない場合は空文字にする
          if (!__isAvailable(available_time)) {
            available_time = '';
          }
        }
        datetimes.push(available_time);
      }
      __available.set(key, datetimes);
    }
  }

  function __isAvailable(datetime) {
    if (__calendar === undefined) {
      __calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
    }
    const stime = new Date(datetime);
    const etime = new Date(stime);
    etime.setMinutes(etime.getMinutes() + settings.config.experimentLength);
    settings.config.closeTime.setFullYear(etime.getFullYear(), etime.getMonth(), etime.getDate());

    // 計算された終了時刻が設定されている終了時刻を超えていないか
    if (etime > settings.config.closeTime) {
      return false;
    }

    // カレンダーの予定と重複しているかどうか
    const events = __calendar.getEvents(stime, etime);
    if (events.length == 0) {
      return true;
    }
    return false;
  }

  return {
    get available() {
      if (TYPE == 3) {
        return multiSchedule.slots();
      }
      if (__available === undefined) {
        __getAvailable(sheets.getValuesOf('空き予定', true));
      }
      return __available;
    },

    set calendar(val) {
      __calendar = val;
    },

    update: function () {
      if (TYPE == 3) {
        multiSchedule.applySnapshot(multiSchedule.buildSnapshot());
        return this;
      }
      __getAvailable(sheets.getValuesOf('空き予定', true));
      return this;
    },

    allocate: function () {
      if (TYPE == 3) {
        return this;
      }
      const table = [];
      for (let [exp_date, exp_times] of __available.entries()) {
        let new_row = exp_times.map((exp_time) => {
          if (is(exp_time, 'Date')) {
            return fmtDate(new Date(exp_time), 'HH:mm');
          }
          return exp_time; // should be blank string
        });
        new_row.splice(0, 0, exp_date);
        table.push(new_row);
      }
      const sh = sheets.getSheetByName('空き予定');
      sh.getRange(1, 1, table.length, table[0].length).setValues(table);
    },

    init: function () {
      if (TYPE == 3) {
        throw new Error('TYPE 3 の空き予定初期化は新形式では行いません。必要な slot を「空き予定」シートに追加してください。');
      }
      const available_array = [];
      for (let now = new Date(settings.config.openDate); now <= settings.config.closeDate; now.setDate(now.getDate() + 1)) {
        const new_row = [];
        new_row.push(new Date(now));
        settings.config.closeTime.setFullYear(now.getFullYear(), now.getMonth(), now.getDate());
        while (now < settings.config.closeTime) {
          new_row.push(new Date(now));
          now.setMinutes(now.getMinutes() + settings.config.experimentLength);
        }
        now.setHours(settings.config.openTime.getHours(), settings.config.openTime.getMinutes(), 0, 0); // for 文の終了条件での比較のため
        available_array.push(new_row);
      }
      __getAvailable(available_array);
      sheets.getSheetByName('空き予定').clearContents();
      this.allocate();
    },
  };
})();

// TYPE 3 の複数実験者・複数実験室スケジューリングを扱うオブジェクト
const multiSchedule = (function () {
  const ACTIVE_STATES = new Set(['仮割当', '確定']);

  function __rooms() {
    const count = Number(settings.config.numberOfRooms);
    if (!Number.isInteger(count) || count < 1 || count > 5) {
      throw new Error('設定シートの「実験室数 (1-5)」には 1 から 5 の整数を入力してください。');
    }
    return Array.from({ length: count }, (_, index) => `Room ${index + 1}`);
  }

  function __backupName() {
    const base = '空き予定_旧形式_backup';
    let name = base;
    let index = 2;
    while (sheets.ss.getSheetByName(name) !== null) {
      name = `${base}_${index}`;
      index += 1;
    }
    return name;
  }

  function __migrateLegacySlots() {
    const sh = sheets.getSheetByName('空き予定');
    const legacy = sh.getDataRange().getValues();
    if (legacy.length === 0 || legacy[0].includes('日付')) {
      return;
    }
    const converted = [['日付', '開始', '終了', '担当候補', '受付可', '空き実験者（参考）', '空き実験室数', '残り枠', '割当済み']];
    legacy.forEach((row) => {
      const date = new Date(row[0]);
      if (date.toString() == 'Invalid Date') {
        return;
      }
      row.slice(1).forEach((time) => {
        const from = dateFromSlotParts(date, time);
        if (from === undefined) {
          return;
        }
        const to = new Date(from);
        to.setMinutes(to.getMinutes() + Number(settings.config.experimentLength));
        converted.push([fmtDate(from, 'yyyy/MM/dd'), fmtDate(from, 'HH:mm'), fmtDate(to, 'HH:mm'), '', 1, '', '', '', '']);
      });
    });
    sh.copyTo(sheets.ss).setName(__backupName());
    sh.clearContents();
    sh.getRange(1, 1, converted.length, converted[0].length).setValues(converted);
    sheets.update();
  }

  function __isLegacySchedule() {
    const table = sheets.getSheetByName('空き予定').getDataRange().getValues();
    return table.length > 0 && !(table[0] || []).includes('日付');
  }

  function __assertSchema() {
    const answerHeaders = sheets.sheets[0].getRange(1, 1, 1, sheets.sheets[0].getLastColumn()).getValues()[0];
    const requiredAnswers = ['実験室', '割当状態', 'CalendarEventId'];
    if (requiredAnswers.some((header) => !answerHeaders.includes(header)) || __isLegacySchedule()) {
      throw new Error('TYPE 3 の複数人共同募集を開始する前に、管理者が migrateType3MultiResourceSchema を一度実行してください。');
    }
  }

  function __assertMigrationIsSafe() {
    if (!__isLegacySchedule()) {
      return;
    }
    const values = sheets.sheets[0].getDataRange().getValues();
    const futureRows = values.slice(1).filter((row) => {
      const slot = parseType3Slot(row[settings.config.colExpDate]);
      return slot !== undefined && slot.from > new Date();
    });
    if (futureRows.length > 0) {
      throw new Error('将来の旧 TYPE 3 予約があるため自動移行を中止しました。既存予約を完了または別途処理してから移行してください。');
    }
  }

  function __ensureAvailabilityColumns() {
    const sh = sheets.getSheetByName('空き予定');
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (!headers.includes('日付')) {
      return;
    }
    const additions = ['空き実験者（参考）', '空き実験室数'].filter((header) => !headers.includes(header));
    if (additions.length > 0) {
      sh.getRange(1, headers.length + 1, 1, additions.length).setValues([additions]);
    }
  }

  function __getSlots() {
    __assertSchema();
    __ensureAvailabilityColumns();
    const table = sheets.getValuesOf('空き予定', true);
    const headers = table[0] || [];
    const dateCol = getHeaderIndex(headers, '日付');
    const startCol = getHeaderIndex(headers, '開始');
    const endCol = getHeaderIndex(headers, '終了');
    const candidateCol = getHeaderIndex(headers, '担当候補');
    const acceptingCol = getHeaderIndex(headers, '受付可');
    const availableMembersCol = getHeaderIndex(headers, '空き実験者（参考）');
    const availableRoomsCol = getHeaderIndex(headers, '空き実験室数');
    const remainingCol = getHeaderIndex(headers, '残り枠');
    const allocatedCol = getHeaderIndex(headers, '割当済み');
    if ([dateCol, startCol, endCol, candidateCol, acceptingCol, availableMembersCol, availableRoomsCol, remainingCol, allocatedCol].some((index) => index < 0)) {
      throw new Error('「空き予定」は 日付、開始、終了、担当候補、受付可、空き実験者（参考）、空き実験室数、残り枠、割当済み の列を持つ必要があります。');
    }
    const slots = [];
    const labels = new Set();
    for (let row = 1; row < table.length; row++) {
      const from = dateFromSlotParts(table[row][dateCol], table[row][startCol]);
      const to = dateFromSlotParts(table[row][dateCol], table[row][endCol]);
      if (from === undefined || to === undefined || to <= from) {
        continue;
      }
      const label = type3SlotLabel(from, to);
      if (labels.has(label)) {
        throw new Error(`「空き予定」に重複した slot ${label} があります。同じ日時は1行だけにしてください。`);
      }
      labels.add(label);
      slots.push({
        row: row + 1,
        from: from,
        to: to,
        label: label,
        candidateKeys: parseKeys(table[row][candidateCol]),
        accepting: isEnabled(table[row][acceptingCol]),
        remainingCol: remainingCol,
        allocatedCol: allocatedCol,
      });
    }
    return {
      table: table,
      slots: slots,
      availableMembersCol: availableMembersCol,
      availableRoomsCol: availableRoomsCol,
      remainingCol: remainingCol,
      allocatedCol: allocatedCol,
    };
  }

  function __activeBookings() {
    const table = sheets.sheets[0].getDataRange().getValues();
    const bookings = [];
    for (let row = 1; row < table.length; row++) {
      if (!ACTIVE_STATES.has(String(table[row][settings.config.colAssignmentStatus]))) {
        continue;
      }
      const slot = parseType3Slot(table[row][settings.config.colExpDate]);
      const assistant = String(zenToHan(table[row][settings.config.colAssistant] || '')).trim();
      const room = String(table[row][settings.config.colRoom] || '').trim();
      if (slot !== undefined && assistant !== '' && room !== '') {
        bookings.push({
          row: row + 1,
          from: slot.from,
          to: slot.to,
          assistant: assistant,
          room: room,
          eventId: String(table[row][settings.config.colCalendarEventId] || ''),
        });
      }
    }
    return bookings;
  }

  function __calendarFor(member) {
    const calendar = CalendarApp.getCalendarById(member.calendarId);
    if (calendar === null) {
      throw new Error(`メンバー ${member.key} のカレンダーIDにアクセスできません。共有設定とIDを確認してください。`);
    }
    return calendar;
  }

  function __isBlockingEvent(event, ignoredEventIds) {
    if (ignoredEventIds.has(String(event.getId()))) {
      return false;
    }
    if (typeof event.getTransparency === 'function' && event.getTransparency() === CalendarApp.EventTransparency.TRANSPARENT) {
      return false;
    }
    return true;
  }

  function __loadAvailabilityContext(options) {
    options = options || {};
    const slotData = __getSlots();
    const members = settings.activeMembers;
    const rooms = __rooms();
    const activeBookings = __activeBookings().filter((item) => item.row !== options.excludedRow);
    const assignmentsByMember = activeBookings.reduce((counts, item) => {
      counts[item.assistant] = (counts[item.assistant] || 0) + 1;
      return counts;
    }, {});
    const ignoredEventIds = new Set((options.ignoredEventIds || []).map(String));
    const eventsByMember = {};
    if (slotData.slots.length > 0) {
      const rangeStart = new Date(Math.min.apply(null, slotData.slots.map((slot) => slot.from.getTime())));
      const rangeEnd = new Date(Math.max.apply(null, slotData.slots.map((slot) => slot.to.getTime())));
      members.forEach((member) => {
        eventsByMember[member.key] = __calendarFor(member)
          .getEvents(rangeStart, rangeEnd)
          .filter((event) => __isBlockingEvent(event, ignoredEventIds))
          .map((event) => ({ from: event.getStartTime(), to: event.getEndTime(), eventId: String(event.getId()) }));
      });
    }
    return {
      slotData: slotData,
      members: members,
      rooms: rooms,
      activeBookings: activeBookings,
      assignmentsByMember: assignmentsByMember,
      eventsByMember: eventsByMember,
    };
  }

  function __calculateAvailabilitySnapshot(context) {
    const states = context.slotData.slots.map((slot) => {
      const overlappingBookings = context.activeBookings.filter((item) => intervalsOverlap(slot.from, slot.to, item.from, item.to));
      const candidates = context.members.filter((member) => slot.candidateKeys.length === 0 || slot.candidateKeys.includes(member.key));
      const availableMembers = candidates.filter((member) => {
        if (overlappingBookings.some((item) => item.assistant === member.key)) {
          return false;
        }
        return !(context.eventsByMember[member.key] || []).some((event) => intervalsOverlap(slot.from, slot.to, event.from, event.to));
      });
      const availableRooms = context.rooms.filter((room) => !overlappingBookings.some((item) => item.room === room));
      availableMembers.sort((a, b) => {
        const difference = (context.assignmentsByMember[a.key] || 0) - (context.assignmentsByMember[b.key] || 0);
        if (difference !== 0) {
          return difference;
        }
        return String(a.key).localeCompare(String(b.key), undefined, { numeric: true });
      });
      return {
        slot: slot,
        bookings: overlappingBookings,
        members: availableMembers,
        rooms: availableRooms,
        remaining: Math.min(availableMembers.length, availableRooms.length),
        allocated: overlappingBookings.length,
      };
    });
    return {
      context: context,
      states: states,
      choices: states
        .filter((state) => state.slot.accepting && state.remaining > 0)
        .filter((state) => settings.config.openDate <= state.slot.from && state.slot.from <= settings.config.closeDate)
        .map((state) => state.slot.label),
    };
  }

  function __applyMemberAssignmentCounts(context) {
    const sh = sheets.getSheetByName('メンバー');
    const table = sh.getDataRange().getValues();
    const headers = table[0] || [];
    const keyCol = getHeaderIndex(headers, 'キー');
    const countCol = getHeaderIndex(headers, '担当回数');
    if (keyCol < 0 || countCol < 0) {
      throw new Error('「メンバー」は キー と 担当回数 の列を持つ必要があります。');
    }
    if (table.length < 2) {
      return;
    }
    const counts = table.slice(1).map((row) => {
      const key = String(zenToHan(row[keyCol] || '')).trim();
      return [key === '' ? '' : context.assignmentsByMember[key] || 0];
    });
    sh.getRange(2, countCol + 1, counts.length, 1).setValues(counts);
  }

  function __applyAvailabilitySnapshot(snapshot) {
    const slotData = snapshot.context.slotData;
    if (slotData.table.length > 1) {
      const resultColumns = [slotData.availableMembersCol, slotData.availableRoomsCol, slotData.remainingCol, slotData.allocatedCol];
      const firstResultCol = Math.min.apply(null, resultColumns);
      const lastResultCol = Math.max.apply(null, resultColumns);
      const results = slotData.table.slice(1).map((row) => row.slice(firstResultCol, lastResultCol + 1));
      snapshot.states.forEach((state) => {
        const row = results[state.slot.row - 2];
        const memberLabels = state.members.map((member) => {
          const name = String(member.name || '').trim();
          return name === '' ? String(member.key) : `${member.key}: ${name}`;
        });
        row[slotData.availableMembersCol - firstResultCol] = memberLabels.length > 0 ? memberLabels.join(', ') : 'なし';
        row[slotData.availableRoomsCol - firstResultCol] = state.rooms.length;
        row[slotData.remainingCol - firstResultCol] = state.remaining;
        row[slotData.allocatedCol - firstResultCol] = state.allocated;
      });
      sheets.getSheetByName('空き予定').getRange(2, firstResultCol + 1, results.length, lastResultCol - firstResultCol + 1).setValues(results);
    }
    __applyMemberAssignmentCounts(snapshot.context);
    form.modify(snapshot);
  }

  function __addBookingToContext(context, row, allocation, eventId) {
    if (!allocation.valid) {
      return;
    }
    const item = {
      row: row,
      from: allocation.slot.from,
      to: allocation.slot.to,
      assistant: allocation.member.key,
      room: allocation.room,
      eventId: eventId,
    };
    context.activeBookings.push(item);
    context.assignmentsByMember[item.assistant] = (context.assignmentsByMember[item.assistant] || 0) + 1;
  }

  function __writeAssignment(row, values) {
    const sh = sheets.sheets[0];
    Object.keys(values).forEach((configKey) => {
      sh.getRange(row, settings.config[configKey] + 1).setValue(values[configKey]);
    });
  }

  function __eventTitle(prefix, name) {
    return `${prefix}: ${name}`;
  }

  return {
    slots: function () {
      return __getSlots().slots;
    },

    loadAvailabilityContext: function (options) {
      return __loadAvailabilityContext(options);
    },

    calculateAvailabilitySnapshot: function (context) {
      return __calculateAvailabilitySnapshot(context);
    },

    buildSnapshot: function (options) {
      return __calculateAvailabilitySnapshot(__loadAvailabilityContext(options));
    },

    applySnapshot: function (snapshot) {
      __applyAvailabilitySnapshot(snapshot);
    },

    assertMigrationIsSafe: function () {
      __assertMigrationIsSafe();
    },

    migrateLegacySlots: function () {
      __migrateLegacySlots();
    },

    prepareTentative: function (values) {
      const requested = parseType3Slot(values[settings.config.colExpDate]);
      const context = __loadAvailabilityContext();
      const snapshot = __calculateAvailabilitySnapshot(context);
      const state = requested === undefined ? undefined : snapshot.states.find((candidate) => candidate.slot.label === type3SlotLabel(requested.from, requested.to));
      const slot = state === undefined ? undefined : state.slot;
      if (slot === undefined || !slot.accepting) {
        return { valid: false, trigger: '重複', context: context };
      }
      if (state.remaining < 1) {
        return { valid: false, trigger: '重複', context: context };
      }
      return { valid: true, trigger: '仮予約', slot: slot, member: state.members[0], room: state.rooms[0], context: context };
    },

    commitTentative: function (row, values, allocation) {
      if (!allocation.valid) {
        sheets.sheets[0].getRange(row, settings.config.colStatus + 1, 1, 4).setValues([[allocation.trigger, 1, 'N/A', 'N/A']]);
        __writeAssignment(row, { colAssignmentStatus: '解放' });
        __applyAvailabilitySnapshot(__calculateAvailabilitySnapshot(allocation.context));
        return;
      }
      const calendar = __calendarFor(allocation.member);
      const event = calendar.createEvent(__eventTitle('仮予約', values[settings.config.colParName]), allocation.slot.from, allocation.slot.to, { location: allocation.room });
      try {
        sheets.sheets[0].getRange(row, settings.config.colStatus + 1, 1, 4).setValues([['', '', '', '']]);
        __writeAssignment(row, {
          colAssistant: allocation.member.key,
          colRoom: allocation.room,
          colAssignmentStatus: '仮割当',
          colCalendarEventId: event.getId(),
        });
      } catch (err) {
        event.deleteEvent();
        throw err;
      }
      __addBookingToContext(allocation.context, row, allocation, event.getId());
      __applyAvailabilitySnapshot(__calculateAvailabilitySnapshot(allocation.context));
    },

    finalize: function (row, values) {
      if (String(values[settings.config.colAssignmentStatus]) !== '仮割当') {
        throw new Error('最終確認できるのは「仮割当」状態の予約だけです。新しい pair の再割当は行いません。');
      }
      const member = settings.memberByKey(values[settings.config.colAssistant]);
      const eventId = String(values[settings.config.colCalendarEventId] || '');
      if (member === undefined || eventId === '') {
        throw new Error('仮割当の実験者または CalendarEventId がありません。予約を確認してください。');
      }
      const event = __calendarFor(member).getEventById(eventId);
      if (event === null) {
        throw new Error('仮予約のカレンダーイベントが見つかりません。重複を避けるため自動再作成は行いません。');
      }
      event.setTitle(__eventTitle('予約完了', values[settings.config.colParName]));
      event.setLocation(values[settings.config.colRoom]);
      __writeAssignment(row, { colAssignmentStatus: '確定' });
    },

    release: function (row, values) {
      if (!ACTIVE_STATES.has(String(values[settings.config.colAssignmentStatus]))) {
        return '';
      }
      const member = settings.memberByKey(values[settings.config.colAssistant]);
      const eventId = String(values[settings.config.colCalendarEventId] || '');
      if (member !== undefined && eventId !== '') {
        const event = __calendarFor(member).getEventById(eventId);
        if (event !== null) {
          event.deleteEvent();
        }
      }
      __writeAssignment(row, { colAssignmentStatus: '解放' });
      return eventId;
    },
  };
})();

///////////////////////////////////////////////////////////////////////////////
// 初期設定に関わる関数
///////////////////////////////////////////////////////////////////////////////

// 設定用のシートおよびその見本を最初に作る関数
settings.init = function () {
  try {
    let buttons = dlg.ui.ButtonSet.OK_CANCEL;
    let start = true;
    let msg;

    // タイプの確認
    if (TYPE == 1) {
      msg = '自由回答形式の設定で初期化を行います';
    } else if (TYPE == 2) {
      msg = '選択形式の設定で初期化を行います';
    } else if (TYPE == 3) {
      msg = '選択形式の設定で初期化を行います';
    } else {
      msg = '半角数字の1,2,3のいずれかを入力して設定の形式を選択してください';
      buttons = dlg.ui.ButtonSet.OK;
      start = false;
    }
    let choice = dlg.alert('設定の初期化', msg, buttons);
    // let choice = Browser.msgBox('設定の初期化', msg, buttons);
    if (choice != dlg.ui.Button.OK) {
      start = false;
    }

    if (sheets.isConfigured && start) {
      msg = '一度設定を行ったことがあるようです。\nもう一度初期化を行いますか？\n';
      msg += 'フォームの回答が一番初めのシートでないとこれまでの情報が失われる場合があります。';
      choice = dlg.alert('設定の初期化を行います', msg, buttons);
      // let choice = Browser.msgBox('設定の初期化を行います', msg, buttons);
      if (choice != dlg.ui.Button.OK) {
        start = false;
      }
    }

    if (start) {
      sheets.ss.setSpreadsheetTimeZone('Asia/Tokyo');
      settings.default.create();
      scriptTriggers.init();
      msg = '初期設定が終了しました。\n';
      msg += '「設定」シートの太枠に囲まれた項目を適切な情報に変更してください。';
      dlg.alert('設定の初期化', msg, dlg.ui.ButtonSet.OK);
      // Browser.msgBox('設定の初期化', msg, Browser.Buttons.OK);
    } else {
      dlg.alert('設定の初期化', '初期化はキャンセルされました', dlg.ui.ButtonSet.OK);
      // Browser.msgBox('設定の初期化', '初期化はキャンセルされました', Browser.Buttons.OK);
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
    // Browser.msgBox('エラーが発生しました', msg, Browser.Buttons.OK);
  }
};

settings.default = (function () {
  const __sheet_answers = sheets.sheets[0];
  const __default = {};

  function __createDefault() {
    const default_timezone = 'Asia/Tokyo';
    const close_date = new Date();
    close_date.setDate(new Date().getDate() + 13);
    __default.config = [
      ['設定項目', 'メール本文内でのキー', '値'],
      ['実験責任者名', 'experimenterName', '実験太郎'],
      ['実験責任者のGmailアドレス', 'experimenterMailAddress', Session.getActiveUser().getEmail()],
      ['実験責任者の電話番号', 'experimenterPhone', 'xxx-xxx-xxx'],
      ['実験の実施場所', 'experimentRoom', '実施場所'],
      ['実験の所要時間', 'experimentLength', 60],
      ['実験開始可能時刻', 'openTime', 9],
      ['実験終了時刻', 'closeTime', 19],
      ['参照するカレンダー', 'workingCalendar', Session.getActiveUser().getEmail()],
      ['実験室数 (1-5)', 'numberOfRooms', 1],
      ['実験開始日', 'openDate', Utilities.formatDate(new Date(), default_timezone, 'yyyy/MM/dd')],
      ['実験最終日', 'closeDate', Utilities.formatDate(close_date, default_timezone, 'yyyy/MM/dd')],
      ['リマインダー送信時刻', 'remindHour', 19],
      ['予約を完了させるトリガー', 'finalizeTrigger', 111],
      ['タイムゾーン設定', 'expTimeZone', 'Asia/Tokyo'],
      ['自動送信メール残数', 'remainingMails', MailApp.getRemainingDailyQuota()],
      ['予約確認メールを自分にも送るか', 'selfBccTentative', 1],
      ['予約完了メールを自分にも送るか', 'selfBccFinalize', 0],
      ['リマインダーを自分にも送るか', 'selfBccReminder', 0],
      ['翌日の実験予定を送るか', 'sendTmrwExps', 1],
      ['フォーム周りの関数を使用するか', 'useFormSystem', 1],
      ['参加者名の列', 'colParName', 'B'],
      ['ふりがなの列', 'colParNameKana', null],
    ];

    __default.config_note_template = '「フォームの回答」シートにある該当の列と一致しているか確認してください';
    __default.config_notes = [
      ['各項目の備考がコメントとして付されています'],
      ['実験責任者の名前を記入してください'], // 実験責任者
      ['変更する必要はありません。実験用のGmailアドレスが入力されています'], // 実験責任者のGmailアドレス
      ['電話番号を記入してください'], // 電話番号
      ['実験の実施場所を記入してください'], // 実施場所
      ['実験の所要時間を記入してください。'], // 実験の所要時間
      ['何時から実験できるかを記入してください（24時間表記）'], // 実験開始時刻
      ['何時まで実験可能かを記入してください（24時間表記）'], // 実験終了時刻
      ['利用したいカレンダーのIDをコピペしてください'], // 参照するカレンダー
      ['TYPE 3 のみ使用します。1-5 の整数を入力すると Room 1...Room N が自動生成されます。'], // 実験室数
      ['実験を開始する日付を記入してください（年/月/日で表記）'], // 実験開始日
      ['実験の終了予定日を記入してください（年/月/日で表記）'], // 実験最終日
      ['リマインダーを送信する時刻を記入してください（24時間表記）。実験終了時刻以後にして下さい。なお指定した時刻から1時間以内に送信されます。'], // リマインダー送信時刻
      ['必要に応じて任意の半角数字列に変更してください。複数指定する場合はカンマで区切ってください。'], // 予約を完了させるトリガー
      ['必要に応じて変更してください。形式は http://joda-time.sourceforge.net/timezones.html を参照してください。'], // タイムゾーン設定
      ['自動で送信できるメールの残数の目安です。「担当」機能を使っていると一気に2減ったりします。1日経つと100に近い値に戻ります。'], // 自動送信メール残数
      ['自分にも予約確認メールを送る場合は1を，送らない場合は0を入力してください。送らない場合は自動送信できる総メール数が増えます（以下同様）。'], // 予約確認メールを自分にも送るか
      ['自分にも予約完了メールを送る場合は1を，送らない場合は0を入力してください。'], // 予約完了メールを自分にも送るか
      ['自分にも参加者と同様のリマインダーを送る場合は1を，送らない場合は0を入力してください。'], // リマインダーを自分にも送るか
      ['翌日の実験予定の一覧を自分にメールする場合は1を，しない場合は0を入力してください。'], // 翌日の実験予定を送るか
      [
        'ここを0にすると，formに関わる関数が動作しなくなります。この項目はスプレッドシートだけからメールの自動送信システムだけを使用したい人を想定しています',
      ], // フォーム周りの関数を使用するか
      [__default.config_note_template], // 参加者名の列
      [__default.config_note_template + 'もし利用しない場合は空欄にしてください。'], // ふりがなの列
    ];

    // メールテンプレート
    const template_bodies = {
      仮予約: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験へのご応募ありがとうございました。',
        '予約の確認メールを自動で送信しております。\n',
        'expDate fromWhen〜toWhen',
        'で予約を受け付けました（まだ確定はしていません)。',
        '後日、予約完了のメールを送信いたします。',
        'もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
        'どうぞよろしくお願いいたします。\n',
        'experimenterName',
      ],
      時間外: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験へのご応募ありがとうございました。',
        '申し訳ありませんが、ご希望いただいた',
        'expDate fromWhen〜toWhen',
        'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openDate〜closeDate）外です。',
        'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
        'experimenterName',
      ],
      重複: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験へのご応募ありがとうございました。',
        '申し訳ありませんが、ご希望いただいた',
        'expDate fromWhen〜toWhen',
        'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
        'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n',
        'experimenterName',
      ],
      予約完了wd: [
        'participantName 様\n',
        'この度は心理学実験へのご応募ありがとうございました。',
        'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
        '場所はexperimentRoomです。当日は直接お越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '当日もよろしくお願いいたします。\n',
        '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
        '当日の連絡はexperimenterPhoneまでお願いいたします。',
      ],
      予約完了we: [
        'participantName 様\n',
        'この度は心理学実験へのご応募ありがとうございました。',
        'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
        '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '当日もよろしくお願いいたします。\n',
        '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
        '当日の連絡はexperimenterPhoneまでお願いいたします。',
      ],
      222: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験へのご応募ありがとうございました。',
        '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '今後ともよろしくお願いします。\n',
        'experimenterName',
      ],
      333: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験へのご応募ありがとうございました。',
        '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
        '今後、次の実験を実施する際に再度応募していただけると幸いです。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '今後ともよろしくお願いいたします。\n',
        'experimenterName',
      ],
      リマインダーwd: [
        'participantName 様\n',
        '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
        '明日 fromWhenから実験に参加していただく予定となっております。',
        '場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
        'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        'それでは明日、よろしくお願いいたします。\n',
        'experimenterName',
      ],
      リマインダーwe: [
        'participantName 様\n',
        '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
        '明日 fromWhenから実験に参加していただく予定となっております。',
        '場所はexperimentRoomです。\n',
        'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
        'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        'それでは明日、よろしくお願いいたします。\n',
        'experimenterName',
      ],
    };

    for (const key in template_bodies) {
      template_bodies[key] = template_bodies[key].join('\n');
    }

    const not_used = '利用する場合はここに本文を記載するとともに土日での変更の数字を1に変えてください。なお，改行は"alt + enter"です';

    __default.templates = [
      ['トリガー', '休日での変更', '題名', '本文（平日）', '本文（土日祝）'],
      ['仮予約', 0, '予約の確認', template_bodies['仮予約'], not_used],
      ['時間外', 0, '実験実施可能時間外です', template_bodies['時間外'], not_used],
      ['重複', 0, '予約が重複しています', template_bodies['重複'], not_used],
      [111, 1, '実験予約が完了いたしました', template_bodies['予約完了wd'], template_bodies['予約完了we']],
      [222, 0, '以前に実験にご参加いただいたことがあります', template_bodies[222], not_used],
      [333, 0, '定員に達してしまいました', template_bodies[333], not_used],
      ['リマインダー', 1, '明日実施の心理学実験のリマインダー', template_bodies['リマインダーwd'], template_bodies['リマインダーwe']],
    ];

    const note =
      '適宜変更してください。参加者名は participantName ，実験実施時間は fromWhen および toWhen に代入されます。その他のキーは設定シートを参照してください。';

    __default.templates_note = __default.templates.map((_, idx) => {
      if (idx == 0) {
        return [null, null];
      }
      return [note, note];
    });

    // メンバー
    const sh_name_answers = __sheet_answers.getName();
    const answerHeaders = __sheet_answers.getRange(1, 1, 1, __sheet_answers.getLastColumn()).getValues()[0];
    const assistantColumn = getHeaderIndex(answerHeaders, '担当') + 1;
    const assistantColNotation = __sheet_answers.getRange(1, assistantColumn).getA1Notation().replace(/\d/, '');
    const formula = `=COUNTIF('${sh_name_answers}'!${assistantColNotation}:${assistantColNotation}, A2)`;
    __default.members = [
      ['キー', '名前', 'アドレス', 'カレンダーID', '有効', '担当回数'],
      [1, 'りんご', 'apple@hogege.com', Session.getActiveUser().getEmail(), 1, formula],
      [2, 'ごりら', 'gorilla@hogege.com', '', 0, ''],
      [3, 'らっぱ', 'horn@hogege.com', '', 0, ''],
    ];

    // 空き予定
    __default.available = [['日付', '開始', '終了', '担当候補', '受付可', '空き実験者（参考）', '空き実験室数', '残り枠', '割当済み']];
    for (const now = new Date(); now <= close_date; now.setDate(now.getDate() + 1)) {
      now.setHours(9, 0, 0);
      const close_time = 19;
      const exp_length = 60;
      for (const cur_time = new Date(now); cur_time.getHours() < close_time; cur_time.setMinutes(cur_time.getMinutes() + exp_length)) {
        const end_time = new Date(cur_time);
        end_time.setMinutes(end_time.getMinutes() + exp_length);
        __default.available.push([
          Utilities.formatDate(cur_time, default_timezone, 'yyyy/MM/dd'),
          Utilities.formatDate(cur_time, default_timezone, 'HH:mm'),
          Utilities.formatDate(end_time, default_timezone, 'HH:mm'),
          '',
          1,
          '',
          '',
        ]);
      }
    }
  }

  function __addNewColNames() {
    const current_colnms = __sheet_answers.getRange(1, 1, 1, __sheet_answers.getLastColumn()).getValues()[0];
    const new_colnms = ['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当'];
    if (TYPE == 3) {
      new_colnms.push('実験室', '割当状態', 'CalendarEventId');
    }
    const colnms = current_colnms.concat(new_colnms.filter((header) => !current_colnms.includes(header)));
    __sheet_answers.getRange(1, 1, 1, colnms.length).setValues([colnms]);
  }

  function __createConfig() {
    sheets.ss.insertSheet('設定');
    const sh = sheets.ss.getSheetByName('設定');
    const last_column = __sheet_answers.getLastColumn() - 1;
    const extra_config_type1 = [
      ['参加者アドレスの列', 'colAddress', numToColumnNotation(last_column - 6)],
      ['希望日時の列', 'colExpDate', numToColumnNotation(last_column - 5)],
    ];

    const extra_config_type2 = [
      ['参加者アドレスの列', 'colAddress', numToColumnNotation(last_column - 7)],
      ['希望日の列', 'colExpDate', numToColumnNotation(last_column - 6)],
      ['希望時間の列', 'colExpTime', numToColumnNotation(last_column - 5)],
    ];

    const extra_config_common = [
      ['予約ステータスの列', 'colStatus', numToColumnNotation(last_column - 4)],
      ['「連絡したか」の列', 'colMailed', numToColumnNotation(last_column - 3)],
      ['リマインド日時の列', 'colRemindDate', numToColumnNotation(last_column - 2)],
      ['「リマインドしたか」の列', 'colReminded', numToColumnNotation(last_column - 1)],
      ['担当の列', 'colAssistant', numToColumnNotation(last_column)],
    ];

    const extra_config_type3 = [
      ['参加者アドレスの列', 'colAddress', numToColumnNotation(last_column - 9)],
      ['希望日時の列', 'colExpDate', numToColumnNotation(last_column - 8)],
      ['予約ステータスの列', 'colStatus', numToColumnNotation(last_column - 7)],
      ['「連絡したか」の列', 'colMailed', numToColumnNotation(last_column - 6)],
      ['リマインド日時の列', 'colRemindDate', numToColumnNotation(last_column - 5)],
      ['「リマインドしたか」の列', 'colReminded', numToColumnNotation(last_column - 4)],
      ['担当の列', 'colAssistant', numToColumnNotation(last_column - 3)],
      ['実験室の列', 'colRoom', numToColumnNotation(last_column - 2)],
      ['割当状態の列', 'colAssignmentStatus', numToColumnNotation(last_column - 1)],
      ['CalendarEventId の列', 'colCalendarEventId', numToColumnNotation(last_column)],
    ];

    let extra_config;
    if (TYPE == 1) {
      extra_config = extra_config_type1.concat(extra_config_common);
    } else if (TYPE == 2) {
      extra_config = extra_config_type2.concat(extra_config_common);
    } else if (TYPE == 3) {
      extra_config = extra_config_type3;
    }

    __default.config.push(...extra_config);
    const extra_config_notes = extra_config.map(() => [__default.config_note_template]);
    __default.config_notes.push(...extra_config_notes);

    // 値の設定
    const nrow = __default.config.length;
    const ncol = __default.config[0].length;
    sh.getRange(1, 1, nrow, ncol).setValues(__default.config);

    // 注釈
    sh.getRange(1, 3, nrow, 1).setNotes(__default.config_notes);

    // 書式の設定
    sh.getRange(15, 3).setFontColor('#FF0000'); // メールの残数のセルを赤色にする
    sh.getRange(2, 2, nrow - 1, 1).setFontColor('#C8C8C8');
    sh.autoResizeColumn(1);
    sh.autoResizeColumn(3);
    sh.getRange(2, 3, 1, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(4, 3, 8, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(16, 3, 5, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  function __createMailTemplate() {
    sheets.ss.insertSheet('テンプレート');
    const sh = sheets.ss.getSheetByName('テンプレート');
    const rng = sh.getRange(1, 1, __default.templates.length, __default.templates[0].length);
    rng.setValues(__default.templates);
    // 体裁を整える
    rng.setVerticalAlignment('top');
    sh.setColumnWidth(4, 500);
    sh.setColumnWidth(5, 500);
    const cell_text_wrap = __default.templates.map(() => [false, false, true, true, true]);
    rng.setWraps(cell_text_wrap);

    // 注釈の設定
    sh.getRange(1, 4, __default.templates_note.length, __default.templates_note[0].length).setNotes(__default.templates_note);
  }

  function __createMembers() {
    // メンバーシートの設定
    sheets.ss.insertSheet('メンバー');
    const sh = sheets.ss.getSheetByName('メンバー');
    sh.getRange(1, 1, __default.members.length, __default.members[0].length).setValues(__default.members);
    sh.getRange(1, 3).setNote('Gmailのアドレスでなくても大丈夫です。');
  }

  function __createAvailable() {
    // 空き予定シートの設定
    sheets.ss.insertSheet('空き予定');
    const sh = sheets.ss.getSheetByName('空き予定');
    sh.getRange(1, 1, __default.available.length, __default.available[0].length).setValues(__default.available);
  }

  return {
    create() {
      if (sheets.length > 2) {
        sheets.sheets.forEach((sh, idx) => {
          if (idx > 0) {
            sheets.ss.deleteSheet(sh);
          }
        });
      }
      // フォームの回答に不足している管理列だけを追加する
      __addNewColNames();
      __createDefault();
      __createConfig();
      __createMailTemplate();
      __createMembers();
      if (TYPE == 3) {
        __createAvailable();
        SpreadsheetApp.getUi()
          .createMenu('カレンダー')
          .addItem('空き予定とフォームを更新', 'updateFormAndSchedule')
          .addItem('旧形式を複数人共同募集形式へ移行', 'migrateType3MultiResourceSchema')
          .addToUi();
      }
      sheets.ss.insertSheet('Cached');
      sheets.update();
      settings.retrieve().save();
      sheets.ss.getSheetByName('設定').activate(); // 設定画面を開く
    },
  };
})();
