/**
 * 구글시트 연동 웹 추첨 프로그램 (Google Apps Script)
 * 
 * 시트 구성:
 * 1. 참여자: 이름 | 전화번호 | 당첨여부 | 당첨상품 | 당첨일시
 * 2. 상품설정: 상품명 | 총수량 | 당첨수량 | 남은수량 | 순서
 * 3. 당첨기록: 당첨일시 | 상품명 | 이름 | 전화번호
 */

const SHEET_NAMES = {
  PARTICIPANTS: '참여자',
  PRIZES: '상품설정',
  WINNERS: '당첨기록'
};

/**
 * 구글시트 열릴 때 상단 메뉴 추가
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎉 추첨 관리')
    .addItem('1. 시트 입력 양식 자동 생성', 'setupSheets')
    .addItem('2. 샘플 데이터 채우기 (테스트용)', 'insertSampleData')
    .addSeparator()
    .addItem('당첨 기록만 초기화 (재추첨 준비)', 'resetAllDrawsMenu')
    .addToUi();
}

/**
 * 웹 앱 접속 (HTML 반환 또는 JSON API 응답)
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action) {
    return handleApiGet_(action, e.parameter);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('🎉 구글시트 연동 추첨 프로그램')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 외부 웹(로컬 HTML 등)에서 POST 요청 시 처리
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    let result;

    switch (action) {
      case 'addParticipant':
        result = addParticipant(body.name, body.phone);
        break;
      case 'addParticipantsBulk':
        result = addParticipantsBulk(body.list || []);
        break;
      case 'deleteParticipant':
        result = deleteParticipant(body.rowIndex);
        break;
      case 'savePrizes':
        result = savePrizes(body.prizes || []);
        break;
      case 'drawWinners':
        result = drawWinners(body.prizeName, body.count || 1);
        break;
      case 'cancelWinner':
        result = cancelWinner(body.rowIndex);
        break;
      case 'resetAllDraws':
        result = resetAllDraws();
        break;
      case 'setupSheets':
        setupSheets(true);
        result = { ok: true, message: '시트 양식이 생성되었습니다.' };
        break;
      default:
        result = { ok: false, error: '알 수 없는 action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiGet_(action, params) {
  let result;
  try {
    if (action === 'getData') {
      result = getAllData();
    } else {
      result = { ok: false, error: '알 수 없는 GET action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 1. 시트 입력 양식 자동 생성
 */
function setupSheets(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 참여자 시트
  let pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  if (!pSheet) {
    pSheet = ss.insertSheet(SHEET_NAMES.PARTICIPANTS);
  }
  if (pSheet.getLastRow() === 0) {
    const pHeaders = [['이름', '전화번호', '당첨여부', '당첨상품', '당첨일시']];
    pSheet.getRange(1, 1, 1, 5).setValues(pHeaders);
    styleHeader_(pSheet, 5, '#4F46E5');
    pSheet.setColumnWidth(1, 130);
    pSheet.setColumnWidth(2, 160);
    pSheet.setColumnWidth(3, 100);
    pSheet.setColumnWidth(4, 180);
    pSheet.setColumnWidth(5, 170);
    // 전화번호 열을 텍스트 서식으로 지정 (앞자리 0 유지)
    pSheet.getRange('B2:B2000').setNumberFormat('@');
    // 당첨여부 드롭다운 유효성 검사
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['미당첨', '당첨'], true)
      .build();
    pSheet.getRange('C2:C2000').setDataValidation(rule);
    pSheet.setFrozenRows(1);
  }

  // 2) 상품설정 시트
  let prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);
  if (!prSheet) {
    prSheet = ss.insertSheet(SHEET_NAMES.PRIZES);
  }
  if (prSheet.getLastRow() === 0) {
    const prHeaders = [['상품명', '총수량', '당첨수량', '남은수량', '비고']];
    prSheet.getRange(1, 1, 1, 5).setValues(prHeaders);
    styleHeader_(prSheet, 5, '#059669');
    prSheet.setColumnWidth(1, 200);
    prSheet.setColumnWidth(2, 90);
    prSheet.setColumnWidth(3, 90);
    prSheet.setColumnWidth(4, 90);
    prSheet.setColumnWidth(5, 200);
    prSheet.setFrozenRows(1);
  }

  // 3) 당첨기록 시트
  let wSheet = ss.getSheetByName(SHEET_NAMES.WINNERS);
  if (!wSheet) {
    wSheet = ss.insertSheet(SHEET_NAMES.WINNERS);
  }
  if (wSheet.getLastRow() === 0) {
    const wHeaders = [['당첨일시', '상품명', '이름', '전화번호']];
    wSheet.getRange(1, 1, 1, 4).setValues(wHeaders);
    styleHeader_(wSheet, 4, '#D97706');
    wSheet.setColumnWidth(1, 170);
    wSheet.setColumnWidth(2, 180);
    wSheet.setColumnWidth(3, 130);
    wSheet.setColumnWidth(4, 160);
    wSheet.getRange('D2:D2000').setNumberFormat('@');
    wSheet.setFrozenRows(1);
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert('✅ 시트 입력 양식(참여자, 상품설정, 당첨기록)이 준비되었습니다!');
  }
}

function styleHeader_(sheet, colCount, bgColor) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground(bgColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
}

/**
 * 테스트용 샘플 데이터 삽입
 */
function insertSampleData() {
  setupSheets(true);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);

  if (pSheet.getLastRow() <= 1) {
    const sampleParticipants = [
      ['김민준', '010-1234-5678', '미당첨', '', ''],
      ['이서연', '010-2345-6789', '미당첨', '', ''],
      ['박도윤', '010-3456-7890', '미당첨', '', ''],
      ['최지우', '010-4567-8901', '미당첨', '', ''],
      ['정하은', '010-5678-9012', '미당첨', '', ''],
      ['강시우', '010-6789-0123', '미당첨', '', ''],
      ['조수아', '010-7890-1234', '미당첨', '', ''],
      ['윤예준', '010-8901-2345', '미당첨', '', '']
    ];
    pSheet.getRange(2, 1, sampleParticipants.length, 5).setValues(sampleParticipants);
  }

  if (prSheet.getLastRow() <= 1) {
    const samplePrizes = [
      ['1등: 에어팟 프로', 1, 0, '=B2-C2', '최고 인기 상품'],
      ['2등: 백화점 상품권 5만원', 2, 0, '=B3-C3', ''],
      ['3등: 스타벅스 아메리카노', 3, 0, '=B4-C4', '']
    ];
    prSheet.getRange(2, 1, samplePrizes.length, 5).setValues(samplePrizes);
  }

  SpreadsheetApp.getUi().alert('✅ 샘플 데이터가 추가되었습니다.');
}

/**
 * 전체 데이터 조회 (웹 앱 초기화 & 새로고침)
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_NAMES.PARTICIPANTS) || !ss.getSheetByName(SHEET_NAMES.PRIZES)) {
    setupSheets(true);
  }

  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);
  const wSheet = ss.getSheetByName(SHEET_NAMES.WINNERS);

  const participants = [];
  if (pSheet.getLastRow() > 1) {
    const rows = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 5).getDisplayValues();
    rows.forEach((row, idx) => {
      const name = (row[0] || '').trim();
      const phone = (row[1] || '').trim();
      if (!name && !phone) return;
      const status = (row[2] || '').trim();
      const isWinner = status === '당첨' || status.toUpperCase() === 'TRUE' || status === 'O';
      participants.push({
        rowIndex: idx + 2,
        name: name,
        phone: phone,
        isWinner: isWinner,
        prize: (row[3] || '').trim(),
        wonAt: (row[4] || '').trim()
      });
    });
  }

  const prizes = [];
  if (prSheet.getLastRow() > 1) {
    const rows = prSheet.getRange(2, 1, prSheet.getLastRow() - 1, 5).getValues();
    rows.forEach((row, idx) => {
      const name = String(row[0] || '').trim();
      if (!name) return;
      const total = Math.max(0, Number(row[1]) || 0);
      const drawn = Math.max(0, Number(row[2]) || 0);
      const remaining = Math.max(0, total - drawn);
      prizes.push({
        rowIndex: idx + 2,
        name: name,
        total: total,
        drawn: drawn,
        remaining: remaining,
        note: String(row[4] || '').trim()
      });
    });
  }

  const winners = [];
  if (wSheet && wSheet.getLastRow() > 1) {
    const rows = wSheet.getRange(2, 1, wSheet.getLastRow() - 1, 4).getDisplayValues();
    rows.forEach((row, idx) => {
      if (!row[1] && !row[2]) return;
      winners.push({
        rowIndex: idx + 2,
        wonAt: row[0],
        prize: row[1],
        name: row[2],
        phone: row[3]
      });
    });
  }

  return {
    ok: true,
    participants: participants,
    prizes: prizes,
    winners: winners.reverse()
  };
}

/**
 * 참여자 단건 등록
 */
function addParticipant(name, phone) {
  const cleanName = String(name || '').trim();
  const cleanPhone = normalizePhone_(phone);

  if (!cleanName) return { ok: false, error: '이름을 입력해주세요.' };
  if (!cleanPhone) return { ok: false, error: '전화번호를 입력해주세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    if (!pSheet) {
      setupSheets(true);
      pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    }

    // 중복 전화번호 체크
    const existing = getAllData().participants;
    const dup = existing.find(p => p.phone.replace(/\D/g, '') === cleanPhone.replace(/\D/g, ''));
    if (dup) {
      return { ok: false, error: `이미 등록된 전화번호입니다 (${dup.name}).` };
    }

    const targetRow = pSheet.getLastRow() + 1;
    pSheet.getRange(targetRow, 2).setNumberFormat('@');
    pSheet.getRange(targetRow, 1, 1, 5).setValues([[cleanName, cleanPhone, '미당첨', '', '']]);
    return { ok: true, data: getAllData() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 참여자 일괄 등록
 */
function addParticipantsBulk(list) {
  if (!Array.isArray(list) || list.length === 0) {
    return { ok: false, error: '등록할 목록이 비어 있습니다.' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    const existingPhones = new Set(
      getAllData().participants.map(p => p.phone.replace(/\D/g, ''))
    );

    const rowsToAdd = [];
    let skipped = 0;

    list.forEach(item => {
      const n = String(item.name || '').trim();
      const ph = normalizePhone_(item.phone);
      const digits = ph.replace(/\D/g, '');
      if (!n || !digits || existingPhones.has(digits)) {
        skipped++;
        return;
      }
      existingPhones.add(digits);
      rowsToAdd.push([n, ph, '미당첨', '', '']);
    });

    if (rowsToAdd.length > 0) {
      const startRow = pSheet.getLastRow() + 1;
      pSheet.getRange(startRow, 2, rowsToAdd.length, 1).setNumberFormat('@');
      pSheet.getRange(startRow, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }

    return {
      ok: true,
      added: rowsToAdd.length,
      skipped: skipped,
      data: getAllData()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 참여자 삭제
 */
function deleteParticipant(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    if (rowIndex >= 2 && rowIndex <= pSheet.getLastRow()) {
      pSheet.deleteRow(rowIndex);
    }
    return { ok: true, data: getAllData() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 상품 및 수량 설정 저장
 */
function savePrizes(prizes) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);
    const lastRow = prSheet.getLastRow();
    if (lastRow > 1) {
      prSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }

    if (Array.isArray(prizes) && prizes.length > 0) {
      const rows = prizes.map((p, idx) => {
        const r = idx + 2;
        const total = Math.max(0, Number(p.total) || 0);
        const drawn = Math.max(0, Math.min(total, Number(p.drawn) || 0));
        return [
          String(p.name || '').trim(),
          total,
          drawn,
          `=B${r}-C${r}`,
          String(p.note || '').trim()
        ];
      }).filter(r => r[0] !== '');

      if (rows.length > 0) {
        prSheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }

    return { ok: true, data: getAllData() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 추첨 실행 (당첨자는 자동으로 참여자 목록에서 제외됨)
 */
function drawWinners(prizeName, count) {
  const drawCount = Math.max(1, Number(count) || 1);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);
    const wSheet = ss.getSheetByName(SHEET_NAMES.WINNERS);

    const data = getAllData();
    const prize = data.prizes.find(p => p.name === prizeName);
    if (!prize) {
      return { ok: false, error: '선택한 상품을 찾을 수 없습니다.' };
    }
    if (prize.remaining <= 0) {
      return { ok: false, error: `"${prizeName}" 상품의 남은 수량이 없습니다.` };
    }

    const actualCount = Math.min(drawCount, prize.remaining);
    const eligible = data.participants.filter(p => !p.isWinner);
    if (eligible.length === 0) {
      return { ok: false, error: '추첨 가능한 미당첨 참여자가 없습니다.' };
    }
    if (eligible.length < actualCount) {
      return {
        ok: false,
        error: `추첨 가능 인원(${eligible.length}명)이 요청 인원(${actualCount}명)보다 적습니다.`
      };
    }

    // Fisher-Yates 셔플로 공정하게 무작위 추첨
    const pool = eligible.slice();
    for (let i = pool.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const temp = pool[i];
      pool[i] = pool[j];
      pool[j] = temp;
    }
    const picked = pool.slice(0, actualCount);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    // 참여자 시트 업데이트 (당첨여부 = 당첨 -> 이후 추첨에서 제외)
    const winnerLogRows = [];
    picked.forEach(person => {
      pSheet.getRange(person.rowIndex, 3, 1, 3).setValues([['당첨', prize.name, nowStr]]);
      winnerLogRows.push([nowStr, prize.name, person.name, person.phone]);
    });

    // 상품설정 시트 당첨수량 증가
    prSheet.getRange(prize.rowIndex, 3).setValue(prize.drawn + picked.length);

    // 당첨기록 시트 추가
    if (winnerLogRows.length > 0) {
      const startRow = wSheet.getLastRow() + 1;
      wSheet.getRange(startRow, 4, winnerLogRows.length, 1).setNumberFormat('@');
      wSheet.getRange(startRow, 1, winnerLogRows.length, 4).setValues(winnerLogRows);
    }

    return {
      ok: true,
      picked: picked,
      prizeName: prize.name,
      wonAt: nowStr,
      data: getAllData()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 특정 참여자 당첨 취소 (다시 미당첨으로 복귀 + 상품 수량 복구)
 */
function cancelWinner(participantRowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);

    const data = getAllData();
    const person = data.participants.find(p => p.rowIndex === participantRowIndex);
    if (!person || !person.isWinner) {
      return { ok: false, error: '당첨 상태인 참여자를 찾을 수 없습니다.' };
    }

    // 상품 당첨수량 1 감소
    const prize = data.prizes.find(p => p.name === person.prize);
    if (prize && prize.drawn > 0) {
      prSheet.getRange(prize.rowIndex, 3).setValue(prize.drawn - 1);
    }

    // 참여자 상태 초기화
    pSheet.getRange(person.rowIndex, 3, 1, 3).setValues([['미당첨', '', '']]);

    return { ok: true, data: getAllData() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 전체 당첨 내역 초기화 (참여자 명단과 상품 총수량은 유지)
 */
function resetAllDraws() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    const prSheet = ss.getSheetByName(SHEET_NAMES.PRIZES);
    const wSheet = ss.getSheetByName(SHEET_NAMES.WINNERS);

    if (pSheet.getLastRow() > 1) {
      const count = pSheet.getLastRow() - 1;
      const resetRows = Array.from({ length: count }, () => ['미당첨', '', '']);
      pSheet.getRange(2, 3, count, 3).setValues(resetRows);
    }

    if (prSheet.getLastRow() > 1) {
      const count = prSheet.getLastRow() - 1;
      const zeroRows = Array.from({ length: count }, () => [0]);
      prSheet.getRange(2, 3, count, 1).setValues(zeroRows);
    }

    if (wSheet && wSheet.getLastRow() > 1) {
      wSheet.getRange(2, 1, wSheet.getLastRow() - 1, 4).clearContent();
    }

    return { ok: true, data: getAllData() };
  } finally {
    lock.releaseLock();
  }
}

function resetAllDrawsMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('당첨 내역 초기화', '모든 참여자를 [미당첨]으로 되돌리고 상품 당첨수량을 0으로 초기화할까요?', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    resetAllDraws();
    ui.alert('초기화되었습니다.');
  }
}

function normalizePhone_(phone) {
  const raw = String(phone || '').trim();
  const digits = raw.replace(/\D/g, '');
  if (digits.length === 11 && digits.startsWith('010')) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`;
  }
  if (digits.length === 10) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`;
  }
  return raw;
}
