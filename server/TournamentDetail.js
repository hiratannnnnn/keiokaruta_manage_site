// ============================================================
// 大会詳細・操作パネル
// ============================================================

// 大会詳細取得（大会名と同名のシートを読む）
//
// シート構造:
//   row[0]       : ヘッダー行。末尾セルにカラム数調整整数 N を格納
//   col 0〜N-1   : フォームカラム
//   col N〜N+2   : 固定管理列（出場回数 / その他 / 振込み済みか）
//   col N+3      : フォームID（システム列）
//   col N+4      : フォームURL（システム列）
//   col N+5 = N  : カラム数調整整数
function getTournamentDetail(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const data = sheet.getDataRange().getValues();
    const rows = data.map(row => row.map(cell => formatCell(cell)));
    if (!rows.length) return JSON.stringify({ name, personHeaders: [], personRows: [], bottomLeft: [], bottomRight: [] });

    // 1 行目の右端の整数 = N
    const headerRow = rows[0];
    const lastVal   = [...headerRow].reverse().find(c => c !== '');
    const N         = (lastVal && /^\d+$/.test(String(lastVal).trim())) ? parseInt(lastVal) : 6;

    const personHeaders = headerRow.slice(0, N + 5);
    const personRows = [];
    let formEndIdx = 1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === '') { formEndIdx = i; break; }
      if (rows[i][2] !== '') personRows.push(rows[i].slice(0, N + 5));
      formEndIdx = i + 1;
    }

    const bottomRows  = rows.slice(formEndIdx).filter(r => r[2] === '');
    const bottomLeft  = bottomRows.filter(r => r[0] !== '').map(r => r.slice(0, N + 3));
    const bottomRight = bottomRows.filter(r => r[N + 3] !== '').map(r => ({ key: r[N + 3], value: r[N + 4] || '' }));

    // 公認/非公認・登録済み判定
    let isOfficial   = true;
    let isRegistered = false;
    for (let i = 1; i < rows.length - 1; i++) {
      if (rows[i][0].trim() === 'registerDatabase') {
        isOfficial   = rows[i + 1][1] === '';
        isRegistered = rows[i + 1][0].includes('登録済み');
        break;
      }
    }

    // グレードサマリー（下部セクションの A〜E 行）
    const gradeSummary = [];
    const gradePattern = /^[A-E]$/;
    for (let i = formEndIdx; i < data.length; i++) {
      const gradeKey = String(data[i][0] || '').trim();
      if (!gradePattern.test(gradeKey)) continue;
      const fee = data[i][1];
      if (typeof fee !== 'number' || fee <= 0) continue;
      const date = rows[i][N + 2] || '';
      // 参加者数カウント（全登録者で grade が一致するもの）
      let count = 0;
      for (let j = 1; j < formEndIdx; j++) {
        if (!data[j][2]) continue;
        const gradeStr = String(data[j][4] || '').trim()
          .replace(/級/g, '')
          .replace(/[Ａ-Ｅ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        if (gradeStr.split('').some(g => g === gradeKey)) count++;
      }
      gradeSummary.push({ grade: gradeKey, fee, count, total: fee * count, date });
    }

    return JSON.stringify({ name, N, isOfficial, isRegistered, personHeaders, personRows, bottomLeft, bottomRight, gradeSummary });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 公認/非公認をトグルする
function toggleOfficialStatus(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length - 1; i++) {
      if (String(data[i][0]).trim() === 'registerDatabase') {
        const isCurrentlyOfficial = String(data[i + 1][1]) === '';
        sheet.getRange(i + 2, 2).setValue(isCurrentlyOfficial ? '非公認' : '');
        return JSON.stringify({ ok: true, isOfficial: !isCurrentlyOfficial });
      }
    }
    return JSON.stringify({ error: '"registerDatabase" 行が見つかりません' });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートの操作トリガー：ラベル行直下の col[0] に値を書き込む
function setTournamentKeyValue(sheetName, key, newValue) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length - 1; i++) {
      if (String(data[i][0]).trim() === key) {
        sheet.getRange(i + 2, 1).setValue(newValue);
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: `ラベル「${key}」が見つかりません` });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートの操作トリガー：ラベル行直下 col[0] の現在値を読み取る
function getTournamentKeyValue(sheetName, key) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length - 1; i++) {
      if (String(data[i][0]).trim() === key) {
        return JSON.stringify({ ok: true, value: formatCell(data[i + 1][0]) });
      }
    }
    return JSON.stringify({ error: `ラベル「${key}」が見つかりません` });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会詳細シートの振込み済みか列（col N+2）を書き換える
function setDetailPayStatus(sheetName, playerName, value, useDeposit) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const N = headerRow.find(c => typeof c === 'number');
    if (N == null) throw new Error('カラム数 (N) が取得できません');

    const allData = sheet.getDataRange().getValues();
    const formEndIdx = getSuitouFormEndIdx_(allData);
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][2]) === playerName) {
        sheet.getRange(i + 1, N + 3).setValue(value); // col N+2 (0-indexed) = N+3 (1-indexed)
        // 済になった場合は出納管理にトランザクションを追加
        if (value === '済') {
          const feeMap   = getSuitouFeeMap_(allData, formEndIdx);
          const gradeStr = String(allData[i][4] || '').trim();
          const fee      = calcFeeFromGrade_(gradeStr, feeMap);
          if (fee > 0) {
            const normalizedName = normalizeName_(playerName);
            if (useDeposit) {
              // デポジットをマイナスで相殺してから参加費をプラス追加
              const suitouSheet = ss.getSheetByName('出納管理');
              if (suitouSheet && suitouSheet.getLastRow() >= 7) {
                const txRows = suitouSheet.getRange(7, 1, suitouSheet.getLastRow() - 6, 3).getValues();
                for (let j = 0; j < txRows.length; j++) {
                  if (normalizeName_(String(txRows[j][0])) === normalizedName &&
                      String(txRows[j][2]).trim() === 'デポジット' &&
                      Number(txRows[j][1]) > 0) {
                    const depositAmt = Number(txRows[j][1]);
                    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
                    const lastRow = Math.max(suitouSheet.getLastRow(), 6);
                    suitouSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[normalizedName, -depositAmt, 'デポジット', today]]);
                    break;
                  }
                }
              }
            }
            appendSuitouTx_(ss, normalizedName, fee, sheetName + '　参加費');
          }
        }
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: '選手が見つかりません' });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会への振込み用：済参加者の人数と合計金額を返す
function getTournamentPaySummary(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const allData    = sheet.getDataRange().getValues();
    const N          = getSuitouN_(allData[0]);
    if (N == null) throw new Error('カラム数 (N) が取得できません');
    const formEndIdx = getSuitouFormEndIdx_(allData);
    const feeMap     = getSuitouFeeMap_(allData, formEndIdx);

    let count = 0, total = 0;
    const gradeMap = {}; // grade -> { fee, names[] }
    for (let i = 1; i < formEndIdx; i++) {
      const payStatus = String(allData[i][N + 2] || '').trim();
      const isPaid    = payStatus === '済' || payStatus === '繰越' || payStatus === 'くりこし';
      if (!isPaid) continue;
      const gradeStr = String(allData[i][4] || '').trim();
      const fee      = calcFeeFromGrade_(gradeStr, feeMap);
      if (fee <= 0) continue;
      const playerName = String(allData[i][2] || '').trim();
      count++;
      total += fee;
      if (!gradeMap[gradeStr]) gradeMap[gradeStr] = { fee, names: [] };
      gradeMap[gradeStr].names.push(playerName);
    }
    const grades = Object.keys(gradeMap).sort().map(g => ({ grade: g, fee: gradeMap[g].fee, names: gradeMap[g].names }));
    return JSON.stringify({ ok: true, count, total, grades });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートのグレード行（下部セクション col[0]=A〜E）の参加費（col[1]）を更新する
function setGradeFee(sheetName, grade, fee) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);
    const data = sheet.getDataRange().getValues();
    const N    = getSuitouN_(data[0]);
    const formEndIdx = getSuitouFormEndIdx_(data);
    for (let i = formEndIdx; i < data.length; i++) {
      if (String(data[i][0]).trim() === grade) {
        sheet.getRange(i + 1, 2).setValue(fee);
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: `グレード「${grade}」の行が見つかりません` });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートのグレード行（A〜E）に大会日を書き込む
// gradeDatesJson: JSON文字列 { A: "2026-05-01", B: "2026-05-01", ... }
function saveTournamentDates(sheetName, gradeDatesJson) {
  try {
    const gradeDates = JSON.parse(gradeDatesJson);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const count = headerRow.find(c => typeof c === 'number' && Number.isFinite(c));
    if (count == null) throw new Error('N が取得できません');

    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      const grade = String(data[i][0]).trim();
      if (/^[A-E]$/.test(grade) && gradeDates[grade]) {
        sheet.getRange(i + 1, count + 3).setValue(new Date(gradeDates[grade]));
      }
    }
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// フォームURL取得（col N+4: フォームURL）
function getFormUrl(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const count = headerRow.find(c => typeof c === 'number' && Number.isFinite(c));
    if (count == null) throw new Error('N が取得できません');
    const formUrl = String(sheet.getRange(1, count + 5).getValue());
    return JSON.stringify({ formUrl });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
