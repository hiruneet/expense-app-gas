// ============================================================
// Sheets.gs - スプレッドシート CRUD
// ============================================================

const RECORDS_SHEET = 'expense_records';

/**
 * 経費レコード保存
 */
function saveRecord(params) {
  try {
    const token      = params.token;
    const groupId    = params.groupId    || params.group_id   || '';
    const userId     = params.userId     || params.user_id    || '';
    const date       = params.date       || params.expenseDate || '';
    const amount     = parseInt(params.amount || 0, 10);
    const payee      = params.payee      || params.expensePayee || '';
    const invoiceNo  = params.invoiceNo  || params.invoice_no  || '';
    const categoryId = params.categoryCode || params.categoryId || params.category_code || '';
    const isAdvance  = !!(params.advancePayment || params.isAdvancePayment || false);
    const comment    = params.comment    || '';
    const imageFileId = params.imageFileId || params.image_file_id || '';
    const imageUrl    = params.imageUrl   || params.image_url   || '';
    const ocrRawText  = params.ocrRawText || params.ocr_raw_text || '';

    // セッション確認
    const session = getSessionData(token);
    if (!session || !session.userId) {
      return { success: false, error: 'セッションが無効です' };
    }

    // カテゴリ名取得
    const catObj = getCategoryById(isAdvance ? 'CAT011' : categoryId);
    const categoryName = catObj ? catObj.name : '';
    const finalCategoryId = isAdvance ? 'CAT011' : categoryId;

    // ユーザー名取得
    const userSession = getSessionData(token);
    const displayName = (userSession && userSession.displayName) ? userSession.displayName : userId;

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(RECORDS_SHEET);
    if (!sheet) return { success: false, error: `${RECORDS_SHEET} シートが見つかりません` };

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const lastRow = sheet.getLastRow();
      const newId   = 'EXP' + String(lastRow).padStart(5, '0');
      const now     = new Date();

      // ★ 列の順番：スプレッドシートのヘッダー行と必ず一致させること
      // record_id, group_id, user_id, user_name, expense_date, amount,
      // payee, invoice_no, category_code, category_name, is_advance_payment,
      // comment, image_file_id, image_url, ocr_raw_text, status, created_at, updated_at
      const row = [
        newId,           // A: record_id
        groupId,         // B: group_id
        userId,          // C: user_id
        displayName,     // D: user_name
        date,            // E: expense_date
        amount,          // F: amount
        payee,           // G: payee
        invoiceNo,       // H: invoice_no
        finalCategoryId, // I: category_code
        categoryName,    // J: category_name
        isAdvance,       // K: is_advance_payment
        comment,         // L: comment
        imageFileId,     // M: image_file_id
        imageUrl,        // N: image_url
        ocrRawText,      // O: ocr_raw_text
        '登録済',         // P: status
        now,             // Q: created_at
        now              // R: updated_at
      ];

      sheet.appendRow(row);
      return { success: true, recordId: newId };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error('saveRecord error:', err);
    return { success: false, error: err.message };
  }
}

/**
 * 経費レコード取得（グループ・月フィルタ）
 */
function getRecords(params) {
  const { token, yearMonth } = params;
  const session = getSessionData(token);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(RECORDS_SHEET);
  const data = sheet.getDataRange().getValues();

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === '') continue;
    // グループフィルタ
    if (row[1] !== session.groupId) continue;
    // 月フィルタ（任意）
    if (yearMonth && !String(row[4]).startsWith(yearMonth)) continue;

    records.push({
      recordId:         row[0],
      groupId:          row[1],
      userId:           row[2],
      displayName:      row[3],
      date:             row[4],
      amount:           row[5],
      payee:            row[6],
      invoiceNo:        row[7],
      categoryId:       row[8],
      categoryName:     row[9],
      isAdvancePayment: row[10],
      comment:          row[11],
      imageFileId:      row[12],
      createdAt:        row[13],
      updatedAt:        row[14],
      rowIndex:         i + 1  // 編集用
    });
  }
  return { success: true, records: records };
}

/**
 * 経費レコード更新
 */
function updateRecord(params) {
  const { token, recordId, updates } = params;
  const session = getSessionData(token);

  // 立替金チェック上書き
  if (updates.isAdvancePayment) {
    updates.categoryId   = 'CAT011';
    updates.categoryName = '立替金';
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(RECORDS_SHEET);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === recordId && data[i][1] === session.groupId) {
        const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        if (updates.date !== undefined)             sheet.getRange(i+1, 5).setValue(updates.date);
        if (updates.amount !== undefined)           sheet.getRange(i+1, 6).setValue(updates.amount);
        if (updates.payee !== undefined)            sheet.getRange(i+1, 7).setValue(updates.payee);
        if (updates.invoiceNo !== undefined)        sheet.getRange(i+1, 8).setValue(updates.invoiceNo);
        if (updates.categoryId !== undefined)       sheet.getRange(i+1, 9).setValue(updates.categoryId);
        if (updates.categoryName !== undefined)     sheet.getRange(i+1, 10).setValue(updates.categoryName);
        if (updates.isAdvancePayment !== undefined) sheet.getRange(i+1, 11).setValue(updates.isAdvancePayment);
        if (updates.comment !== undefined)          sheet.getRange(i+1, 12).setValue(updates.comment);
        sheet.getRange(i+1, 15).setValue(now);
        return { success: true };
      }
    }
    return { success: false, error: 'レコードが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 経費レコード削除（論理削除）
 */
function deleteRecord(params) {
  const { token, recordId } = params;
  const session = getSessionData(token);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(RECORDS_SHEET);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === recordId && data[i][1] === session.groupId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'レコードが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}
