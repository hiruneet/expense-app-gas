// ============================================================
// Sheets.gs - スプレッドシート CRUD
// ============================================================
//
// ★ スプレッドシート列定義（1-indexed = getRange 用）
// -------------------------------------------------------
//  1 (A) record_id          | row[0]
//  2 (B) group_id           | row[1]
//  3 (C) user_id            | row[2]
//  4 (D) user_name          | row[3]
//  5 (E) expense_date       | row[4]
//  6 (F) amount             | row[5]
//  7 (G) payee              | row[6]
//  8 (H) invoice_no         | row[7]
//  9 (I) category_code      | row[8]
// 10 (J) category_name      | row[9]
// 11 (K) is_advance_payment | row[10]
// 12 (L) comment            | row[11]
// 13 (M) image_file_id      | row[12]
// 14 (N) image_url          | row[13]
// 15 (O) ocr_raw_text       | row[14]
// 16 (P) status             | row[15]
// 17 (Q) created_at         | row[16]
// 18 (R) updated_at         | row[17]
// -------------------------------------------------------

const RECORDS_SHEET = 'expense_records';

// ===== 列番号定数（getRange 用：1-indexed） =====
const COL = {
  RECORD_ID:          1,
  GROUP_ID:           2,
  USER_ID:            3,
  USER_NAME:          4,
  EXPENSE_DATE:       5,
  AMOUNT:             6,
  PAYEE:              7,
  INVOICE_NO:         8,
  CATEGORY_CODE:      9,
  CATEGORY_NAME:     10,
  IS_ADVANCE:        11,
  COMMENT:           12,
  IMAGE_FILE_ID:     13,
  IMAGE_URL:         14,
  OCR_RAW_TEXT:      15,
  STATUS:            16,
  CREATED_AT:        17,
  UPDATED_AT:        18
};

// ===== ユーティリティ：日付 → "yyyy-MM" 正規化 =====
function normalizeYearMonth(value) {
  if (!value) return '';

  // Date オブジェクトの場合
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
  }

  // 文字列の場合
  var s = String(value).trim();
  if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);        // "2026-03-04" → "2026-03"
  if (/^\d{4}\/\d{2}/.test(s)) return s.slice(0, 7).replace(/\//g, '-'); // "2026/03/04" → "2026-03"

  return '';
}

// ===== ユーティリティ：日付 → "yyyy-MM-dd" 文字列化 =====
function formatDateValue(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

// ============================================================
// 保存（Create）
// ============================================================
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
    const displayName = session.displayName || userId;

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(RECORDS_SHEET);
    if (!sheet) return { success: false, error: RECORDS_SHEET + ' シートが見つかりません' };

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const lastRow = sheet.getLastRow();
      const newId   = 'EXP' + String(lastRow).padStart(5, '0');
      const now     = new Date();

      // ★ 列の順番：上部の列定義コメントと必ず一致させること（18列）
      const row = [
        newId,           //  1 A: record_id
        groupId,         //  2 B: group_id
        userId,          //  3 C: user_id
        displayName,     //  4 D: user_name
        date,            //  5 E: expense_date
        amount,          //  6 F: amount
        payee,           //  7 G: payee
        invoiceNo,       //  8 H: invoice_no
        finalCategoryId, //  9 I: category_code
        categoryName,    // 10 J: category_name
        isAdvance,       // 11 K: is_advance_payment
        comment,         // 12 L: comment
        imageFileId,     // 13 M: image_file_id
        imageUrl,        // 14 N: image_url
        ocrRawText,      // 15 O: ocr_raw_text
        '登録済',         // 16 P: status
        now,             // 17 Q: created_at
        now              // 18 R: updated_at
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

// ============================================================
// 取得（Read）
// ============================================================
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

    // グループフィルタ（row[1] = group_id）
    if (row[1] !== session.groupId) continue;

    // 月フィルタ（row[4] = expense_date）— Date型・文字列 両対応
    if (yearMonth) {
      const rowYM = normalizeYearMonth(row[4]);
      if (rowYM !== yearMonth) continue;
    }

    records.push({
      recordId:         row[0],   //  1 A: record_id
      groupId:          row[1],   //  2 B: group_id
      userId:           row[2],   //  3 C: user_id
      displayName:      row[3],   //  4 D: user_name
      date:             formatDateValue(row[4]),  //  5 E: expense_date
      amount:           row[5],   //  6 F: amount
      payee:            row[6],   //  7 G: payee
      invoiceNo:        row[7],   //  8 H: invoice_no
      categoryId:       row[8],   //  9 I: category_code
      categoryName:     row[9],   // 10 J: category_name
      isAdvancePayment: row[10],  // 11 K: is_advance_payment
      comment:          row[11],  // 12 L: comment
      imageFileId:      row[12],  // 13 M: image_file_id
      imageUrl:         row[13],  // 14 N: image_url
      ocrRawText:       row[14],  // 15 O: ocr_raw_text
      status:           row[15],  // 16 P: status
      createdAt:        row[16],  // 17 Q: created_at
      updatedAt:        row[17],  // 18 R: updated_at
      rowIndex:         i + 1     // 編集・削除用
    });
  }

  return { success: true, records: records };
}

// ============================================================
// 更新（Update）
// ============================================================
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
      // row[0] = record_id, row[1] = group_id
      if (data[i][0] === recordId && data[i][1] === session.groupId) {
        const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

        // COL定数で列番号を指定（ズレ防止）
        if (updates.date !== undefined)             sheet.getRange(i+1, COL.EXPENSE_DATE).setValue(updates.date);
        if (updates.amount !== undefined)           sheet.getRange(i+1, COL.AMOUNT).setValue(updates.amount);
        if (updates.payee !== undefined)            sheet.getRange(i+1, COL.PAYEE).setValue(updates.payee);
        if (updates.invoiceNo !== undefined)        sheet.getRange(i+1, COL.INVOICE_NO).setValue(updates.invoiceNo);
        if (updates.categoryId !== undefined)       sheet.getRange(i+1, COL.CATEGORY_CODE).setValue(updates.categoryId);
        if (updates.categoryName !== undefined)     sheet.getRange(i+1, COL.CATEGORY_NAME).setValue(updates.categoryName);
        if (updates.isAdvancePayment !== undefined) sheet.getRange(i+1, COL.IS_ADVANCE).setValue(updates.isAdvancePayment);
        if (updates.comment !== undefined)          sheet.getRange(i+1, COL.COMMENT).setValue(updates.comment);

        // ★ updated_at は列18（R列）— 旧コードでは列15（O列=ocr_raw_text）を誤更新していた
        sheet.getRange(i+1, COL.UPDATED_AT).setValue(now);

        return { success: true };
      }
    }
    return { success: false, error: 'レコードが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 削除（Delete）— 物理削除
// ============================================================
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
      // row[0] = record_id, row[1] = group_id
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
