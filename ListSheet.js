// ============================================================
// ListSheet.gs - 経費一覧スプレッドシート生成
// ============================================================

/**
 * 月次経費一覧スプレッドシート生成
 */
function generateListSheet(session, yearMonth, records) {
  const year  = yearMonth.substring(0, 4);
  const month = yearMonth.substring(4, 6);
  const title = `経費一覧_${session.groupName}_${yearMonth}`;

  // 新規スプレッドシート作成
  const ss    = SpreadsheetApp.create(title);
  const sheet = ss.getActiveSheet();
  sheet.setName('経費一覧');

  // ===== ヘッダー =====
  const headerTitle = `${year}年${month}月 経費一覧　${session.groupName}`;
  sheet.getRange(1, 1, 1, 10).merge();
  sheet.getRange(1, 1).setValue(headerTitle)
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // ===== 列ヘッダー =====
  const headers = ['No.', '日付', '担当者', '支払先', 'インボイス番号', '費目', '金額', '立替', 'コメント', '画像No.'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#e8f0fe').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // ===== データ行 =====
  let totalAmount = 0;
  const rows = records.map((rec, idx) => {
    totalAmount += Number(rec.amount) || 0;
    return [
      idx + 1,
      rec.date,
      rec.displayName,
      rec.payee,
      rec.invoiceNo,
      rec.categoryName,
      Number(rec.amount) || 0,
      rec.isAdvancePayment ? '✓' : '',
      rec.comment,
      idx + 1  // 画像PDFとの紐付け番号
    ];
  });
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  }

  // ===== 合計行 =====
  const totalRow = records.length + 3;
  sheet.getRange(totalRow, 1, 1, 6).merge();
  sheet.getRange(totalRow, 1).setValue('合　計').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(totalRow, 7).setValue(totalAmount).setFontWeight('bold');

  // ===== 書式設定 =====
  sheet.setColumnWidth(1, 50);   // No.
  sheet.setColumnWidth(2, 100);  // 日付
  sheet.setColumnWidth(3, 90);   // 担当者
  sheet.setColumnWidth(4, 160);  // 支払先
  sheet.setColumnWidth(5, 160);  // インボイス番号
  sheet.setColumnWidth(6, 110);  // 費目
  sheet.setColumnWidth(7, 90);   // 金額
  sheet.setColumnWidth(8, 45);   // 立替
  sheet.setColumnWidth(9, 200);  // コメント
  sheet.setColumnWidth(10, 70);  // 画像No.

  // 金額列に通貨フォーマット
  if (rows.length > 0) {
    sheet.getRange(3, 7, rows.length + 1, 1)
      .setNumberFormat('¥#,##0');
  }

  // ===== Drive の exports フォルダへ移動 =====
  const exportsFolder = getExportsFolder(session.groupId, yearMonth);
  const file = DriveApp.getFileById(ss.getId());
  exportsFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // マイドライブのルートから削除

  return {
    url:    ss.getUrl(),
    fileId: ss.getId()
  };
}
