// ============================================================
// Report.gs - レポート生成メイン
// ============================================================

/**
 * 月次レポート生成
 */
function generateReport(params) {
  const { token, yearMonth, reportType } = params;
  const session = getSessionData(token);

  // 対象レコード取得
  const recordsResult = getRecords({ token, yearMonth });
  if (!recordsResult.success || recordsResult.records.length === 0) {
    return { success: false, error: '対象データがありません' };
  }
  const records = recordsResult.records;

  const results = {};

  // ① 経費一覧スプレッドシート
  if (!reportType || reportType === 'list' || reportType === 'both') {
    const listResult = generateListSheet(session, yearMonth, records);
    results.listUrl = listResult.url;
    results.listFileId = listResult.fileId;
  }

  // ② 画像一覧PDF
  if (!reportType || reportType === 'pdf' || reportType === 'both') {
    const pdfResult = generateImagesPdf(session, yearMonth, records);
    results.pdfUrl = pdfResult.url;
    results.pdfFileId = pdfResult.fileId;
  }

  return { success: true, ...results, recordCount: records.length };
}
