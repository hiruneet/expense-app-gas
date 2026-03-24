// ============================================================
// Category.gs - 費目自動判定
// ============================================================

/**
 * 費目自動判定（キーワードマッチング）
 */
function detectCategory(payee, rawText) {
  const target = ((payee || '') + ' ' + (rawText || '')).toLowerCase();

  // 駐車場 → CAT001（交通費）※最優先
  const parkingKeywords = [
    '駐車場', 'パーキング', 'parking', 'park',
    'タイムズ', 'times', 'リパーク', 'repark',
    'npc', 'エコロ', 'パラカ', 'paraca'
  ];
  if (parkingKeywords.some(k => target.includes(k))) return 'CAT001';

  // 交通（電車・バス・タクシー）
  if (/電車|バス|タクシー|鉄道|新幹線|jr|metro|モノレール|乗車|運賃/.test(target)) return 'CAT001';

  // 宿泊
  if (/ホテル|旅館|宿泊|hotel|inn/.test(target)) return 'CAT002';

  // 飲食・接待
  if (/レストラン|食堂|居酒屋|カフェ|喫茶|ランチ|ディナー|飲食|食事/.test(target)) return 'CAT003';

  // 会議・セミナー
  if (/会議|セミナー|研修|会場|貸し室/.test(target)) return 'CAT004';

  // 消耗品・事務用品
  if (/文具|事務|コンビニ|スーパー|ドラッグ|薬局|雑貨/.test(target)) return 'CAT005';

  // 通信
  if (/通信|電話|internet|wifi|携帯|スマホ/.test(target)) return 'CAT006';

  // 書籍
  if (/書籍|本|雑誌|書店|amazon/.test(target)) return 'CAT007';

  return ''; // 判定不能は空欄（手動選択）
}


/**
 * 費目一覧取得
 */
function getCategories() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('categories_master');
    if (!sheet) return { success: false, error: 'categories_masterシートが見つかりません' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, categories: [] };

    const categories = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // E列(is_active)がTRUEのもののみ
      if (row[4] === true || row[4] === 'TRUE' || row[4] === 1) {
        categories.push({
          code:     row[0],
          name:     row[1],
          keywords: row[2],
          order:    row[3]
        });
      }
    }
    return { success: true, categories: categories };
  } catch (err) {
    console.error('getCategories error:', err);
    return { success: false, error: err.message };
  }
}
function getCategoryById(code) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('categories_master');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === code) {
        return { code: data[i][0], name: data[i][1] };
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

