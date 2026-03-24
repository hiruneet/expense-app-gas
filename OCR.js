// ============================================================
// OCR.gs - Vision API OCR処理
// ============================================================

/**
 * OCR メイン処理
 */
function processOCR(params) {
  const token = params.token;
  const base64Image = params.imageBase64 || params.base64Image; // 両方に対応

  // セッション確認
  const session = getSessionData(token);

  const apiKey = getSettings('vision_api_key');
  if (!apiKey) return { success: false, error: 'Vision API キーが設定されていません' };

  // Vision API 呼び出し
  const url = `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
  const requestBody = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: 'TEXT_DETECTION', maxResults: 1 }]
    }]
  };

  let rawText = '';
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (result.responses && result.responses[0].textAnnotations) {
      rawText = result.responses[0].textAnnotations[0].description || '';
    }
  } catch (err) {
    console.error('Vision API error:', err);
    return { success: false, error: 'OCR 処理に失敗しました: ' + err.message };
  }

  // テキストから各項目を抽出
  const extracted = {
    date:      extractDate(rawText),
    amount:    extractAmount(rawText),
    payee:     extractPayee(rawText),
    invoiceNo: extractInvoiceNo(rawText),
    rawText:   rawText
  };

  // 費目自動判定
  extracted.category = detectCategory(extracted.payee, rawText);

  return { success: true, data: extracted };
}

/**
 * 日付抽出
 */
function extractDate(text) {
  const patterns = [
    // 西暦 yyyy/mm/dd, yyyy-mm-dd, yyyy年mm月dd日
    { re: /(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/, type: 'ymd' },
    // 令和 nn年mm月dd日
    { re: /令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/, type: 'reiwa' },
    // R nn.mm.dd
    { re: /R\s*(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{1,2})/, type: 'reiwa' },
    // 'nn.mm.dd'（年2桁）
    { re: /(\d{2})[\.\-\/](\d{1,2})[\.\-\/](\d{1,2})/, type: 'yy' }
  ];

  for (const { re, type } of patterns) {
    const m = text.match(re);
    if (!m) continue;

    let y = parseInt(m[1]);
    const mo = String(parseInt(m[2])).padStart(2, '0');
    const dy = String(parseInt(m[3])).padStart(2, '0');

    if (type === 'reiwa') y = y + 2018;
    if (type === 'yy')    y = y + 2000;

    // 月・日の妥当性チェック
    if (parseInt(mo) < 1 || parseInt(mo) > 12) continue;
    if (parseInt(dy) < 1 || parseInt(dy) > 31) continue;

    return `${y}-${mo}-${dy}`;
  }

  return null; // ← 読めなかったら空欄（今日の日付を入れない）
}


/**
 * 金額抽出（最大金額を採用）
 */
function extractAmount(text) {
  // 優先度1：合計・お支払い等キーワード直後の数値
  const priorityPatterns = [
    /合\s*計\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /総\s*合\s*計\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /お\s*支\s*払\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /請\s*求\s*額\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /税\s*込\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /小\s*計\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/,
    /TOTAL\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/i,
    /AMOUNT\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/i,
  ];

  for (const pat of priorityPatterns) {
    const m = text.match(pat);
    if (m) {
      const val = parseInt(m[1].replace(/[,，]/g, ''), 10);
      if (!isNaN(val) && val > 0) return val;
    }
  }

  // 優先度2：¥ または ￥ 記号の直後の数値
  const yenPat = /[¥￥]\s*([\d,，]+)/g;
  let maxYen = 0, m2;
  while ((m2 = yenPat.exec(text)) !== null) {
    const val = parseInt(m2[1].replace(/[,，]/g, ''), 10);
    if (!isNaN(val) && val > maxYen && val < 10000000) maxYen = val;
  }
  if (maxYen > 0) return maxYen;

  // 優先度3：3〜6桁の数値（ただし年っぽい1900〜2099は除外）
  const nums = text.match(/\d{3,6}/g);
  if (nums) {
    let maxVal = 0;
    nums.forEach(n => {
      const v = parseInt(n, 10);
      if (v >= 1900 && v <= 2099) return; // 年を除外 ← これが今回のバグ修正
      if (v > maxVal && v < 1000000) maxVal = v;
    });
    if (maxVal > 0) return maxVal;
  }

  return null;
}

/**
 * インボイス番号抽出 (T + 13桁)
 */
function extractInvoiceNo(text) {
  const match = text.match(/T\d{13}/);
  return match ? match[0] : '';
}

/**
 * 支払先（会社名）抽出
 */
function extractPayee(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // 優先度1：駐車場ブランド名（固有名詞）
  const parkingBrands = [
    /タイムズ[^\n]*/i,
    /Times[^\n]*/i,
    /リパーク[^\n]*/i,
    /Repark[^\n]*/i,
    /NPC[^\n]*/i,
    /エコロ[^\n]*/i,
    /パラカ[^\n]*/i,
    /Paraca[^\n]*/i,
    /[Pp]ark(?:ing)?[^\n]*/,
    /パーキング[^\n]*/,
    /駐\s*車\s*場[^\n]*/
  ];

  for (const line of lines) {
    for (const brand of parkingBrands) {
      const m = line.match(brand);
      if (m) return m[0].trim().substring(0, 40);
    }
  }

  // 優先度2：会社形態を含む行（株式会社など）
  const companyPattern = /(株式会社|有限会社|合同会社|㈱|㈲|（株）).{0,20}|([\w\s]{2,20})(株式会社|有限会社)/;
  for (const line of lines) {
    if (companyPattern.test(line) && line.length <= 40) {
      return line.replace(/[　\s]+/g, ' ').trim();
    }
  }

  // 優先度3：除外ワード以外の最初の有効行
  const excludeWords = /領収書|領収証|receipt|invoice|合計|小計|税|金額|お買い上げ|ありがとう|^\d+$/i;
  for (const line of lines.slice(0, 6)) {
    if (!excludeWords.test(line) && line.length >= 2 && line.length <= 30) {
      return line;
    }
  }

  return '';
}



  // 最小限のテスト画像（「TEST 1234」と書かれた1x1のbase64画像）
  // 実際には小さなテキスト画像で試す
  const testUrl = 'https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/PNG_transparency_demonstration_1.png/280px-PNG_transparency_demonstration_1.png';
  
  try {
    const imageBlob = UrlFetchApp.fetch(testUrl).getBlob();
    const base64 = Utilities.base64Encode(imageBlob.getBytes());
    
    const url = `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
    const body = {
      requests: [{
        image: { content: base64 },
        features: [{ type: 'TEXT_DETECTION', maxResults: 1 }]
      }]
    };

    const res = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(body)
    });

    const json = JSON.parse(res.getContentText());
    Logger.log('Vision APIレスポンス全体: ' + JSON.stringify(json));

    if (json.responses && json.responses[0]) {
      const r = json.responses[0];
      if (r.error) {
        Logger.log('❌ APIエラー: ' + r.error.message);
      } else {
        Logger.log('✅ Vision API 正常動作');
        Logger.log('rawText: ' + (r.fullTextAnnotation ? r.fullTextAnnotation.text : '（テキストなし）'));
      }
    }
  } catch (err) {
    Logger.log('❌ 例外エラー: ' + err.message);
  }
}
