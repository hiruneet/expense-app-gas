// ============================================================
// OCR.gs - Vision API OCR処理
// ============================================================

/**
 * OCR メイン処理
 */
function processOCR(params) {
  const token = params.token;
  const base64Image = params.imageBase64 || params.base64Image;

  // セッション確認
  const session = getSessionData(token);

  const apiKey = getSettings('vision_api_key');
  if (!apiKey) return { success: false, error: 'Vision API キーが設定されていません' };

  // Vision API 呼び出し
  const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
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

// ============================================================
// 日付抽出
// ============================================================
function extractDate(text) {
  const patterns = [
    // 西暦 yyyy/mm/dd, yyyy-mm-dd, yyyy年mm月dd日
    { re: /(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/, type: 'ymd' },
    // 令和 nn年mm月dd日
    { re: /令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/, type: 'reiwa' },
    // R nn.mm.dd
    { re: /R\s*(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{1,2})/, type: 'reiwa' },
    // nn-mm-dd（年2桁、駐車場レシートで多い形式）
    { re: /(\d{2})[\-\.\/](\d{1,2})[\-\.\/](\d{1,2})/, type: 'yy' }
  ];

  for (const { re, type } of patterns) {
    const m = text.match(re);
    if (!m) continue;

    let y = parseInt(m[1]);
    const mo = String(parseInt(m[2])).padStart(2, '0');
    const dy = String(parseInt(m[3])).padStart(2, '0');

    if (type === 'reiwa') y = y + 2018;
    if (type === 'yy')    y = y + 2000;

    if (parseInt(mo) < 1 || parseInt(mo) > 12) continue;
    if (parseInt(dy) < 1 || parseInt(dy) > 31) continue;

    return y + '-' + mo + '-' + dy;
  }

  return null;
}

// ============================================================
// 金額抽出（駐車場レシート対応強化版）
// ============================================================
function extractAmount(text) {
  // ★ 前処理: インボイス番号（T + 13桁）を除去してから金額を探す
  //    これにより T7010001068319 内の "701000" を誤検出しなくなる
  var cleanText = text.replace(/T\d{13}/g, '');

  // ===== 優先度0: 「〇〇円」パターン（駐車場レシートで最も多い表記）=====
  // 駐車料金 220円、現金 220円、合計 1,500円 など
  var yenPatterns = [
    /駐\s*車\s*料\s*金\s*[：:\s]*([\d,，]+)\s*円/,
    /合\s*計\s*[：:\s]*([\d,，]+)\s*円/,
    /お\s*支\s*払\s*[：:\s]*([\d,，]+)\s*円/,
    /請\s*求\s*額\s*[：:\s]*([\d,，]+)\s*円/,
    /税\s*込\s*[：:\s]*([\d,，]+)\s*円/,
    /現\s*金\s*[：:\s]*([\d,，]+)\s*円/,
    /小\s*計\s*[：:\s]*([\d,，]+)\s*円/
  ];

  for (var i = 0; i < yenPatterns.length; i++) {
    var m = cleanText.match(yenPatterns[i]);
    if (m) {
      var val = parseInt(m[1].replace(/[,，]/g, ''), 10);
      if (!isNaN(val) && val > 0 && val < 10000000) return val;
    }
  }

  // ===== 優先度1: キーワード + ¥/￥ + 数値 =====
  var priorityPatterns = [
    /合\s*計\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /総\s*合\s*計\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /お\s*支\s*払\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /請\s*求\s*額\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /税\s*込\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /小\s*計\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /駐\s*車\s*料\s*金\s*[：:\s]*[¥￥\\]\s*([\d,，]+)/,
    /TOTAL\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/i,
    /AMOUNT\s*[：:\s]*[¥￥\\]?\s*([\d,，]+)/i
  ];

  for (var j = 0; j < priorityPatterns.length; j++) {
    var m2 = cleanText.match(priorityPatterns[j]);
    if (m2) {
      var val2 = parseInt(m2[1].replace(/[,，]/g, ''), 10);
      if (!isNaN(val2) && val2 > 0) return val2;
    }
  }

  // ===== 優先度2: ¥/￥ 記号の直後の数値 =====
  var yenSymbolPat = /[¥￥]\s*([\d,，]+)/g;
  var maxYen = 0, m3;
  while ((m3 = yenSymbolPat.exec(cleanText)) !== null) {
    var val3 = parseInt(m3[1].replace(/[,，]/g, ''), 10);
    if (!isNaN(val3) && val3 > maxYen && val3 < 10000000) maxYen = val3;
  }
  if (maxYen > 0) return maxYen;

  // ===== 優先度3: 数値 + 「円」（キーワードなし） =====
  var enPat = /([\d,，]+)\s*円/g;
  var maxEn = 0, m4;
  while ((m4 = enPat.exec(cleanText)) !== null) {
    var val4 = parseInt(m4[1].replace(/[,，]/g, ''), 10);
    // 0円、消費税率の10は除外
    if (!isNaN(val4) && val4 > 0 && val4 < 10000000 && val4 > maxEn) {
      maxEn = val4;
    }
  }
  if (maxEn > 0) return maxEn;

  // ===== 優先度4: 3〜6桁の裸の数値（最後の手段） =====
  var nums = cleanText.match(/\d{3,6}/g);
  if (nums) {
    var maxVal = 0;
    for (var k = 0; k < nums.length; k++) {
      var v = parseInt(nums[k], 10);
      if (v >= 1900 && v <= 2099) continue;  // 年を除外
      if (v > maxVal && v < 1000000) maxVal = v;
    }
    if (maxVal > 0) return maxVal;
  }

  return null;
}

// ============================================================
// インボイス番号抽出 (T + 13桁)
// ============================================================
function extractInvoiceNo(text) {
  var match = text.match(/T\d{13}/);
  return match ? match[0] : '';
}

// ============================================================
// 駐車場名抽出（2行結合対応）
// ============================================================
function extractPayee(text) {
  var lines = text.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });

  // ★ 駐車場ブランド名のパターン（優先度最高）
  //    マッチした行 + 次の行を結合して駐車場名を構成する
  var parkingBrandPatterns = [
    /NPC\s*24\s*H/i,
    /タイムズ/i,
    /Times/i,
    /リパーク/i,
    /Repark/i,
    /エコロ/i,
    /パラカ/i,
    /Paraca/i,
    /三井のリパーク/i,
    /名鉄/i
  ];

  for (var i = 0; i < lines.length; i++) {
    for (var j = 0; j < parkingBrandPatterns.length; j++) {
      if (parkingBrandPatterns[j].test(lines[i])) {
        var name = lines[i];

        // ★ 次の行が駐車場の場所名っぽければ結合する
        //    条件: 次の行が存在し、金額・日付・除外ワードでなく、短い名前
        if (i + 1 < lines.length) {
          var nextLine = lines[i + 1];
          var isExclude = /^\d+円|^\d{2,4}[\-\/\.]|領収|精算|料金|割引|消費税|適用|株式会社|有限会社|合同会社/.test(nextLine);
          if (!isExclude && nextLine.length >= 2 && nextLine.length <= 30) {
            name = name + ' ' + nextLine;
          }
        }

        // 「NPC」ロゴの誤読重複を除去（例: "NPC NPC24H" → "NPC24H"）
        name = name.replace(/^NPC\s+(NPC)/i, '$1');

        return name.trim().substring(0, 50);
      }
    }
  }

  // ★ 「〇〇パーキング」「〇〇P」「〇〇駐車場」を含む行を探す
  var parkingSuffixPat = /(.{2,20})(パーキング|駐\s*車\s*場|Parking)/i;
  for (var k = 0; k < lines.length; k++) {
    var pm = lines[k].match(parkingSuffixPat);
    if (pm) {
      var parkName = lines[k];
      // 前の行にブランド名があれば結合
      if (k > 0 && lines[k - 1].length <= 20) {
        var prevExclude = /領収|精算|料金|駐\s*車\s*券|^\d/.test(lines[k - 1]);
        if (!prevExclude) {
          parkName = lines[k - 1] + ' ' + parkName;
        }
      }
      return parkName.trim().substring(0, 50);
    }
  }

  // ★ 末尾が「P」で終わる行（駐車場の略称、例: "下総中山P"）
  //    ただし、前の行にブランド名があればそちらで既にマッチしているはず
  for (var n = 0; n < Math.min(lines.length, 8); n++) {
    if (/^.{2,15}P$/.test(lines[n]) && !/消費税|適用|精算/.test(lines[n])) {
      // 前の行と結合
      if (n > 0 && /NPC|タイムズ|Times|リパーク|パラカ/i.test(lines[n - 1])) {
        var combined = lines[n - 1] + ' ' + lines[n];
        combined = combined.replace(/^NPC\s+(NPC)/i, '$1');
        return combined.trim().substring(0, 50);
      }
      return lines[n].trim();
    }
  }

  // ★ 会社形態を含む行（株式会社など）— 駐車場以外の場合のフォールバック
  var companyPattern = /(株式会社|有限会社|合同会社|㈱|㈲|（株）).{0,20}|([\w\s]{2,20})(株式会社|有限会社)/;
  for (var p = 0; p < lines.length; p++) {
    if (companyPattern.test(lines[p]) && lines[p].length <= 40) {
      return lines[p].replace(/[　\s]+/g, ' ').trim();
    }
  }

  // ★ 除外ワード以外の最初の有効行
  var excludeWords = /領収書|領収証|receipt|invoice|合計|小計|税|金額|お買い上げ|ありがとう|駐\s*車\s*券|^\d+$/i;
  for (var q = 0; q < Math.min(lines.length, 6); q++) {
    if (!excludeWords.test(lines[q]) && lines[q].length >= 2 && lines[q].length <= 30) {
      return lines[q];
    }
  }

  return '';
}
