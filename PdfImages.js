// ============================================================
// PdfImages.gs - 領収書画像PDF生成（Google Slides使用）
// ============================================================

/**
 * 領収書画像PDF生成（A4、1ページ8画像）
 */
function generateImagesPdf(session, yearMonth, records) {
  const year  = yearMonth.substring(0, 4);
  const month = yearMonth.substring(4, 6);

  // 画像を持つレコードのみ抽出（連番付き）
  const imageRecords = records.map((rec, idx) => ({ ...rec, no: idx + 1 }))
                               .filter(rec => rec.imageFileId);

  if (imageRecords.length === 0) {
    return { url: null, fileId: null, message: '画像データなし' };
  }

  // ===== Slides 作成 =====
  const title = `領収書画像集_${session.groupName}_${yearMonth}`;
  const presentation = SlidesApp.create(title);
  const firstSlide   = presentation.getSlides()[0];

  // A4 サイズ設定（EMU: 1cm = 914400）
  const A4_W = Math.round(21.0  * 914400);
  const A4_H = Math.round(29.7  * 914400);
  presentation.setPageSize(A4_W, A4_H);

  // ===== レイアウト定数 (pt: 1pt = 12700 EMU) =====
  // A4 = 595pt × 842pt
  const MARGIN   = 20;   // pt
  const COLS     = 2;
  const ROWS_PER_PAGE = 4;
  const IMGS_PER_PAGE = COLS * ROWS_PER_PAGE;
  const GAP_X    = 10;   // pt
  const GAP_Y    = 8;    // pt
  const HEADER_H = 30;   // pt（ページタイトル）
  const LABEL_H  = 14;   // pt（No.+コメント）
  const IMG_AREA_W = (595 - MARGIN * 2 - GAP_X) / COLS;  // ~267pt
  const IMG_AREA_H = (842 - MARGIN * 2 - HEADER_H - GAP_Y * (ROWS_PER_PAGE - 1) - LABEL_H * ROWS_PER_PAGE) / ROWS_PER_PAGE; // ~178pt

  // 最初のスライドは再利用
  let slides = [firstSlide];

  // ===== ページ分割処理 =====
  const totalPages = Math.ceil(imageRecords.length / IMGS_PER_PAGE);

  for (let p = 0; p < totalPages; p++) {
    const slide = p === 0 ? firstSlide : presentation.appendSlide();

    // 背景白
    slide.getBackground().setSolidFill('#ffffff');

    // ページヘッダー
    const headerBox = slide.insertTextBox(
      `${year}年${month}月経費精算　領収書画像集　${session.groupName}　(${p+1}/${totalPages})`,
      MARGIN * 12700, MARGIN * 12700,
      (595 - MARGIN * 2) * 12700, HEADER_H * 12700
    );
    const headerStyle = headerBox.getText().getTextStyle();
    headerStyle.setFontSize(10).setBold(true);
    headerBox.getFill().setSolidFill('#1a73e8');
    headerBox.getText().getTextStyle().setForegroundColor('#ffffff');

    // 画像配置
    const pageRecords = imageRecords.slice(p * IMGS_PER_PAGE, (p + 1) * IMGS_PER_PAGE);
    pageRecords.forEach((rec, idx) => {
      const col = idx % COLS;
      const row = Math.floor(idx / COLS);

      const x = MARGIN + col * (IMG_AREA_W + GAP_X);
      const y = MARGIN + HEADER_H + GAP_Y + row * (IMG_AREA_H + LABEL_H + GAP_Y);

      // ラベル（No. + コメント）
      const labelText = `No.${rec.no}　${rec.comment || ''}`;
      slide.insertTextBox(
        labelText,
        x * 12700, y * 12700,
        IMG_AREA_W * 12700, LABEL_H * 12700
      ).getText().getTextStyle().setFontSize(8);

      // 画像
      try {
        const blob = getReceiptImageBlob(rec.imageFileId);
        if (blob) {
          slide.insertImage(
            blob,
            x * 12700, (y + LABEL_H) * 12700,
            IMG_AREA_W * 12700, IMG_AREA_H * 12700
          );
        }
      } catch (e) {
        // 画像取得失敗時はプレースホルダー
        slide.insertTextBox(
          '画像取得失敗',
          x * 12700, (y + LABEL_H) * 12700,
          IMG_AREA_W * 12700, IMG_AREA_H * 12700
        );
      }
    });
  }

  // ===== PDF エクスポート =====
  const pdfBlob = presentation.getBlob().setName(`領収書画像集_${session.groupName}_${yearMonth}.pdf`);
  const exportsFolder = getExportsFolder(session.groupId, yearMonth);
  const pdfFile = exportsFolder.createFile(pdfBlob);

  // 一時 Slides 削除
  DriveApp.getFileById(presentation.getId()).setTrashed(true);

  return {
    url:    pdfFile.getUrl(),
    fileId: pdfFile.getId()
  };
}
