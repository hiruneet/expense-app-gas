// ============================================================
// Drive.gs - Drive フォルダ・画像管理
// ============================================================

/**
 * 領収書画像保存
 */
function saveReceiptImage(params) {
  try {
    const token       = params.token;
    const base64Image = params.imageBase64 || params.base64Image || '';
    const mimeType    = params.mimeType    || 'image/jpeg';
    const fileName    = params.fileName    || `receipt_${Date.now()}.jpg`;
    const groupId     = params.groupId     || params.group_id || 'default';

    if (!base64Image) {
      return { success: false, error: '画像データが空です' };
    }

    const session = getSessionData(token);
    if (!session || !session.userId) {
      return { success: false, error: 'セッションが無効です' };
    }

    // root_folder_id があればそれを使い、なければスプレッドシートの親フォルダを使う
    let rootFolder;
    const rootFolderId = getSettings('root_folder_id');
    if (rootFolderId) {
      rootFolder = DriveApp.getFolderById(rootFolderId);
    } else {
      const ss         = getSpreadsheet();
      const ssFile     = DriveApp.getFileById(ss.getId());
      const parentIter = ssFile.getParents();
      rootFolder       = parentIter.hasNext() ? parentIter.next() : DriveApp.getRootFolder();
    }

    // /receipts/{groupId}/{YYYY_MM}/ フォルダ取得または作成
    const receiptsFolder = getOrCreateFolder(rootFolder, 'receipts');
    const groupFolder    = getOrCreateFolder(receiptsFolder, groupId);
    const now            = new Date();
    const monthFolder    = getOrCreateFolder(
      groupFolder,
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy_MM')
    );

    // 画像を保存
    const ext      = mimeType === 'image/png' ? 'png' : 'jpg';
    const saveName = fileName.includes('.') ? fileName : `${fileName}.${ext}`;
    const blob     = Utilities.newBlob(Utilities.base64Decode(base64Image), mimeType, saveName);
    const file     = monthFolder.createFile(blob);

    const fileId  = file.getId();
    // ★ 変更①：getUrl() で Drive の正規 URL を取得
    const fileUrl = file.getUrl();

    // ★ 変更②：setSharing を独立した try-catch に分離
    //    組織ポリシーで外部共有が禁止されていても success: true を返せるようにする
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      console.warn('共有設定スキップ（組織ポリシー）:', sharingErr.message);
    }

    return { success: true, fileId: fileId, fileUrl: fileUrl, fileName: file.getName() };

  } catch (err) {
    console.error('saveReceiptImage error:', err);
    return { success: false, error: err.message };
  }
}


/**
 * フォルダ取得または作成
 */
function getOrCreateFolder(parentFolder, folderName) {
  const iter = parentFolder.getFoldersByName(folderName);
  if (iter.hasNext()) return iter.next();
  return parentFolder.createFolder(folderName);
}

/**
 * エクスポートフォルダ取得または作成
 */
function getExportsFolder(groupId, yearMonth) {
  const rootFolderId = getSettings('root_folder_id');
  const rootFolder   = DriveApp.getFolderById(rootFolderId);
  const exportsFolder = getOrCreateFolder(rootFolder, 'exports');
  const groupFolder   = getOrCreateFolder(exportsFolder, groupId);
  return getOrCreateFolder(groupFolder, yearMonth);
}

/**
 * 画像ファイル取得（Blob）
 */
function getReceiptImageBlob(fileId) {
  try {
    return DriveApp.getFileById(fileId).getBlob();
  } catch (e) {
    console.warn('画像取得失敗: ' + fileId);
    return null;
  }
}
