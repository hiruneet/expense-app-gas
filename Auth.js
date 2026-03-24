// ============================================================
// Auth.gs - 認証・セッション管理
// ============================================================

const SESSION_PREFIX = 'session_';
const SESSION_EXPIRE_HOURS = 8;

/**
 * ログイン処理
 */
function login(params) {
  const { groupId, userId, password } = params;
  if (!groupId || !userId || !password) {
    return { success: false, error: 'グループ・ユーザーID・パスワードを入力してください' };
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('users_groups');
  const data = sheet.getDataRange().getValues();

  // ヘッダー: group_id, group_name, user_id, display_name, password_hash, role, created_at, is_active
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === groupId && row[2] === userId && row[7] === true) {
      const inputHash = hashPassword(password);
      if (row[4] === inputHash) {
        // セッション生成
        const token = Utilities.getUuid();
        const expireTime = new Date().getTime() + (SESSION_EXPIRE_HOURS * 3600 * 1000);
        const sessionData = JSON.stringify({
          groupId: row[0],
          groupName: row[1],
          userId: row[2],
          displayName: row[3],
          role: row[5],
          expireTime: expireTime
        });
        PropertiesService.getScriptProperties().setProperty(SESSION_PREFIX + token, sessionData);
        return {
          success: true,
          token: token,
          groupId: row[0],
          groupName: row[1],
          userId: row[2],
          displayName: row[3],
          role: row[5]
        };
      } else {
        return { success: false, error: 'パスワードが正しくありません' };
      }
    }
  }
  return { success: false, error: 'ユーザーが見つかりません' };
}

/**
 * セッション検証
 */
function validateSession(params) {
  const { token } = params;
  if (!token) return { success: false, error: 'トークンがありません' };

  const sessionStr = PropertiesService.getScriptProperties().getProperty(SESSION_PREFIX + token);
  if (!sessionStr) return { success: false, error: 'セッションが無効です' };

  const session = JSON.parse(sessionStr);
  if (new Date().getTime() > session.expireTime) {
    PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + token);
    return { success: false, error: 'セッションが期限切れです' };
  }
  return { success: true, session: session };
}

/**
 * ログアウト
 */
function logout(params) {
  const { token } = params;
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + token);
  }
  return { success: true };
}

/**
 * グループ一覧取得
 */
function getGroups() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('users_groups');
  const data = sheet.getDataRange().getValues();
  const groups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[7] === true && !groups[row[0]]) {
      groups[row[0]] = { groupId: row[0], groupName: row[1] };
    }
  }
  return { success: true, groups: Object.values(groups) };
}

/**
 * ユーザー一覧取得（グループ内）
 */
function getUsers(params) {
  const { groupId } = params;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('users_groups');
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === groupId && row[7] === true) {
      users.push({ userId: row[2], displayName: row[3] });
    }
  }
  return { success: true, users: users };
}

/**
 * SHA-256 ハッシュ生成
 */
function hashPassword(password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * セッション情報取得（内部用）
 */
function getSessionData(token) {
  const result = validateSession({ token });
  if (!result.success) throw new Error(result.error);
  return result.session;
}
