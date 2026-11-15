/**
 * 筑附剣道部管理システム - Google Apps Script バックエンド
 * index.html と同じGASプロジェクトに配置してください。
 *
 * データ: スプレッドシート / 添付ファイル: 非公開のGoogle Driveフォルダ
 *
 * セキュリティ方針:
 *  - 認証はアカウントごとのソルト付き SHA-256 ハッシュで行う（シートに平文は置かない）
 *  - ブラウザへパスワードを自動送信しない（作成者が明示的に「表示」した時だけ返す）
 *  - 表示用の控えはスクリプトプロパティにのみ保存し、作成者以外には返さない
 */
const CONFIG = {
  DATA_SHEET: 'SystemData',
  SESSION_SECONDS: 21600,              // 6時間
  MAX_FILE_BYTES: 12 * 1024 * 1024,    // 添付ファイル上限 12MB
  MAX_PDF_BYTES: 12 * 1024 * 1024,
  SESSION_PREFIX: 'club_session_',
  CHUNK_SIZE: 40000,                   // 1セル上限(50,000文字)対策の分割保存サイズ
  MAX_VALUE_CHARS: 4 * 1024 * 1024,
  LOGIN_MAX_FAILS: 10,
  LOGIN_FAIL_SECONDS: 600,
  ADVISOR_RESET_MONTH: 4,
  ADVISOR_RESET_DAY: 7,
  TREASURER_DEFAULT_ID: 'Contabilità/director@13456',
  TREASURER_DEFAULT_PASSWORD: 'bilancio@12456'
};

/** 保護パスワードの初期値（初回のみScript Propertiesへ複写。以降はプロパティ側が正） */
const SECONDARY_PASS_DEFAULTS = {
  toinkai: 'Paulownia.135',
  dm: 'contatto.246',
  change_save: 'conoscenza.790',
  contact_view: 'rapporto.912',
  update_exec: 'aggiornamento.356',
  official_broadcast: 'ufficiale.13579',
  pass_change: 'cambiare.789',
  advisor_event: 'struttura.109',
  captain_event: 'vacanza.286',
  advisor_delete_user: 'Scomparendo.370',
  captain_upload: 'ombra.583',
  maker_delete_doc: 'Rivista.760',
  maker_passcode: 'Primo.37890',
  advisor_dm_broadcast: 'atmosferico.279',
  vice_broadcast: 'anonimo.4703'
};
const SECONDARY_PASS_PREFIX = 'SEC_PASS_';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('部活動管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 【廃止】共有の会計アカウントは作りません。
 * 会計は「三役」で選ばれた本人（system_sanyi_data.vice2Id）のアカウントを使います。
 * 旧バージョンが作った共有アカウントが残っている場合のみ、この関数で削除できます。
 */
function removeLegacyTreasurerAccount_() {
  const accounts = readData_('system_accounts', []);
  const legacy = accounts.filter(a => a && a.isTreasurer === true);
  if (!legacy.length) return { removed: 0 };
  legacy.forEach(a => vaultDelete_(String(a.id)));
  writeData_('system_accounts', accounts.filter(a => !(a && a.isTreasurer === true)));
  bumpRev_('system_accounts');
  return { removed: legacy.length, ids: legacy.map(a => a.id) };
}

/** 初回のみ実行：データ用スプレッドシート・非公開フォルダ・年次トリガーを作成（何度実行しても安全） */
function setupClubSystem(setupSecret) {
  requireSetupSecret_(setupSecret);
  const props = PropertiesService.getScriptProperties();
  const result = {
    ok: true, spreadsheetUrl: '', manualFolderId: '', attachmentFolderId: '',
    createdSpreadsheet: false, createdFolder: false, trigger: 'ok', warnings: []
  };

  // 1) スプレッドシート
  let spreadsheetId = props.getProperty('CLUB_SPREADSHEET_ID');
  let ss = null;
  if (spreadsheetId) {
    try { ss = SpreadsheetApp.openById(spreadsheetId); }
    catch (e) {
      result.warnings.push('既存スプレッドシートを開けませんでした（削除・権限変更の可能性）。新規に作り直します。');
      ss = null;
    }
  }
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('部活動管理システム データ');
    } catch (e) {
      throw new Error('スプレッドシートを作成できませんでした。Google ドライブの利用権限と保存容量を確認してください。（詳細: ' + e + '）');
    }
    spreadsheetId = ss.getId();
    props.setProperty('CLUB_SPREADSHEET_ID', spreadsheetId);
    result.createdSpreadsheet = true;
  }
  try { ss.setSpreadsheetTimeZone('Asia/Tokyo'); } catch (e) {}

  // 2) データシート（無ければ作成／1行目ヘッダーを修復）
  let sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(CONFIG.DATA_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'JSON Value']]);
    sheet.setFrozenRows(1);
  }
  try {
    const head = sheet.getRange(1, 1, 1, 2).getValues()[0];
    if (String(head[0]) !== 'Key' || String(head[1]) !== 'JSON Value') {
      sheet.getRange(1, 1, 1, 2).setValues([['Key', 'JSON Value']]);
    }
  } catch (e) {}
  result.spreadsheetUrl = ss.getUrl();

  // 3) マニュアルPDF用の非公開フォルダ
  let folderId = props.getProperty('CLUB_MANUAL_FOLDER_ID');
  if (folderId) {
    try { DriveApp.getFolderById(folderId); } catch (e) { folderId = ''; }
  }
  if (!folderId) {
    try {
      const folder = DriveApp.createFolder('部活動管理システム マニュアルPDF（非公開）');
      folderId = folder.getId();
      props.setProperty('CLUB_MANUAL_FOLDER_ID', folderId);
      result.createdFolder = true;
    } catch (e) {
      throw new Error('フォルダを作成できませんでした。Google ドライブの利用権限を確認してください。（詳細: ' + e + '）');
    }
  }
  result.manualFolderId = folderId;

  // 4) 添付ファイル用フォルダ
  try {
    result.attachmentFolderId = getAttachmentFolder_().getId();
  } catch (e) {
    result.warnings.push('添付用フォルダの作成に失敗: ' + e);
  }

  // 5) 年次トリガー（失敗してもセットアップ自体は成功扱い）
  try { installAnnualAdvisorResetTrigger(); }
  catch (e) { result.trigger = 'failed: ' + e; result.warnings.push('年次トリガー作成に失敗しました（後で単体実行してください）'); }
  try { installMonthlyAttendanceReportTrigger(); }
  catch (e) { result.warnings.push('月次出席率レポートのトリガー作成に失敗しました（後で installMonthlyAttendanceReportTrigger を単体実行してください）'); }

  // 6) スクリプトプロパティの既定値（保護パスワード等）を投入
  try { ensureSecondaryPassDefaults_(); } catch (e) { result.warnings.push('保護パスワード初期化スキップ: ' + e); }

  Logger.log('=== セットアップ結果 ===');
  Logger.log('スプレッドシート: ' + result.spreadsheetUrl);
  Logger.log('マニュアルフォルダ: https://drive.google.com/drive/folders/' + folderId);
  if (result.warnings.length) Logger.log('注意: ' + result.warnings.join(' / '));
  Logger.log('次に createInitialCreator を1回実行してください。');
  return result;
}

/** 現在の設定状態を点検して、どこで止まっているかを返す（エディタで単体実行） */
function checkClubSetup() {
  const props = PropertiesService.getScriptProperties();
  const out = {
    setupSecret: !!props.getProperty('SETUP_SECRET'),
    spreadsheetId: props.getProperty('CLUB_SPREADSHEET_ID') || '',
    spreadsheetOk: false,
    sheetOk: false,
    dataRows: 0,
    accounts: 0,
    manualFolderOk: false,
    attachmentFolderOk: false,
    triggers: 0,
    timeZone: ''
  };
  if (out.spreadsheetId) {
    try {
      const ss = SpreadsheetApp.openById(out.spreadsheetId);
      out.spreadsheetOk = true;
      out.timeZone = ss.getSpreadsheetTimeZone();
      const sh = ss.getSheetByName(CONFIG.DATA_SHEET);
      if (sh) {
        out.sheetOk = true;
        out.dataRows = Math.max(0, sh.getLastRow() - 1);
      }
    } catch (e) {}
  }
  try { out.accounts = readData_('system_accounts', []).length; } catch (e) {}
  const mf = props.getProperty('CLUB_MANUAL_FOLDER_ID');
  if (mf) { try { DriveApp.getFolderById(mf); out.manualFolderOk = true; } catch (e) {} }
  const af = props.getProperty('CLUB_ATTACHMENT_FOLDER_ID');
  if (af) { try { DriveApp.getFolderById(af); out.attachmentFolderOk = true; } catch (e) {} }
  try { out.triggers = ScriptApp.getProjectTriggers().length; } catch (e) {}

  Logger.log(JSON.stringify(out, null, 2));
  const problems = [];
  if (!out.spreadsheetOk) problems.push('スプレッドシートが未作成／開けません → setupClubSystem を実行');
  if (!out.sheetOk) problems.push('データシートがありません → setupClubSystem を実行');
  if (!out.accounts) problems.push('アカウント0件 → createInitialCreator を実行');
  if (!out.manualFolderOk) problems.push('マニュアル用フォルダがありません → setupClubSystem を実行');
  Logger.log(problems.length ? ('要対応: ' + problems.join(' / ')) : 'すべて正常です。');
  return out;
}

/** 初回の作成者アカウントを作る。サンプルアカウントは作成しません。 */
function createInitialCreator(setupSecret) {
  requireSetupSecret_(setupSecret);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const accounts = readData_('system_accounts', []);
    if (accounts.length) throw new Error('既にアカウントが存在します。初期作成者は作成済みです。');
    const id = String(Math.floor(100000 + Math.random() * 900000));
    const password = randomPassword_(12);
    const user = {
      id: id, name: '作成者', isAdvisor: false,
      isMaker: true, role: '作成者', grade: '', class: '',
      createdAt: new Date().toISOString()
    };
    setAccountPassword_(user, password);
    writeData_('system_accounts', [user]);
    PropertiesService.getScriptProperties().deleteProperty('SETUP_SECRET');
    Logger.log('初期作成者 ID: ' + id);
    Logger.log('初期作成者 パスワード（この画面を閉じると再表示できません）: ' + password);
    return { id: id, password: password };
  } finally {
    lock.releaseLock();
  }
}

function loginWithCredentials(id, password) {
  id = String(id || '').replace(/[\u3000\s]+/g, ' ').trim();
  password = String(password == null ? '' : password);
  if (!id || !password) return { ok: false, message: 'IDとパスワードを入力してください。' };

  const cache = CacheService.getScriptCache();
  const failKey = 'login_fail_' + id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CONFIG.LOGIN_MAX_FAILS) {
    return { ok: false, message: 'ログインの失敗回数が上限に達しました。しばらく待ってから再度お試しください。' };
  }

  const accounts = readData_('system_accounts', []);
  let user = accounts.find(a => String(a.id).trim() === id);
  if (!user) user = accounts.find(a => String(a.id).trim().toLowerCase() === id.toLowerCase());
  if (!user || !verifyAccountPassword_(user, password)) {
    cache.put(failKey, String(fails + 1), CONFIG.LOGIN_FAIL_SECONDS);
    return {
      ok: false,
      message: accounts.length
        ? 'IDまたはパスワードが正しくありません。'
        : '初期設定が未完了です。管理者がsetupClubSystemとcreateInitialCreatorを実行してください。'
    };
  }
  cache.remove(failKey);

  // 退部・卒業・停止アカウントはログイン不可（記録は保持）
  const acctStatus = String(user.status || 'active');
  if (acctStatus !== 'active') {
    const label = acctStatus === 'retired' ? '退部' : (acctStatus === 'graduated' ? '卒業' : '停止');
    return { ok: false, message: 'このアカウントは「' + label + '」のためログインできません。顧問または作成者にご連絡ください。' };
  }
  try { appendAudit_(user.id, 'ログイン', String(user.id), String(user.name || '')); } catch (e) {}

  // 旧バージョンの平文パスワードが残っている場合は、ログイン成功時にハッシュへ移行する
  if (!user.passHash || !user.salt || typeof user.pass === 'string') {
    setAccountPassword_(user, password);
    writeData_('system_accounts', accounts);
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  const sanyi = readData_('system_sanyi_data', {});
  const isMaker = !!(user.isMaker || user.role === '作成者');
  const isCaptain = isMaker || [sanyi.captainId, sanyi.vice1Id, sanyi.vice2Id].some(v => String(v || '') === String(user.id));
  // 会計権限の交代は毎年1月1日（三役の交代では即時には移さない）
  try { applyTreasurerHandoverIfDue_(); } catch (e) {}
  const treasurerId = resolveTreasurerId_();
  const isTreasurer = String(user.id) === treasurerId;
  const formerIds = Array.isArray(sanyi.formerCaptainIds) ? sanyi.formerCaptainIds.map(String) : [];
  const isFormerCaptain = !isMaker && (!!user.toinkaiOnly || String(user.role || '') === '前主将')
    && String(user.id) !== String(sanyi.captainId || '');
  const session = {
    id: String(user.id), isMaker: isMaker, isAdvisor: !!user.isAdvisor,
    isCaptain: isCaptain, isTreasurer: isTreasurer, role: String(user.role || ''),
    toinkaiOnly: isFormerCaptain
  };
  CacheService.getScriptCache().put(CONFIG.SESSION_PREFIX + token, JSON.stringify(session), CONFIG.SESSION_SECONDS);

  const safeUser = publicAccount_(user);
  safeUser.isCaptain = isCaptain;
  safeUser.isTreasurer = isTreasurer;
  safeUser.toinkaiOnly = isFormerCaptain;
  if (isFormerCaptain) safeUser.isCaptain = false;
  return { ok: true, token: token, user: safeUser, data: dataForSession_(session) };
}

/** パスワード変更。条件検証とセッション失効を必ずサーバー側で実施します。 */
function changeAccountPassword(token, currentPassword, newPassword) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  currentPassword = String(currentPassword == null ? '' : currentPassword);
  newPassword = String(newPassword == null ? '' : newPassword);
  if (newPassword.length < 8) throw new Error('新しいパスワードは8文字以上にしてください。');
  if (!/[!-/:-@\[-`{-~]/.test(newPassword)) throw new Error('新しいパスワードには記号を1つ以上含めてください。');
  if (!/[0-9]/.test(newPassword)) throw new Error('新しいパスワードには数字を1つ以上含めてください。');
  if (!/[A-Z]/.test(newPassword)) throw new Error('新しいパスワードにはアルファベットの大文字を1つ以上含めてください。');
  if (newPassword === currentPassword) throw new Error('新しいパスワードは現在と異なるものにしてください。');

  const accounts = readData_('system_accounts', []);
  const account = accounts.find(a => String(a.id) === String(session.id));
  if (!account || !verifyAccountPassword_(account, currentPassword)) {
    throw new Error('現在のパスワードが正しくありません。');
  }
  setAccountPassword_(account, newPassword);
  writeData_('system_accounts', accounts);

  CacheService.getScriptCache().remove(CONFIG.SESSION_PREFIX + String(token));
  return { ok: true };
}

/** 作成者によるパスワード再発行。新しいパスワードはこの戻り値でのみ1回返す。 */
function resetAccountPassword(token, targetId) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!session.isMaker) throw new Error('パスワードの再発行は作成者のみ実行できます。');
  const accounts = readData_('system_accounts', []);
  const account = accounts.find(a => String(a.id) === String(targetId));
  if (!account) throw new Error('対象のアカウントが見つかりません。');
  const password = randomPassword_(12);
  setAccountPassword_(account, password);
  writeData_('system_accounts', accounts);
  return { ok: true, id: String(account.id), name: account.name, password: password };
}

/**
 * 作成者ツール用：登録済みパスワードを表示（作成者のみ）。
 * 自動では一切返さず、この関数を明示的に呼んだときだけ返します。
 */
function revealAccountPassword(token, targetId) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!session.isMaker) throw new Error('パスワードの表示は作成者のみ実行できます。');
  const accounts = readData_('system_accounts', []);
  const account = accounts.find(a => String(a.id) === String(targetId));
  if (!account) throw new Error('対象のアカウントが見つかりません。');
  const stored = vaultGet_(account.id);
  if (!stored) {
    return {
      ok: false,
      message: 'このアカウントは旧形式で保存されているため表示できません。「再発行」で新しいパスワードを発行してください。'
    };
  }
  return { ok: true, id: String(account.id), name: account.name, password: stored };
}


/**
 * アカウント専用保存API。
 * knownIds = 端末が最後に同期したID一覧。
 *  - 送られてきたアカウントは更新
 *  - knownIds にあり今回送られてこなかったIDのみ削除
 *  - knownIds にない（他端末で追加された）アカウントは保持
 * 不正な行があっても全体を失敗させず、結果を返す。
 */
/** 登録画面用：登録済みアカウントの有無と、ログイン中ユーザーの登録権限をサーバー側で確認 */
function getClubRegistrationStatus(token) {
  const accounts = readData_('system_accounts', []);
  const safeAccounts = Array.isArray(accounts) ? accounts : [];
  const session = requireSession_(token);
  const isMaker = !!(session && session.isMaker);
  return {
    ok: true,
    hasAccounts: safeAccounts.length > 0,
    accountCount: safeAccounts.length,
    isMaker: isMaker,
    // 作成者がログイン中の場合だけ、ID採番用に公開アカウント情報を返す
    accounts: isMaker ? safeAccounts.map(publicAccount_) : []
  };
}

/**
 * 部員・顧問の新規登録API。
 * 初回は作成者パスコードで最初の作成者だけを作成し、
 * 2人目以降はログイン済み作成者だけが登録できる。
 * ブラウザ側localStorageだけに登録して終わる旧処理を置き換える。
 */
function registerClubAccount(token, jsonAccount, bootstrapPasscode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let incoming;
    try { incoming = JSON.parse(String(jsonAccount || '{}')); }
    catch (e) { throw new Error('登録データの形式が正しくありません。'); }
    if (!incoming || typeof incoming !== 'object' || Array.isArray(incoming)) {
      throw new Error('登録データの形式が正しくありません。');
    }

    const accounts = readData_('system_accounts', []);
    const firstAccount = !Array.isArray(accounts) || accounts.length === 0;
    let session = null;

    if (firstAccount) {
      const props = PropertiesService.getScriptProperties();
      const expected = String(props.getProperty('SEC_PASS_maker_passcode') || SECONDARY_PASS_DEFAULTS.maker_passcode || '');
      const given = String(bootstrapPasscode || '').trim();
      if (!expected || given !== expected) throw new Error('作成者アカウント作成用パスコードが正しくありません。');
    } else {
      session = requireSession_(token);
      if (!session) throw new Error('部員登録には作成者アカウントでログインしてください。');
      if (!session.isMaker) throw new Error('部員登録は作成者のみ実行できます。');
    }

    const id = String(incoming.id || '').trim();
    const name = String(incoming.name || '').trim();
    const password = String(incoming.initialPassword || '');
    if (!id || !name || !password) throw new Error('ID・名前・初期パスワードが不足しています。');
    if (name.length > 80 || id.length > 40 || password.length < 8) throw new Error('登録内容が不正です。');
    if (accounts.some(a => String(a.id) === id)) throw new Error('このIDは既に使用されています。');
    if (accounts.some(a => String(a.name || '').trim() === name)) throw new Error('同じ名前のアカウントが既に登録されています。');

    const isAdvisor = !!incoming.isAdvisor;
    if (isAdvisor && accounts.filter(a => a && a.isAdvisor).length >= 2) {
      throw new Error('顧問アカウントは2名までです。');
    }

    const user = {
      id: id,
      name: name,
      isAdvisor: isAdvisor,
      isMaker: firstAccount,
      grade: isAdvisor ? '' : String(incoming.grade || '1'),
      class: isAdvisor ? '' : String(incoming.class || '1組'),
      gender: String(incoming.gender || ''),
      role: firstAccount ? '作成者' : (isAdvisor ? '顧問' : '一般'),
      status: 'active',
      createdAt: new Date().toISOString()
    };
    if (!isAdvisor && incoming.entryYear != null) {
      user.entryYear = Number(incoming.entryYear);
      user.kisei = String(incoming.kisei || '');
      user.joinedAt = String(incoming.joinedAt || new Date().toISOString());
    }

    setAccountPassword_(user, password);
    accounts.push(user);
    writeData_('system_accounts', accounts);
    saveAccountsBackup_(accounts);
    const rev = bumpRev_('system_accounts');
    if (session) {
      try { appendAudit_(session.id, 'アカウント登録', id, name); } catch (e) {}
    }
    return {
      ok: true,
      firstAccount: firstAccount,
      rev: rev,
      accounts: accounts.map(publicAccount_),
      user: publicAccount_(user)
    };
  } finally {
    lock.releaseLock();
  }
}

/** 初期パスワード生成（Pass@＋3桁の数字） */
function makeSelfRegPassword_() {
  return 'Pass@' + String(Math.floor(Math.random() * 1000)).padStart(3, '0');
}

/** 公開登録：部員本人が自分の一般アカウントを作成する。作成者・顧問権限は付与しない。
 *  kind に「顧問」を指定すると顧問アカウント（IDは000001／000002、最大2名）を作成する。
 *  初期パスワードは入力させず、サーバー側で「Pass@＋3桁の数字」を自動生成し、戻り値で1回だけ返す。 */
function selfRegisterClubMember(name, grade, className, gender, kind) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    name = String(name || '').replace(/[\u3000\s]+/g, ' ').trim();
    grade = String(grade == null ? '' : grade).trim();
    className = String(className || '').trim();
    gender = String(gender || '').trim();
    kind = String(kind || '部員').trim();
    const isAdvisor = (kind === '顧問');
    if (!name || name.length > 80) throw new Error('氏名を正しく入力してください。');
    if (!gender) throw new Error('性別を選択してください。');
    if (!isAdvisor) {
      grade = Number(grade);
      if (![1, 2, 3].includes(grade)) throw new Error('学年の指定が正しくありません。');
      if (!/^[1-5]組$/.test(className)) throw new Error('クラスの指定が正しくありません。');
    }

    const accounts = readData_('system_accounts', []);
    if (!Array.isArray(accounts) || accounts.length === 0) throw new Error('システムの初期設定が未完了です。作成者の登録後にお試しください。');
    if (accounts.some(a => String(a.name || '').replace(/\s+/g, '') === name.replace(/\s+/g, '') && String(a.status || 'active') === 'active')) {
      throw new Error('同じ氏名のアカウントが既に登録されています。顧問または作成者に確認してください。');
    }

    const password = makeSelfRegPassword_();
    const now = new Date();
    let id = '';
    const user = {
      id: '', name: name, isAdvisor: isAdvisor, isMaker: false,
      role: isAdvisor ? '顧問' : '一般',
      grade: isAdvisor ? '顧問' : String(grade), class: isAdvisor ? '' : className,
      gender: gender,
      joinedAt: now.toISOString(), createdAt: now.toISOString(), status: 'active'
    };

    if (isAdvisor) {
      const used = accounts.filter(a => a && a.isAdvisor).map(a => Number(a.id)).filter(n => !isNaN(n));
      let n = 1;
      while (used.indexOf(n) >= 0 && n <= 2) n++;
      if (n > 2) throw new Error('顧問アカウントは2名までです（000001／000002）。');
      id = String(n).padStart(6, '0');
    } else {
      const fy = Number(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy')) - (Number(Utilities.formatDate(now, 'Asia/Tokyo', 'M')) < 4 ? 1 : 0);
      const entryYear = fy - (grade - 1);
      const seq = nextMemberSeq_(accounts, entryYear);
      if (seq > 99) throw new Error('この代の登録上限（99名）に達しています。');
      id = String(entryYear % 100).padStart(2, '0') + String(kiseiOfEntryYear_(entryYear)).padStart(2, '0') + String(seq).padStart(2, '0');
      user.entryYear = entryYear;
      user.kisei = String(kiseiOfEntryYear_(entryYear));
    }
    if (accounts.some(a => String(a.id) === id)) throw new Error('IDの採番が重複しました。もう一度登録してください。');
    user.id = id;

    setAccountPassword_(user, password);
    accounts.push(user);
    writeData_('system_accounts', accounts);
    saveAccountsBackup_(accounts);
    bumpRev_('system_accounts');
    try { appendAudit_('self-registration', '本人による登録', id, name); } catch (e) {}
    return { ok: true, id: id, name: name, kind: kind, password: password, user: publicAccount_(user), total: accounts.length };
  } finally {
    lock.releaseLock();
  }
}

function apiSaveAccounts(token, jsonAccounts, jsonKnownIds) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!session.isMaker) throw new Error('アカウントを保存する権限がありません。');
  let incoming;
  try { incoming = JSON.parse(jsonAccounts || '[]'); } catch (e) { throw new Error('アカウントデータの形式が正しくありません。'); }
  if (!Array.isArray(incoming)) throw new Error('アカウントデータの形式が正しくありません。');
  if (!incoming.length) return { ok: false, message: 'アカウントが空のため保存しませんでした（既存データは保持）' };
  let known = [];
  try { known = JSON.parse(jsonKnownIds || '[]'); } catch (e) { known = []; }
  if (!Array.isArray(known)) known = [];
  const knownSet = {};
  known.forEach(id => { if (id != null && id !== '') knownSet[String(id)] = true; });

  const stored = readData_('system_accounts', []);
  const byId = {};
  stored.forEach(a => { if (a && a.id != null) byId[String(a.id)] = a; });

  const merged = [];
  const seen = {};
  let skipped = 0;
  incoming.forEach(raw => {
    const acc = Object.assign({}, raw || {});
    const id = String(acc.id == null ? '' : acc.id).trim();
    if (!id || seen[id]) { skipped++; return; }
    seen[id] = true;
    acc.id = id;
    const existing = byId[id];
    if (!session.isMaker) {
      // 作成者以外は管理者権限の付与・剥奪をできない
      if (existing) { acc.isMaker = !!existing.isMaker; acc.isAdvisor = !!existing.isAdvisor; }
      else { acc.isMaker = false; acc.isAdvisor = false; }
    }
    const freshPassword = String(acc.initialPassword || acc.pass || '');
    delete acc.initialPassword;
    delete acc.pass;
    if (existing && existing.passHash && existing.salt) {
      acc.passHash = existing.passHash;
      acc.salt = existing.salt;
      acc.passHistory = Array.isArray(existing.passHistory) ? existing.passHistory : [];
      acc.passwordChangedAt = existing.passwordChangedAt || '';
    } else if (existing && typeof existing.pass === 'string' && existing.pass !== '') {
      setAccountPassword_(acc, existing.pass);
    } else if (freshPassword) {
      setAccountPassword_(acc, freshPassword);
    }
    merged.push(acc);
  });

  let kept = 0;
  stored.forEach(a => {
    if (!a || a.id == null) return;
    const id = String(a.id);
    if (seen[id]) return;
    if (knownSet[id]) { merged.push(a); kept++; return; } // 明示的な削除API以外では、退部者以外を消さない
    merged.push(a);
    kept++;
  });

  const beforeCount = stored.length;
  writeData_('system_accounts', merged);
  saveAccountsBackup_(merged);
  const rev = bumpRev_('system_accounts');
  try { appendAudit_(session.id, '名簿を保存', 'system_accounts', '保存' + merged.length + '件（変更前' + beforeCount + '件）'); } catch (e) {}
  return { ok: true, rev: rev, saved: merged.length, kept: kept, skipped: skipped, accounts: merged.map(publicAccount_) };
}


/** 三役の確定状態（未ログイン＝ゲストでも参照できる） */
function apiSanyiState() {
  const s = readData_('system_sanyi_data', {}) || {};
  return {
    ok: true,
    locked: !!s.locked,
    confirmedAt: String(s.confirmedAt || ''),
    year: Number(s.year || 0),
    hasSanyi: !!(s.captainId && s.vice1Id && s.vice2Id)
  };
}

const SANYI_LOCKED_MSG = '三役は確定済みです。変更は9月の引継ぎまでできません。';

/** 在席（ログイン中）ビーコン。通話の着信可否に使う */
function apiHeartbeat(token) {
  const session = requireSession_(token);
  if (!session) return { ok: false };
  CacheService.getScriptCache().put('PRESENCE_' + String(session.id), String(Date.now()), 130);
  return { ok: true };
}

function isOnline_(id) {
  try { return !!CacheService.getScriptCache().get('PRESENCE_' + String(id)); } catch (e) { return false; }
}

/** 添付ファイルのAIチェック（Gemini APIキーがあれば画像・PDFも判定） */
function apiCheckAttachment(base64, fileName, mimeType, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  const toxic = /殺す|殴る|ボコボコ|殴打|暴力|暴行|脅迫|脅す|晒す|死ね|犯罪|違法|万引き|カツアゲ|放火|爆破|薬物|ナイフ|包丁|凶器/;
  const name = String(fileName || '');
  let score = 0, reasons = [], engine = 'none', text = '';
  if (toxic.test(name)) { score += 60; reasons.push('ファイル名'); }

  // ---- AIキー不要の検査（拡張子・偽装・危険な形式） ----
  const lower = name.toLowerCase();
  const extM = lower.match(/\.([a-z0-9]+)$/);
  const ext = extM ? extM[1] : '';
  const allowedExt = ['pdf', 'jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'];
  if (allowedExt.indexOf(ext) < 0) {
    score += 70;
    reasons.push('対応していない形式（.' + (ext || '拡張子なし') + '）');
  }
  if (/\.(exe|bat|cmd|js|vbs|scr|jar|msi|apk|zip|rar|7z|html|htm|php|ps1|sh)\./.test(lower)
      || /\.(exe|bat|cmd|vbs|scr|jar|msi|apk|zip|rar|7z|html|htm|php|ps1|sh)$/.test(lower)) {
    score += 70;
    reasons.push('実行ファイル／書庫の疑い（二重拡張子）');
  }
  const sigMap = { pdf: 'JVBERi0', jpg: '/9j/', jpeg: '/9j/', png: 'iVBORw0KGgo', gif: 'R0lGOD', webp: 'UklGR' };
  const b64str = String(base64 || '');
  if (b64str && sigMap[ext] && b64str.indexOf(sigMap[ext]) !== 0) {
    score += 70;
    reasons.push('中身が拡張子と一致しません（偽装の可能性）');
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const configured = props.getProperty('GEMINI_MODEL') || '';
  const models = [];
  [configured, 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-3.8-flash'].forEach(function (m) {
    if (m && models.indexOf(m) < 0) models.push(m);
  });
  let aiError = '';
  if (apiKey && base64 && base64.length < 5500000) {
    const prompt = 'あなたは学校部活動の連絡ツールの安全監視です。添付画像/PDFに、暴力・脅迫・いじめ・わいせつ・犯罪示唆・個人情報の晒しがあれば検出してください。'
      + 'JSONのみで回答: {"dangerous": true/false, "score": 0-100, "reasons": ["..."]}';
    const payload = {
      contents: [{ parts: [
        { text: prompt },
        { inline_data: { mime_type: String(mimeType || 'application/octet-stream'), data: String(base64) } }
      ] }],
      generationConfig: { temperature: 0, maxOutputTokens: 200 }
    };
    for (let mi = 0; mi < models.length && engine === 'none'; mi++) {
      try {
        const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(models[mi]) + ':generateContent', {
          method: 'post',
          contentType: 'application/json',
          headers: { 'x-goog-api-key': String(apiKey) },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        const code = res.getResponseCode();
        const body = res.getContentText();
        if (code !== 200) { aiError = 'HTTP ' + code + ' / ' + models[mi] + ' / ' + String(body).slice(0, 160); continue; }
        const parsed = JSON.parse(body);
        const cand = (parsed.candidates || [])[0] || {};
        const parts = (cand.content || {}).parts || [];
        text = (parts[0] || {}).text || '';
        const m = text.match(/\{[\s\S]*\}/);
        if (m) {
          const j = JSON.parse(m[0]);
          engine = 'gemini:' + models[mi];
          if (j.dangerous) { score = Math.max(score, Number(j.score || 70)); }
          (j.reasons || []).forEach(function (r) { reasons.push(String(r)); });
        } else {
          aiError = 'no-json / ' + models[mi] + ' / ' + String(text).slice(0, 120);
        }
      } catch (e) { aiError = String(e); }
    }
  }
  if (engine === 'none') engine = 'rule';
  return { ok: true, engine: engine, score: score, dangerous: score >= 50, reasons: reasons, error: aiError, ai: !!apiKey };
}

/** 添付ファイルの表示URL（Driveの権限に従う） */
function apiAttachmentUrl(fileId, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  const id = String(fileId || '');
  if (!id) throw new Error('ファイルが指定されていません。');
  const file = DriveApp.getFileById(id);
  return { ok: true, url: file.getUrl(), name: file.getName() };
}

/** ゲストでも三役登録できるよう、公開してよい項目だけ返す */
function apiPublicMembers() {
  const accounts = readData_('system_accounts', []);
  const members = accounts.filter(function (a) {
    return a && !a.isAdvisor && String(a.grade || '') !== '卒業';
  }).map(function (a) {
    // クラスはログインなしでも表示できるよう公開する（パスワード等は返さない）
    return {
      id: String(a.id), name: String(a.name || ''), grade: String(a.grade || ''),
      class: String(a.class || ''), role: String(a.role || '一般'), isAdvisor: !!a.isAdvisor
    };
  });
  return { ok: true, members: members };
}

/** ゲスト向け公開結果（閲覧のみ）。 */
function apiPublicResults() {
  const tournaments = readData_('system_tournaments', []);
  const exams = readData_('system_exams', []);
  return {
    ok: true,
    tournaments: Array.isArray(tournaments) ? tournaments : [],
    exams: Array.isArray(exams) ? exams : []
  };
}

/** 三役の3名を検証して名簿ごと返す */
function sanyiIds_(captainId, vice1Id, vice2Id) {
  const ids = [String(captainId || '').trim(), String(vice1Id || '').trim(), String(vice2Id || '').trim()];
  if (ids.some(function (v) { return !v; })) throw new Error('主将・副将・会計の全員を指定してください。');
  if (new Set(ids).size !== 3) throw new Error('同じ部員を複数の役に重複して指定できません。');
  const accounts = readData_('system_accounts', []);
  const byId = {};
  accounts.forEach(function (a) { if (a && a.id != null) byId[String(a.id)] = a; });
  ids.forEach(function (id) {
    const acc = byId[id];
    if (!acc) throw new Error('指定された部員が名簿に見つかりません。');
    if (acc.isAdvisor) throw new Error('顧問は三役に指定できません。');
    if (String(acc.grade || '') === '卒業') throw new Error('卒業した部員は三役に指定できません。');
  });
  return { ids: ids, accounts: accounts };
}

/** 三役を確定保存（ロック）。以後は引継ぎ以外で変更できない */
function sanyiCommit_(captainId, vice1Id, vice2Id, actorLabel, doc, kind) {
  const v = sanyiIds_(captainId, vice1Id, vice2Id);
  const ids = v.ids, accounts = v.accounts;
  const now = readData_('system_sanyi_data', {}) || {};
  const jst = jstParts_();
  const history = Array.isArray(now.history) ? now.history.slice(-9) : [];
  history.push({
    at: new Date().toISOString(),
    by: String(actorLabel || ''),
    kind: String(kind || 'register'),
    captainId: ids[0], vice1Id: ids[1], vice2Id: ids[2],
    doc: doc || null
  });
  const prevCaptain = String(now.captainId || '');
  let formerCaptainIds = Array.isArray(now.formerCaptainIds) ? now.formerCaptainIds.map(String) : [];
  if (String(kind) === 'handover' && prevCaptain && prevCaptain !== String(ids[0])) {
    formerCaptainIds = formerCaptainIds.filter(function (x) { return x !== prevCaptain; });
    formerCaptainIds.push(prevCaptain);
    formerCaptainIds = formerCaptainIds.slice(-5);   // 直近5代まで保持
  }
  const next = Object.assign({}, now, {
    captainId: ids[0], vice1Id: ids[1], vice2Id: ids[2],
    formerCaptainIds: formerCaptainIds,
    locked: true,
    confirmedAt: new Date().toISOString(),
    confirmedBy: String(actorLabel || ''),
    year: jst.y,
    history: history
  });
  writeData_('system_sanyi_data', next);
  bumpRev_('system_sanyi_data');

  accounts.forEach(function (a) {
    if (!a || a.id == null) return;
    const id = String(a.id);
    if (id === ids[0]) { a.role = '主将'; a.toinkaiOnly = false; }
    else if (id === ids[1]) { a.role = '副将'; a.toinkaiOnly = false; }
    else if (id === ids[2]) { a.role = '会計'; a.toinkaiOnly = false; }
    else if (String(kind) === 'handover' && prevCaptain && id === prevCaptain) { a.role = '前主将'; a.toinkaiOnly = true; }
    else if (a.role !== '作成者' && a.role !== '前主将' && a.role !== '旧主将' && a.role !== '顧問') a.role = '一般';
  });
  writeData_('system_accounts', accounts);
  bumpRev_('system_accounts');

  registerTreasurerCandidate_(ids[2]);
  return { ok: true, sanyi: next, treasurer: treasurerState_() };
}

/** 三役を登録（ログイン済みなら誰でも可。確定後は不可） */
function apiSetSanyi(token, captainId, vice1Id, vice2Id) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  const now = readData_('system_sanyi_data', {}) || {};
  if (now.locked) throw new Error(SANYI_LOCKED_MSG);
  return sanyiCommit_(captainId, vice1Id, vice2Id, String(session.name || session.id || ''), null, 'register');
}

/** ゲスト（未ログイン）でも登録できる。確定後は不可 */
function apiSetSanyiGuest(captainId, vice1Id, vice2Id) {
  const now = readData_('system_sanyi_data', {}) || {};
  if (now.locked) throw new Error(SANYI_LOCKED_MSG);
  return sanyiCommit_(captainId, vice1Id, vice2Id, 'ゲスト', null, 'register_guest');
}

/** 三役引継ぎ（9月〜・主将または作成者）。資料を添えて新しい三役へ交代する */
function apiHandoverSanyi(token, captainId, vice1Id, vice2Id, docJson) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  const sanyi = readData_('system_sanyi_data', {}) || {};
  const isMaker = !!session.isMaker;
  const isCaptain = [sanyi.captainId, sanyi.vice1Id, sanyi.vice2Id]
    .some(function (v) { return String(v || '') === String(session.id); });
  if (!isMaker && !isCaptain) throw new Error('三役引継ぎは主将（または作成者）のみ実行できます。');
  const jst = jstParts_();
  if (!isMaker && jst.m < 9) throw new Error('三役引継ぎは9月から実行できます。');
  let doc = null;
  try { doc = docJson ? JSON.parse(String(docJson)) : null; } catch (e) { doc = null; }
  const res = sanyiCommit_(captainId, vice1Id, vice2Id, String(session.name || session.id || ''), doc, 'handover');
  notifyAdvisors('【三役引継ぎ】新しい三役が確定しました。\n主将: ' + String(captainId)
    + ' / 副将: ' + String(vice1Id) + ' / 会計: ' + String(vice2Id)
    + '\n資料: ' + (doc && doc.fileName ? doc.fileName : 'なし'), '【公式】三役引継ぎ');
  return res;
}

/** 音声通話のシグナリング（GASにソケットは無いためキャッシュを中継に使う） */
function apiCall(token, action, payloadJson) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  const uid = String(session.id);
  let p = {};
  try { p = JSON.parse(payloadJson || '{}'); } catch (e) { p = {}; }
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  const key = id => 'CALLQ_' + String(id);
  const qGet = id => { try { return JSON.parse(cache.get(key(id)) || '[]'); } catch (e) { return []; } };
  const qPut = (id, q) => cache.put(key(id), JSON.stringify(q.slice(-30)), 300);
  const qPush = (id, item) => { const q = qGet(id); q.push(item); qPut(id, q); };

  if (action === 'poll') {
    lock.waitLock(3000);
    try {
      const q = qGet(uid);
      if (q.length) qPut(uid, []);
      return { ok: true, items: q };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  }
  if (action === 'invite') {
    const room = String(p.room || '');
    if (room.indexOf('sanyi') === 0) throw new Error('三役専用グループDMは通話の対象外です。');
    const targets = (Array.isArray(p.targets) ? p.targets : []).map(String).filter(t => t && t !== uid).slice(0, 3);
    if (!targets.length) throw new Error('通話相手がいません。');
    // 1対1の通話は、相手がログインしていないと着信させない（グループ通話は対象外）
    const isGroupCall = !!p.isGroup;
    if (!isGroupCall && targets.length === 1 && !isOnline_(targets[0])) {
      const accs = readData_('system_accounts', []);
      const t = accs.filter(function (a) { return a && String(a.id) === targets[0]; })[0];
      throw new Error((t ? String(t.name) : '相手') + ' さんはログインしていません。着信できません。');
    }
    lock.waitLock(3000);
    try {
      targets.forEach(t => qPush(t, {
        type: 'invite', from: uid, fromName: String(p.fromName || ''), room: room,
        label: String(p.label || ''), targets: targets.concat([uid]), ts: Date.now()
      }));
    } finally { try { lock.releaseLock(); } catch (e) {} }
    return { ok: true, room: room };
  }
  if (action === 'signal') {
    const to = String(p.to || '');
    if (!to || to === uid) return { ok: false };
    lock.waitLock(3000);
    try { qPush(to, { type: 'signal', from: uid, room: String(p.room || ''), data: p.data || null, ts: Date.now() }); }
    finally { try { lock.releaseLock(); } catch (e) {} }
    return { ok: true };
  }
  if (action === 'accept' || action === 'decline' || action === 'leave') {
    const to = String(p.to || '');
    if (!to) return { ok: false };
    lock.waitLock(3000);
    try { qPush(to, { type: action, from: uid, room: String(p.room || ''), ts: Date.now() }); }
    finally { try { lock.releaseLock(); } catch (e) {} }
    return { ok: true };
  }
  return { ok: false };
}

function apiLoadData(token) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログイン期限が切れました。再ログインしてください。' };
  try { CacheService.getScriptCache().put('PRESENCE_' + String(session.id), String(Date.now()), 130); } catch (e) {}
  return { ok: true, data: dataForSession_(session), revs: getAllRevs_() };
}

function apiSaveData(key, jsonValue, token, baseRev) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  key = String(key || '');
  if (!/^system_[a-zA-Z0-9_]+$/.test(key)) throw new Error('保存できないデータ名です。');
  if (String(jsonValue == null ? '' : jsonValue).length > CONFIG.MAX_VALUE_CHARS) {
    throw new Error('データ量が大きすぎます。添付ファイルはGoogleドライブ保存をご利用ください。');
  }

  const isMaker = !!session.isMaker;
  // 旧主将アカウントは「桐蔭会雑誌原稿」の提出のみ
  if (session.toinkaiOnly && key !== 'system_toinkai_docs') {
    throw new Error('旧主将アカウントは桐蔭会雑誌原稿の提出のみ利用できます。');
  }
  const permissionMap = {
    system_accounts: isMaker || !!session.isCaptain,
    system_manual_versions: isMaker,
    system_sanyi_data: isMaker || !!session.isCaptain,
    system_schedule_events: isMaker || session.isAdvisor || session.isCaptain,
    system_tournaments: isMaker || session.isCaptain,
    system_exams: isMaker || session.isAdvisor,
    system_yearly_plans: isMaker || session.isAdvisor,
    system_toinkai_docs: isMaker || session.isCaptain || !!session.toinkaiOnly,
    system_handover_docs: isMaker || session.isCaptain,
    system_meetings: isMaker || session.isAdvisor,
    system_practices: isMaker || session.isAdvisor,
    system_last_promotion_year: isMaker,
    system_advisor_reset_year: isMaker,
    system_accounting_docs: isMaker || !!session.isTreasurer || session.isAdvisor || isAdminToolUnlocked_('accounting', session.id),
    system_accounting_state: isMaker || !!session.isTreasurer || isAdminToolUnlocked_('accounting', session.id),
    system_attendance_alert_state: isMaker || session.isAdvisor || session.isCaptain,
    system_monthly_reports: isMaker
  };
  if (Object.prototype.hasOwnProperty.call(permissionMap, key) && !permissionMap[key]) {
    throw new Error('このデータを変更する権限がありません。');
  }

  // 楽観ロック：別端末で更新済みなら上書きせず、最新値を返して画面側で差し替える
  // ※ accounts / messages はサーバー側でマージするため競合判定の対象外（同時送信で消えないように）
  // アカウント名簿は空配列での上書きを拒否（コード差し替え時の消失防止）
  if (key === 'system_accounts') {
    let incomingAccounts = [];
    try { incomingAccounts = JSON.parse(jsonValue || '[]'); } catch (e) { incomingAccounts = []; }
    if (!Array.isArray(incomingAccounts) || incomingAccounts.length === 0) {
      const existing = readData_('system_accounts', []);
      if (Array.isArray(existing) && existing.length) {
        return { ok: false, message: 'アカウントが空のため保存しませんでした（既存データを保持）', data: existing, rev: getRev_(key) };
      }
    }
  }

  const MERGE_KEYS = { system_accounts: true, system_messages: true };
  const currentRev = getRev_(key);
  if (!MERGE_KEYS[key] && baseRev !== undefined && baseRev !== null && baseRev !== '' && Number(baseRev) !== currentRev) {
    return { ok: false, conflict: true, data: readRaw_(key), rev: currentRev };
  }

  if (key === 'system_accounts') {
    writeRawData_(key, jsonValue ? mergeAccountsWithSecrets_(String(jsonValue)) : '');
    return { ok: true, rev: bumpRev_(key) };
  }

  if (key === 'system_messages' && !isMaker) {
    const oldMessages = readData_(key, []);
    let newMessages;
    try { newMessages = jsonValue ? JSON.parse(jsonValue) : []; } catch (e) { throw new Error('通知データの形式が正しくありません。'); }
    const protectedMessages = oldMessages.filter(m =>
      String(m.senderId) !== String(session.id) &&
      String(m.receiverId) !== String(session.id) &&
      String(m.receiverId) !== 'ALL'
    );
    const merged = protectedMessages.concat((newMessages || []));
    writeData_(key, merged);
    return { ok: true, rev: bumpRev_(key) };
  }
  if (key === 'system_sanyi_data' && jsonValue) {
    try {
      const sd = JSON.parse(String(jsonValue));
      if (sd && sd.vice2Id) registerTreasurerCandidate_(sd.vice2Id);
    } catch (e) { /* 形式不正は保存自体を妨げない */ }
  }
  writeRawData_(key, String(jsonValue == null ? '' : jsonValue));
  const newRev = bumpRev_(key);
  const AUDIT_KEYS = {
    system_sanyi_data: '三役を変更',
    system_accounting_docs: '会計書類を変更',
    system_accounting_state: '会計状態を変更',
    system_schedule_events: '予定を変更',
    system_yearly_plans: '年間予定を変更',
    system_exams: '審査予定を変更',
    system_practices: '稽古予定を変更',
    system_meetings: '会議を変更',
    system_handover_docs: '引継ぎ資料を変更',
    system_toinkai_docs: '桐蔭会原稿を変更',
    system_manual_versions: 'マニュアルを変更',
  };
  if (AUDIT_KEYS[key]) {
    let teaser = '';
    if (key === 'system_sanyi_data') {
      try { const sd = JSON.parse(String(jsonValue || '{}')); teaser = '主将:' + sd.captainId + ' / 副将:' + sd.vice1Id + ' / 会計:' + sd.vice2Id; } catch (e) {}
    }
    try { appendAudit_(session.id, AUDIT_KEYS[key], key, teaser); } catch (e) {}
  }
  return { ok: true, rev: newRev };
}

// ---- ファイル（Googleドライブ・非公開） ----
function getAttachmentFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CLUB_ATTACHMENT_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 再作成へ */ }
  }
  const folder = DriveApp.createFolder('部活動管理システム 添付ファイル（非公開）');
  props.setProperty('CLUB_ATTACHMENT_FOLDER_ID', folder.getId());
  return folder;
}

/** 添付ファイルをDriveへ保存（年度予定・桐蔭会原稿・引継資料など） */
function uploadProtectedFile(base64, fileName, mimeType, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!base64 || typeof base64 !== 'string') throw new Error('ファイルデータがありません。');

  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > CONFIG.MAX_FILE_BYTES) throw new Error('ファイルは12MB以下にしてください。');

  const isPdf = bytes.length > 4 && bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70;
  const type = String(mimeType || '');
  const isImage = /^image\//.test(type);
  if (!isPdf && !isImage) throw new Error('PDFまたは画像ファイルのみアップロードできます。');

  const safeName = String(fileName || (isPdf ? 'file.pdf' : 'image')).replace(/[\\/:*?"<>|]/g, '_');
  const finalType = isPdf ? 'application/pdf' : type;
  const blob = Utilities.newBlob(bytes, finalType, safeName);
  const file = getAttachmentFolder_().createFile(blob);
  file.setDescription('部活動管理システムのログインユーザー向け添付ファイル');

  return { ok: true, fileId: file.getId(), fileName: file.getName(), mimeType: finalType };
}

/** 添付ファイルを取得（ログインユーザー限定）。base64で返す。 */
function getProtectedFile(fileId, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ファイルを閲覧するにはログインしてください。');

  const props = PropertiesService.getScriptProperties();
  const allowedFolders = [
    props.getProperty('CLUB_ATTACHMENT_FOLDER_ID'),
    props.getProperty('CLUB_MANUAL_FOLDER_ID')
  ].filter(Boolean);

  const id = String(fileId || '');
  const file = DriveApp.getFileById(id);
  let allowed = false;
  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (allowedFolders.indexOf(parent.getId()) >= 0) { allowed = true; break; }
  }
  // 会計書類として登録済みのファイルは、ログイン済みユーザーなら閲覧可能。
  // ファイルIDがサーバー上の会計一覧に存在する場合に限定し、Drive全体は公開しない。
  if (!allowed) {
    const accountingDocs = readData_('system_accounting_docs', []);
    allowed = Array.isArray(accountingDocs) && accountingDocs.some(function (doc) {
      return doc && String(doc.driveFileId || '') === id;
    });
  }
  if (!allowed) throw new Error('このファイルを閲覧する権限がありません。');

  const blob = file.getBlob();
  return {
    ok: true,
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: file.getName(),
    mimeType: blob.getContentType()
  };
}

/** PDFをDriveへ保存（マニュアル用）。Driveファイルは公開せず、認証済みユーザーにだけ返します。 */
function uploadManualPdfToDrive(base64, fileName, token) {
  const session = requireSession_(token);
  if (!session || !session.isMaker) throw new Error('マニュアル登録は作成者のみ実行できます。');
  if (!base64 || typeof base64 !== 'string') throw new Error('PDFデータがありません。');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > CONFIG.MAX_PDF_BYTES) throw new Error('PDFは12MB以下にしてください。');
  if (bytes.length < 5 || bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70 || bytes[4] !== 45) {
    throw new Error('PDFファイルとして認識できません。');
  }
  const folderId = PropertiesService.getScriptProperties().getProperty('CLUB_MANUAL_FOLDER_ID');
  if (!folderId) throw new Error('先にsetupClubSystemを実行してください。');
  const safeName = String(fileName || 'manual.pdf').replace(/[\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(bytes, MimeType.PDF, safeName);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  file.setDescription('部活動管理システムのログインユーザー向けマニュアルPDF');
  return { ok: true, fileId: file.getId(), fileName: file.getName() };
}

function getProtectedManualPdf(fileId, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('PDFを閲覧するにはログインしてください。');
  const versions = readData_('system_manual_versions', []);
  const live = versions.some(v => String(v.driveFileId || '') === String(fileId));
  if (live) {
    const file = DriveApp.getFileById(String(fileId));
    const blob = file.getBlob();
    return { ok: true, base64: Utilities.base64Encode(blob.getBytes()), fileName: file.getName() };
  }
  return getProtectedFile(fileId, token);
}

function logoutSession(token) {
  if (token) CacheService.getScriptCache().remove(CONFIG.SESSION_PREFIX + String(token));
  return { ok: true };
}

// ---- 顧問アカウントの年次リセット（毎年4月7日） ----
function installAnnualAdvisorResetTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'annualAdvisorReset');
  if (!exists) {
    ScriptApp.newTrigger('annualAdvisorReset')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
  }
  return { ok: true, installed: !exists };
}

/** トリガー実行用：毎日1時に起動し、4/7以降の初回だけ削除する（作成者は保護） */
function annualAdvisorReset() {
  try { runTreasurerHandoverCore_(); } catch (e) { Logger.log('会計引継ぎ失敗: ' + e); }
  const result = runAdvisorResetForToday_();
  if (result && result.removed) Logger.log('顧問アカウント年次削除: ' + result.removed + '件');
  return result;
}

function runAdvisorResetForToday_() {
  if (!isAdvisorResetDue_(new Date())) return { ok: true, skipped: true, reason: 'not-due', removed: 0 };
  const props = PropertiesService.getScriptProperties();
  const year = String(jstParts_().y);
  if (props.getProperty('CLUB_ADVISOR_RESET_YEAR') === year) {
    return { ok: true, skipped: true, alreadyDone: true, removed: 0 };
  }
  const result = deleteAdvisorAccountsCore_();
  return { ok: true, skipped: false, removed: result.removed, year: year };
}

/** 画面から実行：4月7日を過ぎていて今年度未処理の場合のみ削除 */
function runScheduledAdvisorReset(token) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログインしてください。' };
  const now = new Date();
  if (!isAdvisorResetDue_(now)) return { ok: true, skipped: true, removed: 0 };
  const year = String(now.getFullYear());
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLUB_ADVISOR_RESET_YEAR') === year) {
    return { ok: true, skipped: true, alreadyDone: true, removed: 0 };
  }
  const result = deleteAdvisorAccountsCore_();
  return { ok: true, skipped: false, removed: result.removed, year: year };
}

/** 作成者が手動で顧問アカウントを一括削除 */
function deleteAdvisorAccounts(token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!session.isMaker) throw new Error('顧問アカウントの一括削除は作成者のみ実行できます。');
  return deleteAdvisorAccountsCore_();
}

function deleteAdvisorAccountsCore_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const accounts = readData_('system_accounts', []);
    const removed = accounts.filter(a => a && a.isAdvisor && !a.isMaker);
    const kept = accounts.filter(a => !(a && a.isAdvisor && !a.isMaker));
    if (removed.length) {
      removed.forEach(a => vaultDelete_(a.id));
      writeData_('system_accounts', kept);
    }
    PropertiesService.getScriptProperties()
      .setProperty('CLUB_ADVISOR_RESET_YEAR', String(jstParts_().y));
    return { ok: true, removed: removed.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 会計アカウントの年次引き継ぎ
 * 12/31までは前の会計が保持。1月になったら三役データの「会計」へID・パスワードを再発行してDMで通知。
 */
// ---- 会計（管理者ツール）アカウント ----
// 会計は「三役」で選ばれるが、権限の交代は毎年1月1日。
// それまでは前年度の会計がそのまま会計ツールを使える。
const TREASURER_ID_KEY = 'TREASURER_ID';
const TREASURER_PENDING_KEY = 'TREASURER_PENDING_ID';
const TREASURER_PENDING_YEAR_KEY = 'TREASURER_PENDING_YEAR';

function treasurerState_() {
  const p = PropertiesService.getScriptProperties();
  return {
    id: String(p.getProperty(TREASURER_ID_KEY) || ''),
    pendingId: String(p.getProperty(TREASURER_PENDING_KEY) || ''),
    pendingYear: Number(p.getProperty(TREASURER_PENDING_YEAR_KEY) || 0)
  };
}

/** 現在の会計ID。未設定なら三役の会計（vice2Id）をそのまま採用（2026年度の登録済み状態） */
function resolveTreasurerId_() {
  const p = PropertiesService.getScriptProperties();
  let id = String(p.getProperty(TREASURER_ID_KEY) || '');
  if (!id) {
    const sanyi = readData_('system_sanyi_data', {});
    id = String((sanyi && sanyi.vice2Id) || '');
    if (id) {
      p.setProperty(TREASURER_ID_KEY, id);
      p.setProperty('TREASURER_EFFECTIVE_FROM', new Date().toISOString());
    }
  }
  return id;
}

/** 三役登録時に「次の会計」を予約する（交代は次の1月1日） */
function registerTreasurerCandidate_(vice2Id) {
  const p = PropertiesService.getScriptProperties();
  const id = String(vice2Id || '');
  if (!id) return;
  const current = resolveTreasurerId_();
  if (!current) { p.setProperty(TREASURER_ID_KEY, id); return; }
  if (id === current) {
    p.deleteProperty(TREASURER_PENDING_KEY);
    p.deleteProperty(TREASURER_PENDING_YEAR_KEY);
    return;
  }
  const now = jstParts_();
  const year = (now.m === 1 && now.d === 1) ? now.y : now.y + 1;
  p.setProperty(TREASURER_PENDING_KEY, id);
  p.setProperty(TREASURER_PENDING_YEAR_KEY, String(year));
}

/** 1月1日を過ぎていれば会計権限を交代する */
function applyTreasurerHandoverIfDue_() {
  const p = PropertiesService.getScriptProperties();
  const st = treasurerState_();
  if (!st.pendingId || !st.pendingYear) return { changed: false };
  const now = jstParts_();
  if (now.y < st.pendingYear) return { changed: false, dueYear: st.pendingYear };
  const oldId = st.id;
  p.setProperty(TREASURER_ID_KEY, st.pendingId);
  p.setProperty('TREASURER_EFFECTIVE_FROM', new Date().toISOString());
  p.deleteProperty(TREASURER_PENDING_KEY);
  p.deleteProperty(TREASURER_PENDING_YEAR_KEY);

  const accounts = readData_('system_accounts', []);
  let oldName = '', newName = '';
  accounts.forEach(a => {
    if (!a || a.id == null) return;
    const id = String(a.id);
    if (id === st.pendingId) { a.role = '会計'; newName = a.name || id; }
    else if (id === oldId && a.role === '会計') { a.role = '一般'; oldName = a.name || id; }
  });
  writeData_('system_accounts', accounts);
  bumpRev_('system_accounts');
  try {
    notifyAdvisors('【会計】会計ツールの権限を ' + (newName || st.pendingId) + ' さんへ引き継ぎました。' + (oldName ? '（前年度: ' + oldName + '）' : ''), '【公式】システム');
  } catch (e) {}
  return { changed: true, from: oldId, to: st.pendingId };
}

function runTreasurerHandoverCore_() { return applyTreasurerHandoverIfDue_(); }

/** 画面から：1月で未処理なら引き継ぎを実行 */
function runTreasurerHandoverIfDue(token) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログインしてください。' };
  return runTreasurerHandoverCore_();
}

function isAdvisorResetDue_(date) {
  const p = jstParts_(date);
  return p.m > CONFIG.ADVISOR_RESET_MONTH ||
    (p.m === CONFIG.ADVISOR_RESET_MONTH && p.d >= CONFIG.ADVISOR_RESET_DAY);
}

// ---- パスワード（ハッシュ＋表示用控え） ----
function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ':' + String(password),
    Utilities.Charset.UTF_8
  );
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    const v = (bytes[i] + 256) % 256;
    hex += (v < 16 ? '0' : '') + v.toString(16);
  }
  return hex;
}

function setAccountPassword_(account, password) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  const hash = hashPassword_(password, salt);
  account.salt = salt;
  account.passHash = hash;
  delete account.pass;
  const history = Array.isArray(account.passHistory)
    ? account.passHistory.filter(h => typeof h === 'string' && /^[0-9a-f]{64}$/.test(h))
    : [];
  history.push(hash);
  account.passHistory = history.slice(-6);
  account.passwordChangedAt = new Date().toISOString();
  vaultSet_(account.id, password);
  return account;
}

function verifyAccountPassword_(account, password) {
  if (!account) return false;
  const p = String(password == null ? '' : password);
  if (account.passHash && account.salt) return hashPassword_(p, account.salt) === account.passHash;
  // 旧バージョン（平文保存）との互換
  if (typeof account.pass === 'string' && account.pass !== '') return account.pass === p;
  return false;
}

function publicAccount_(account) {
  const copy = Object.assign({}, account || {});
  delete copy.pass;
  delete copy.initialPassword;
  delete copy.passHash;
  delete copy.salt;
  delete copy.passHistory;
  return copy;
}

// 表示用パスワード控え（スクリプトプロパティのみ／シートやブラウザには保存しない）
function vaultSet_(id, password) {
  if (!id || !password) return;
  try {
    PropertiesService.getScriptProperties().setProperty('ACCOUNT_PASS_' + String(id), String(password));
  } catch (e) { Logger.log('パスワード控えの保存に失敗: ' + e); }
}
function vaultGet_(id) {
  try {
    return PropertiesService.getScriptProperties().getProperty('ACCOUNT_PASS_' + String(id)) || '';
  } catch (e) { return ''; }
}
function vaultDelete_(id) {
  try { PropertiesService.getScriptProperties().deleteProperty('ACCOUNT_PASS_' + String(id)); } catch (e) {}
}

function mergeAccountsWithSecrets_(json) {
  let incoming;
  try { incoming = json ? JSON.parse(json) : []; } catch (e) { throw new Error('アカウントデータの形式が正しくありません。'); }
  if (!Array.isArray(incoming)) throw new Error('アカウントデータの形式が正しくありません。');

  const stored = readData_('system_accounts', []);
  const byId = {};
  stored.forEach(a => { if (a && a.id != null) byId[String(a.id)] = a; });

  const merged = incoming.map(raw => {
    const acc = Object.assign({}, raw || {});
    const id = String(acc.id == null ? '' : acc.id);
    if (!id) throw new Error('IDのないアカウントは保存できません。');
    acc.id = id;

    const existing = byId[id];
    const freshPassword = String(acc.initialPassword || acc.pass || '');
    delete acc.initialPassword;
    delete acc.pass;

    if (existing && existing.passHash && existing.salt) {
      acc.passHash = existing.passHash;
      acc.salt = existing.salt;
      acc.passHistory = Array.isArray(existing.passHistory) ? existing.passHistory : [];
      acc.passwordChangedAt = existing.passwordChangedAt || '';
    } else if (existing && typeof existing.pass === 'string' && existing.pass !== '') {
      setAccountPassword_(acc, existing.pass);
    } else if (freshPassword) {
      setAccountPassword_(acc, freshPassword);
    }
    return acc;
  });

  // 削除されたアカウントの控えを掃除
  const keepIds = {};
  merged.forEach(a => { keepIds[String(a.id)] = true; });
  Object.keys(byId).forEach(id => { if (!keepIds[id]) vaultDelete_(id); });

  return JSON.stringify(merged);
}

// ---- 内部処理 ----
// ---- タイムゾーン（トリガーがUTCで動いても日本時間で判定する） ----
function jstParts_(date) {
  const d = date || new Date();
  return {
    y: Number(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy')),
    m: Number(Utilities.formatDate(d, 'Asia/Tokyo', 'M')),
    d: Number(Utilities.formatDate(d, 'Asia/Tokyo', 'd'))
  };
}

// ---- 保存リビジョン（複数端末の同時編集を検知） ----
function getRev_(key) {
  const v = PropertiesService.getScriptProperties().getProperty('REV_' + key);
  return v ? Number(v) : 0;
}
function getAllRevs_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  Object.keys(all).forEach(k => {
    if (k.indexOf('REV_') === 0) out[k.slice(4)] = Number(all[k] || 0);
  });
  return out;
}
function bumpRev_(key) {
  const n = getRev_(key) + 1;
  PropertiesService.getScriptProperties().setProperty('REV_' + key, String(n));
  return n;
}
function readRaw_(key) {
  const raw = readAllData_()[key];
  return raw ? String(raw) : '';
}

// ---- 保護パスワード（画面には値を返さず、合否のみ） ----
function secondaryPass_(key) {
  const props = PropertiesService.getScriptProperties();
  const propKey = SECONDARY_PASS_PREFIX + key;
  let value = props.getProperty(propKey);
  if (!value && Object.prototype.hasOwnProperty.call(SECONDARY_PASS_DEFAULTS, key)) {
    value = SECONDARY_PASS_DEFAULTS[key];
    props.setProperty(propKey, value);
  }
  return value || '';
}

function verifySecondaryPass(token, key, pass) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログインしてください。' };
  const expected = secondaryPass_(String(key || ''));
  if (!expected) return { ok: false, message: '未設定の認可キーです。' };
  return { ok: String(pass || '') === expected };
}

/** 作成者のみ：保護パスワードを変更（初期値から変更したい場合） */
function changeSecondaryPassword(token, key, currentPass, newPass) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログインしてください。' };
  if (!session.isMaker) return { ok: false, message: '作成者のみ変更できます。' };
  if (!verifySecondaryPass(token, key, currentPass).ok) return { ok: false, message: '現在のパスワードが違います。' };
  const next = String(newPass || '');
  if (next.length < 6) return { ok: false, message: '6文字以上にしてください。' };
  PropertiesService.getScriptProperties().setProperty(SECONDARY_PASS_PREFIX + String(key), next);
  return { ok: true };
}

function requireSetupSecret_(secret) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('SETUP_SECRET');
  const provided = String(secret == null ? '' : secret).trim();
  const ssId = props.getProperty('CLUB_SPREADSHEET_ID');
  let hasAccounts = false;
  try { hasAccounts = ssId ? readData_('system_accounts', []).length > 0 : false; } catch (e) {}

  // 既にシークレットがある場合：値が渡されたときだけ照合する。
  // 引数なし（エディタの「実行」ボタン）でも通す。setupClubSystem は冪等で破壊的変更をしないため。
  if (expected) {
    if (provided && provided !== expected) throw new Error('初期設定用シークレットが正しくありません。');
    return;
  }

  // 未設定：初回のみ自動発行（アカウント作成済みの環境では勝手に作り直さない）
  if (hasAccounts) throw new Error('初期設定用シークレットが未設定です。スクリプトプロパティ SETUP_SECRET を確認してください。');
  const generated = provided || ('setup-' + Utilities.getUuid().slice(0, 8));
  props.setProperty('SETUP_SECRET', generated);
  Logger.log('初期設定用シークレットを自動設定しました: ' + generated);
}

function dataForSession_(session) {
  try { autoRestoreAccountsIfLost_(); } catch (e) {}
  const data = readAllData_();

  // パスワード関連の情報は自動では返さない（作成者でも「表示」操作をした時だけ別APIで取得）
  if (data.system_accounts) {
    try {
      const accounts = JSON.parse(data.system_accounts);
      data.system_accounts = JSON.stringify(accounts.map(publicAccount_));
    } catch (e) {
      delete data.system_accounts;
    }
  }

  // 三役専用グループDMは、サーバー側でも三役以外（作成者を含む）に配信しない
  let protectedGroupIds = [];
  try {
    const sanyi = readData_('system_sanyi_data', {}) || {};
    const isSanyiMember = [sanyi.captainId, sanyi.vice1Id, sanyi.vice2Id]
      .some(v => String(v || '') === String(session && session.id));
    if (!isSanyiMember && data.system_group_dms) {
      try {
        const groups = JSON.parse(data.system_group_dms);
        if (Array.isArray(groups)) {
          const isSanyiGroup = function (g) {
            return g && (String(g.id || '').indexOf('sanyi') === 0
              || String(g.name || '').indexOf('三役') >= 0
              || g.isSanyi === true || g.sanyi === true);
          };
          protectedGroupIds = groups.filter(isSanyiGroup).map(function (g) { return String(g.id || ''); });
          data.system_group_dms = JSON.stringify(groups.filter(function (g) { return !isSanyiGroup(g); }));
        }
      } catch (e) {}
    }
  } catch (e) {}

  if (!session || session.isMaker) return data;

  if (data.system_messages) {
    try {
      const messages = JSON.parse(data.system_messages);
      data.system_messages = JSON.stringify(messages.filter(m =>
        protectedGroupIds.indexOf(String(m.receiverId)) < 0 &&
        (String(m.receiverId) === String(session.id) ||
        String(m.senderId) === String(session.id) ||
        String(m.receiverId) === 'ALL')
      ));
    } catch (e) {}
  }
  return data;
}

function requireSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(CONFIG.SESSION_PREFIX + String(token));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/**
 * 未初期化の環境でデータ読み書きが起きた場合だけ、冪等な初期セットアップ
 * （setupClubSystem）を自動実行する。
 * ※ setupClubSystem は何度実行しても安全（既存環境があれば何もしない）。
 *   これにより「アカウント登録の途中で 初期設定が未完了 エラーになる」問題を防ぐ。
 */
var ENSURE_ENV_ATTEMPTED_ = false;
function ensureEnvironment_() {
  if (ENSURE_ENV_ATTEMPTED_) return;
  ENSURE_ENV_ATTEMPTED_ = true;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLUB_SPREADSHEET_ID')) return;
  try {
    Logger.log('CLUB_SPREADSHEET_ID が未設定のため、初期セットアップ（setupClubSystem）を自動実行します。');
    setupClubSystem();
  } catch (e) {
    Logger.log('自動セットアップに失敗しました: ' + ((e && e.message) || e));
  }
}

function getDataSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('CLUB_SPREADSHEET_ID');
  if (!id) {
    // 初期設定が未完了の場合は、ここで自動セットアップしてから処理を続ける
    ensureEnvironment_();
    id = PropertiesService.getScriptProperties().getProperty('CLUB_SPREADSHEET_ID');
  }
  if (!id) throw new Error('初期設定が未完了です。setupClubSystemを実行してください。');
  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DATA_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'JSON Value']]);
  }
  return sheet;
}

function readAllData_() {
  const sheet = getDataSheet_();
  const last = sheet.getLastRow();
  const out = {};
  if (last < 2) return out;
  const rows = sheet.getRange(2, 1, last - 1, 2).getValues();
  const chunks = {};

  rows.forEach(row => {
    const k = String(row[0] || '');
    if (!k) return;
    const v = row[1] === null || row[1] === undefined ? '' : String(row[1]);
    const m = k.match(/^(.*)#(\d+)$/);
    if (m) {
      if (!chunks[m[1]]) chunks[m[1]] = {};
      chunks[m[1]][Number(m[2])] = v;
      return;
    }
    out[k] = v;
  });

  Object.keys(chunks).forEach(base => {
    const parts = chunks[base];
    const count = Object.keys(parts).length;
    let joined = '';
    for (let i = 0; i < count; i++) joined += (parts[i] || '');
    out[base] = joined;
  });
  return out;
}

function readData_(key, fallback) {
  const raw = readAllData_()[key];
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}

function writeData_(key, value) { writeRawData_(key, JSON.stringify(value)); }

function writeRawData_(key, value) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getDataSheet_();
    const last = sheet.getLastRow();

    if (last >= 2) {
      const keys = sheet.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || ''));
      const targets = [];
      keys.forEach((k, i) => {
        if (k === String(key) || k.indexOf(String(key) + '#') === 0) targets.push(i + 2);
      });
      targets.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
    }

    if (value === '') return;

    if (String(value).length <= CONFIG.CHUNK_SIZE) {
      sheet.appendRow([key, value]);
      return;
    }
    let index = 0;
    for (let i = 0; i < value.length; i += CONFIG.CHUNK_SIZE) {
      sheet.appendRow([key + '#' + index, value.slice(i, i + CONFIG.CHUNK_SIZE)]);
      index++;
    }
  } finally {
    lock.releaseLock();
  }
}

function randomPassword_(length) {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghijkmnopqrstuvwxyz';
  const digits = '23456789';
  const symbols = '!@#$%&*+-_?';
  const all = upper + lower + digits + symbols;
  const pick = chars => chars.charAt(Math.floor(Math.random() * chars.length));
  const out = [pick(upper), pick(digits), pick(symbols)];
  while (out.length < Math.max(8, length)) out.push(pick(all));
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = out[i]; out[i] = out[j]; out[j] = tmp;
  }
  return out.join('');
}

/** 保護パスワードの初期値を未設定分だけ投入する */
function ensureSecondaryPassDefaults_() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(SECONDARY_PASS_DEFAULTS).forEach(function (k) {
    if (!props.getProperty(SECONDARY_PASS_PREFIX + k)) {
      props.setProperty(SECONDARY_PASS_PREFIX + k, SECONDARY_PASS_DEFAULTS[k]);
    }
  });
}

/* ============================================================
 * 年度予定（画像/PDF）から長期休みを抽出する
 *  - 追加のAPIキー不要。Google ドライブのOCR（Drive API 拡張サービス）を使用
 *  - Drive API が無く Gemini キーがあれば Gemini の画像読解を使用
 * ============================================================ */
function apiExtractVacations(base64, fileName, mimeType, year, token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  if (!session.isAdvisor && !session.isMaker) throw new Error('顧問または作成者のみ実行できます。');
  const y = Number(year) || new Date().getFullYear();

  let text = '', engine = '', err = '';

  // 1) Google ドライブ OCR（Drive API 拡張サービス）
  try {
    text = ocrTextFromBlob_(base64, fileName, mimeType);
    if (text) engine = 'drive-ocr';
  } catch (e) { err = 'OCR: ' + e; }

  // 2) Gemini（キーがある場合のみ）
  if (!text) {
    try {
      const r = geminiReadPlan_(base64, mimeType);
      if (r) { text = r; engine = 'gemini'; }
    } catch (e) { err += ' / Gemini: ' + e; }
  }

  if (!text) {
    return {
      ok: false, engine: 'none', periods: [],
      message: '画像から文字を読み取れませんでした。',
      error: err
    };
  }

  const periods = parseVacationText_(text, y);
  return { ok: periods.length > 0, engine: engine, periods: periods, textSample: String(text).slice(0, 1200), error: err };
}

/** 画像/PDFをGoogleドキュメントにOCR変換してテキストを返す */
function ocrTextFromBlob_(base64, fileName, mimeType) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Drive API（拡張サービス）が未設定です。エディタ左の「サービス」→「Drive API」を追加してください。');
  }
  if (!base64) throw new Error('ファイルデータがありません。');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > CONFIG.MAX_FILE_BYTES) throw new Error('ファイルが大きすぎます（12MB以下）。');
  const blob = Utilities.newBlob(bytes, mimeType || 'image/png', fileName || 'yearly_plan');
  let docId = '';
  try {
    const res = Drive.Files.insert(
      { title: 'OCR_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
      blob,
      { ocr: true, ocrLanguage: 'ja' }
    );
    docId = res.id;
    const body = DocumentApp.openById(docId).getBody().getText();
    return String(body || '');
  } finally {
    if (docId) { try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {} }
  }
}

/** Gemini で画像を読む（GEMINI_API_KEY がある時だけ） */
function geminiReadPlan_(base64, mimeType) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key || !base64 || base64.length > 5500000) return '';
  const prompt = 'この年度予定表の画像から「夏休み」「冬休み」「春休み」（夏季休業・冬季休業・春季休業）の開始日と終了日だけを抜き出し、'
    + 'JSONのみで返してください: {"periods":[{"label":"夏休み","start":"7/10","end":"8/31"}]}。見つからない項目は省いてください。';
  const models = [];
  const configured = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || '';
  [configured, 'gemini-2.5-flash', 'gemini-2.0-flash'].forEach(function (m) { if (m && models.indexOf(m) < 0) models.push(m); });
  for (let i = 0; i < models.length; i++) {
    try {
      const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(models[i]) + ':generateContent', {
        method: 'post', contentType: 'application/json',
        headers: { 'x-goog-api-key': String(key) },
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType || 'image/png', data: base64 } }] }],
          generationConfig: { temperature: 0, maxOutputTokens: 300 }
        }),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) continue;
      const body = JSON.parse(res.getContentText());
      const parts = (((body.candidates || [])[0] || {}).content || {}).parts || [];
      const txt = (parts[0] || {}).text || '';
      if (txt) return txt;
    } catch (e) {}
  }
  return '';
}

/** テキストから長期休み期間を抽出 */
function parseVacationText_(text, year) {
  const lines = String(text || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
  const defs = [
    { label: '夏休み', re: /夏(季|期)?\s*(休み|休業|休暇)|夏季休業/ },
    { label: '冬休み', re: /冬(季|期)?\s*(休み|休業|休暇)|冬季休業/ },
    { label: '春休み', re: /春(季|期)?\s*(休み|休業|休暇)|春季休業/ }
  ];
  const out = [];
  defs.forEach(function (d) {
    for (let i = 0; i < lines.length; i++) {
      if (!d.re.test(lines[i])) continue;
      const windowText = [lines[i], lines[i + 1] || '', lines[i + 2] || ''].join(' ');
      const dates = extractMonthDays_(windowText);
      if (dates.length >= 2) {
        const s = toISODate_(d.label, dates[0], year, false);
        const e = toISODate_(d.label, dates[1], year, true);
        if (s && e) {
          out.push({ label: d.label, startDate: s, endDate: e, raw: windowText.slice(0, 100) });
          break;
        }
      }
    }
  });
  return out;
}

function extractMonthDays_(s) {
  const re = /(\d{1,2})\s*[月\/]\s*(\d{1,2})/g;
  const res = [];
  let m;
  while ((m = re.exec(String(s))) !== null) {
    const mo = Number(m[1]), da = Number(m[2]);
    if (mo >= 1 && mo <= 12 && da >= 1 && da <= 31) res.push({ m: mo, d: da });
    if (res.length >= 4) break;
  }
  return res;
}

function toISODate_(label, md, year, isEnd) {
  const y = Number(year);
  let yy = y;
  if (label === '夏休み') yy = (md.m >= 4) ? y : y + 1;
  else if (label === '冬休み') yy = (md.m >= 9) ? y : y + 1;
  else if (label === '春休み') yy = (md.m >= 4) ? y + 1 : y + 1;
  const mm = String(md.m).padStart(2, '0');
  const dd = String(md.d).padStart(2, '0');
  const iso = yy + '-' + mm + '-' + dd;
  return isNaN(new Date(iso + 'T00:00:00').getTime()) ? '' : iso;
}

/* ============================================================
 * アカウント消失対策：保存のたびに5世代バックアップを残す
 * ============================================================ */
function saveAccountsBackup_(accounts) {
  try {
    if (!Array.isArray(accounts) || !accounts.length) return;
    const props = PropertiesService.getScriptProperties();
    const idx = (Number(props.getProperty('ACCOUNTS_BACKUP_IDX') || -1) + 1) % 5;
    props.setProperty('ACCOUNTS_BACKUP_IDX', String(idx));
    const at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    writeData_('accounts_backup_' + idx, { at: at, count: accounts.length, accounts: accounts });
    props.setProperty('ACCOUNTS_BACKUP_TS_' + idx, at);
  } catch (e) { Logger.log('バックアップ失敗: ' + e); }
}

/** バックアップ一覧（作成者のみ） */
function apiListAccountBackups(token) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  if (!session.isMaker) throw new Error('作成者のみ実行できます。');
  const items = [];
  for (let i = 0; i < 5; i++) {
    const b = readData_('accounts_backup_' + i, null);
    if (b && Array.isArray(b.accounts) && b.accounts.length) {
      items.push({ slot: i, at: b.at || '', count: b.count || b.accounts.length });
    }
  }
  items.sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); });
  return { ok: true, items: items };
}

/** バックアップから復元（作成者のみ） */
function apiRestoreAccounts(token, slot) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  if (!session.isMaker) throw new Error('作成者のみ実行できます。');
  const b = readData_('accounts_backup_' + Number(slot), null);
  if (!b || !Array.isArray(b.accounts) || !b.accounts.length) throw new Error('そのバックアップは空です。');
  // 復元前の現状も退避
  const current = readData_('system_accounts', []);
  if (Array.isArray(current) && current.length) saveAccountsBackup_(current);
  writeData_('system_accounts', b.accounts);
  bumpRev_('system_accounts');
  return { ok: true, count: b.accounts.length, at: b.at || '' };
}

/** 名簿が空になっていたら、直近のバックアップから自動復元する */
function autoRestoreAccountsIfLost_() {
  const cur = readData_('system_accounts', []);
  if (Array.isArray(cur) && cur.length) return;
  // 意図的なリセット直後は、古いバックアップから勝手に復元しない
  const reset = readData_('system_reset', null);
  if (reset && reset.at) { Logger.log('リセット後のため自動復元をスキップ（' + reset.at + '）'); return; }
  let best = null;
  for (let i = 0; i < 5; i++) {
    const b = readData_('accounts_backup_' + i, null);
    if (b && Array.isArray(b.accounts) && b.accounts.length) {
      if (!best || String(b.at) > String(best.at)) best = b;
    }
  }
  if (best) {
    writeData_('system_accounts', best.accounts);
    bumpRev_('system_accounts');
    Logger.log('名簿が空のため、バックアップ（' + best.at + '）から自動復元しました。');
  }
}

/** アカウント削除（明示操作のみ）。作成者＝全員／顧問＝部員のみ削除可 */
function apiDeleteAccount(token, targetId) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  const id = String(targetId || '').trim();
  if (!id) throw new Error('対象IDが不正です。');
  const accounts = readData_('system_accounts', []);
  const target = accounts.find(a => String(a.id) === id);
  if (!target) throw new Error('対象アカウントが見つかりません。');
  if (target.isMaker) throw new Error('作成者アカウントは削除できません。');
  if (!session.isMaker && !session.isAdvisor) throw new Error('削除する権限がありません。');
  if (!session.isMaker && target.isAdvisor) throw new Error('顧問アカウントは作成者のみ削除できます。');
  const next = accounts.filter(a => String(a.id) !== id);
  writeData_('system_accounts', next);
  saveAccountsBackup_(next);
  bumpRev_('system_accounts');
  try { vaultDelete_(id); } catch (e) {}
  try { appendAudit_(session.id, 'アカウントを削除', id, String(target.name || '')); } catch (e) {}
  return { ok: true, deleted: id, remaining: next.length };
}


/* ============================================================
 * 監査ログ / 退部フラグ / 名簿CSV一括登録
 * ============================================================ */
const AUDIT_MAX = 800;
const ACCOUNT_STATUS_LABEL = { active: '在籍', retired: '退部', graduated: '卒業', left: '停止' };

function auditActorName_(id) {
  try {
    const a = readData_('system_accounts', []).find(x => String(x.id) === String(id));
    return a ? String(a.name || '') : '';
  } catch (e) { return ''; }
}

function appendAudit_(actorId, action, target, detail) {
  try {
    const list = readData_('system_audit_log', []);
    const arr = Array.isArray(list) ? list : [];
    arr.push({
      at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      actorId: String(actorId || ''),
      actorName: auditActorName_(actorId),
      action: String(action || ''),
      target: String(target || ''),
      detail: String(detail || '').slice(0, 400)
    });
    while (arr.length > AUDIT_MAX) arr.shift();
    writeData_('system_audit_log', arr);
    return true;
  } catch (e) { return false; }
}

function apiListAuditLog(token, limit) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  if (!session.isMaker && !session.isAdvisor) throw new Error('監査ログを閲覧する権限がありません。');
  const n = Math.min(Math.max(Number(limit) || 200, 1), AUDIT_MAX);
  const list = readData_('system_audit_log', []);
  const arr = Array.isArray(list) ? list : [];
  return { ok: true, total: arr.length, entries: arr.slice(-n).reverse() };
}

function apiClearAuditLog(token) {
  const session = requireSession_(token);
  if (!session || !session.isMaker) throw new Error('作成者のみ実行できます。');
  const before = (readData_('system_audit_log', []) || []).length;
  writeData_('system_audit_log', []);
  appendAudit_(session.id, '監査ログを全削除', 'system_audit_log', '削除件数:' + before);
  return { ok: true, cleared: before };
}

function apiSetAccountStatus(token, targetId, status) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  status = String(status || '');
  if (!ACCOUNT_STATUS_LABEL[status]) throw new Error('状態の指定が不正です。');
  const id = String(targetId || '').trim();
  const accounts = readData_('system_accounts', []);
  const target = accounts.find(a => String(a.id) === id);
  if (!target) throw new Error('対象アカウントが見つかりません。');
  if (target.isMaker) throw new Error('作成者アカウントの状態は変更できません。');
  if (!session.isMaker && !session.isAdvisor) throw new Error('状態変更の権限がありません。');
  if (!session.isMaker && target.isAdvisor) throw new Error('顧問の状態は作成者のみ変更できます。');
  const before = String(target.status || 'active');
  target.status = status;
  target.statusChangedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  target.statusChangedBy = String(session.id);
  writeData_('system_accounts', accounts);
  saveAccountsBackup_(accounts);
  bumpRev_('system_accounts');
  appendAudit_(session.id, 'アカウント状態変更', id, String(target.name || '') + '：' + (ACCOUNT_STATUS_LABEL[before] || before) + ' → ' + ACCOUNT_STATUS_LABEL[status]);
  return { ok: true, id: id, status: status, statusLabel: ACCOUNT_STATUS_LABEL[status] };
}

function kiseiOfEntryYear_(y) { return Number(y) - 1975; }

function nextMemberSeq_(accounts, entryYear) {
  const k = kiseiOfEntryYear_(entryYear);
  const prefix = String(Number(entryYear) % 100).padStart(2, '0') + String(k).padStart(2, '0');
  let max = 0;
  (accounts || []).forEach(a => {
    const id = String((a && a.id) || '');
    if (id.length === 6 && id.indexOf(prefix) === 0) {
      const n = Number(id.slice(4));
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max + 1;
}

function makeInitialPassword_() {
  const up = 'ABCDEFGHJKLMNPQRSTUVWXYZ', low = 'abcdefghijkmnpqrstuvwxyz', dg = '23456789';
  const pick = s => s.charAt(Math.floor(Math.random() * s.length));
  return 'Kd' + pick(dg) + pick(dg) + pick(up) + pick(low) + pick(up) + '!' + pick(dg) + pick(dg);
}

/** CSV（名前,学年,入学年度,区分）で名簿を一括登録し、初期パスワードを1回だけ返す */
function apiBulkRegisterAccounts(token, jsonRows) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログインしてください。');
  if (!session.isMaker && !session.isAdvisor) throw new Error('名簿の一括登録は顧問・作成者のみ実行できます。');
  let rows;
  try { rows = JSON.parse(jsonRows || '[]'); } catch (e) { throw new Error('CSVデータの形式が正しくありません。'); }
  if (!Array.isArray(rows) || !rows.length) throw new Error('登録する行がありません。');
  if (rows.length > 200) throw new Error('一度に登録できるのは200件までです。');

  const accounts = readData_('system_accounts', []);
  const created = [], skipped = [];
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const d = new Date();
  const fy = (d.getMonth() + 1) >= 4 ? d.getFullYear() : d.getFullYear() - 1;

  rows.forEach((r, i) => {
    const name = String((r && r.name) || '').replace(/[\u3000\s]+/g, ' ').trim();
    const line = i + 1;
    if (!name) { skipped.push({ line: line, reason: '名前が空です' }); return; }
    const norm = name.replace(/\s+/g, '');
    if (accounts.some(a => String((a && a.name) || '').replace(/\s+/g, '') === norm)) {
      skipped.push({ line: line, name: name, reason: '同じ名前のアカウントが既にあります' });
      return;
    }
    const kindRaw = String((r && r.kind) || '').trim();
    const isAdvisor = kindRaw === '顧問' || kindRaw.toLowerCase() === 'advisor';
    const pass = makeInitialPassword_();
    const acc = { name: name, createdAt: nowStr, status: 'active' };

    if (isAdvisor) {
      const used = accounts.filter(a => a.isAdvisor).map(a => Number(a.id)).filter(n => !isNaN(n));
      let n = 1;
      while (used.indexOf(n) >= 0 && n <= 2) n++;
      if (n > 2) { skipped.push({ line: line, name: name, reason: '顧問アカウントは2名までです（000001／000002）' }); return; }
      acc.id = String(n).padStart(6, '0');
      acc.isAdvisor = true; acc.isMaker = false; acc.role = '顧問'; acc.grade = '顧問';
      acc.joinedAt = nowStr;
    } else {
      const g = Number(String((r && r.grade) || '').replace(/[^0-9]/g, ''));
      if (!(g >= 1 && g <= 3)) { skipped.push({ line: line, name: name, reason: '学年は1〜3で入力してください' }); return; }
      const entryYear = Number(String((r && r.entryYear) || '').replace(/[^0-9]/g, '')) || (fy - (g - 1));
      const seq = nextMemberSeq_(accounts, entryYear);
      if (seq > 99) { skipped.push({ line: line, name: name, reason: '同一年度の登録上限（99名）に達しています' }); return; }
      acc.id = String(entryYear % 100).padStart(2, '0') + String(kiseiOfEntryYear_(entryYear)).padStart(2, '0') + String(seq).padStart(2, '0');
      acc.isAdvisor = false; acc.isMaker = false; acc.role = '一般'; acc.grade = String(g);
      acc.entryYear = entryYear; acc.kisei = kiseiOfEntryYear_(entryYear); acc.joinedAt = nowStr;
    }
    setAccountPassword_(acc, pass);
    accounts.push(acc);
    created.push({ id: acc.id, name: name, kind: isAdvisor ? '顧問' : '部員', grade: acc.grade || '', entryYear: acc.entryYear || '', password: pass });
  });

  if (created.length) {
    writeData_('system_accounts', accounts);
    saveAccountsBackup_(accounts);
    bumpRev_('system_accounts');
    appendAudit_(session.id, '名簿を一括登録', 'system_accounts', created.length + '件登録 / ' + skipped.length + '件スキップ');
  }
  return { ok: true, created: created, skipped: skipped, total: accounts.length };
}

/* ===== 2026年度限定：桐蔭会原稿・三役引継ぎ資料のリマインド停止 ===== */
const DOC_REMINDER_EXEMPT_FISCAL_YEARS = [2026];
function currentFiscalYearForNotice_() {
  const d = new Date();
  return (d.getMonth() + 1) >= 4 ? d.getFullYear() : d.getFullYear() - 1;
}
function isDocReminderSuppressed_(text) {
  if (DOC_REMINDER_EXEMPT_FISCAL_YEARS.indexOf(currentFiscalYearForNotice_()) < 0) return false;
  return /桐蔭会|原稿|引継ぎ資料|引き継ぎ資料|雑誌/.test(String(text || ''));
}


/** 旧主将（原稿提出のみ）の指定・解除。作成者のみ */
function apiSetFormerCaptain(token, targetId, on) {
  const session = requireSession_(token);
  if (!session || !session.isMaker) throw new Error('作成者のみ実行できます。');
  const id = String(targetId || '').trim();
  const accounts = readData_('system_accounts', []);
  const target = accounts.find(a => String(a.id) === id);
  if (!target) throw new Error('対象アカウントが見つかりません。');
  if (target.isMaker) throw new Error('作成者アカウントは変更できません。');
  const sanyi = readData_('system_sanyi_data', {}) || {};
  if (String(sanyi.captainId || '') === id) throw new Error('現在の主将は旧主将にできません。');
  const enable = (on === true || String(on) === 'true');
  let former = Array.isArray(sanyi.formerCaptainIds) ? sanyi.formerCaptainIds.map(String) : [];
  if (enable) {
    target.toinkaiOnly = true;
    target.role = '前主将';
    if (former.indexOf(id) < 0) former.push(id);
  } else {
    target.toinkaiOnly = false;
    if (String(target.role || '') === '前主将') target.role = '一般';
    former = former.filter(function (x) { return x !== id; });
  }
  sanyi.formerCaptainIds = former.slice(-5);
  writeData_('system_accounts', accounts);
  writeData_('system_sanyi_data', sanyi);
  saveAccountsBackup_(accounts);
  bumpRev_('system_accounts');
  bumpRev_('system_sanyi_data');
  appendAudit_(session.id, enable ? '旧主将（原稿のみ）に指定' : '旧主将の指定を解除', id, String(target.name || ''));
  return { ok: true, id: id, toinkaiOnly: !!target.toinkaiOnly, role: target.role };
}


/* ============================================================
 * 管理者ツール入室用 専用ID／パスワード
 * Script Properties の ADMIN_ID_<TOOL> / ADMIN_PASS_<TOOL> で上書き可
 * ============================================================ */
const ADMIN_TOOL_CREDENTIALS = {
  accounting: { id: 'Contabilità/director@13456', pass: 'bilancio@12456' },
  captain:    { id: 'director/capitano@456',      pass: 'pass.24680' },
  advisor:    { id: 'director/consulente@123',    pass: 'pass.13579' },
  maker:      { id: 'Maker/Creatore@789',         pass: 'pass.124578' },
  vice:       { id: 'Lavoro/director@43021',      pass: 'Pass@7431' }
};

function adminToolCredential_(tool) {
  const t = String(tool || '');
  const base = ADMIN_TOOL_CREDENTIALS[t];
  if (!base) return null;
  const props = PropertiesService.getScriptProperties();
  return {
    id: String(props.getProperty('ADMIN_ID_' + t.toUpperCase()) || base.id),
    pass: String(props.getProperty('ADMIN_PASS_' + t.toUpperCase()) || base.pass)
  };
}

function verifyAdminToolCredentials(token, tool, id, password) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログイン期限が切れました。再ログインしてください。' };
  const key = String(tool || '');
  const cred = adminToolCredential_(key);
  if (!cred) return { ok: false, message: '対象の管理者ツールが見つかりません。' };
  const cache = CacheService.getScriptCache();
  const failKey = 'adm_fail_' + key + '_' + session.id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) return { ok: false, message: '認証の失敗が続いています。5分ほど待ってから再度お試しください。' };
  const norm = function (v) { return String(v == null ? '' : v).trim().normalize('NFC'); };
  const okId = norm(id) === norm(cred.id);
  const okPass = String(password == null ? '' : password) === cred.pass;
  if (!okId || !okPass) {
    cache.put(failKey, String(fails + 1), 300);
    appendAudit_(session.id, '管理者ツール認証失敗', key, '');
    return { ok: false, message: '専用IDまたはパスワードが正しくありません。' };
  }
  cache.remove(failKey);
  try { cache.put('ADMTOOL_' + key + '_' + String(session.id), '1', 1800); } catch (e) {}
  appendAudit_(session.id, '管理者ツール認証', key, '');
  return { ok: true, tool: key };
}

/** 管理者ツールの専用ID認証が有効か（30分） */
function isAdminToolUnlocked_(tool, userId) {
  try {
    return CacheService.getScriptCache().get('ADMTOOL_' + String(tool) + '_' + String(userId)) === '1';
  } catch (e) { return false; }
}


/* ============================================================
 * 名簿の全削除（再構築用）／復元
 * GASエディタから手動実行してください。
 * ============================================================ */
function resetAllAccountsForRebuild(keepMakerId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const accounts = readData_('system_accounts', []);
    if (!Array.isArray(accounts) || !accounts.length) {
      return { ok: true, deleted: 0, kept: 0, message: 'アカウントは既に空です。' };
    }
    // 復元用に固定キーへ退避（＋通常の5世代バックアップにも保存）
    writeData_('system_accounts_pre_reset', {
      at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      count: accounts.length,
      accounts: accounts
    });
    saveAccountsBackup_(accounts);

    const resetAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    writeData_('system_reset', { at: resetAt, by: 'rebuild' });
    const keep = String(keepMakerId || '').trim();
    const next = keep ? accounts.filter(function (a) { return String(a.id) === keep; }) : [];
    accounts.forEach(function (a) {
      if (keep && String(a.id) === keep) return;
      try { vaultDelete_(String(a.id)); } catch (e) {}
    });
    writeData_('system_accounts', next);
    bumpRev_('system_accounts');

    try { writeData_('system_sanyi_data', {}); bumpRev_('system_sanyi_data'); } catch (e) {}
    try { writeData_('system_dm_warnings', []); } catch (e) {}
    appendAudit_('system', '全アカウント削除（再構築）', 'system_accounts',
      '削除' + (accounts.length - next.length) + '件 / 残' + next.length + '件');
    return { ok: true, deleted: accounts.length - next.length, kept: next.length, backupKey: 'system_accounts_pre_reset' };
  } finally {
    lock.releaseLock();
  }
}

/** resetAllAccountsForRebuild の直前に戻す */
function restoreAccountsAfterReset() {
  const bak = readData_('system_accounts_pre_reset', null);
  if (!bak || !Array.isArray(bak.accounts) || !bak.accounts.length) {
    throw new Error('復元できるバックアップがありません（system_accounts_pre_reset）。');
  }
  writeData_('system_accounts', bak.accounts);
  saveAccountsBackup_(bak.accounts);
  bumpRev_('system_accounts');
  appendAudit_('system', '全削除を復元', 'system_accounts', bak.count + '件（' + bak.at + '）');
  return { ok: true, restored: bak.accounts.length, at: bak.at };
}


/**
 * 【再構築用・一括】全アカウントを削除し、新しい作成者を1件だけ作る。
 * GASエディタから引数なしで実行 → 実行ログに新しい作成者ID／パスワードが出ます。
 * 事前に system_accounts_pre_reset へ自動退避します（restoreAccountsAfterReset で復元可）。
 */
function resetAllAccountsAndCreateCreatorForRebuild() {
  const reset = resetAllAccountsForRebuild();   // keepMakerId なし＝全削除
  const created = createInitialCreator();        // アカウント0件のときだけ成功
  Logger.log('=== 再構築完了 ===');
  Logger.log('削除件数: ' + reset.deleted);
  Logger.log('新しい作成者ID: ' + created.id);
  Logger.log('新しい作成者パスワード: ' + created.password);
  return {
    ok: true,
    deleted: reset.deleted,
    creatorId: created.id,
    creatorPassword: created.password,
    note: 'パスワードはこの実行ログにのみ表示されます。必ず控えてください。'
  };
}


/* ============================================================
 * ログインできないときの復旧用（GASエディタから実行・ログイン不要）
 * ============================================================ */

/** 登録済みアカウントのID・名前・権限を一覧表示（パスワードは出ません） */
function listAllAccountsForRecovery() {
  const accounts = readData_('system_accounts', []);
  const rows = (accounts || []).map(function (a) {
    return {
      id: String(a.id || ''),
      name: String(a.name || ''),
      isMaker: !!a.isMaker,
      isAdvisor: !!a.isAdvisor,
      role: String(a.role || ''),
      grade: String(a.grade || ''),
      status: String(a.status || 'active')
    };
  });
  Logger.log('登録アカウント数: ' + rows.length);
  Logger.log(JSON.stringify(rows, null, 2));
  return rows;
}

/** 指定IDのパスワードを強制再発行し、ログインできる状態に戻す */
function resetPasswordForAccountById(targetId) {
  const id = String(targetId || '').trim();
  if (!id) throw new Error('引数にIDを指定してください（例: resetPasswordForAccountById("244901")）');
  const accounts = readData_('system_accounts', []);
  const target = (accounts || []).find(function (a) { return String(a.id) === id; });
  if (!target) throw new Error('ID「' + id + '」が見つかりません。先に listAllAccountsForRecovery でIDを確認してください。');
  const pass = randomPassword_(12);
  setAccountPassword_(target, pass);
  target.status = 'active';
  target.toinkaiOnly = false;
  writeData_('system_accounts', accounts);
  saveAccountsBackup_(accounts);
  bumpRev_('system_accounts');
  appendAudit_('system', 'パスワード強制再発行（エディタ）', id, String(target.name || ''));
  Logger.log('=== 再発行完了 ===');
  Logger.log('ID: ' + id + '（' + String(target.name || '') + '）');
  Logger.log('新しいパスワード: ' + pass);
  return { id: id, name: String(target.name || ''), password: pass };
}

/** 指定IDを作成者（全権限）に昇格させる */
function makeAccountMakerById(targetId) {
  const id = String(targetId || '').trim();
  const accounts = readData_('system_accounts', []);
  const target = (accounts || []).find(function (a) { return String(a.id) === id; });
  if (!target) throw new Error('ID「' + id + '」が見つかりません。');
  target.isMaker = true;
  target.role = '作成者';
  target.status = 'active';
  target.toinkaiOnly = false;
  writeData_('system_accounts', accounts);
  saveAccountsBackup_(accounts);
  bumpRev_('system_accounts');
  appendAudit_('system', '作成者へ昇格（エディタ）', id, String(target.name || ''));
  Logger.log('ID「' + id + '」を作成者にしました。');
  return { ok: true, id: id, name: String(target.name || '') };
}


/**
 * 【完全消去】名簿と全バックアップを空にする（GASエディタから実行・ログイン不要）
 * 復元されないよう system_reset 印も同時に記録します。
 * 実行後は誰もログインできません。createInitialCreator か
 * resetAllAccountsAndCreateCreatorForRebuild で作り直してください。
 */
function wipeAllAccountsHard() {
  const accounts = readData_('system_accounts', []);
  const at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  if (Array.isArray(accounts) && accounts.length) {
    writeData_('system_accounts_pre_reset', { at: at, count: accounts.length, accounts: accounts });
    accounts.forEach(function (a) { try { vaultDelete_(String(a.id)); } catch (e) {} });
  }
  writeData_('system_accounts', []);
  // 5世代バックアップも空にする（自動復元の元を断つ）
  for (let i = 0; i < 5; i++) { try { writeData_('accounts_backup_' + i, { at: '', count: 0, accounts: [] }); } catch (e) {} }
  // 意図的な初期化の印
  writeData_('system_reset', { at: at, by: 'hard-wipe' });
  try { writeData_('system_sanyi_data', {}); } catch (e) {}
  try { writeData_('system_dm_warnings', []); } catch (e) {}
  bumpRev_('system_accounts');
  bumpRev_('system_sanyi_data');
  Logger.log('名簿を完全に空にしました（' + at + '）。アカウント数: 0');
  return { ok: true, deleted: Array.isArray(accounts) ? accounts.length : 0, at: at };
}


/**
 * 【旧名簿の復活を防ぐ】いま接続中のスプレッドシートから名簿を完全に消す。
 * 新プロジェクトで旧CLUB_SPREADSHEET_IDを設定してしまった場合でも、
 * これを1回実行すれば旧名簿は消えます（バックアップと復元印も同時に処理）。
 */
function purgeAccountsInCurrentSpreadsheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('CLUB_SPREADSHEET_ID');
  if (!ssId) throw new Error('CLUB_SPREADSHEET_ID が未設定です（先に setupClubSystem を実行してください）。');
  const result = wipeAllAccountsHard();
  Logger.log('接続中のスプレッドシート: https://docs.google.com/spreadsheets/d/' + ssId);
  Logger.log('削除件数: ' + result.deleted + ' / 実行日時: ' + result.at);
  return result;
}

/**
 * 【完全に別環境にする】旧スプレッドシート・フォルダとの接続を切る。
 * 実行後に setupClubSystem を実行すると、新しい空のシート・フォルダが作られます。
 * ※旧シート自体は削除しません（Driveに残ります）。
 */
function detachOldSpreadsheetForFreshStart() {
  const props = PropertiesService.getScriptProperties();
  const before = {
    spreadsheet: props.getProperty('CLUB_SPREADSHEET_ID') || '',
    manualFolder: props.getProperty('CLUB_MANUAL_FOLDER_ID') || '',
    attachmentFolder: props.getProperty('CLUB_ATTACH_FOLDER_ID') || ''
  };
  props.deleteProperty('CLUB_SPREADSHEET_ID');
  props.deleteProperty('CLUB_MANUAL_FOLDER_ID');
  props.deleteProperty('CLUB_ATTACH_FOLDER_ID');
  Logger.log('旧環境との接続を解除しました。');
  Logger.log(JSON.stringify(before, null, 2));
  Logger.log('次に setupClubSystem を実行すると、新しい空の環境が作られます。');
  return { ok: true, detached: before };
}


/** 初回（未ログイン）の作成者登録用パスコード照合。ログイン不要。 */
function verifyMakerBootstrapPasscode(pass) {
  try {
    const accounts = readData_('system_accounts', []);
    if (Array.isArray(accounts) && accounts.length) {
      return { ok: false, message: '既にアカウントが登録されています。' };
    }
    const cache = CacheService.getScriptCache();
    const failKey = 'boot_maker_pass_fail';
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= 8) return { ok: false, message: '試行回数が上限に達しました。しばらく待ってから再度お試しください。' };
    const props = PropertiesService.getScriptProperties();
    const expected = String(props.getProperty('SEC_PASS_maker_passcode') || SECONDARY_PASS_DEFAULTS.maker_passcode || '');
    const given = String(pass == null ? '' : pass).trim();
    if (!expected || given !== expected) {
      cache.put(failKey, String(fails + 1), 300);
      return { ok: false, message: '作成者アカウント作成用パスコードが正しくありません。' };
    }
    cache.remove(failKey);
    return { ok: true };
  } catch (e) {
    return { ok: false, message: '照合できませんでした: ' + ((e && e.message) || e) };
  }
}


/** 顧問へ公式メッセージを送る（三役引継ぎ・会計引継ぎなどで使用） */
function notifyAdvisors(body, title) {
  try {
    const accounts = readData_('system_accounts', []);
    const advisors = (accounts || []).filter(function (a) {
      return a && a.isAdvisor && String(a.status || 'active') === 'active';
    });
    if (!advisors.length) return { ok: false, sent: 0 };
    const msgs = readData_('system_messages', []);
    const arr = Array.isArray(msgs) ? msgs : [];
    const at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    advisors.forEach(function (a) {
      arr.unshift({
        id: 'msg_' + Date.now() + '_official',
        senderId: 'official',
        senderName: String(title || '【公式】築附中学剣道部'),
        receiverId: String(a.id),
        content: String(body || ''),
        date: at,
        isOfficial: true
      });
    });
    writeData_('system_messages', arr);
    bumpRev_('system_messages');
    return { ok: true, sent: advisors.length };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}


/* ============================================================
 * 運営（作成者）から個人への個別連絡
 *  - 画面ではなくサーバー側で「作成者セッション」を検証する
 *  - 送信は公式（運営）名義のメッセージとして保存する
 * ============================================================ */
function pushClubMessage_(receiverId, content, senderId, senderName, extra) {
  const raw = readData_('system_messages', []);
  const list = Array.isArray(raw) ? raw.slice() : [];
  list.unshift(Object.assign({
    id: 'msg_' + Utilities.getUuid(),
    senderId: String(senderId || 'official'),
    senderName: String(senderName || '🏫 運営'),
    receiverId: String(receiverId || ''),
    content: String(content || ''),
    date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd H:mm:ss'),
    isOfficial: true
  }, extra || {}));
  writeData_('system_messages', list.slice(0, 600));
  bumpRev_('system_messages');
  return true;
}

/** 運営（＝作成者）から特定の部員・顧問へ個別連絡 */
function apiCreatorSendToMember(token, targetId, content) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログイン期限が切れました。再ログインしてください。' };
  if (!session.isMaker) return { ok: false, message: 'この機能は作成者のみ利用できます。' };
  const target = String(targetId || '').trim();
  const text = String(content || '').trim();
  if (!target) return { ok: false, message: '送信先を選択してください。' };
  if (!text) return { ok: false, message: '連絡内容を入力してください。' };
  if (text.length > 4000) return { ok: false, message: '本文が長すぎます（4000文字以内）。' };
  const accounts = readData_('system_accounts', []);
  const acc = (Array.isArray(accounts) ? accounts : []).filter(function (a) { return a && String(a.id) === target; })[0];
  if (!acc) return { ok: false, message: '送信先が見つかりません。' };
  pushClubMessage_(target, text, 'official', '🏫 運営（作成者）', { fromMaker: true });
  try { appendAudit_(session.id, '運営から個別連絡', target, String(acc.name || '')); } catch (e) {}
  return { ok: true, message: String(acc.name || '') + ' さんへ運営から連絡を送信しました。' };
}

/* ============================================================
 * 月次出席率レポート（PDF）
 *  - 3年生は夏休み（7/10）以降を集計から除外
 *  - 事前連絡のある欠席は欠席扱いにせず分母からも除外
 * ============================================================ */
const SUMMER_FREE_FROM_MONTH = 7;
const SUMMER_FREE_FROM_DAY = 10;

function isAfterSummerStartServer_(dateStr) {
  const s = String(dateStr || '');
  const m = Number(s.slice(5, 7)), d = Number(s.slice(8, 10));
  if (!m) return false;
  return (m === SUMMER_FREE_FROM_MONTH && d >= SUMMER_FREE_FROM_DAY) || m > SUMMER_FREE_FROM_MONTH || m <= 3;
}
function gradeNumberServer_(g) {
  const m = String(g == null ? '' : g).match(/[0-9]+/);
  return m ? parseInt(m[0], 10) : 0;
}
function escHtmlServer_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function monthKeyOf_(y, m) {
  return String(y) + '-' + ('0' + Number(m)).slice(-2);
}

/** 指定月の出席率を集計する（分母＝稽古回数−事前連絡欠席） */
function computeAttendanceReport_(year, month) {
  const accounts = (readData_('system_accounts', []) || []).filter(function (a) {
    return a && !a.isAdvisor && String(a.grade || '') !== '卒業' && String(a.status || 'active') !== 'retired';
  });
  const practicesRaw = readData_('system_practices', []);
  const noticesRaw = readData_('system_attendance', []);
  const practices = (Array.isArray(practicesRaw) ? practicesRaw : []).filter(function (p) { return p && p.date; });
  const notices = Array.isArray(noticesRaw) ? noticesRaw : [];
  const prefix = monthKeyOf_(year, month);
  const monthPractices = practices.filter(function (p) { return String(p.date || '').indexOf(prefix) === 0; });

  const excusedOn = function (name, dateKey) {
    return notices.some(function (n) {
      return n && String(n.name) === String(name) && String(n.status) === '欠席'
        && String(n.dateKey || '') === String(dateKey);
    });
  };
  const calc = function (acc, list) {
    const third = gradeNumberServer_(acc.grade) === 3;
    const counted = list.filter(function (p) { return !(third && isAfterSummerStartServer_(p.date)); });
    let attended = 0, excused = 0;
    counted.forEach(function (p) {
      const names = (p.attendedMembers || []).map(String);
      if (names.indexOf(String(acc.name)) >= 0) attended++;
      else if (excusedOn(acc.name, p.date)) excused++;
    });
    const denom = counted.length - excused;
    return {
      total: counted.length, attended: attended, excused: excused, denom: denom, thirdGrade: third,
      rate: denom > 0 ? Math.min(100, Math.round((attended / denom) * 1000) / 10) : null
    };
  };

  const rows = accounts.map(function (acc) {
    const m = calc(acc, monthPractices);
    const c = calc(acc, practices);
    return {
      id: String(acc.id), name: String(acc.name || ''), grade: String(acc.grade || ''),
      class: String(acc.class || ''), role: String(acc.role || ''),
      monthRate: m.rate, monthAttended: m.attended, monthDenom: m.denom,
      monthExcused: m.excused, monthTotal: m.total, thirdGrade: m.thirdGrade,
      totalRate: c.rate, totalAttended: c.attended, totalDenom: c.denom
    };
  });
  rows.sort(function (a, b) {
    return gradeNumberServer_(a.grade) - gradeNumberServer_(b.grade)
      || String(a.class).localeCompare(String(b.class), 'ja')
      || String(a.name).localeCompare(String(b.name), 'ja');
  });
  return { year: Number(year), month: Number(month), practiceCount: monthPractices.length, rows: rows };
}

/** 出席率レポートPDFを生成して非公開フォルダへ保存 */
function buildAttendanceReportPdf_(report) {
  const y = report.year, m = report.month;
  const label = y + '年' + m + '月';
  const rateCell = function (r) {
    if (r == null) return '<span style="color:#888;">対象外</span>';
    const color = r >= 80 ? '#1b7f3a' : (r < 50 ? '#c0392b' : '#b9770e');
    return '<span style="color:' + color + '; font-weight:bold;">' + Number(r).toFixed(1) + '%</span>';
  };
  const rows = report.rows.map(function (r, i) {
    return '<tr>'
      + '<td style="text-align:center;">' + (i + 1) + '</td>'
      + '<td style="text-align:center;">' + escHtmlServer_(r.grade || '-') + '</td>'
      + '<td style="text-align:center;">' + escHtmlServer_(r.class || '-') + '</td>'
      + '<td>' + escHtmlServer_(r.name) + escHtmlServer_(r.thirdGrade ? '（3年・夏以降自由参加）' : '') + '</td>'
      + '<td style="text-align:center;">' + r.monthAttended + ' / ' + r.monthDenom + '</td>'
      + '<td style="text-align:center;">' + rateCell(r.monthRate) + '</td>'
      + '<td style="text-align:center;">' + r.totalAttended + ' / ' + r.totalDenom + '</td>'
      + '<td style="text-align:center;">' + rateCell(r.totalRate) + '</td>'
      + '</tr>';
  }).join('');

  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:"Helvetica Neue",Arial,"Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;color:#111;margin:18px;}'
    + 'h1{font-size:17px;margin:0 0 2px;}'
    + '.sub{font-size:10px;color:#555;margin-bottom:10px;}'
    + 'table{width:100%;border-collapse:collapse;font-size:10px;}'
    + 'th{background:#f0f0f0;border:1px solid #bbb;padding:5px;font-size:10px;}'
    + 'td{border:1px solid #ccc;padding:4px 5px;}'
    + '.note{font-size:9px;color:#555;margin-top:10px;line-height:1.6;}'
    + '</style></head><body>'
    + '<h1>筑附中学剣道部　出席率レポート（' + label + '）</h1>'
    + '<div class="sub">対象：' + report.rows.length + '名 ／ 当月の稽古回数：' + report.practiceCount
    + '回 ／ 作成日：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd') + '</div>'
    + '<table><thead><tr>'
    + '<th style="width:5%;">#</th><th style="width:8%;">学年</th><th style="width:10%;">クラス</th>'
    + '<th>氏名</th><th style="width:13%;">当月 出席/分母</th><th style="width:11%;">当月出席率</th>'
    + '<th style="width:13%;">累計 出席/分母</th><th style="width:11%;">累計出席率</th>'
    + '</tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="note">※ 分母＝稽古回数−事前連絡のあった欠席回数（連絡済みの欠席は欠席扱いにしていません）。<br>'
    + '※ 3年生は夏休み（7/10）以降の稽古を集計から除外しています（自由参加のため）。<br>'
    + '※ このPDFは部活動管理システムが自動生成したものです。</div>'
    + '</body></html>';

  const fileName = '出席率レポート_' + monthKeyOf_(y, m) + '.pdf';
  const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(fileName);
  const file = getAttachmentFolder_().createFile(blob);
  return { fileId: file.getId(), fileName: fileName };
}

/** レポートを生成して保存（既にあればそれを返す） */
function ensureMonthlyReport_(year, month) {
  const key = monthKeyOf_(year, month);
  const raw = readData_('system_monthly_reports', []);
  const list = Array.isArray(raw) ? raw.slice() : [];
  const exists = list.filter(function (r) { return r && String(r.monthKey) === key; })[0];
  if (exists) return exists;
  const data = computeAttendanceReport_(year, month);
  const built = buildAttendanceReportPdf_(data);
  const item = {
    monthKey: key, fileId: built.fileId, fileName: built.fileName,
    createdAt: new Date().toISOString(), practiceCount: data.practiceCount,
    memberCount: data.rows.length, sentAt: null, sentBy: ''
  };
  list.unshift(item);
  writeData_('system_monthly_reports', list.slice(0, 60));
  bumpRev_('system_monthly_reports');
  return item;
}

function viceIds_() {
  const sanyi = readData_('system_sanyi_data', {}) || {};
  return { captainId: String(sanyi.captainId || ''), vice1Id: String(sanyi.vice1Id || ''), vice2Id: String(sanyi.vice2Id || '') };
}

function isViceSession_(session) {
  if (!session) return false;
  const ids = viceIds_();
  return String(session.id) === ids.vice1Id;
}

/** 副将ツール：生成済みレポートと未送信状況 */
function apiMonthlyReportStatus(token) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログイン期限が切れました。再ログインしてください。' };
  const raw = readData_('system_monthly_reports', []);
  const list = (Array.isArray(raw) ? raw.slice() : []).sort(function (a, b) {
    return String((b && b.monthKey) || '').localeCompare(String((a && a.monthKey) || ''));
  });
  const now = jstParts_();
  const currentMonthKey = monthKeyOf_(now.y, now.m);
  const pending = list.filter(function (r) { return r && !r.sentAt; }).map(function (r) { return r.monthKey; });
  return { ok: true, reports: list, currentMonthKey: currentMonthKey, pending: pending };
}

/** 副将ツール：対象月のレポートPDFを生成（副将・作成者・顧問） */
function apiGenerateMonthlyReport(token, monthKey) {
  const session = requireSession_(token);
  if (!session) throw new Error('ログイン期限が切れました。再ログインしてください。');
  if (!(session.isMaker || session.isAdvisor || session.isCaptain || isViceSession_(session))) {
    throw new Error('副将・主将・顧問・作成者のみ実行できます。');
  }
  const mk = String(monthKey || '');
  let y, m;
  if (/^[0-9]{4}-[0-9]{2}$/.test(mk)) { y = Number(mk.slice(0, 4)); m = Number(mk.slice(5, 7)); }
  else { const p = jstParts_(); y = p.y; m = p.m; }
  const item = ensureMonthlyReport_(y, m);
  try { appendAudit_(session.id, '月次出席率レポート生成', monthKeyOf_(y, m), ''); } catch (e) {}
  return { ok: true, report: item };
}

/** 副将ツール：一斉送信済みの記録 */
function markMonthlyReportSent(token, monthKey) {
  const session = requireSession_(token);
  if (!session) return { ok: false, message: 'ログイン期限が切れました。再ログインしてください。' };
  if (!(session.isMaker || session.isAdvisor || isViceSession_(session))) {
    return { ok: false, message: '副将または作成者のみ記録できます。' };
  }
  const key = String(monthKey || '');
  const raw = readData_('system_monthly_reports', []);
  const list = Array.isArray(raw) ? raw.slice() : [];
  let hit = null;
  list.forEach(function (r) {
    if (r && String(r.monthKey) === key) { r.sentAt = new Date().toISOString(); r.sentBy = String(session.id); hit = r; }
  });
  writeData_('system_monthly_reports', list.slice(0, 60));
  bumpRev_('system_monthly_reports');
  try { appendAudit_(session.id, '月次出席率レポート送信', key, ''); } catch (e) {}
  return { ok: true, report: hit };
}

/** 副将・作成者へシステム通知を送る */
function notifyViceAndMaker_(text, label) {
  const ids = viceIds_();
  const targets = [ids.vice1Id, ids.captainId].filter(Boolean);
  const accs = readData_('system_accounts', []);
  (Array.isArray(accs) ? accs : []).forEach(function (a) {
    if (a && (a.isMaker || String(a.role || '') === '作成者') && a.id != null) targets.push(String(a.id));
  });
  const uniq = {};
  targets.forEach(function (id) {
    const k = String(id || '');
    if (!k || uniq[k]) return;
    uniq[k] = true;
    pushClubMessage_(k, text, 'official_system', label || '📣 【公式】出席率レポート', { isOfficial: true });
  });
  return Object.keys(uniq).length;
}

/** 未送信レポートのリマインド（1日1回まで） */
function remindUnsentReports_() {
  const raw = readData_('system_monthly_reports', []);
  const list = Array.isArray(raw) ? raw : [];
  const pending = list.filter(function (r) { return r && !r.sentAt; });
  if (!pending.length) return { reminded: 0 };
  const now = jstParts_();
  const todayKey = String(now.y) + '-' + ('0' + now.m).slice(-2) + '-' + ('0' + now.d).slice(-2);
  const props = PropertiesService.getScriptProperties();
  const propKey = 'REPORT_REMIND_' + todayKey;
  if (props.getProperty(propKey)) return { reminded: 0, skipped: true };
  const keys = pending.map(function (r) { return r.monthKey; }).join('、');
  notifyViceAndMaker_(
    '【月末タスク】出席率レポート（' + keys + '）がまだ全員へ送信されていません。\n'
    + '管理者ツール（副将）から「一斉送信」でPDFを添付して送信してください。',
    '📣 【公式】出席率レポート未送信'
  );
  props.setProperty(propKey, String(Date.now()));
  return { reminded: pending.length, months: keys };
}

/**
 * 毎日20時台に実行されるトリガー本体。
 * GASには「月末」トリガーが無いため「翌日が1日か」で月末を判定する。
 */
function monthlyAttendanceReportJob() {
  const out = { ok: true, monthEnd: false, created: null, reminded: 0 };
  try {
    const now = new Date();
    const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    const tp = jstParts_(tomorrow);
    if (tp.d === 1) {
      const p = jstParts_();
      out.monthEnd = true;
      const item = ensureMonthlyReport_(p.y, p.m);
      out.created = item.monthKey;
      notifyViceAndMaker_(
        '【月末】' + item.monthKey + ' の出席率レポート（PDF）を自動生成しました。\n'
        + '管理者ツール（副将）を開き、「全員への一斉送信」からPDFを添付して必ず送信してください。',
        '📣 【公式】月次出席率レポート'
      );
    }
  } catch (e) { out.error = String((e && e.message) || e); }
  try { out.reminded = (remindUnsentReports_() || {}).reminded || 0; } catch (e) {}
  return out;
}

/** トリガー設置（1回だけ実行してください） */
function installMonthlyAttendanceReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function (t) { return t.getHandlerFunction() === 'monthlyAttendanceReportJob'; });
  if (!exists) {
    ScriptApp.newTrigger('monthlyAttendanceReportJob')
      .timeBased().everyDays(1).atHour(20).create();
  }
  return { ok: true, installed: !exists, trigger: 'monthlyAttendanceReportJob' };
}

/** 手動テスト用：今すぐ当月分のレポートを生成してみる（エディタから実行） */
function testGenerateCurrentMonthReport() {
  const p = jstParts_();
  const item = ensureMonthlyReport_(p.y, p.m);
  Logger.log('生成: ' + item.monthKey + ' / ' + item.fileName + ' / ' + item.fileId);
  return item;
}
