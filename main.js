const NOTION_VERSION = '2022-06-28';
const NOTION_BASE = 'https://api.notion.com/v1';

function processAssignMails() {
  const props = PropertiesService.getScriptProperties();
  const NOTION_TOKEN = props.getProperty('NOTION_TOKEN');
  const NOTION_DB_ID = props.getProperty('NOTION_DB_ID');

  const LABEL_NAME = 'Notion登録済み';
  const label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);

  const query = 'from:noreply@v1.herp.cloud '
              + 'subject:選考予定 '
              + '-label:' + LABEL_NAME;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log('対象メールなし');
    return;
  }

  // この実行内でアーカイブしたUUIDを記録（同一実行内の復活防止）
  const archivedUuids = new Set();

  threads.forEach(thread => {
    let allOk = true;

    thread.getMessages().forEach(message => {
      const data = parseMail(message.getPlainBody());

      if (data && data.name) {
        let ok;
        if (data.isDelete) {
          // 削除 → アーカイブ
          ok = archiveFromNotion(NOTION_TOKEN, NOTION_DB_ID, data);
          if (data.uuid) archivedUuids.add(data.uuid);
        } else if (data.uuid && archivedUuids.has(data.uuid)) {
          // この実行内で既に削除済みのUUID → 復活させない
          Logger.log('同一実行内で削除済みのためスキップ: ' + data.name);
          ok = true;
        } else if (isMyInterview(data)) {
          // 自分が担当（カレンダーに予定あり）→ 無ければ作る・あれば更新
          ok = upsertToNotion(NOTION_TOKEN, NOTION_DB_ID, data, true);
        } else {
          // 担当外
          Logger.log('担当外のためスキップ: ' + data.name);
          ok = true;
        }

        if (ok) {
          Logger.log('処理成功: ' + data.name);
        } else {
          allOk = false;
          Logger.log('処理失敗: ' + data.name);
        }
      } else {
        allOk = false;
        Logger.log('パース失敗（スキップ）');
      }
    });

    if (allOk) {
      thread.addLabel(label);
      Logger.log('ラベル付与 完了');
    }
  });
}

// 本文から各項目を抜き出す
function parseMail(body) {
  const pick = (label) => {
    const m = body.match(new RegExp(label + ':\\s*(.+)'));
    return m ? m[1].trim() : null;
  };

  const isDelete = /選考予定を削除しました/.test(body);
  const isAssign = /選考予定にあなたをアサインしました/.test(body);

  const rawName = pick('応募者氏名');
  const name = rawName ? rawName.replace(/　/g, ' ') : null;

  // 日時：開始時刻のみ取得（末尾の「- HH:MM」で選考予定の日時行だと判別）
  const dateLine = pick('日時');
  let start = null;
  if (dateLine) {
    const m = dateLine.match(/(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2})\s*-\s*\d{2}:\d{2}/);
    if (m) {
      const day = m[1].replace(/\//g, '-');
      start = day + 'T' + m[2] + ':00+09:00';
    }
  }

  const urlMatch = body.match(/選考予定URL:\s*(https?:\/\/\S+)/);
  const url = urlMatch ? urlMatch[1] : null;

  const uuidMatch = body.match(/candidacies\/id\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/);
  const uuid = uuidMatch ? uuidMatch[1] : null;

  return {
    name, start, url, uuid, isDelete, isAssign,
    format: pick('選考予定名'),
    position: pick('職種'),
    route: pick('応募経路'),
  };
}

// 面談日のカレンダー予定に該当UUIDを含むものがあるか（＝自分が担当か）
function isMyInterview(data) {
  if (!data.uuid || !data.start) {
    Logger.log('UUIDまたは日時なし → 担当判定不可: ' + data.name);
    return false;
  }

  const startDate = new Date(data.start);
  const dayStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const dayEnd = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);

  const events = CalendarApp.getDefaultCalendar().getEvents(dayStart, dayEnd);
  const uuidRegex = new RegExp(data.uuid);
  Logger.log(dayStart + 'の面談予定: ' + data.name);
  return events.some(event => uuidRegex.test(event.getDescription() || ''));
}

// Notion APIリクエストの共通処理
function notionFetch(method, path, token, payload) {
  const options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
      'Notion-Version': NOTION_VERSION,
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);
  return UrlFetchApp.fetch(NOTION_BASE + path, options);
}

// UUIDでNotionを検索し、既存ページのIDを返す（なければ null）
function findPageByUuid(token, dbId, uuid) {
  const payload = { filter: { property: 'UUID', rich_text: { equals: uuid } } };
  const response = notionFetch('post', '/databases/' + dbId + '/query', token, payload);

  if (response.getResponseCode() !== 200) {
    Logger.log('検索エラー: ' + response.getContentText());
    return null;
  }
  const result = JSON.parse(response.getContentText());
  return result.results.length > 0 ? result.results[0].id : null;
}

// UUIDで照合。あれば更新。無ければ allowCreate が true のときだけ新規作成
function upsertToNotion(token, dbId, data, allowCreate) {
  const existingPageId = data.uuid ? findPageByUuid(token, dbId, data.uuid) : null;

  // 既存が無く、新規作成も許可されていない → 何もしない
  if (!existingPageId && !allowCreate) {
    Logger.log('既存なし＆新規作成しない設定のためスキップ: ' + data.name);
    return true;
  }

  const properties = {
    Name:     { title: [{ text: { content: data.name || '（氏名不明）' } }] },
    Url:      { url: data.url },
    UUID:     { rich_text: [{ text: { content: data.uuid || '' } }] },
    Format:   { select: { name: data.format || '未設定' } },
    Position: { select: { name: data.position || '未設定' } },
    Route:    { select: { name: data.route || '未設定' } },
  };
  if (data.start) {
    properties.Date = { date: { start: data.start } };
  }

  const response = existingPageId
    ? notionFetch('patch', '/pages/' + existingPageId, token, { properties })
    : notionFetch('post', '/pages', token, { parent: { database_id: dbId }, properties });

  const code = response.getResponseCode();
  Logger.log('ステータス: ' + code + '（' + (existingPageId ? '更新' : '新規') + '）');
  if (code === 200) return true;
  Logger.log('エラー: ' + response.getContentText());
  return false;
}

// UUIDで照合し、該当ページがあればアーカイブ
function archiveFromNotion(token, dbId, data) {
  if (!data.uuid) {
    Logger.log('UUIDなし → アーカイブ不可（スキップ）');
    return true;
  }
  const pageId = findPageByUuid(token, dbId, data.uuid);
  if (!pageId) {
    Logger.log('該当ページなし（既に無い or 未登録）: ' + data.name);
    return true;
  }

  const response = notionFetch('patch', '/pages/' + pageId, token, { archived: true });
  const code = response.getResponseCode();
  Logger.log('ステータス: ' + code + '（アーカイブ）: ' + data.name);
  if (code === 200) return true;
  Logger.log('エラー: ' + response.getContentText());
  return false;
}
