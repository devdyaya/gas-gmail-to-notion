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

  if (threads.length === 0) return;  // 対象なしは無言で終了

  const archivedUuids = new Set();

  threads.forEach(thread => {
    let allOk = true;

    thread.getMessages().forEach(message => {
      const data = parseMail(message.getPlainBody());

      if (data && data.name) {
        let ok;
        if (data.isDelete) {
          ok = archiveFromNotion(NOTION_TOKEN, NOTION_DB_ID, data);
          if (data.uuid) archivedUuids.add(data.uuid);
        } else if (data.uuid && archivedUuids.has(data.uuid)) {
          ok = true;  // 同一実行内で削除済み → 復活させない（正常系、無言）
        } else if (isMyInterview(data)) {
          ok = upsertToNotion(NOTION_TOKEN, NOTION_DB_ID, data, true);
        } else {
          ok = true;  // 担当外（正常系、無言）
        }
        if (!ok) allOk = false;
      } else {
        allOk = false;
        Logger.log('パース失敗: ' + (data && data.name ? data.name : '(氏名取得不可)'));
      }
    });

    if (allOk) thread.addLabel(label);
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
  if (!data.uuid || !data.start) return false;  // 判定不可（無言、false）

  const startDate = new Date(data.start);
  const dayStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const dayEnd = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);

  const events = CalendarApp.getDefaultCalendar().getEvents(dayStart, dayEnd);
  const uuidRegex = new RegExp(data.uuid);
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

  if (!existingPageId && !allowCreate) return true;  // 更新対象なし（正常系、無言）

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
  if (code === 200) {
    Logger.log((existingPageId ? '更新' : '新規') + ': ' + data.name);
    return true;
  }
  Logger.log('登録失敗(' + code + '): ' + data.name + ' / ' + response.getContentText());
  return false;
}

// UUIDで照合し、該当ページがあればアーカイブ
function archiveFromNotion(token, dbId, data) {
  if (!data.uuid) return true;  // UUIDなし（無言）
  const pageId = findPageByUuid(token, dbId, data.uuid);
  if (!pageId) return true;  // 対象なし（正常系、無言）

  const response = notionFetch('patch', '/pages/' + pageId, token, { archived: true });
  const code = response.getResponseCode();
  if (code === 200) {
    Logger.log('アーカイブ: ' + data.name);
    return true;
  }
  Logger.log('アーカイブ失敗(' + code + '): ' + data.name + ' / ' + response.getContentText());
  return false;
}
