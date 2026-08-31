function processAssignMails() {
  const props = PropertiesService.getScriptProperties();
  const NOTION_TOKEN = props.getProperty('NOTION_TOKEN');
  const NOTION_DB_ID = props.getProperty('NOTION_DB_ID');

  const LABEL_NAME = 'Notion登録済み';

  // ラベルを取得（なければ作る）
  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(LABEL_NAME);
  }

  // 差出人＋件名 かつ「Notion登録済み」ラベルが付いていないもの
  const query = 'from:noreply@v1.herp.cloud '
              + 'subject:選考予定 '
              + '-label:' + LABEL_NAME;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log('対象メールなし');
    return;
  }

  threads.forEach(thread => {
    let allOk = true; // このスレッド内が全部成功したか

    thread.getMessages().forEach(message => {
      const body = message.getPlainBody();
      const data = parseMail(body);

      if (data && data.name) {
        // 削除メールならアーカイブ、それ以外は新規/更新
        const ok = data.isDelete
          ? archiveFromNotion(NOTION_TOKEN, NOTION_DB_ID, data)
          : upsertToNotion(NOTION_TOKEN, NOTION_DB_ID, data);

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

    // 全部成功したスレッドだけラベルを付ける
    if (allOk) {
      thread.addLabel(label);
      Logger.log('ラベル付与 完了');
    }
  });
}

// 本文から各項目を抜き出す
function parseMail(body) {
  // ヘルパー：ラベルに続く値を1行分取る
  const pick = (label) => {
    const m = body.match(new RegExp(label + ':\\s*(.+)'));
    return m ? m[1].trim() : null;
  };

  // 削除メールかどうか（本文に「削除しました」が含まれるか）
  const isDelete = /選考予定を削除しました/.test(body);

  // 氏名：全角スペースを半角に変換
  const rawName = pick('応募者氏名');
  const name = rawName ? rawName.replace(/　/g, ' ') : null;

  // 日時：「2026/09/02 10:30 - 11:30」から開始と終了を取る
  const dateLine = pick('日時');
  let start = null, end = null;
  if (dateLine) {
    const m = dateLine.match(/(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/);
    if (m) {
      const day = m[1].replace(/\//g, '-'); // Notion用に YYYY-MM-DD へ
      start = day + 'T' + m[2] + ':00+09:00';
      end   = day + 'T' + m[3] + ':00+09:00';
    }
  }

  // 選考予定URL：https から行末まで
  const urlMatch = body.match(/選考予定URL:\s*(https?:\/\/\S+)/);
  const url = urlMatch ? urlMatch[1] : null;

  // UUID：URLの candidacies/id/ の直後にある32桁（8-4-4-4-12形式）
  const uuidMatch = body.match(/candidacies\/id\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/);
  const uuid = uuidMatch ? uuidMatch[1] : null;

  return {
    name: name,
    start: start,
    end: end,
    url: url,
    uuid: uuid,
    isDelete: isDelete,
    format: pick('選考予定名'),
    position: pick('職種'),
    route: pick('応募経路'),
  };
}

// UUIDでNotionを検索し、既存ページのIDを返す（なければ null）
function findPageByUuid(token, dbId, uuid) {
  const payload = {
    filter: {
      property: 'UUID',
      rich_text: { equals: uuid }
    }
  };

  const response = UrlFetchApp.fetch(
    'https://api.notion.com/v1/databases/' + dbId + '/query',
    {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );

  if (response.getResponseCode() !== 200) {
    Logger.log('検索エラー: ' + response.getContentText());
    return null;
  }

  const result = JSON.parse(response.getContentText());
  // 見つかれば最初の1件のページIDを返す
  return result.results.length > 0 ? result.results[0].id : null;
}

// UUIDで照合し、あれば更新・なければ新規作成
function upsertToNotion(token, dbId, data) {
  // まず properties を組み立て（新規・更新で共通）
  const properties = {
    Name:   { title: [{ text: { content: data.name || '（氏名不明）' } }] },
    Url:    { url: data.url },
    UUID:   { rich_text: [{ text: { content: data.uuid || '' } }] },
    Format: { select: { name: data.format || '未設定' } },
    Position: { select: { name: data.position || '未設定' } },
    Route:    { select: { name: data.route || '未設定' } },
  };

  if (data.start) {
    properties.Date = { date: { start: data.start} };
  }

  // UUIDで既存ページを検索
  const existingPageId = data.uuid
    ? findPageByUuid(token, dbId, data.uuid)
    : null;

  let url, method;
  if (existingPageId) {
    // あった → 更新（pages.update は PATCH）
    url = 'https://api.notion.com/v1/pages/' + existingPageId;
    method = 'patch';
  } else {
    // なかった → 新規作成（pages.create は POST）
    url = 'https://api.notion.com/v1/pages';
    method = 'post';
  }

  // 更新時は parent 不要、新規時のみ parent を付ける
  const payload = existingPageId
    ? { properties: properties }
    : { parent: { database_id: dbId }, properties: properties };

  const response = UrlFetchApp.fetch(url, {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  Logger.log('ステータス: ' + code + '（' + (existingPageId ? '更新' : '新規') + '）');
  if (code === 200) {
    return true;
  } else {
    Logger.log('エラー: ' + response.getContentText());
    return false;
  }
}

// UUIDで照合し、該当ページがあればアーカイブ（NotionのDBリストから削除）
function archiveFromNotion(token, dbId, data) {
  if (!data.uuid) {
    Logger.log('UUIDなし → アーカイブ不可（スキップ）');
    return true; // UUIDが取れない時は処理対象外として成功扱い
  }

  // UUIDで既存ページを検索
  const pageId = findPageByUuid(token, dbId, data.uuid);

  if (!pageId) {
    // Notion上で見つからない場合は何もしない
    Logger.log('該当ページなし（既に無い or 未登録）: ' + data.name);
    return true;
  }

  // archived: true でアーカイブ（PATCH）
  const payload = { archived: true };

  const response = UrlFetchApp.fetch(
    'https://api.notion.com/v1/pages/' + pageId,
    {
      method: 'patch',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );

  const code = response.getResponseCode();
  Logger.log('ステータス: ' + code + '（アーカイブ）: ' + data.name);
  if (code === 200) {
    return true;
  } else {
    Logger.log('エラー: ' + response.getContentText());
    return false;
  }
}

