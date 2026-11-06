// ===== 面談枠ブロック自動化（カレンダー更新トリガー版）=====

// カレンダー更新トリガーから呼ばれるエントリポイント
function onCalendarChange() {
  // イベント駆動は短時間に複数発火しうる。さらにブロック作成自体が
  // 再びトリガーを呼ぶため、ロックで直列化して重複実行を防ぐ。
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    // 他の実行が走行中。走査は冪等なので取りこぼしにならず、スキップしてよい
    return;
  }
  try {
    blockOverbookedDays();
  } finally {
    lock.releaseLock();
  }
}

function blockOverbookedDays() {
  const props = PropertiesService.getScriptProperties();
  const THRESHOLD  = parseInt(props.getProperty('BLOCK_THRESHOLD'), 10); // 面談上限数
  const DAYS_AHEAD = parseInt(props.getProperty('BLOCK_DAYS_AHEAD'), 10); // 何日先まで監視するか
  const KEEP_DAYS  = parseInt(props.getProperty('KEEP_BLOCK_DAYS'), 10); // 何日先まではブロック後に面談が減ってもブロックを解除しないか
  const BLOCK_START = props.getProperty('BLOCK_START'); // ブロック予定開始時刻、"09:00"
  const BLOCK_END   = props.getProperty('BLOCK_END');   // ブロック予定終了時刻、"21:00"
  const BLOCK_TITLE = props.getProperty('BLOCK_TITLE');  // 予定タイトル

  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();

  // 今日から DAYS_AHEAD 日後まで、1日ずつ見ていく
  for (let offset = 0; offset <= DAYS_AHEAD; offset++) {
    const targetDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + offset);
    processDay(calendar, targetDate, offset, THRESHOLD, KEEP_DAYS, BLOCK_START, BLOCK_END, BLOCK_TITLE);
  }
}

// 1日分の処理
function processDay(calendar, date, offset, threshold, keepDays, blockStart, blockEnd, blockTitle) {
  const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const dayEnd   = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);

  const events = calendar.getEvents(dayStart, dayEnd);

  // 面談件数を数える（説明欄にHERPのUUID-URLを含む予定）
  const interviewCount = events.filter(isInterviewEvent).length;

  // その日に既にあるブロック予定（タイトル一致）
  const blockEvents = events.filter(e => e.getTitle() === blockTitle);

  const dateStr = Utilities.formatDate(date, 'JST', 'MM/dd');

  // 重複ブロックの自己修復：2件以上あれば1件に揃える
  if (blockEvents.length > 1) {
    blockEvents.slice(1).forEach(e => e.deleteEvent());
    Logger.log(dateStr + ': 重複ブロック' + (blockEvents.length - 1) + '件を削除');
  }
  const hasBlock = blockEvents.length > 0;

  if (!hasBlock) {
    // ブロック無し → 閾値以上なら追加
    if (interviewCount >= threshold) {
      createBlock(calendar, date, blockStart, blockEnd, blockTitle);
      Logger.log(dateStr + ': 面談' + interviewCount + '件 → 面談ブロック追加');
    }
  } else {
    // ブロック有り かつ 面談が閾値未満に減った
    if (interviewCount < threshold) {
      if (offset > keepDays) {
        // 猶予ゾーン → 削除して開ける
        blockEvents.forEach(e => e.deleteEvent());
        Logger.log(dateStr + ': 面談' + interviewCount + '件（調整猶予期間）→ 面談ブロック削除');
      }
      // 直近ゾーン（offset <= keepDays）→ 残す（無言）
    }
    // 面談が閾値以上 → ブロック維持（無言）
  }
}

// 面談予定かどうか（説明欄にHERPのUUID-URLを含み、終日予定でない）
function isInterviewEvent(event) {
  if (event.isAllDayEvent()) return false;
  const desc = event.getDescription() || '';
  return /candidacies\/id\/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/.test(desc);
}

// ブロック予定を作成（その日の blockStart〜blockEnd）
function createBlock(calendar, date, blockStart, blockEnd, blockTitle) {
  const [sh, sm] = blockStart.split(':').map(n => parseInt(n, 10));
  const [eh, em] = blockEnd.split(':').map(n => parseInt(n, 10));

  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), sh, sm);
  const end   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), eh, em);

  calendar.createEvent(blockTitle, start, end, {
    description: '面談件数が上限に達したため自動作成されたブロック予定です。'
  });
}
