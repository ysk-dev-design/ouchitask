// ===================================================
// おうちタスク GAS コード（全文差し替え版）
// スプレッドシートのスクリプトエディタに貼り付けてください
// 貼り付け後：「デプロイ」→「デプロイを管理」→「編集（鉛筆アイコン）」
//             →「バージョン：新しいバージョン」→「デプロイ」
// ===================================================

// ── GETリクエスト：データ読み込み ──
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  const raw = sheet.getRange('A2').getValue();
  let data = [];
  if (raw) {
    try { data = JSON.parse(raw); } catch(err) { data = []; }
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POSTリクエスト：データ保存 or LINE Webhook ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // LINE Webhookの場合（eventsキーが存在）
    if (body.events) {
      return handleLineWebhook(body);
    }

    // アプリからのデータ保存
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');

    // 今週の目標の保存（別キー送信の場合）
    if (body.weeklyGoals !== undefined) {
      sheet.getRange('C1').setValue(JSON.stringify(body.weeklyGoals));
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const incoming = body.data;
    if (!Array.isArray(incoming)) throw new Error('dataが配列ではありません');

    const currentRaw = sheet.getRange('A2').getValue();
    let current = [];
    if (currentRaw) {
      try { current = JSON.parse(currentRaw); } catch(err) { current = []; }
    }

    // updatedAt基準でタスク単位にマージ
    const merged = mergeByUpdatedAt(current, incoming);
    sheet.getRange('A2').setValue(JSON.stringify(merged));

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', data: merged }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── マージ処理：updatedAtが新しいタスクを採用 ──
// ・グループ/タスクの追加・削除・並び順 → incoming（送信元）を優先
// ・各タスクのhistory/updatedAt → updatedAtが新しいほうを採用
function mergeByUpdatedAt(current, incoming) {
  // currentのタスクをidでMapに変換
  var baseTaskMap = {};
  (current || []).forEach(function(g) {
    (g.tasks || []).forEach(function(t) {
      baseTaskMap[t.id] = t;
    });
  });

  return (incoming || []).map(function(inGroup) {
    return Object.assign({}, inGroup, {
      tasks: (inGroup.tasks || []).map(function(inTask) {
        var baseTask = baseTaskMap[inTask.id];
        if (!baseTask) return inTask; // 新規タスクはそのまま
        var baseTime = baseTask.updatedAt || '';
        var inTime   = inTask.updatedAt   || '';
        // incomingのほうが新しい（or同じ）ならincoming採用、古ければbase採用
        return inTime >= baseTime ? inTask : baseTask;
      })
    });
  });
}

// ── LINE Webhook処理（妻のユーザーID取得用） ──
function handleLineWebhook(body) {
  try {
    var events = body.events || [];
    if (events.length > 0 && events[0].source) {
      var userId = events[0].source.userId;
      if (userId) {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
        sheet.getRange('B1').setValue(userId);
      }
    }
  } catch(err) {}
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LINE通知：期限切れタスクを毎朝送信 ──
// トリガー設定：この関数を「時間ベース → 午前7〜8時」で実行
function checkAndNotifyOverdueTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  var raw = sheet.getRange('A2').getValue();
  if (!raw) return;

  var groups = [];
  try { groups = JSON.parse(raw); } catch(e) { return; }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var overdue = [];
  groups.forEach(function(g) {
    (g.tasks || []).forEach(function(t) {
      var history = t.history || [];
      var lastDone = history.length > 0
        ? history.slice().sort().slice(-1)[0]
        : null;
      var daysSince = lastDone
        ? Math.floor((today - new Date(lastDone)) / 86400000)
        : null;
      if (daysSince === null || daysSince >= t.intervalDays) {
        overdue.push({
          name: t.name,
          icon: t.icon,
          days: daysSince,
          interval: t.intervalDays
        });
      }
    });
  });

  if (overdue.length === 0) return;

  var lines = overdue.map(function(t) {
    return t.icon + ' ' + t.name + '（' + t.interval + '日ごと・'
      + (t.days === null ? '未実施' : t.days + '日経過') + '）';
  });
  var message = '🏠 おうちタスク通知\n要対応：' + overdue.length + '件\n\n' + lines.join('\n');

  // 今週の目標を取得してメッセージに追加
  try {
    var goalsRaw = sheet.getRange('C1').getValue();
    if (goalsRaw) {
      var goals = JSON.parse(goalsRaw);
      if (Array.isArray(goals) && goals.length > 0) {
        var goalLines = goals.map(function(g) { return '• ' + g.text; });
        message += '\n\n🎯 今週の目標\n' + goalLines.join('\n');
      }
    }
  } catch(e) {}

  var props = PropertiesService.getScriptProperties();
  var token  = props.getProperty('LINE_TOKEN');
  var userId  = props.getProperty('LINE_USER_ID');
  var userId2 = props.getProperty('LINE_USER_ID_2');

  if (token && userId)  sendLineMessage(token, userId,  message);
  if (token && userId2) sendLineMessage(token, userId2, message);
}

// ── LINE メッセージ送信 ──
function sendLineMessage(token, userId, message) {
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = JSON.stringify({
    to: userId,
    messages: [{ type: 'text', text: message }]
  });
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: payload,
    muteHttpExceptions: true
  });
}
