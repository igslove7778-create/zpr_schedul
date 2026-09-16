function zNotifySchedule_(schedule, title, skipTelegramId) {
  // 동일 일정/알림의 짧은 시간 중복 발송만 막는다.
  // 다른 편집 처리에서 이미 DocumentLock을 잡을 수 있으므로 여기서는 락을 다시 잡지 않는다.
  var dedupeKey = 'TG_SCHEDULE_' + String(schedule && schedule.id || '') + '_' + String(title || '');
  var cache = CacheService.getScriptCache();
  if (cache.get(dedupeKey)) return { sent: [], failed: [] };
  cache.put(dedupeKey, '1', 30);

  var sent = [];
  var failed = [];
  var recipients = {};
  var sentChats = {};
  zRecipients_(schedule).forEach(function(member) { recipients[member.name] = member; });
  Object.keys(recipients).map(function(name) { return recipients[name]; }).forEach(function(member) {
    if (!member.telegramId) return;
    var chatKey = String(member.telegramId);
    if (skipTelegramId && chatKey === String(skipTelegramId)) return;
    if (sentChats[chatKey]) return;
    sentChats[chatKey] = true;
    try {
      zSendTelegram_(member.telegramId, title + '\n' + zScheduleText_(schedule));
      sent.push(member.name);
    } catch (error) {
      failed.push(member.name);
      zLog_('오류', schedule.id, '텔레그램', '실패', error);
    }
  });
  return { sent: sent, failed: failed };
}

function zAlarmResult_(notification, hasDirectReply) {
  if (notification && notification.failed && notification.failed.length) return ZEPHYRUS.alarm.failed;
  if (notification && notification.sent && notification.sent.length) return ZEPHYRUS.alarm.sent;
  return hasDirectReply ? ZEPHYRUS.alarm.sent : '발송대상없음';
}

function zGenerateAuthCodes() {
  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  var map = zHeaders_(sheet);
  var count = 0;
  zMembers_().forEach(function(member) {
    if (member.telegramId || member.code) return;
    sheet.getRange(member.row, map[ZEPHYRUS.memberCol.code]).setValue(String(Math.floor(100000 + Math.random() * 900000)));
    count++;
  });
  zNotice_(count + '명에게 인증코드를 만들었습니다.');
}

// Normally Telegram calls the Apps Script URL directly.  When the optional
// immediate-response proxy is enabled, TELEGRAM_WEBHOOK_URL holds that public
// Worker URL instead.  The schedule data still stays in Apps Script/Sheets.
function zTelegramWebhookUrl_() {
  return String(zProp_('TELEGRAM_WEBHOOK_URL') || zRequireProp_('WEB_APP_URL')).replace(/\/$/, '');
}

function zRegisterWebhook() {
  var baseUrl = zTelegramWebhookUrl_();
  var key = zProp_('WEBHOOK_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    zProps_().setProperty('WEBHOOK_KEY', key);
  }
  zTelegramCall_('setWebhook', { url: baseUrl + '?key=' + encodeURIComponent(key), allowed_updates: ['message'] });
  zNotice_('텔레그램 웹훅을 등록했습니다.');
}

// Recovery diagnostic: checks the outbound Telegram connection without
// creating an event, changing a calendar, or relying on a webhook update.
function zTestTelegramReply() {
  var member = zMembers_().filter(function(item) {
    return item.name === '김형원' && item.telegramId;
  })[0] || null;
  if (!member) throw new Error('김형원 구성원의 텔레그램 ID를 찾지 못했습니다.');
  try {
    zSendTelegram_(member.telegramId, '[제피로스 점검] 텔레그램 발송 연결이 정상입니다.');
    zLog_('점검', '', '텔레그램', '성공', '김형원 텔레그램 직접 발송 성공');
    zNotice_('김형원 텔레그램으로 점검 문구를 보냈습니다.');
  } catch (error) {
    zLog_('점검', '', '텔레그램', '실패', '김형원 직접 발송 오류: ' + error);
    throw error;
  }
}

// One-time fallback for linking an account when Telegram has queued old
// /start updates.  This does not change any calendar or schedule data.
function zBeginTelegramManualLink() {
  zTelegramCall_('deleteWebhook', { drop_pending_updates: true });
  zGenerateAuthCodes();
  zNotice_('이제 텔레그램 봇에 시트의 6자리 인증번호를 한 번 보내고, zFinishTelegramManualLink를 실행하세요.');
}

function zFinishTelegramManualLink() {
  var updates = zTelegramCall_('getUpdates', { timeout: 0 }).result || [];
  var members = zMembers_();
  var found = null;
  updates.forEach(function(update) {
    var message = update && update.message;
    var code = message ? String(message.text || '').replace(/[^0-9]/g, '') : '';
    if (!message || !message.chat || !/^\d{6}$/.test(code)) return;
    var member = members.filter(function(item) {
      return String(item.code || '').replace(/[^0-9]/g, '') === code;
    })[0] || null;
    if (member) found = { update: update, message: message, member: member };
  });
  if (!found) throw new Error('방금 보낸 인증번호를 찾지 못했습니다. 먼저 zBeginTelegramManualLink를 실행한 뒤 번호를 한 번 보내세요.');

  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  var map = zHeaders_(sheet);
  sheet.getRange(found.member.row, map[ZEPHYRUS.memberCol.telegramId]).setValue(String(found.message.chat.id));
  sheet.getRange(found.member.row, map[ZEPHYRUS.memberCol.code]).setValue('');
  zTelegramCall_('getUpdates', { offset: Number(found.update.update_id) + 1, timeout: 0 });
  zRegisterWebhook();
  zSendTelegram_(String(found.message.chat.id), found.member.name + '님 연결이 완료되었습니다.');
  zLog_('인증', '', '텔레그램', '성공', found.member.name + ' 텔레그램 연결 완료');
  zNotice_('텔레그램 연결을 완료했습니다.');
}

function doGet() {
  return ContentService.createTextOutput('Zephyrus webhook is running.');
}

// Kept for backward compatibility with earlier deployments.  The active
// webhook flow below sends the reply through the Bot API instead: Apps Script
// ContentService uses a redirect, and carrying Telegram's inline reply through
// that redirect can intermittently make Telegram retry the same update.
function zWebhookTelegramReply_(chatId, text) {
  return ContentService.createTextOutput(JSON.stringify({
    method: 'sendMessage',
    chat_id: String(chatId),
    text: String(text || '')
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(event) {
  try {
    var suppliedKey = event && event.parameter ? String(event.parameter.key || '') : '';
    if (!suppliedKey || suppliedKey !== zRequireProp_('WEBHOOK_KEY')) {
      zLog_('수신', '', '텔레그램', '거부', '웹훅 키가 없거나 일치하지 않습니다.');
      return ContentService.createTextOutput('forbidden');
    }
    var update = JSON.parse((event.postData && event.postData.contents) || '{}');
    var message = update.message;
    if (!message || !message.chat) return ContentService.createTextOutput('ok');
    var updateId = String(update.update_id || '');
    var lastUpdateId = Number(zProp_('LAST_TELEGRAM_UPDATE_ID') || 0);
    var numericUpdateId = Number(updateId || 0);
    // Telegram update IDs increase in order.  Store the last completed one
    // only after the handler finishes; unlike a temporary cache this cannot
    // make a freshly sent message look like an old duplicate.
    if (numericUpdateId && numericUpdateId <= lastUpdateId) {
      zLog_('수신', '', '텔레그램', '중복', '이미 처리한 업데이트: ' + updateId);
      return ContentService.createTextOutput('ok');
    }
    var chatId = String(message.chat.id);
    zLog_('수신', '', '텔레그램', '시작', '업데이트 수신: ' + (updateId || '번호 없음'));
    // Generate the reply first, then store the completed update ID before
    // sending.  This prevents a network retry from registering one schedule
    // twice, while the actual chat reply is delivered through Telegram's Bot
    // API rather than the fragile inline-webhook response path.
    var replyText = zHandleTelegramMessage_(chatId, String(message.text || '').trim(), true);
    if (numericUpdateId) zProps_().setProperty('LAST_TELEGRAM_UPDATE_ID', String(numericUpdateId));
    zLog_('수신', '', '텔레그램', '성공', '업데이트 처리 완료: ' + (updateId || '번호 없음'));
    if (replyText) {
      try {
        zSendTelegram_(chatId, replyText);
      } catch (sendError) {
        // Return 200 to stop Telegram from re-sending a schedule-registration
        // update.  The user can safely send a new query if a rare send error
        // occurs, and the failure is retained in the log for checking.
        zLog_('오류', '', '텔레그램', '실패', '답장 발신: ' + sendError);
      }
    }
    return ContentService.createTextOutput('ok');
  } catch (error) {
    zLog_('오류', '', '텔레그램', '실패', '웹훅: ' + error);
    return ContentService.createTextOutput('error');
  }
}

function zIsFullViewer_(member) {
  return member && (member.role === '관리자' || member.role === '전체열람');
}

function zIsScheduleVisibleToMember_(schedule, member) {
  if (zIsFullViewer_(member)) return true;
  if (schedule.registrant === member.name) return true;
  var target = String(schedule.targets || '').trim();
  if (target === '전부' || target === '전체') return true;
  return target.split(/[\s,]+/).filter(Boolean).indexOf(member.name) !== -1;
}

function zVisibleSchedulesForMember_(member, date, members) {
  return zSchedules_(members).filter(function(schedule) {
    return schedule.status !== ZEPHYRUS.status.deleted && (!date || schedule.date === date) && zIsScheduleVisibleToMember_(schedule, member);
  });
}

