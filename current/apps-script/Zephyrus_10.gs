function zTelegramTime_(period, hour, minute) {
  if (hour === '' || hour === undefined || hour === null) return '';
  var h = Number(hour);
  var m = Number(minute || 0);
  if (period === '오후' && h < 12) h += 12;
  if (period === '오전' && h === 12) h = 0;
  // In ordinary Korean work-schedule input, an unqualified 1~7시는 usually
  // intended as afternoon.  오전/오후 or 24-hour input remains explicit.
  if (!period && h > 0 && h < 8) h += 12;
  return zNormalizeTime_(('0' + h).slice(-2) + ':' + ('0' + m).slice(-2));
}

function zExtractTelegramTargets_(titleText, members) {
  var text = String(titleText || '').trim();
  if (!text) return { title: '', targets: '' };

  // '전부' / '전체'가 맨 끝에 있으면 전사 일정으로 처리한다.
  var allMatch = text.match(/^(.*?)(?:\s+)(전부|전체)$/);
  if (allMatch && String(allMatch[1] || '').trim()) {
    return { title: String(allMatch[1] || '').trim(), targets: allMatch[2] };
  }

  var known = {};
  (members || zMembers_()).forEach(function(member) {
    if (member && member.name) known[String(member.name).trim()] = true;
  });

  // 맨 뒤에서부터 구성원 이름을 최대한 많이 떼어낸다.
  // 예: "회의 김형원, 성도형" -> title "회의", targets "김형원, 성도형"
  var normalized = text.replace(/\s*,\s*/g, ',');
  var tokens = normalized.split(/\s+/);
  var targetNames = [];

  while (tokens.length > 1) {
    var tail = tokens[tokens.length - 1];
    var parts = tail.split(',').filter(Boolean);
    var allKnown = parts.length && parts.every(function(name) { return !!known[name]; });
    if (!allKnown) break;

    targetNames = parts.concat(targetNames);
    tokens.pop();

    // 앞 토큰이 "김형원,"처럼 끝나는 경우도 이어서 처리한다.
    while (tokens.length > 1) {
      var prev = tokens[tokens.length - 1];
      var prevParts = prev.split(',').filter(Boolean);
      if (!prevParts.length || !prevParts.every(function(name) { return !!known[name]; })) break;
      targetNames = prevParts.concat(targetNames);
      tokens.pop();
    }
    break;
  }

  if (!targetNames.length) return { title: text, targets: '' };

  return {
    title: tokens.join(' ').trim(),
    targets: targetNames.filter(function(name, index, arr) {
      return arr.indexOf(name) === index;
    }).join(', ')
  };
}

function zParseTelegramTimeAndTitle_(rest, members) {
  var source = String(rest || '').trim();

  // 지원 예:
  // 오후 2시~3시 회의
  // 오후 2시~오후 4시 회의
  // 오전 11시~오후 8시 POC
  // 오후 10시20분~오후 11시 연휴
  // 오후 10시 20분~오후 11시 연휴
  // 오후 10:20~오후 11:00 연휴
  var timeMatch = source.match(
    /^(오전|오후)?\s*(\d{1,2})(?:(?::(\d{1,2}))|\s*시(?:\s*(\d{1,2})\s*분?)?)\s*(?:[~～]\s*(오전|오후)?\s*(\d{1,2})(?:(?::(\d{1,2}))|\s*시(?:\s*(\d{1,2})\s*분?)?))?\s+(.+)$/
  );

  if (timeMatch) {
    var parsed = zExtractTelegramTargets_(String(timeMatch[9] || '').trim(), members);
    return {
      time: zTelegramTime_(timeMatch[1] || '', timeMatch[2], timeMatch[3] || timeMatch[4]),
      endTime: timeMatch[6] ? zTelegramTime_(timeMatch[5] || '', timeMatch[6], timeMatch[7] || timeMatch[8]) : '',
      title: parsed.title,
      targets: parsed.targets
    };
  }

  // 시간이 없는 일정도 기존처럼 허용한다.
  var noTimeParsed = zExtractTelegramTargets_(source, members);
  return {
    time: '',
    endTime: '',
    title: noTimeParsed.title,
    targets: noTimeParsed.targets
  };
}

function zParseTelegramSchedule_(text, members) {
  var source = String(text || '').trim().replace(/^일정\s*/i, '');

  // YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
  var full = source.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s+(.+)$/);
  if (full) {
    var fullBody = zParseTelegramTimeAndTitle_(full[4], members);
    if (!fullBody.title) return null;
    return {
      date: zNormalizeDate_(full[1] + '-' + full[2] + '-' + full[3]),
      time: fullBody.time,
      endTime: fullBody.endTime,
      title: fullBody.title,
      targets: fullBody.targets
    };
  }

  // 9월 15일 / 9/15 / 9.15
  var short = source.match(/^(\d{1,2})(?:\s*월\s*|[/.])(\d{1,2})\s*(?:일)?\s+(.+)$/);
  if (short) {
    var shortBody = zParseTelegramTimeAndTitle_(short[3], members);
    if (!shortBody.title) return null;
    return {
      date: zNormalizeDate_(new Date().getFullYear() + '-' + short[1] + '-' + short[2]),
      time: shortBody.time,
      endTime: shortBody.endTime,
      title: shortBody.title,
      targets: shortBody.targets
    };
  }

  return null;
}

function zTelegramUsageText_() {
  return '9월 15일 오후2시~3시 회의 대상자,대상자\n'
    + '(본인 일정: 대상자 생략)\n'
    + '전체 일정: 전부 또는 전체\n'
    + '일정 보기: 오늘 입력\n'
    + '문법 보기: .';
}

function zHandleTelegramMessage_(chatId, text, replyViaWebhook) {
  text = String(text || '').replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
  var compactText = text.normalize('NFC').replace(/\s+/g, '');
  var members = zMembers_();
  var member = members.filter(function(item) { return item.telegramId === chatId; })[0] || null;
  function reply(message) {
    if (replyViaWebhook) return String(message);
    zSendTelegram_(chatId, message);
    return '';
  }
  if (text === '/start') {
    return reply(member ? zTelegramUsageText_() : '제피로스 일정관리입니다. 받은 6자리 인증코드를 입력하세요.');
  }
  var enteredCode = text.replace(/[^0-9]/g, '');
  if (/^\d{6}$/.test(enteredCode)) {
    var match = members.filter(function(item) {
      return String(item.code || '').replace(/[^0-9]/g, '') === enteredCode;
    })[0] || null;
    if (!match) {
      return reply('인증코드를 찾지 못했습니다.');
    }
    var sheet = zSheet_(ZEPHYRUS.sheet.members);
    var map = zHeaders_(sheet);
    sheet.getRange(match.row, map[ZEPHYRUS.memberCol.telegramId]).setValue(chatId);
    sheet.getRange(match.row, map[ZEPHYRUS.memberCol.code]).setValue('');
    zLog_('인증', '', '텔레그램', '성공', match.name + ' 텔레그램 연결 완료');
    return reply(match.name + '님 연결이 완료되었습니다.\n\n' + zTelegramUsageText_());
  }
  if (!member) {
    return reply('/start 후 인증코드를 먼저 입력하세요.');
  }
  if (text === '.' || compactText === '사용법' || compactText === '도움말' || text === '/help') {
    return reply(zTelegramUsageText_());
  }
  // Telegram clients can retain an invisible formatting character even when
  // the bubble visibly says only "오늘".  Treat any otherwise ordinary text
  // containing 오늘 as the today-query command.
  if (compactText === '오늘일정' || compactText === '오늘') {
    var items = zVisibleSchedulesForMember_(member, zToday_(), members);
    zLog_('조회', '', '텔레그램', '성공', member.name + ' / 오늘 / ' + items.length + '건');
    return reply(items.length ? items.map(zScheduleText_).join('\n\n') : '오늘 일정이 없습니다.');
  }
  var add = zParseTelegramSchedule_(text, members);
  if (add && add.date && add.title) {
    // 텔레그램 등록은 다른 편집 작업의 DocumentLock을 기다리지 않는다.
    // 고유 일정ID를 사용하므로 동시 등록 시에도 자신의 행을 다시 찾아 처리한다.
    // 정렬은 주기 동기화/후속 작업에 맡겨 웹훅 응답을 가볍게 유지한다.
    var schedule = zAppendSchedule_({
      date: add.date, time: add.time, endTime: add.endTime, title: add.title,
      targets: add.targets || member.name, registrant: member.name, channel: '텔레그램',
      deferSort: true
    });

    // 사용자는 먼저 즉시 접수 사실을 확인한다.
    // 웹훅 처리 중 캘린더 API가 조금 느려도 텔레그램 체감 응답은 먼저 오게 한다.
    if (replyViaWebhook) {
      try {
        zSendTelegram_(chatId, '등록했습니다.\n' + zScheduleText_(schedule));
      } catch (quickReplyError) {
        zLog_('오류', schedule.id, '텔레그램', '실패', '즉시 등록답장: ' + quickReplyError);
      }
    }

    // 새 일정은 공용 + 대상자 개인캘린더만 즉시 반영 시도.
    // 실패해도 기존 주기 동기화가 나중에 복구한다.
    var immediate = zImmediateCalendarSyncForNewSchedule_(schedule);
    if (immediate && immediate.schedule) schedule = immediate.schedule;

    // 대상자 알림은 기존 방식 그대로 보낸다.
    // 보낸 사람은 위 등록확인 메시지를 이미 받았으므로 중복 알림을 제외한다.
    var notified = zNotifySchedule_(schedule, '[새 일정]', chatId);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.alarm, zAlarmResult_(notified, true));

    if (replyViaWebhook) {
      // 이미 직접 답장을 보냈으므로 doPost에서 다시 보내지 않는다.
      return '';
    }

    return reply(
      '등록했습니다.\n' +
      zScheduleText_(schedule) +
      (immediate && immediate.ok ? '\n\n캘린더 반영 완료' : '\n\n캘린더는 자동 동기화됩니다.')
    );
  }
  return reply(zTelegramUsageText_());
}

function zFileIndexSheet_() {
  var ss = zSpreadsheet_();
  var sheet = ss.getSheetByName(ZEPHYRUS.sheet.fileIndex);
  if (sheet) return sheet;
  sheet = ss.insertSheet(ZEPHYRUS.sheet.fileIndex);
  sheet.appendRow(['일정ID', '폴더ID', '첨부파일ID목록']);
  sheet.hideSheet();
  return sheet;
}

function zFileIndexRecord_(scheduleId) {
  var sheet = zFileIndexSheet_();
  if (sheet.getLastRow() < 2) return { row: 0, folderId: '', fileIds: [] };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (var index = 0; index < rows.length; index++) {
    if (String(rows[index][0] || '') !== String(scheduleId || '')) continue;
    return { row: index + 2, folderId: String(rows[index][1] || ''), fileIds: String(rows[index][2] || '').split(',').filter(Boolean) };
  }
  return { row: 0, folderId: '', fileIds: [] };
}

