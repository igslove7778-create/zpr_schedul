function zSetFileIndexRecord_(scheduleId, folderId, fileIds) {
  var sheet = zFileIndexSheet_();
  var record = zFileIndexRecord_(scheduleId);
  var row = record.row || sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 3).setValues([[String(scheduleId || ''), String(folderId || ''), (fileIds || []).join(',')]]);
}

function zFileRootFolder_() {
  var props = zProps_();
  var knownId = String(props.getProperty('SCHEDULE_FILE_ROOT_FOLDER_ID') || '');
  if (knownId) { try { return DriveApp.getFolderById(knownId); } catch (ignore) {} }
  var folder = DriveApp.createFolder('제피로스 일정자료');
  props.setProperty('SCHEDULE_FILE_ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function zScheduleFileFolderName_(schedule) {
  var title = String(schedule.title || '제목 없음').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  return (schedule.date || '날짜미정') + '_' + title + '_' + schedule.id;
}

function zScheduleFileFolder_(schedule, createIfMissing) {
  if (!schedule || !schedule.id) return null;
  var record = zFileIndexRecord_(schedule.id);
  if (record.folderId) { try { return DriveApp.getFolderById(record.folderId); } catch (ignore) {} }
  if (!createIfMissing) return null;
  var folder = zFileRootFolder_().createFolder(zScheduleFileFolderName_(schedule));
  zSetFileIndexRecord_(schedule.id, folder.getId(), record.fileIds);
  return folder;
}

function zScheduleFileFolderUrl_(scheduleId) {
  if (!scheduleId) return '';
  var record = zFileIndexRecord_(scheduleId);
  if (!record.folderId) return '';
  try { return DriveApp.getFolderById(record.folderId).getUrl(); } catch (ignore) { return ''; }
}

function zFileRecipientEmails_(schedule) {
  var seen = {};
  var recipients = zRecipients_(schedule);
  // File access must work before a person's private calendar exists.
  zMembers_().forEach(function(member) {
    if (member.name === schedule.registrant) recipients.push(member);
  });
  return recipients.map(function(member) { return String(member.email || '').trim().toLowerCase(); }).filter(function(email) {
    if (!email || seen[email]) return false;
    seen[email] = true;
    return true;
  });
}

function zSyncScheduleFolderViewers_(folder, schedule) {
  var wanted = {};
  zFileRecipientEmails_(schedule).forEach(function(email) { wanted[email] = true; });
  folder.getViewers().forEach(function(user) {
    var email = String(user.getEmail() || '').toLowerCase();
    if (email && !wanted[email]) folder.removeViewer(email);
  });
  Object.keys(wanted).forEach(function(email) {
    if (folder.getAccess(email) === DriveApp.Permission.NONE) folder.addViewer(email);
  });
}

function zCalendarAttachments_(event) {
  return ((event && event.attachments) || []).map(function(attachment) {
    return { id: String(attachment.fileId || ''), title: String(attachment.title || '첨부파일'), url: String(attachment.fileUrl || '') };
  }).filter(function(attachment) { return attachment.id && attachment.url; });
}

function zFolderHasShortcut_(folder, targetId) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try { if (file.getTargetId && String(file.getTargetId() || '') === String(targetId)) return true; } catch (ignore) {}
  }
  return false;
}

function zRemoveStaleFolderShortcuts_(folder, fileIds) {
  var expected = {};
  (fileIds || []).forEach(function(id) { expected[String(id)] = true; });
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      var targetId = String(file.getTargetId ? file.getTargetId() || '' : '');
      if (targetId && !expected[targetId]) file.setTrashed(true);
    } catch (ignore) {}
  }
}

function zSyncAttachmentShortcuts_(folder, schedule, attachments) {
  var recipients = zFileRecipientEmails_(schedule);
  zRemoveStaleFolderShortcuts_(folder, attachments.map(function(attachment) { return attachment.id; }));
  attachments.forEach(function(attachment) {
    try {
      var source = DriveApp.getFileById(attachment.id);
      recipients.forEach(function(email) {
        if (source.getAccess(email) === DriveApp.Permission.NONE) source.addViewer(email);
      });
      if (!zFolderHasShortcut_(folder, attachment.id)) DriveApp.createShortcut(attachment.id).setName(attachment.title).moveTo(folder);
    } catch (error) {
      zLog_('오류', schedule.id, '파일', '공유 실패', attachment.title + ': ' + error);
    }
  });
}

function zSetScheduleFileCell_(schedule, folder, count) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var column = zHeaders_(sheet)[ZEPHYRUS.col.file];
  if (!column) return;
  var cell = sheet.getRange(schedule.row, column);
  if (!count) { cell.clearContent(); return; }
  var label = '파일 ' + count + '개 보기';
  var value = SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(folder.getUrl()).build();
  cell.setRichTextValue(value);
}

function zSyncScheduleFiles_(schedule, event) {
  if (!schedule || !schedule.id) return '';
  try {
    var attachments = event ? zCalendarAttachments_(event) : null;
    var record = zFileIndexRecord_(schedule.id);
    if (!record.folderId && (!attachments || !attachments.length)) return '';
    var folder = zScheduleFileFolder_(schedule, !!(attachments && attachments.length));
    if (!folder) return '';
    zSyncScheduleFolderViewers_(folder, schedule);
    if (attachments !== null) {
      zSyncAttachmentShortcuts_(folder, schedule, attachments);
      zSetFileIndexRecord_(schedule.id, folder.getId(), attachments.map(function(attachment) { return attachment.id; }));
      record.fileIds = attachments.map(function(attachment) { return attachment.id; });
    }
    zSetScheduleFileCell_(schedule, folder, record.fileIds.length);
    return folder.getUrl();
  } catch (error) {
    zLog_('오류', schedule.id, '파일', '실패', error);
    return '';
  }
}

function zCalendarTargetInfo_(description) {
  var text = String(description || '').trim();
  var found = false;
  var values = [];

  // 기존 방식도 그대로 지원한다.
  // 예: "대상: 김형원, 김명재" / "대상자: 성도형"
  text.split(/\r?\n/).forEach(function(line) {
    var match = String(line).match(/^\s*대상(?:자)?\s*[:：]\s*(.+?)\s*$/);
    if (!match) return;
    found = true;
    values.push(match[1]);
  });

  if (found) {
    var explicitNames = zTargetNames_(values.join(', '));
    if (explicitNames === '전부' || explicitNames === '전체') {
      return { found: true, targets: explicitNames };
    }
    return {
      found: true,
      targets: explicitNames.filter(function(name, index) {
        return explicitNames.indexOf(name) === index;
      }).join(', ')
    };
  }

  // 새 방식: 설명란 전체가 등록된 구성원 이름들만으로 이루어졌으면
  // "대상:" 없이도 대상자로 인식한다.
  // 쉼표와 공백을 모두 허용한다.
  // 예: "김형원", "김형원, 김명재, 성도형", "김형원 김명재 성도형"
  // 반대로 "김명재 부장과 회의 후 자료 전달"처럼 일반 문장이 섞이면
  // 대상자로 처리하지 않는다.
  if (!text) return { found: false, targets: '' };

  var plainNames = zTargetNames_(text);
  if (plainNames === '전부' || plainNames === '전체') {
    return { found: true, targets: plainNames };
  }
  if (!plainNames.length) return { found: false, targets: '' };

  var known = {};
  zMembers_().forEach(function(member) {
    if (member && member.name) known[String(member.name).trim()] = true;
  });

  var allKnown = plainNames.every(function(name) { return !!known[name]; });
  if (!allKnown) return { found: false, targets: '' };

  return {
    found: true,
    targets: plainNames.filter(function(name, index) {
      return plainNames.indexOf(name) === index;
    }).join(', ')
  };
}

function zEventDescription_(schedule) {
  var lines = [];
  if (schedule.targets) lines.push('대상: ' + schedule.targets);
  var folderUrl = zScheduleFileFolderUrl_(schedule.id);
  if (folderUrl) lines.push('자료: ' + folderUrl);
  if (schedule.registrant) lines.push('등록: ' + schedule.registrant);
  lines.push('[일정ID:' + schedule.id + ']');
  lines.push(ZEPHYRUS.markerPrefix + schedule.id + ']');
  return lines.join('\n');
}

function zEventTimes_(schedule) {
  var day = zDateObject_(schedule.date);
  if (!day) throw new Error('날짜 형식 오류: ' + schedule.date);
  if (!schedule.time) return { allDay: true, start: day, end: null };
  var parts = schedule.time.split(':').map(Number);
  var start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), parts[0], parts[1]);
  if (schedule.endTime) {
    var endParts = schedule.endTime.split(':').map(Number);
    var enteredEnd = new Date(day.getFullYear(), day.getMonth(), day.getDate(), endParts[0], endParts[1]);
    // An end before the start is an overnight schedule.
    if (enteredEnd.getTime() <= start.getTime()) enteredEnd = new Date(enteredEnd.getTime() + 86400000);
    return { allDay: false, start: start, end: enteredEnd };
  }
  var minutes = Number(zSettings_()['기본일정길이분'] || 60) || 60;
  return { allDay: false, start: start, end: new Date(start.getTime() + minutes * 60000) };
}

