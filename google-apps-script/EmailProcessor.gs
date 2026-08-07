// ═══════════════════════════════════════════════════════════════════
// Email Processor — detects labelled emails, converts to HTML, files
// ═══════════════════════════════════════════════════════════════════

function processIncomingEmails() {
  var incomingLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL_INCOMING);
  if (!incomingLabel) return;

  var filedLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL_FILED);
  if (!filedLabel) {
    filedLabel = GmailApp.createLabel(CONFIG.GMAIL_LABEL_FILED);
  }

  var threads = incomingLabel.getThreads(0, CONFIG.MAX_EMAILS_PER_RUN);
  if (threads.length === 0) return;

  var rootFolderId = getRootFolderId();
  if (!rootFolderId) {
    Logger.log('Root folder not configured. Run setupKnowledgebase() first.');
    return;
  }

  var rootFolder = DriveApp.getFolderById(rootFolderId);
  var userEmail = getCurrentUserEmail();
  var userName = getCurrentUserName();
  var userFolder = getOrCreateSubfolder(rootFolder, sanitizeFolderName(userName + ' (' + userEmail + ')'));

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var messages = thread.getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      try {
        fileEmail(message, userFolder, userEmail, userName);
      } catch (e) {
        Logger.log('Error filing email ' + message.getId() + ': ' + e.toString());
      }
    }

    thread.removeLabel(incomingLabel);
    thread.addLabel(filedLabel);
  }
}

function fileEmail(message, userFolder, userEmail, userName) {
  var existingData = loadData();
  var existingIds = existingData.entries.map(function(e) { return e.emailMessageId; });
  if (existingIds.indexOf(message.getId()) !== -1) return;

  var subject = message.getSubject() || '(No Subject)';
  var from = message.getFrom();
  var to = message.getTo();
  var cc = message.getCc() || '';
  var date = message.getDate();
  var body = message.getBody();
  var plainBody = message.getPlainBody();

  var attachments = message.getAttachments();
  var savedAttachments = [];
  var inlineImages = {};

  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var safeSubject = sanitizeFileName(subject).substring(0, 60);
  var emailFolderName = dateStr + '_' + safeSubject;
  var emailFolder = getOrCreateSubfolder(userFolder, emailFolderName);

  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    var attName = att.getName() || ('attachment_' + i);
    var contentId = att.getContentId ? att.getContentId() : null;

    try {
      var savedFile = emailFolder.createFile(att.copyBlob().setName(attName));

      var attRecord = {
        name: attName,
        mimeType: att.getContentType(),
        driveFileId: savedFile.getId(),
        sizeBytes: att.getSize()
      };
      savedAttachments.push(attRecord);

      if (contentId && att.getContentType().indexOf('image/') === 0) {
        var bytes = att.getBytes();
        var b64 = Utilities.base64Encode(bytes);
        inlineImages['cid:' + contentId] = 'data:' + att.getContentType() + ';base64,' + b64;
      }
    } catch (e) {
      Logger.log('Error saving attachment ' + attName + ': ' + e.toString());
    }
  }

  var processedHtml = buildEmailHtml(message, body, inlineImages);
  var htmlFileName = dateStr + '_' + safeSubject + '.html';
  var htmlFile = emailFolder.createFile(htmlFileName, processedHtml, MimeType.HTML);

  var snippet = (plainBody || '').substring(0, 300).replace(/\s+/g, ' ').trim();

  var entry = {
    id: generateId(),
    emailMessageId: message.getId(),
    emailThreadId: message.getThread().getId(),
    title: subject,
    subject: subject,
    from: parseEmailAddress(from),
    to: (to || '').split(',').map(function(e) { return e.trim(); }),
    cc: cc ? cc.split(',').map(function(e) { return e.trim(); }) : [],
    date: date.toISOString(),
    snippet: snippet,
    htmlFileId: htmlFile.getId(),
    htmlFileName: htmlFileName,
    emailFolderId: emailFolder.getId(),
    attachments: savedAttachments,
    submittedBy: {
      name: userName,
      email: userEmail
    },
    submittedAt: new Date().toISOString(),
    tags: [],
    contractStage: '',
    clientName: '',
    region: '',
    notes: '',
    images: [],
    starred: false,
    lastEditedBy: {
      name: userName,
      email: userEmail
    },
    lastEditedAt: new Date().toISOString()
  };

  withLock(function() {
    var data = loadData();
    data.entries.push(entry);
    saveData(data);

    addChangelogEntry(entry.id, 'created', null, null, null,
      'Entry created from email: ' + subject,
      userEmail, userName
    );
  });

  Logger.log('Filed: ' + subject);
}

function buildEmailHtml(message, body, inlineImages) {
  var subject = message.getSubject() || '(No Subject)';
  var from = message.getFrom();
  var to = message.getTo();
  var cc = message.getCc() || '';
  var date = message.getDate();

  var processedBody = body || '';
  for (var cid in inlineImages) {
    processedBody = processedBody.replace(
      new RegExp(cid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'),
      inlineImages[cid]
    );
  }

  var html = '<!DOCTYPE html>\n<html>\n<head>\n';
  html += '<meta charset="utf-8">\n';
  html += '<meta name="viewport" content="width=device-width, initial-scale=1">\n';
  html += '<title>' + escapeHtml(subject) + '</title>\n';
  html += '<style>\n';
  html += 'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; ';
  html += 'max-width: 900px; margin: 0 auto; padding: 20px; color: #1a1a2e; background: #fff; }\n';
  html += '.email-header { background: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 8px; ';
  html += 'padding: 16px 20px; margin-bottom: 24px; }\n';
  html += '.email-header .field { margin: 4px 0; font-size: 14px; }\n';
  html += '.email-header .label { font-weight: 600; color: #475569; min-width: 50px; display: inline-block; }\n';
  html += '.email-header .subject { font-size: 18px; font-weight: 700; color: #1e293b; margin-bottom: 12px; }\n';
  html += '.email-body { line-height: 1.6; }\n';
  html += '.email-body img { max-width: 100%; height: auto; }\n';
  html += '.filed-notice { background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 6px; ';
  html += 'padding: 10px 14px; margin-bottom: 16px; font-size: 12px; color: #065f46; }\n';
  html += '</style>\n</head>\n<body>\n';

  html += '<div class="filed-notice">Filed to Sales Knowledgebase on ';
  html += Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  html += '</div>\n';

  html += '<div class="email-header">\n';
  html += '<div class="subject">' + escapeHtml(subject) + '</div>\n';
  html += '<div class="field"><span class="label">From:</span> ' + escapeHtml(from) + '</div>\n';
  html += '<div class="field"><span class="label">To:</span> ' + escapeHtml(to) + '</div>\n';
  if (cc) {
    html += '<div class="field"><span class="label">CC:</span> ' + escapeHtml(cc) + '</div>\n';
  }
  html += '<div class="field"><span class="label">Date:</span> ' +
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE, dd MMM yyyy HH:mm') + '</div>\n';
  html += '</div>\n';

  html += '<div class="email-body">\n' + processedBody + '\n</div>\n';
  html += '</body>\n</html>';

  return html;
}

// ── HELPERS ──────────────────────────────────────────────────────

function getOrCreateSubfolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function sanitizeFileName(name) {
  return name.replace(/[^a-zA-Z0-9 _\-]/g, '').replace(/\s+/g, '_');
}

function sanitizeFolderName(name) {
  return name.replace(/[\/\\:*?"<>|]/g, '_');
}

function parseEmailAddress(raw) {
  var match = raw.match(/^(.*?)\s*<(.+?)>$/);
  if (match) {
    return { name: match[1].replace(/"/g, '').trim(), email: match[2].trim() };
  }
  return { name: raw.trim(), email: raw.trim() };
}

function escapeHtml(str) {
  return (str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
