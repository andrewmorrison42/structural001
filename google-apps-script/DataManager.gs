// ═══════════════════════════════════════════════════════════════════
// Data Manager — CRUD operations for entries and changelog
// Uses Google Drive JSON files with script-level locking
// ═══════════════════════════════════════════════════════════════════

// ── LOCKING ──────────────────────────────────────────────────────

function withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ── DATA FILE OPERATIONS ─────────────────────────────────────────

function loadData() {
  var fileId = getDataFileId();
  if (!fileId) return { version: '1.0', entries: [] };
  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString();
    return JSON.parse(content);
  } catch (e) {
    Logger.log('Error loading data: ' + e.toString());
    return { version: '1.0', entries: [] };
  }
}

function saveData(data) {
  var fileId = getDataFileId();
  if (!fileId) throw new Error('Data file not configured. Run setupKnowledgebase() first.');
  var file = DriveApp.getFileById(fileId);
  file.setContent(JSON.stringify(data, null, 2));
}

function loadChangelog() {
  var fileId = getChangelogFileId();
  if (!fileId) return { version: '1.0', changes: [] };
  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString();
    return JSON.parse(content);
  } catch (e) {
    return { version: '1.0', changes: [] };
  }
}

function saveChangelog(changelog) {
  var fileId = getChangelogFileId();
  if (!fileId) return;
  var file = DriveApp.getFileById(fileId);
  file.setContent(JSON.stringify(changelog, null, 2));
}

// ── CHANGELOG ────────────────────────────────────────────────────

function addChangelogEntry(entryId, action, field, oldValue, newValue, summary, email, name) {
  var changelog = loadChangelog();
  changelog.changes.push({
    id: generateChangeId(),
    entryId: entryId,
    action: action,
    timestamp: new Date().toISOString(),
    user: { name: name || getCurrentUserName(), email: email || getCurrentUserEmail() },
    details: {
      field: field || null,
      oldValue: oldValue || null,
      newValue: newValue || null,
      summary: summary || ''
    }
  });
  if (changelog.changes.length > 5000) {
    changelog.changes = changelog.changes.slice(-4000);
  }
  saveChangelog(changelog);
}

// ── CLIENT-CALLABLE FUNCTIONS ────────────────────────────────────

function getEntries() {
  var data = loadData();
  return data.entries.map(function(e) {
    var copy = JSON.parse(JSON.stringify(e));
    delete copy.htmlFileId;
    return copy;
  });
}

function getEntry(entryId) {
  var data = loadData();
  for (var i = 0; i < data.entries.length; i++) {
    if (data.entries[i].id === entryId) return data.entries[i];
  }
  return null;
}

function getEmailHtml(entryId) {
  var entry = getEntry(entryId);
  if (!entry || !entry.htmlFileId) return '<p>Email HTML not found.</p>';
  try {
    var file = DriveApp.getFileById(entry.htmlFileId);
    return file.getBlob().getDataAsString();
  } catch (e) {
    return '<p>Error loading email: ' + e.toString() + '</p>';
  }
}

function updateEntry(entryId, updates) {
  return withLock(function() {
    var data = loadData();
    var userEmail = getCurrentUserEmail();
    var userName = getCurrentUserName();

    for (var i = 0; i < data.entries.length; i++) {
      if (data.entries[i].id !== entryId) continue;

      var entry = data.entries[i];
      var changedFields = [];

      var editableFields = [
        'title', 'tags', 'contractStage', 'clientName',
        'region', 'notes', 'starred'
      ];

      for (var f = 0; f < editableFields.length; f++) {
        var field = editableFields[f];
        if (updates.hasOwnProperty(field) && JSON.stringify(entry[field]) !== JSON.stringify(updates[field])) {
          var oldVal = JSON.stringify(entry[field]);
          entry[field] = updates[field];
          changedFields.push(field);

          addChangelogEntry(entryId, 'edited', field, oldVal,
            JSON.stringify(updates[field]),
            'Updated ' + field, userEmail, userName
          );
        }
      }

      if (changedFields.length > 0) {
        entry.lastEditedBy = { name: userName, email: userEmail };
        entry.lastEditedAt = new Date().toISOString();
        saveData(data);
      }

      return entry;
    }
    throw new Error('Entry not found: ' + entryId);
  });
}

function deleteEntry(entryId) {
  return withLock(function() {
    var data = loadData();
    var userEmail = getCurrentUserEmail();
    var userName = getCurrentUserName();

    for (var i = 0; i < data.entries.length; i++) {
      if (data.entries[i].id !== entryId) continue;

      var entry = data.entries[i];
      addChangelogEntry(entryId, 'deleted', null, null, null,
        'Deleted entry: ' + entry.title, userEmail, userName
      );
      data.entries.splice(i, 1);
      saveData(data);
      return true;
    }
    throw new Error('Entry not found: ' + entryId);
  });
}

function getChangelogForEntry(entryId) {
  var changelog = loadChangelog();
  return changelog.changes.filter(function(c) {
    return c.entryId === entryId;
  }).reverse();
}

function getRecentChanges(limit) {
  var changelog = loadChangelog();
  var changes = changelog.changes.slice(-(limit || 50)).reverse();
  return changes;
}

function uploadImage(entryId, base64Data, fileName, mimeType) {
  return withLock(function() {
    var data = loadData();
    var userEmail = getCurrentUserEmail();
    var userName = getCurrentUserName();

    for (var i = 0; i < data.entries.length; i++) {
      if (data.entries[i].id !== entryId) continue;

      var entry = data.entries[i];
      var folderId = entry.emailFolderId;
      if (!folderId) {
        var rootFolder = DriveApp.getFolderById(getRootFolderId());
        var userFolder = getOrCreateSubfolder(rootFolder,
          sanitizeFolderName(userName + ' (' + userEmail + ')'));
        var emailFolder = getOrCreateSubfolder(userFolder, 'images');
        folderId = emailFolder.getId();
      }

      var folder = DriveApp.getFolderById(folderId);
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, mimeType || 'image/png', fileName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      var imageRecord = {
        name: fileName,
        driveFileId: file.getId(),
        url: 'https://drive.google.com/uc?id=' + file.getId(),
        uploadedBy: userEmail,
        uploadedAt: new Date().toISOString()
      };

      if (!entry.images) entry.images = [];
      entry.images.push(imageRecord);
      entry.lastEditedBy = { name: userName, email: userEmail };
      entry.lastEditedAt = new Date().toISOString();
      saveData(data);

      addChangelogEntry(entryId, 'edited', 'images', null, fileName,
        'Added image: ' + fileName, userEmail, userName
      );

      return imageRecord;
    }
    throw new Error('Entry not found: ' + entryId);
  });
}

function removeImage(entryId, driveFileId) {
  return withLock(function() {
    var data = loadData();
    var userEmail = getCurrentUserEmail();
    var userName = getCurrentUserName();

    for (var i = 0; i < data.entries.length; i++) {
      if (data.entries[i].id !== entryId) continue;
      var entry = data.entries[i];
      if (!entry.images) return;

      for (var j = 0; j < entry.images.length; j++) {
        if (entry.images[j].driveFileId === driveFileId) {
          var removed = entry.images.splice(j, 1)[0];
          entry.lastEditedBy = { name: userName, email: userEmail };
          entry.lastEditedAt = new Date().toISOString();
          saveData(data);

          addChangelogEntry(entryId, 'edited', 'images', removed.name, null,
            'Removed image: ' + removed.name, userEmail, userName
          );
          return true;
        }
      }
    }
  });
}

function createManualEntry(title, notes, tags, contractStage, clientName) {
  return withLock(function() {
    var data = loadData();
    var userEmail = getCurrentUserEmail();
    var userName = getCurrentUserName();

    var entry = {
      id: generateId(),
      emailMessageId: null,
      emailThreadId: null,
      title: title || 'Untitled',
      subject: '',
      from: { name: userName, email: userEmail },
      to: [],
      cc: [],
      date: new Date().toISOString(),
      snippet: (notes || '').substring(0, 300),
      htmlFileId: null,
      htmlFileName: null,
      emailFolderId: null,
      attachments: [],
      submittedBy: { name: userName, email: userEmail },
      submittedAt: new Date().toISOString(),
      tags: tags || [],
      contractStage: contractStage || '',
      clientName: clientName || '',
      region: '',
      notes: notes || '',
      images: [],
      starred: false,
      lastEditedBy: { name: userName, email: userEmail },
      lastEditedAt: new Date().toISOString()
    };

    data.entries.push(entry);
    saveData(data);

    addChangelogEntry(entry.id, 'created', null, null, null,
      'Manual entry created: ' + title, userEmail, userName
    );

    return entry;
  });
}

function getAllTags() {
  var data = loadData();
  var tagCounts = {};
  data.entries.forEach(function(e) {
    (e.tags || []).forEach(function(tag) {
      tagCounts[tag] = (tagCounts[tag] || 0) + 1;
    });
  });
  return tagCounts;
}

function getAllStages() {
  var data = loadData();
  var stages = {};
  data.entries.forEach(function(e) {
    if (e.contractStage) {
      stages[e.contractStage] = (stages[e.contractStage] || 0) + 1;
    }
  });
  return stages;
}

function getAllSubmitters() {
  var data = loadData();
  var submitters = {};
  data.entries.forEach(function(e) {
    var key = e.submittedBy.email;
    if (!submitters[key]) {
      submitters[key] = { name: e.submittedBy.name, email: key, count: 0 };
    }
    submitters[key].count++;
  });
  return submitters;
}

function getAppConfig() {
  return {
    userEmail: getCurrentUserEmail(),
    userName: getCurrentUserName(),
    rootFolderId: getRootFolderId(),
    contractStages: [
      'Lead',
      'Initial Contact',
      'Discovery',
      'Proposal',
      'Negotiation',
      'Contract Review',
      'Closed Won',
      'Closed Lost',
      'Post-Sale'
    ]
  };
}
