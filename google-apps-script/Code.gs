// ═══════════════════════════════════════════════════════════════════
// Sales Knowledgebase — Google Apps Script
// Main entry point: setup, configuration, triggers, web app serving
// ═══════════════════════════════════════════════════════════════════

// ── CONFIGURATION ────────────────────────────────────────────────
// Change these to match your environment
var CONFIG = {
  GMAIL_LABEL_INCOMING: 'Sales-KB',
  GMAIL_LABEL_FILED:    'Sales-KB-Filed',
  ROOT_FOLDER_NAME:     'Sales Knowledgebase',
  DATA_FILE_NAME:       'knowledgebase_data.json',
  CHANGELOG_FILE_NAME:  'changelog.json',
  MAX_EMAILS_PER_RUN:   10,
  TRIGGER_MINUTES:      5
};

// Cached folder/file IDs (set during setup or first run)
var PROPS = PropertiesService.getScriptProperties();

// ── SETUP ────────────────────────────────────────────────────────
// Run this once manually to create folder structure and labels
function setupKnowledgebase() {
  var root = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  PROPS.setProperty('ROOT_FOLDER_ID', root.getId());

  var dataFile = root.createFile(
    CONFIG.DATA_FILE_NAME,
    JSON.stringify({ version: '1.0', entries: [] }, null, 2),
    MimeType.PLAIN_TEXT
  );
  PROPS.setProperty('DATA_FILE_ID', dataFile.getId());

  var changelogFile = root.createFile(
    CONFIG.CHANGELOG_FILE_NAME,
    JSON.stringify({ version: '1.0', changes: [] }, null, 2),
    MimeType.PLAIN_TEXT
  );
  PROPS.setProperty('CHANGELOG_FILE_ID', changelogFile.getId());

  var incomingLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL_INCOMING);
  if (!incomingLabel) {
    GmailApp.createLabel(CONFIG.GMAIL_LABEL_INCOMING);
  }
  var filedLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL_FILED);
  if (!filedLabel) {
    GmailApp.createLabel(CONFIG.GMAIL_LABEL_FILED);
  }

  Logger.log('Setup complete.');
  Logger.log('Root folder: https://drive.google.com/drive/folders/' + root.getId());
  Logger.log('Data file ID: ' + dataFile.getId());
  Logger.log('Share the root folder with your team for shared access.');
  Logger.log('Now run installTrigger() to start automatic email filing.');
}

// Run this after setupKnowledgebase to install the recurring trigger
function installTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processIncomingEmails')
    .timeDriven()
    .everyMinutes(CONFIG.TRIGGER_MINUTES)
    .create();
  Logger.log('Trigger installed: checking every ' + CONFIG.TRIGGER_MINUTES + ' minutes.');
}

// Remove the trigger
function uninstallTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log('Trigger removed.');
}

// ── MANUAL CONFIG ────────────────────────────────────────────────
// If the root folder already exists, set its ID manually
function setRootFolderId(folderId) {
  PROPS.setProperty('ROOT_FOLDER_ID', folderId);
  var folder = DriveApp.getFolderById(folderId);
  var dataFiles = folder.getFilesByName(CONFIG.DATA_FILE_NAME);
  if (dataFiles.hasNext()) {
    PROPS.setProperty('DATA_FILE_ID', dataFiles.next().getId());
  }
  var changelogFiles = folder.getFilesByName(CONFIG.CHANGELOG_FILE_NAME);
  if (changelogFiles.hasNext()) {
    PROPS.setProperty('CHANGELOG_FILE_ID', changelogFiles.next().getId());
  }
  Logger.log('Root folder set to: ' + folderId);
}

// ── WEB APP ──────────────────────────────────────────────────────
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'index';
  if (page === 'email') {
    return HtmlService.createTemplateFromFile('EmailViewer')
      .evaluate()
      .setTitle('Email Viewer')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Knowledgebase')
    .evaluate()
    .setTitle('Sales Knowledgebase')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── UTILITY ──────────────────────────────────────────────────────
function getRootFolderId() {
  return PROPS.getProperty('ROOT_FOLDER_ID');
}

function getDataFileId() {
  return PROPS.getProperty('DATA_FILE_ID');
}

function getChangelogFileId() {
  return PROPS.getProperty('CHANGELOG_FILE_ID');
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getCurrentUserName() {
  var email = getCurrentUserEmail();
  return email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, function(c) {
    return c.toUpperCase();
  });
}

function generateId() {
  return 'entry_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
}

function generateChangeId() {
  return 'chg_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
}
