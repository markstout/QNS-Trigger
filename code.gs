/**
 * Quick Notes Suite - Trigger Code
 * Copyright 2026 Mark A. Stout
 * Last Updated : 6-2-2026
 * see https://sites.google.com/view/quick-notes-suite
 * * @fileoverview Google App Script functions for the QuickNoteSuite Trigger Handler.
 * This script is intended to be a separate project from the main web app.
 * Its sole purpose is to run background tasks based on time-driven triggers.
 *
 * To use this script:
 * 1. Create a new, standalone Apps Script project or a script bound to a spreadsheet.
 * 2. Paste this entire code into the Code.gs file.
 * 3. After a user has completed the web app setup, run the `setupTriggers` 
 * function MANUALLY from the script editor ONCE for that user's account.
 */

// --- GLOBAL CONSTANTS ---
const PREFERENCE_FILE_NAME = 'QuickNoteSuitePreferences-Do-not-Delete-or-Edit';
const PREFERENCE_PARENT_FOLDER_NAME = 'Notes2';
const PREFERENCE_SUB_FOLDER_NAME = 'System Files - DO NOT DELETE OR EDIT';
const PREFERENCE_SPREADSHEET_NAME = 'CapIt-Configuration';
const PREFERENCE_SHEET_TAB_NAME = 'Config';

// --- MASTER SETUP FUNCTION ---

/**
 * Deletes all existing triggers for this project and creates new ones.
 * This function should be run manually once after the web app setup is complete.
 */
function setupTriggers() {
  const ui = SpreadsheetApp.getUi(); // Assumes this is run from a spreadsheet

  // Safety Check: Do not run if the user hasn't completed the web app setup.
  if (!checkPreferencesFileExists()) {
    ui.alert("Preferences file not found. Please complete the web app setup before creating triggers. Aborting.");
    return;
  }

  const removedCount = deleteAllTriggers();
  ui.alert(`${removedCount} existing trigger(s) have been deleted. Press OK to regenerate the triggers with updated versions.`);

  let createdCount = 0;

  const triggerFunctions = [
    'triggered_processEmailsToDoc',
    'triggered_convertTextNotesToDoc',
    'triggered_MoveKeepNotes'
  ];

  triggerFunctions.forEach(function(functionName) {
    try {
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyMinutes(5)
        .withFailureNotificationFrequency(ScriptApp.FailureNotificationFrequency.HOURLY)
        .create();
      Logger.log(`Successfully created 5-minute trigger for ${functionName} with hourly error notifications.`);
      createdCount++;
    } catch (e) {
      Logger.log(`Warning: Could not set failure notification frequency for ${functionName}. Error: ${e.message}. Creating trigger without notification setting.`);
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyMinutes(5)
        .create();
      Logger.log(`Successfully created 5-minute trigger for ${functionName} (without error notifications).`);
      createdCount++;
    }
  });

  try {
    ScriptApp.newTrigger('triggered_makeTagIndex')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .withFailureNotificationFrequency(ScriptApp.FailureNotificationFrequency.HOURLY)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_makeTagIndex with hourly error notifications.`);
    createdCount++;
  } catch (e) {
    Logger.log(`Warning: Could not set failure notification frequency for triggered_makeTagIndex. Error: ${e.message}. Creating trigger without notification setting.`);
    ScriptApp.newTrigger('triggered_makeTagIndex')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_makeTagIndex (without error notifications).`);
    createdCount++;
  }

  try {
    ScriptApp.newTrigger('triggered_dailyReport')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .withFailureNotificationFrequency(ScriptApp.FailureNotificationFrequency.HOURLY)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_dailyReport with hourly error notifications.`);
    createdCount++;
  } catch (e) {
    Logger.log(`Warning: Could not set failure notification frequency for triggered_dailyReport. Error: ${e.message}. Creating trigger without notification setting.`);
    ScriptApp.newTrigger('triggered_dailyReport')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_dailyReport (without error notifications).`);
    createdCount++;
  }

  try {
    ScriptApp.newTrigger('triggered_backupSetupFiles')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .withFailureNotificationFrequency(ScriptApp.FailureNotificationFrequency.HOURLY)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_backupSetupFiles with hourly error notifications.`);
    createdCount++;
  } catch (e) {
    Logger.log(`Warning: Could not set failure notification frequency for triggered_backupSetupFiles. Error: ${e.message}. Creating trigger without notification setting.`);
    ScriptApp.newTrigger('triggered_backupSetupFiles')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_backupSetupFiles (without error notifications).`);
    createdCount++;
  }
  
  // After setting triggers, rename and move this script file for organization.
  try {
    Logger.log("Starting post-setup file cleanup...");
    const preferences = loadPreferences();
    if (preferences && preferences.notesFolder) {
      const notesFolderName = preferences.notesFolder;
      const systemFolderName = PREFERENCE_SUB_FOLDER_NAME;

      const parentFolders = DriveApp.getFoldersByName(notesFolderName);
      if (parentFolders.hasNext()) {
        const parentFolder = parentFolders.next();
        const systemFolders = parentFolder.getFoldersByName(systemFolderName);

        if (systemFolders.hasNext()) {
          const targetFolder = systemFolders.next();
          const filesToMove = DriveApp.getFilesByName("Copy of Trigger Code");
          
          if (filesToMove.hasNext()) {
            const file = filesToMove.next();
            file.setName("Trigger Code");
            targetFolder.addFile(file);
            DriveApp.getRootFolder().removeFile(file); // Assumes the file copy is in the root folder
            Logger.log(`Successfully renamed and moved "Trigger Code" to the system folder.`);
          } else {
            Logger.log('Note: "Copy of Trigger Code" not found. Skipping file organization step.');
          }
        } else {
          Logger.log(`Warning: System folder "${systemFolderName}" not found. Cannot move trigger code file.`);
        }
      } else {
        Logger.log(`Warning: Notes folder "${notesFolderName}" not found. Cannot move trigger code file.`);
      }
    }
  } catch(e) {
      Logger.log(`An error occurred during file cleanup: ${e.toString()}`);
      ui.alert(`Warning: Triggers were created, but the script file could not be automatically organized. Error: ${e.message}`);
  }

  ui.alert(`All ${createdCount} triggers have been set up and completed successfully.\n\nYou may now close this spreadsheet.`);
}


// --- TRIGGER-BASED FUNCTIONS (These run in the background) ---

/**
 * Processes emails based on preferences, generating Google Docs for each.
 */
function triggered_processEmailsToDoc() {
  if (!checkPreferencesFileExists()) { return; } 
  Logger.log("Starting email processing for QNS-Triggered.");
  
  const preferences = loadPreferences();
  if (!preferences) {
    Logger.log("Failed to load or parse preferences. Aborting email processing.");
    return;
  }

  const notesEmailAddress = preferences.notesEmail;
  const yourGmailAddress = preferences.yourGmailAddress;
  const rawNotesEmailLabel = preferences.gmailPendingFolder;
  const noteDocFolder = preferences.notesFolder;
  const notesEmailArchiveFolder = preferences.notesEmailArchiveFolder;
  
  if (!notesEmailAddress || !yourGmailAddress || !rawNotesEmailLabel || !noteDocFolder || !notesEmailArchiveFolder) {
    const missing = [];
    if (!notesEmailAddress) missing.push("notesEmail");
    if (!yourGmailAddress) missing.push("yourGmailAddress");
    if (!rawNotesEmailLabel) missing.push("gmailPendingFolder");
    if (!noteDocFolder) missing.push("notesFolder");
    if (!notesEmailArchiveFolder) missing.push("notesEmailArchiveFolder");
    Logger.log(`Missing one or more required preferences: ${missing.join(", ")}. Aborting.`);
    Logger.log(`Available preference keys loaded: [${Object.keys(preferences).join(", ")}]`);
    return;
  }
  
  // Default to "Attachments" if preference is missing, and ensure folder creation.
  const attachmentFolderName = preferences.attachmentFolder || "Attachments"; 
  const targetDocFolder = getOrCreateFolder(DriveApp, noteDocFolder);
  const targetAttachmentFolder = getOrCreateFolder(targetDocFolder, attachmentFolderName);
  
  const notesEmailLabelForSearch = rawNotesEmailLabel.replace(/ /g, '-');
  const pendingLabel = getOrCreateLabel(rawNotesEmailLabel);
  const archiveLabel = getOrCreateLabel('Notes Processed');

  if (!targetDocFolder || !targetAttachmentFolder || !pendingLabel || !archiveLabel) {
    Logger.log("Could not find or create a required folder or label. Aborting.");
    return;
  }

  const searchQuery = `from:${yourGmailAddress} to:${notesEmailAddress} label:${notesEmailLabelForSearch}`;
  Logger.log(`Final Gmail search query: "${searchQuery}"`);

  try {
    const threads = GmailApp.search(searchQuery);
    if (threads.length === 0) {
      Logger.log("No matching email threads found.");
      return;
    }

    Logger.log(`Found ${threads.length} email threads.`);
    threads.forEach(thread => {
      const message = thread.getMessages()[thread.getMessageCount() - 1];
      const sender = message.getFrom();
      const date = message.getDate();
      const subject = message.getSubject();
      const attachments = message.getAttachments();
      const messageBodyPlain = message.getPlainBody();
      const threadUrl = thread.getPermalink();

      const docTitle = `Email - ${subject.substring(0, 50)} - ${Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/YYYY HH:mm")}`;
      const newDoc = DocumentApp.create(docTitle);
      const docBody = newDoc.getBody();

      addStandardDocHeader(newDoc, date, subject, threadUrl, sender);
      
      docBody.appendParagraph(messageBodyPlain);
      docBody.appendParagraph("");

      const MAX_IMAGE_WIDTH = 500;
      if (attachments.length > 0) {
        const attachmentsHeading = docBody.appendParagraph("--- Attachments ---");
        const headingStyle = {};
        headingStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
        headingStyle[DocumentApp.Attribute.BOLD] = true;
        attachmentsHeading.setAttributes(headingStyle);

        attachments.forEach(attachment => {
          const attachmentName = attachment.getName();
          const attachmentType = attachment.getContentType();
          try {
            if (attachmentType.startsWith('image/')) {
              const blob = attachment.copyBlob();
              const inlineImage = docBody.appendImage(blob);
              if (inlineImage.getWidth() > MAX_IMAGE_WIDTH) {
                const newHeight = inlineImage.getHeight() * (MAX_IMAGE_WIDTH / inlineImage.getWidth());
                inlineImage.setWidth(MAX_IMAGE_WIDTH).setHeight(newHeight);
              }
              docBody.appendParagraph(`Embedded Image: ${attachmentName}`);
            } else {
              // Save to the ensured attachment folder
              const file = targetAttachmentFolder.createFile(attachment);
              const fileUrl = file.getUrl();
              const attachmentParagraph = docBody.appendParagraph(`Attachment: ${attachmentName} - ${fileUrl}`);
              
              const text = attachmentParagraph.editAsText();
              const textLen = text.getText().length;
              const urlLen = fileUrl.length;
              
              // Apply link only to the URL part at the end
              text.setLinkUrl(textLen - urlLen, textLen - 1, fileUrl);
            }
          } catch (attachError) {
            Logger.log(`ERROR: Failed to process attachment "${attachmentName}": ${attachError.toString()}`);
            docBody.appendParagraph(`Could not process attachment: ${attachmentName}. Error: ${attachError.message}`);
          }
        });
      } else {
        docBody.appendParagraph("No attachments found.");
      }

      addStandardDocFooter(newDoc);
      const file = DriveApp.getFileById(newDoc.getId());
      targetDocFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      Logger.log(`Created Google Doc for email "${subject}" at: ${newDoc.getUrl()}`);

      thread.addLabel(archiveLabel);
      thread.removeLabel(pendingLabel);
    });
    Logger.log("Email processing completed.");
  } catch (e) {
    Logger.log(`An error occurred during email processing: ${e.toString()}`);
  }
}

function triggered_convertTextNotesToDoc() {
  if (!checkPreferencesFileExists()) { return; } 
  Logger.log("Starting file processing for QNS-Triggered notes folder.");
  
  const preferences = loadPreferences();
  if (!preferences || !preferences.notesFolder) {
    Logger.log("Notes folder preference (notesFolder) is missing or preferences failed to load. Aborting.");
    if (preferences) {
      Logger.log(`Available preference keys loaded: [${Object.keys(preferences).join(", ")}]`);
    }
    return;
  }

  const targetFolder = getOrCreateFolder(DriveApp, preferences.notesFolder);
  if (!targetFolder) {
    Logger.log(`Could not find or create target notes folder. Aborting.`);
    return;
  }

  try {
    const files = targetFolder.getFilesByType('text/plain');
    let filesProcessed = 0;
    while (files.hasNext()) {
      const file = files.next();
      try {
        const fileContent = file.getBlob().getDataAsString();
        const fileName = file.getName();
        const fileUrl = file.getUrl();
        const fileDate = file.getDateCreated();
        
        const docTitle = `${fileName.replace(/\.txt$/i, '')} - ${Utilities.formatDate(fileDate, Session.getScriptTimeZone(), "MM/dd/YYYY HH:mm")}`;
        const newDoc = DocumentApp.create(docTitle);
        const docBody = newDoc.getBody();

        addStandardDocHeader(newDoc, fileDate, fileName, fileUrl, "File Conversion");
        const fullContentHeading = docBody.appendParagraph("--- Full File Content ---");
        
        const headingStyle = {};
        headingStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
        headingStyle[DocumentApp.Attribute.BOLD] = true;
        fullContentHeading.setAttributes(headingStyle);

        const contentParagraph = docBody.appendParagraph(fileContent);

        const textElement = contentParagraph.editAsText();
        const urlRegex = /(https?:\/\/[^\s]+)/g;
        const match = urlRegex.exec(fileContent);

        if (match && match[0]) {
          const linkUrl = match[0];
          textElement.setLinkUrl(0, textElement.getText().length - 1, linkUrl);
        }

        addStandardDocFooter(newDoc);
        const newDocFile = DriveApp.getFileById(newDoc.getId());
        targetFolder.addFile(newDocFile);
        DriveApp.getRootFolder().removeFile(newDocFile);
        Logger.log(`Created Google Doc from "${fileName}" at: ${newDoc.getUrl()}`);
        
        file.setTrashed(true);
        Logger.log(`Trashed original file: "${fileName}"`);

        filesProcessed++;
      } catch (fileProcessError) {
        Logger.log(`ERROR: Could not process file "${file.getName()}": ${fileProcessError.toString()}`);
      }
    }
    Logger.log(`File processing completed. ${filesProcessed} files converted.`);
  } catch (e) {
    Logger.log(`An error occurred during notes folder file processing: ${e.toString()}`);
  }
}

function triggered_MoveKeepNotes() {
  if (!checkPreferencesFileExists()) { return; }
  Logger.log("Starting triggered_MoveKeepNotes function.");
  
  const preferences = loadPreferences();
  if (!preferences || !preferences.notesFolder) {
    Logger.log("Notes folder preference (notesFolder) is missing or preferences failed to load. Aborting.");
    if (preferences) {
      Logger.log(`Available preference keys loaded: [${Object.keys(preferences).join(", ")}]`);
    }
    return;
  }

  const targetFolder = getOrCreateFolder(DriveApp, preferences.notesFolder);
  if (!targetFolder) {
    Logger.log(`Could not find or create target notes folder. Aborting.`);
    return;
  }

  const keepDocTitle = "Google Keep Document";
  try {
    const files = DriveApp.searchFiles(`title = '${keepDocTitle}' and 'root' in parents`);
    let movedCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.document') {
        try {
          targetFolder.addFile(file);
          DriveApp.getRootFolder().removeFile(file);
          Logger.log(`Moved "${file.getName()}" to folder "${preferences.notesFolder}".`);
          movedCount++;
        } catch (moveError) {
          Logger.log(`ERROR: Could not move file "${file.getName()}": ${moveError.toString()}`);
        }
      }
    }
    Logger.log(`triggered_MoveKeepNotes completed. ${movedCount} Google Keep Documents moved.`);
  } catch (e) {
    Logger.log(`An error occurred during triggered_MoveKeepNotes: ${e.toString()}`);
  }
}



// --- HELPER & DIAGNOSTIC FUNCTIONS ---

function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const count = allTriggers.length;
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log(count + ' triggers deleted.');
  return count;
}

function addStandardDocHeader(doc, date, title, url, source) {
  const docBody = doc.getBody();
  
  const titleParagraph = docBody.appendParagraph("Note Title : " + title);
  
  const style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  titleParagraph.setAttributes(style);

  if (url) {
    const urlParagraph = docBody.appendParagraph("URL: ");
    urlParagraph.appendText(url).setLinkUrl(url);

    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() == 200) {
        const htmlContent = response.getContentText();
        const titleMatch = htmlContent.match(/<title>(.*?)<\/title>/i);
        if (titleMatch && titleMatch[1]) {
          const pageTitle = titleMatch[1].trim();
          docBody.appendParagraph("Page Title: " + pageTitle);
        }
      } else {
        Logger.log(`Could not fetch page title from ${url}. Response code: ${response.getResponseCode()}`);
      }
    } catch (e) {
      Logger.log(`Error fetching URL for page title: ${e.toString()}`);
    }
  }

  docBody.appendParagraph(`Date: ${date.toLocaleString()}`);
  docBody.appendParagraph(`Source: ${source}`);
  docBody.appendParagraph("").appendHorizontalRule();
  docBody.appendParagraph("");
}

function addStandardDocFooter(doc) {
  Logger.log("addStandardDocFooter: Adding footer to document.");
  const docBody = doc.getBody();
  docBody.appendParagraph("");
  docBody.appendHorizontalRule();
  const footerParagraph = docBody.appendParagraph("Created with ");
  footerParagraph.appendText("Quick Notes Suite").setLinkUrl("https://sites.google.com/view/quick-notes-suite");
  footerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footerParagraph.setBold(true);
  Logger.log("addStandardDocFooter: Footer added.");
}

function checkPreferencesFileExists() {
  // This improved function directly checks if the preferences can be loaded.
  return loadPreferences() !== null;
}

function loadPreferences() {
  try {
    let ss = null;

    // 1. Try checking the active spreadsheet first, in case the script is container-bound to CapIt-Configuration or the trigger sheet
    try {
      const activeSS = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSS) {
        if (activeSS.getName() === PREFERENCE_SPREADSHEET_NAME || activeSS.getSheetByName(PREFERENCE_SHEET_TAB_NAME)) {
          ss = activeSS;
        }
      }
    } catch (e) {
      Logger.log("Could not access active spreadsheet or not container-bound: " + e.toString());
    }

    // 2. If not found or didn't have the Config tab, search Google Drive for the CapIt-Configuration spreadsheet
    if (!ss) {
      const files = DriveApp.getFilesByName(PREFERENCE_SPREADSHEET_NAME);
      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          try {
            const tempSS = SpreadsheetApp.open(file);
            if (tempSS.getSheetByName(PREFERENCE_SHEET_TAB_NAME)) {
              ss = tempSS;
              break;
            }
          } catch (openError) {
            Logger.log(`Failed to open file "${file.getName()}" as Spreadsheet: ${openError.toString()}`);
          }
        }
      }
    }

    if (!ss) {
      Logger.log(`Preferences Error: Spreadsheet "${PREFERENCE_SPREADSHEET_NAME}" with tab "${PREFERENCE_SHEET_TAB_NAME}" was not found.`);
      return null;
    }

    const configSheet = ss.getSheetByName(PREFERENCE_SHEET_TAB_NAME);
    if (!configSheet) {
      Logger.log(`Preferences Error: Tab "${PREFERENCE_SHEET_TAB_NAME}" was not found in "${PREFERENCE_SPREADSHEET_NAME}" spreadsheet.`);
      return null;
    }

    const range = configSheet.getDataRange();
    const values = range.getValues();
    const preferences = {};

    for (let i = 0; i < values.length; i++) {
      const key = values[i][0];
      const val = values[i][1];
      if (key !== null && key !== undefined) {
        const trimmedKey = String(key).trim();
        // Skip potential headers or empty keys
        if (trimmedKey && trimmedKey.toLowerCase() !== "parameter" && trimmedKey.toLowerCase() !== "key" && trimmedKey.toLowerCase() !== "property") {
          let processedVal = val;
          if (typeof val === 'string') {
            const trimmedVal = val.trim();
            if (trimmedVal.toLowerCase() === "true") {
              processedVal = true;
            } else if (trimmedVal.toLowerCase() === "false") {
              processedVal = false;
            } else {
              processedVal = trimmedVal;
            }
          }
          preferences[trimmedKey] = processedVal;

          // Normalize keys to camelCase (e.g. "Notes Email" -> "notesEmail") for maximum robustness
          const camelCaseKey = trimmedKey
            .replace(/[^a-zA-Z0-9\s-_]/g, '')
            .replace(/[\s-_]+(.)/g, (_, c) => c.toUpperCase())
            .replace(/^(.)/, c => c.toLowerCase());
          if (camelCaseKey && camelCaseKey !== trimmedKey) {
            preferences[camelCaseKey] = processedVal;
          }
        }
      }
    }

    if (Object.keys(preferences).length === 0) {
      Logger.log(`Preferences Error: No configuration entries found in tab "${PREFERENCE_SHEET_TAB_NAME}".`);
      return null;
    }

    return preferences;
  } catch (e) {
    Logger.log(`A critical error occurred while loading preferences from sheet: ${e.toString()}`);
    return null;
  }
}

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parent.createFolder(name);
  }
}

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/**
 * A diagnostic function to list all plain text (.txt) files in the notes folder.
 * This is intended to be run manually from the script editor for debugging.
 */
function listTextFilesInNotesFolder() {
  Logger.log("--- Starting Diagnostic: Listing Text Files ---");

  const preferences = loadPreferences();
  if (!preferences || !preferences.notesFolder) {
    Logger.log("❌ ABORTING: 'notesFolder' setting is missing from your preferences file or preferences could not be loaded.");
    return;
  }
  const notesFolderName = preferences.notesFolder;
  Logger.log(`🔎 Preference loaded. Will search in folder named: "${notesFolderName}"`);

  const folders = DriveApp.getFoldersByName(notesFolderName);
  if (!folders.hasNext()) {
    Logger.log(`❌ ERROR: Could not find any folder with the name "${notesFolderName}".`);
    Logger.log("Please make sure the folder name in your preferences file matches the folder in Drive exactly.");
    return;
  }
  const targetFolder = folders.next();
  Logger.log(`✅ Successfully found folder: "${targetFolder.getName()}"`);

  const files = targetFolder.getFilesByType(MimeType.PLAIN_TEXT);

  if (!files.hasNext()) {
    Logger.log("📂 No text (.txt) files were found in this folder.");
    return;
  }

  let fileCount = 0;
  Logger.log("--- Found Files ---");
  while (files.hasNext()) {
    const file = files.next();
    Logger.log(`📄 Name: "${file.getName()}"`);
    fileCount++;
  }
  Logger.log("-------------------");
  Logger.log(`✅ Diagnostic complete. Found ${fileCount} text file(s).`);
}


/**
 * Recursively scans a folder and its subfolders for files containing the tag marker in their title.
 * Excludes folders listed in excludeFolderNames.
 */
function getFilesRecursive(folder, taggedFiles, excludeFolderNames, marker) {
  const folderName = folder.getName();
  if (excludeFolderNames.indexOf(folderName) !== -1) {
    return;
  }

  const escapedMarker = marker.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  const tagRegex = new RegExp(escapedMarker + '(\\w+)', 'g');

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const title = file.getName();
    if (title.includes(marker)) {
      const tags = title.match(tagRegex);
      if (tags) {
        Logger.log(`Found tagged file: "${title}" in folder "${folderName}" with tags: ${tags.join(', ')}`);
        tags.forEach(tag => {
          const tagName = tag.substring(marker.length).toUpperCase();
          if (!taggedFiles[tagName]) {
            taggedFiles[tagName] = [];
          }
          if (!taggedFiles[tagName].some(f => f.url === file.getUrl())) {
            taggedFiles[tagName].push({title: title, url: file.getUrl()});
          }
        });
      }
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    getFilesRecursive(subfolders.next(), taggedFiles, excludeFolderNames, marker);
  }
}

function triggered_makeTagIndex() {
  const startTime = new Date();
  
  const preferences = loadPreferences();
  if (!preferences || !preferences.notesFolder) {
    Logger.log("Notes folder preference (notesFolder) is missing or preferences failed to load. Aborting triggered_makeTagIndex.");
    if (preferences) {
      Logger.log(`Available preference keys loaded: [${Object.keys(preferences).join(", ")}]`);
    }
    return;
  }
  const folderName = preferences.notesFolder;
  const folders = DriveApp.getFoldersByName(folderName);

  if (!folders.hasNext()) {
    Logger.log(`Folder '${folderName}' not found.`);
    return;
  }

  const notesFolder = folders.next();
  Logger.log(`Target notes folder: "${notesFolder.getName()}" (ID: ${notesFolder.getId()})`);
  
  const marker = preferences.indexTagMarker || '##';
  const taggedFiles = {};
  const excludeFolders = [
    preferences.dailyReports || "Daily Reports",
    preferences.logFolder || "Logs",
    "System Files - DO NOT DELETE OR EDIT"
  ];

  getFilesRecursive(notesFolder, taggedFiles, excludeFolders, marker);

  const sortedTags = Object.keys(taggedFiles).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  if (sortedTags.length === 0) {
    Logger.log("No tagged files found.");
    return;
  }

  const docName = 'Tag Index';
  const existingFiles = notesFolder.getFilesByName(docName);
  let doc;

  if (existingFiles.hasNext()) {
    const file = existingFiles.next();
    doc = DocumentApp.openById(file.getId());
    // Trash other files with the same name in the same folder
    while(existingFiles.hasNext()){
      existingFiles.next().setTrashed(true);
    }
    Logger.log(`Opened existing file: "${docName}"`);
  } else {
    doc = DocumentApp.create(docName);
    const file = DriveApp.getFileById(doc.getId());
    notesFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    Logger.log(`Created new file: "${docName}"`);
  }
  
  // --- FIX START ---
  const body = doc.getBody();
  
  // 1. Robust Clear: Append safe paragraph, delete old content.
  const originalBodyChildren = body.getNumChildren();
  body.appendParagraph("");
  for (let i = 0; i < originalBodyChildren; i++) {
    body.getChild(0).removeFromParent();
  }
  
  // 2. Reuse that existing empty paragraph for the Title.
  const titleParagraph = body.getParagraphs()[0];
  titleParagraph.setText('Tag Index');
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
  body.appendHorizontalRule();

  // Footer Logic
  let footer = doc.getFooter();
  if (!footer) {
    footer = doc.addFooter();
  }
  
  // 1. Robust Clear: Append safe paragraph, delete old content.
  const originalFooterChildren = footer.getNumChildren();
  footer.appendParagraph("");
  for (let i = 0; i < originalFooterChildren; i++) {
    footer.getChild(0).removeFromParent();
  }
  
  // 2. Reuse that existing empty paragraph for the text.
  const footerTextPara = footer.getParagraphs()[0];
  footerTextPara.setText('Generated by Quick Note Suite');
  footerTextPara.setLinkUrl("https://sites.google.com/view/quick-notes-suite/home");
  footerTextPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  // 3. Insert the Horizontal Rule BEFORE the text paragraph (at index 0).
  footer.insertHorizontalRule(0);
  
  // --- FIX END ---

  const startTimeText = 'Index Created at : ' + Utilities.formatDate(startTime, Session.getScriptTimeZone(), "MM/dd/yyyy, h:mm:ss a");
  body.appendParagraph(startTimeText);

  const endTime = new Date();
  const duration = (endTime.getTime() - startTime.getTime()) / 1000;
  const durationText = 'Duration of Run: ' + duration + ' seconds';
  body.appendParagraph(durationText);

  sortedTags.forEach(tag => {
    body.appendParagraph(tag).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    taggedFiles[tag].forEach(fileInfo => {
      const listItem = body.appendListItem(fileInfo.title);
      listItem.setLinkUrl(fileInfo.url);
    });
  });

  doc.saveAndClose();
}

/**
 * Trigger wrapper for Daily Report.
 * Determines the target date based on current time (before 6 AM = yesterday, else today).
 */
function triggered_dailyReport() {
  const now = new Date();
  const currentHour = now.getHours();
  let targetDate = new Date(now);

  // If before 6 AM, Set target to Yesterday
  if (currentHour < 6) {
    targetDate.setDate(now.getDate() - 1);
  }
  // Otherwise, target remains Today

  // Set start to 00:00:00
  targetDate.setHours(0, 0, 0, 0);
  const startDateTime = new Date(targetDate);

  // Set end to 23:59:59
  targetDate.setHours(23, 59, 59, 999);
  const endDateTime = new Date(targetDate);

  Logger.log(`Running Daily Report for date: ${startDateTime.toLocaleDateString()} (Window: ${startDateTime.toLocaleString()} - ${endDateTime.toLocaleString()})`);
  
  createDailyReport(startDateTime, endDateTime);
  
  Logger.log("Creating/updating Tag Index...");
  triggered_makeTagIndex();
}

/**
 * Trigger wrapper for backing up setup files.
 * Copies all system files (except the "Trigger Code" spreadsheet)
 * into a folder in the "Capit System Backups" folder in root Drive.
 */
function triggered_backupSetupFiles() {
  if (!checkPreferencesFileExists()) {
    Logger.log("Preferences file not found. Skipping backup.");
    return;
  }
  
  Logger.log("Starting triggered_backupSetupFiles...");
  
  try {
    const subFolderName = PREFERENCE_SUB_FOLDER_NAME;
    const systemFolders = DriveApp.getFoldersByName(subFolderName);
    let systemFolderCount = 0;
    let systemFolder = null;
    while (systemFolders.hasNext()) {
      systemFolder = systemFolders.next();
      systemFolderCount++;
    }
    
    if (systemFolderCount === 0) {
      Logger.log(`Backup Error: The system folder "${subFolderName}" was not found.`);
      return;
    }
    if (systemFolderCount > 1) {
      Logger.log(`Backup Error: The system folder "${subFolderName}" exists more than once. Aborting backup.`);
      return;
    }
    
    const files = systemFolder.getFiles();
    const filesToBackup = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().trim() === "Trigger Code") {
        Logger.log("Skipping backup of the 'Trigger Code' spreadsheet.");
        continue;
      }
      filesToBackup.push(file);
    }
    
    if (filesToBackup.length === 0) {
      Logger.log("No files found in the system folder to backup.");
      return;
    }
    
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const backupFolderName = "Capit System Backup-" + formattedDate;
    
    // Get or create "Capit System Backups" in the root of Google Drive
    const rootFolder = DriveApp.getRootFolder();
    const backupFolders = rootFolder.getFoldersByName("Capit System Backups");
    let backupFolder;
    if (backupFolders.hasNext()) {
      backupFolder = backupFolders.next();
    } else {
      backupFolder = rootFolder.createFolder("Capit System Backups");
      Logger.log("Created folder 'Capit System Backups' in root.");
    }
    
    const targetFolder = backupFolder.createFolder(backupFolderName);
    for (const file of filesToBackup) {
      file.makeCopy(targetFolder);
    }
    
    Logger.log(`Successfully backed up system files to folder "${backupFolderName}" in folder "Capit System Backups"`);
  } catch (e) {
    Logger.log(`An error occurred during triggered_backupSetupFiles: ${e.toString()}`);
  }
}

/**
 * Generates a Daily Report Doc listing files created and modified in the given range.
 * Time [HH:mm] is plain text and not part of the clickable link.
 */
function createDailyReport(startDateTime, endDateTime) {
  // Default to Today if no arguments provided (e.g. manual run)
  if (!startDateTime || !endDateTime) {
    const now = new Date();
    startDateTime = new Date(now);
    startDateTime.setHours(0, 0, 0, 0);
    endDateTime = new Date(now);
    endDateTime.setHours(23, 59, 59, 999);
    Logger.log("Manual execution detected. Defaulting to Today: " + startDateTime.toLocaleDateString());
  }

  const startTime = new Date(); 
  
  const preferences = loadPreferences();
  if (!preferences || !preferences.notesFolder) {
    Logger.log("Notes folder preference (notesFolder) is missing or preferences failed to load. Aborting Daily Report.");
    if (preferences) {
      Logger.log(`Available preference keys loaded: [${Object.keys(preferences).join(", ")}]`);
    }
    return;
  }
  
  const notesFolderName = preferences.notesFolder;
  const folders = DriveApp.getFoldersByName(notesFolderName);
  if (!folders.hasNext()) {
    Logger.log(`Folder '${notesFolderName}' not found.`);
    return;
  }
  const notesFolder = folders.next();
  
  // Daily Reports sub-folder
  const reportFolderName = preferences.dailyReports || "Daily Reports";
  const reportFolder = getOrCreateFolder(notesFolder, reportFolderName);
  
  const yyyy = startDateTime.getFullYear();
  const mm = String(startDateTime.getMonth() + 1).padStart(2, '0');
  const dd = String(startDateTime.getDate()).padStart(2, '0');
  const dateString = `${yyyy}-${mm}-${dd}`;
  const docTitle = `Daily Chronicle - ${dateString}`;
  
  const startStr = `${yyyy}-${mm}-${dd}T00:00:00Z`;
  const endStr = `${yyyy}-${mm}-${dd}T23:59:59Z`;
  
  // We collect chronological events in this array
  const events = [];
  const allTasks = [];
  
  // Resolve Log folder
  const logFolderName = preferences.logFolder || "Logs";
  const logFolders = DriveApp.getFoldersByName(logFolderName);
  let logFolderId = null;
  if (logFolders.hasNext()) {
    logFolderId = logFolders.next().getId();
  }
  
  const systemNames = ["TasksAndLists", "TasksAndLists.json", "Calendar Events", "Calendar Events.json", "Preferences", "Preferences.json", "LogCategories", "LogCategories.json"];
  
  // --- STEP 3: Tasks and Lists Parsing (Do first to collect allTasks for deep scan) ---
  const systemFolderName = PREFERENCE_SUB_FOLDER_NAME;
  const systemFolders = DriveApp.getFoldersByName(systemFolderName);
  if (systemFolders.hasNext()) {
    const systemFolder = systemFolders.next();
    
    let taskFile = null;
    let taskFiles = systemFolder.getFilesByName("TasksAndLists");
    if (taskFiles.hasNext()) {
      taskFile = taskFiles.next();
    } else {
      taskFiles = systemFolder.getFilesByName("TasksAndLists.json");
      if (taskFiles.hasNext()) {
        taskFile = taskFiles.next();
      }
    }
    
    if (taskFile) {
      try {
        const content = taskFile.getBlob().getDataAsString();
        const taskData = JSON.parse(content);
        const categories = taskData.categories || [];
        categories.forEach(cat => {
          const items = cat.items || [];
          items.forEach(item => {
            allTasks.push(item);
            
            if (item.date_created) {
              const d = new Date(item.date_created);
              if (d >= startDateTime && d <= endDateTime) {
                events.push({
                  timestamp: d,
                  text: `Task Created: ${item.text}`
                });
              }
            }
            if (item.date_completed) {
              const d = new Date(item.date_completed);
              if (d >= startDateTime && d <= endDateTime) {
                events.push({
                  timestamp: d,
                  text: `Task Completed: ${item.text}`
                });
              }
            }
            if (item.date_pending) {
              const d = new Date(item.date_pending);
              if (d >= startDateTime && d <= endDateTime) {
                events.push({
                  timestamp: d,
                  text: `Task Moved to Pending: ${item.text}`
                });
              }
            }
          });
        });
      } catch (e) {
        Logger.log(`Error parsing TasksAndLists: ${e.toString()}`);
      }
    }
    
    // --- STEP 4: Calendar Parsing ---
    let calFile = null;
    let calFiles = systemFolder.getFilesByName("Calendar Events");
    if (calFiles.hasNext()) {
      calFile = calFiles.next();
    } else {
      calFiles = systemFolder.getFilesByName("Calendar Events.json");
      if (calFiles.hasNext()) {
        calFile = calFiles.next();
      }
    }
    
    if (calFile) {
      try {
        const content = calFile.getBlob().getDataAsString();
        const calData = JSON.parse(content);
        const eventsArray = Array.isArray(calData) ? calData : (calData.events || []);
        eventsArray.forEach(evt => {
          const title = evt.summary || evt.title || evt.text || evt.name || evt.subject || "Unnamed Event";
          
          if (evt.created_at || evt.created) {
            const d = new Date(evt.created_at || evt.created);
            if (d >= startDateTime && d <= endDateTime) {
              events.push({
                timestamp: d,
                text: `Calendar Event Created: ${title}`
              });
            }
          }
          if (evt.completed_at || evt.completed) {
            const d = new Date(evt.completed_at || evt.completed);
            if (d >= startDateTime && d <= endDateTime) {
              events.push({
                timestamp: d,
                text: `Calendar Event Completed: ${title}`
              });
            }
          }
          if (evt.cancelled_at || evt.cancelled) {
            const d = new Date(evt.cancelled_at || evt.cancelled);
            if (d >= startDateTime && d <= endDateTime) {
              events.push({
                timestamp: d,
                text: `Calendar Event Cancelled: ${title}`
              });
            }
          }
        });
      } catch (e) {
        Logger.log(`Error parsing Calendar Events system file: ${e.toString()}`);
      }
    }
  }
  
  // --- STEP 1: Document & Spreadsheet Creations ---
  const creationsQuery = `modifiedDate >= '${yyyy}-${mm}-${dd}T00:00:00' and trashed = false`;
  const creationsIter = DriveApp.searchFiles(creationsQuery);
  while (creationsIter.hasNext()) {
    const file = creationsIter.next();
    const createdDate = file.getDateCreated();
    if (createdDate < startDateTime || createdDate > endDateTime) continue;
    
    const fileName = file.getName();
    
    // Exclusions
    if (systemNames.indexOf(fileName) !== -1 || fileName === PREFERENCE_FILE_NAME || fileName === PREFERENCE_SPREADSHEET_NAME) continue;
    
    // Exclude if in log folder
    const parents = file.getParents();
    let inLogFolder = false;
    while (parents.hasNext()) {
      if (parents.next().getId() === logFolderId) {
        inLogFolder = true;
        break;
      }
    }
    if (inLogFolder) continue;
    
    let typeLabel = "File";
    const mime = file.getMimeType();
    if (mime === 'application/vnd.google-apps.document') {
      typeLabel = "Document";
    } else if (mime === 'application/vnd.google-apps.spreadsheet') {
      typeLabel = "Spreadsheet";
    }
    
    events.push({
      timestamp: file.getDateCreated(),
      text: `${typeLabel} Created: ${fileName}`,
      url: file.getUrl()
    });
  }
  
  // --- STEP 2: Spreadsheet Updates (Logging) ---
  if (logFolderId) {
    const logFolder = DriveApp.getFolderById(logFolderId);
    const sheetsIter = logFolder.getFilesByType('application/vnd.google-apps.spreadsheet');
    while (sheetsIter.hasNext()) {
      const sheetFile = sheetsIter.next();
      const updated = sheetFile.getLastUpdated();
      if (updated >= startDateTime && updated <= endDateTime) {
        const category = sheetFile.getName();
        try {
          const spreadsheet = SpreadsheetApp.openById(sheetFile.getId());
          const sheets = spreadsheet.getSheets();
          sheets.forEach(sh => {
            const range = sh.getDataRange();
            const values = range.getValues();
            values.forEach(row => {
              const parsed = parseDateTimeAndText(row);
              if (parsed) {
                const ts = parsed.timestamp;
                if (ts >= startDateTime && ts <= endDateTime) {
                  events.push({
                    timestamp: ts,
                    text: `Log:${category} - ${parsed.text}`
                  });
                }
              }
            });
          });
        } catch (e) {
          Logger.log(`Error parsing spreadsheet log ${category}: ${e.toString()}`);
        }
      }
    }
  }
  
  // --- STEP 5: Advanced Document Parsing (Deep Scan) ---
  const modifiedQuery = `mimeType = 'application/vnd.google-apps.document' and modifiedDate >= '${startStr}' and modifiedDate <= '${endStr}' and trashed = false`;
  const modifiedIter = DriveApp.searchFiles(modifiedQuery);
  while (modifiedIter.hasNext()) {
    const docFile = modifiedIter.next();
    const docName = docFile.getName();
    
    // Check if created today (already handled in Step 1)
    const createdDate = docFile.getDateCreated();
    if (createdDate >= startDateTime && createdDate <= endDateTime) {
      continue;
    }
    
    // Skip system files
    if (systemNames.indexOf(docName) !== -1 || docName === PREFERENCE_FILE_NAME || docName === PREFERENCE_SPREADSHEET_NAME) {
      continue;
    }
    
    // Skip if in log folder
    const parents = docFile.getParents();
    let inLogFolder = false;
    while (parents.hasNext()) {
      if (parents.next().getId() === logFolderId) {
        inLogFolder = true;
        break;
      }
    }
    if (inLogFolder) continue;
    
    try {
      const doc = DocumentApp.openById(docFile.getId());
      const docText = doc.getBody().getText();
      
      // 1. Location Notes
      if (docName.startsWith("Location-")) {
        const lines = docText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        const addressInfo = lines.join(", ");
        events.push({
          timestamp: docFile.getLastUpdated(),
          text: `Location Update (${docName}): ${addressInfo}`,
          url: docFile.getUrl()
        });
      }
      
      // 2. Linked Notes
      allTasks.forEach(task => {
        if (task.text && (docText.includes(task.id) || docText.includes(task.text))) {
          events.push({
            timestamp: docFile.getLastUpdated(),
            text: `Linked Note: Document "${docName}" is associated with Task "${task.text}"`,
            url: docFile.getUrl()
          });
        }
      });
      
      // 3. Timestamped Entries
      const dateRegex = /Date:\s*(\d{4})[-/](\d{2})[-/](\d{2})\s+(\d{2}):(\d{2})([^\n]*)/g;
      let match;
      while ((match = dateRegex.exec(docText)) !== null) {
        const year = parseInt(match[1], 10);
        const month = parseInt(match[2], 10) - 1;
        const day = parseInt(match[3], 10);
        const hour = parseInt(match[4], 10);
        const minute = parseInt(match[5], 10);
        const entryTime = new Date(year, month, day, hour, minute);
        
        if (entryTime >= startDateTime && entryTime <= endDateTime) {
          const entryText = match[6].trim() || "Notes Entry";
          events.push({
            timestamp: entryTime,
            text: `Doc Entry (${docName}): ${entryText}`,
            url: docFile.getUrl()
          });
        }
      }
    } catch (e) {
      Logger.log(`Error deep scanning doc ${docName}: ${e.toString()}`);
    }
  }
  
  // Sort all events chronologically (24-hour normalized)
  events.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());
  
  // --- Create Google Doc Report ---
  const doc = DocumentApp.create(docTitle);
  const file = DriveApp.getFileById(doc.getId());
  
  reportFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  const body = doc.getBody();
  
  // Style page background (light theme: #FFFFFF)
  body.setBackgroundColor("#FFFFFF");
  
  // Title Setup
  const titleParagraph = body.getParagraphs()[0];
  titleParagraph.setText(`Daily Chronicle: ${dateString}`);
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
  const titleText = titleParagraph.editAsText();
  titleText.setForegroundColor("#000000");
  titleText.setFontSize(24);
  titleText.setBold(true);
  
  // Horizontal Rule (styled as standard divider)
  const hr = body.appendHorizontalRule();
  
  // Time and duration stats
  const endTime = new Date();
  const duration = (endTime.getTime() - startTime.getTime()) / 1000;
  
  const statsPara = body.appendParagraph(`Report Generated: ${new Date().toLocaleString()} | Run Duration: ${duration}s`);
  const statsText = statsPara.editAsText();
  statsText.setForegroundColor("#666666");
  statsText.setFontSize(10);
  statsText.setItalic(true);
  
  body.appendParagraph(""); // Space
  
  // Append all events
  if (events.length === 0) {
    const emptyPara = body.appendParagraph("No activity recorded today.");
    const emptyText = emptyPara.editAsText();
    emptyText.setForegroundColor("#000000");
    emptyText.setFontSize(11);
  } else {
    events.forEach(evt => {
      const timeStr = Utilities.formatDate(evt.timestamp, Session.getScriptTimeZone(), "HH:mm");
      const p = body.appendParagraph("");
      
      // [HH:mm] in legible Green (#2E7D32)
      const timeRun = p.appendText(`[${timeStr}] `);
      timeRun.setForegroundColor("#2E7D32");
      timeRun.setBold(true);
      timeRun.setFontSize(11);
      
      // Description in Black (#000000)
      const textRun = p.appendText(evt.text);
      textRun.setForegroundColor("#000000");
      textRun.setFontSize(11);
      
      // Link in legible Blue (#1976D2)
      if (evt.url) {
        const linkRun = p.appendText(" ");
        const linkText = p.appendText("(Link)");
        linkText.setLinkUrl(evt.url);
        linkText.setForegroundColor("#1976D2");
        linkText.setUnderline(true);
        linkText.setFontSize(11);
      }
    });
  }
  
  // Footer
  body.appendParagraph("").appendHorizontalRule();
  const footerPara = body.appendParagraph("Generated by Quick Note Suite");
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const footerText = footerPara.editAsText();
  footerText.setLinkUrl("https://sites.google.com/view/quick-notes-suite/home");
  footerText.setForegroundColor("#888888");
  footerText.setFontSize(10);
  footerText.setBold(true);
  
  doc.saveAndClose();
}

/**
 * Robust datetime parser that parses rows of log spreadsheets.
 * Supports MM-dd-yyyy HH:mm:ss and variations with / or .
 */
function parseDateTimeAndText(row) {
  if (!row || row.length === 0) return null;
  let timestamp = null;
  
  for (let i = 0; i < row.length; i++) {
    let val = row[i];
    if (val instanceof Date) {
      if (!isNaN(val.getTime())) {
        timestamp = val;
        break;
      }
    } else if (typeof val === 'string' && val.trim() !== '') {
      const match = val.trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
      if (match) {
        const month = parseInt(match[1], 10) - 1;
        const day = parseInt(match[2], 10);
        const year = parseInt(match[3], 10);
        const hour = parseInt(match[4], 10);
        const minute = parseInt(match[5], 10);
        const second = parseInt(match[6], 10);
        const d = new Date(year, month, day, hour, minute, second);
        if (!isNaN(d.getTime())) {
          timestamp = d;
          break;
        }
      }
    }
  }

  if (timestamp) {
    let bodyParts = [];
    for (let i = 0; i < row.length; i++) {
      let val = row[i];
      if (val instanceof Date) continue;
      if (typeof val === 'string' && val.trim() !== '') {
        const match = val.trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
        if (match) continue;
        bodyParts.push(val);
      } else if (val !== null && val !== undefined && val !== '') {
        bodyParts.push(val.toString());
      }
    }
    const text = bodyParts.join(" - ");
    return { timestamp: timestamp, text: text };
  }
  return null;
}

/**
 * Simple trigger that runs when the spreadsheet is opened.
 * Reads preferences and shows the Notes folder in cell A4 and the Date Format in cell A5.
 */
function onOpen() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Set static preference metadata in A9 and A10
    sheet.getRange("A9").setValue(PREFERENCE_SPREADSHEET_NAME);
    sheet.getRange("A10").setValue(PREFERENCE_SHEET_TAB_NAME);

    const preferences = loadPreferences();
    if (preferences) {
      sheet.getRange("A4").setValue(preferences.notesFolder || "");
      sheet.getRange("A5").setValue(preferences.dateFormat || "");
    } else {
      sheet.getRange("A4").setValue("Error: Preferences file not found");
      sheet.getRange("A5").setValue("");
    }
  } catch (e) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet.getRange("A4").setValue("Error: " + e.message);
      sheet.getRange("A5").setValue("");
    } catch (sheetError) {
      Logger.log("Failed to write error to sheet: " + sheetError.toString());
    }
  }
}