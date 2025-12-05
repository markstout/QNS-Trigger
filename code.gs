/**
 * Quick Notes Suite - Trigger Code
 * Copyright 2025 Mark A. Stout
 * Last Updated : 11-29-2025
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

  deleteAllTriggers();
  ui.alert("Existing triggers have been removed.\nPlease wait while I create the new triggers.");

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
    } catch (e) {
      Logger.log(`Warning: Could not set failure notification frequency for ${functionName}. Error: ${e.message}. Creating trigger without notification setting.`);
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyMinutes(5)
        .create();
      Logger.log(`Successfully created 5-minute trigger for ${functionName} (without error notifications).`);
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
  } catch (e) {
    Logger.log(`Warning: Could not set failure notification frequency for triggered_makeTagIndex. Error: ${e.message}. Creating trigger without notification setting.`);
    ScriptApp.newTrigger('triggered_makeTagIndex')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .nearMinute(0)
      .create();
    Logger.log(`Successfully created daily 3 AM to 4 AM trigger for triggered_makeTagIndex (without error notifications).`);
  }
  
  // After setting triggers, rename and move this script file for organization.
  try {
    Logger.log("Starting post-setup file cleanup...");
    const preferences = loadPreferences();
    if (preferences && preferences.notesFolder) {
      const notesFolderName = preferences.notesFolder;
      const systemFolderName = 'System Files - DO NOT DELETE OR EDIT';

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

  ui.alert('All triggers have been set up successfully. They will now run automatically every 5 minutes.\n\nYou may now close this spreadsheet.');
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
  
  // Default to "Attachments" if preference is missing, and ensure folder creation.
  const attachmentFolderName = preferences.attachmentFolder || "Attachments"; 
  const targetDocFolder = getOrCreateFolder(DriveApp, noteDocFolder);
  const targetAttachmentFolder = getOrCreateFolder(targetDocFolder, attachmentFolderName);
  
  if (!notesEmailAddress || !yourGmailAddress || !rawNotesEmailLabel || !noteDocFolder || !notesEmailArchiveFolder) {
    Logger.log("Missing one or more required preferences. Aborting.");
    return;
  }
  
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
              const attachmentParagraph = docBody.appendParagraph(`Attachment: ${attachmentName} - `);
              attachmentParagraph.appendText(fileUrl).setLinkUrl(fileUrl);
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
    Logger.log("Notes folder preference is missing. Aborting.");
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
    Logger.log("Notes folder preference is missing. Aborting.");
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
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log(allTriggers.length + ' triggers deleted.');
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
  footerParagraph.appendText("QuickNoteSuite").setLinkUrl("https://sites.google.com/view/notesondrive/home?authuser=0");
  footerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footerParagraph.setBold(true);
  Logger.log("addStandardDocFooter: Footer added.");
}

function checkPreferencesFileExists() {
  // This improved function directly checks if the preferences can be loaded.
  return loadPreferences() !== null;
}

function loadPreferences() {
  const fileName = 'QuickNoteSuitePreferences-Do-not-Delete-or-Edit';
  const parentFolderName = 'Notes';
  const subFolderName = 'System Files - DO NOT DELETE OR EDIT';

  try {
    const parentFolders = DriveApp.getFoldersByName(parentFolderName);
    if (!parentFolders.hasNext()) {
      Logger.log(`Preferences Error: The primary folder "${parentFolderName}" was not found.`);
      return null;
    }

    const subFolders = parentFolders.next().getFoldersByName(subFolderName);
    if (!subFolders.hasNext()) {
      Logger.log(`Preferences Error: The sub-folder "${subFolderName}" was not found inside "${parentFolderName}".`);
      return null;
    }

    const preferenceFiles = subFolders.next().getFilesByName(fileName);
    if (preferenceFiles.hasNext()) {
      const file = preferenceFiles.next();
      const content = file.getBlob().getDataAsString();
      return JSON.parse(content);
    } else {
      Logger.log(`Preferences Error: The file "${fileName}" was not found inside "${parentFolderName}/${subFolderName}".`);
      return null;
    }
  } catch (e) {
    Logger.log(`A critical error occurred while loading preferences: ${e.toString()}`);
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


function triggered_makeTagIndex() {
  const startTime = new Date();
  const folderName = "Notes";
  const folders = DriveApp.getFoldersByName(folderName);

  if (!folders.hasNext()) {
    Logger.log(`Folder '${folderName}' not found.`);
    return;
  }

  const notesFolder = folders.next();
  const files = notesFolder.getFiles();
  const taggedFiles = {};

  while (files.hasNext()) {
    const file = files.next();
    const title = file.getName();

    if (title.includes('##')) {
      const tags = title.match(/##(\w+)/g);
      if (tags) {
        tags.forEach(tag => {
          const tagName = tag.substring(2);
          if (!taggedFiles[tagName]) {
            taggedFiles[tagName] = [];
          }
          taggedFiles[tagName].push({title: title, url: file.getUrl()});
        });
      }
    }
  }

  const sortedTags = Object.keys(taggedFiles).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  if (sortedTags.length === 0) {
    Logger.log("No tagged files found.");
    return;
  }

  const docName = 'Tags Index';
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
  
  // 1. Clear the body. This leaves exactly ONE empty paragraph.
  body.clear();
  
  // 2. Reuse that existing empty paragraph for the Title.
  const titleParagraph = body.getParagraphs()[0];
  titleParagraph.setText('Tags Index');
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
  body.appendHorizontalRule();

  // Footer Logic
  let footer = doc.getFooter();
  if (!footer) {
    footer = doc.addFooter();
  }
  
  // 1. Clear the footer. This leaves exactly ONE empty paragraph.
  footer.clear();
  
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