// Google Apps Script: FeedbackMenu.gs

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Speech Feedback Tools');

  const activeSpeechType = getActiveSpeechType(); // from Code.gs
  const typeSettingMenu = ui.createMenu('Set Active Speech Type');
  const availableSpeechTypes = getAvailableSpeechTypes(); // from Code.gs

  console.log("onOpen: Available Speech Types for Menu: ", JSON.stringify(availableSpeechTypes));

  if (Object.keys(availableSpeechTypes).length > 0) {
    Object.keys(availableSpeechTypes).forEach(type => {
      const displayName = type.charAt(0).toUpperCase() + type.slice(1) +
                         (type === activeSpeechType ? ' ✓' : '');
      
      // For setActiveType, we can still try the dynamic global function,
      // as it's a simpler action and might be less prone to the timing issue.
      // If this also fails, setActiveType would also need the dialog approach.
      const functionNameForSetActive = `CALL_setActive_${type.replace(/\W/g, '_')}`; // Sanitize
      console.log(`onOpen: Creating setActive function: ${functionNameForSetActive} for type: ${type}`);
      
      globalThis[functionNameForSetActive] = () => { 
        const currentType = type; 
        console.log(`Executing ${functionNameForSetActive} to set active type to: ${currentType}`);
        const result = setActiveSpeechType(currentType);
        if (result.success) {
          SpreadsheetApp.getUi().alert('Success', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
          onOpen(); // Refresh menu
        } else {
          SpreadsheetApp.getUi().alert('Error', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
        }
      };
      typeSettingMenu.addItem(displayName, functionNameForSetActive);
    });
  } else {
    typeSettingMenu.addItem("No speech types configured", "noop_placeholder");
  }
  menu.addSubMenu(typeSettingMenu);
  menu.addSeparator();

  if (Object.keys(availableSpeechTypes).length > 0) {
    Object.keys(availableSpeechTypes).forEach(speechTypeKey => {
      const speechTypeDisplayName = speechTypeKey.charAt(0).toUpperCase() + speechTypeKey.slice(1);
      
      // Each menu item will now call a generic initiator function,
      // and we'll pass the speechType to that initiator to build the dialog.
      menu.addSubMenu(ui.createMenu(`${speechTypeDisplayName} Feedback`)
        .addItem('Send Feedback to All Students', `initiateSendAllFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`)
        .addItem('Send Feedback to Selected Student', `initiateSendSelectedFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`)
        .addItem('Preview Feedback for Selected Student', `initiatePreviewFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`));

      // Now, dynamically create these initiator functions globally
      const currentSpeechType = speechTypeKey; // Closure
      globalThis[`initiateSendAllFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`] = () => {
        showConfirmationDialog_SendAll(currentSpeechType);
      };
      globalThis[`initiateSendSelectedFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`] = () => {
        sendSelectedStudentFeedbackForType(currentSpeechType); // This already shows a dialog for student selection
      };
      globalThis[`initiatePreviewFeedbackProcess_${speechTypeKey.replace(/\W/g, '_')}`] = () => {
        previewSelectedStudentFeedbackForType(currentSpeechType); // This also shows a dialog
      };

    });
  } else {
    menu.addItem("Configure Speech Types in 'Templates'", "noop_placeholder");
  }
  
  menu.addSeparator();
  menu.addItem('Configure Email Settings (Teacher CC)', 'showEmailSettingsDialog');
  menu.addToUi();
  console.log("onOpen: Menu setup complete.");
}

// Define a global no-op function if used by menu items
globalThis["noop_placeholder"] = () => { /* Does nothing */ };

function doSendAllFeedback(speechType) {
  console.log("doSendAllFeedback called with: " + speechType);
  sendAllFeedbackEmailsForType(speechType);
}

function doSendSelectedFeedback(speechType) {
  console.log("doSendSelectedFeedback called with: " + speechType);
  sendSelectedStudentFeedbackForType(speechType);
}

function doPreviewFeedback(speechType) {
  console.log("doPreviewFeedback called with: " + speechType);
  previewSelectedStudentFeedbackForType(speechType);
}

function doSetActiveType(speechType) {
    console.log("doSetActiveType called with: " + speechType);
    const result = setActiveSpeechType(type); // 'type' should be 'speechType' here
    if (result.success) {
      SpreadsheetApp.getUi().alert('Success', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
      onOpen(); 
    } else {
      SpreadsheetApp.getUi().alert('Error', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * Sends feedback emails to all unique presenters for a specific speech type.
 * @param {string} speechType The speech type to process.
 */
function sendAllFeedbackEmailsForType(speechType) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Send All Feedback for ${speechType.toUpperCase()}`,
    `Are you sure you want to send feedback emails to ALL students evaluated for the "${speechType}" speech? This cannot be undone.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const presenters = getUniquePresenters(speechType); // Now takes speechType
  if (presenters.length === 0) {
    ui.alert('No Presenters', `No evaluation data found for any presenters for the "${speechType}" speech.`, ui.ButtonSet.OK);
    return;
  }

  let successCount = 0;
  let failureCount = 0;
  const failures = [];

  presenters.forEach(presenter => {
    try {
      // `sendFeedbackEmail` is in FeedbackProcessor.gs and now also takes speechType
      const result = sendFeedbackEmail(presenter, speechType); 
      if (result.success) {
        successCount++;
      } else {
        failureCount++;
        failures.push(`${presenter}: ${result.message}`);
      }
    } catch (error) {
      console.error(`Error sending email to ${presenter} for ${speechType}:`, error, error.stack);
      failureCount++;
      failures.push(`${presenter} (for ${speechType}): ${error.toString()}`);
    }
  });

  let summaryMessage = `For "${speechType}" speeches:\nSuccessfully sent ${successCount} feedback email(s).`;
  if (failureCount > 0) {
    summaryMessage += `\n${failureCount} email(s) failed to send.\n\nFailures:\n${failures.join('\n')}`;
    ui.alert('Completed with Errors', summaryMessage, ui.ButtonSet.OK);
  } else {
    ui.alert('Success', summaryMessage, ui.ButtonSet.OK);
  }
}

/**
 * Shows a dialog to select a student and send feedback for a specific speech type.
 * @param {string} speechType The speech type for which to send feedback.
 */
function sendSelectedStudentFeedbackForType(speechType) {
  const ui = SpreadsheetApp.getUi();
  const presenters = getUniquePresenters(speechType); // Takes speechType

  if (presenters.length === 0) {
    ui.alert('No Presenters', `No evaluation data found for any presenters for the "${speechType}" speech.`, ui.ButtonSet.OK);
    return;
  }

  const optionsHtml = presenters.map(p => `<option value="${encodeURIComponent(p)}">${p}</option>`).join('');
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top">
        <style> body {font-family: Arial, sans-serif; margin: 20px;} select, button {width: 100%; padding: 10px; margin-bottom:10px; border-radius: 4px; border: 1px solid #ccc;} button { background-color: #4285f4; color: white; cursor: pointer;} button:hover{background-color:#3367d6;} </style>
      </head>
      <body>
        <h3>Select Student for "${speechType}" Feedback</h3>
        <select id="presenterSelect">${optionsHtml}</select>
        <button onclick="processSelection()">Send Feedback</button>
        <script>
          function processSelection() {
            const selectedPresenter = decodeURIComponent(document.getElementById('presenterSelect').value);
            google.script.run
              .withSuccessHandler(function(result) {
                google.script.host.close();
                if(result.success) SpreadsheetApp.getUi().alert('Success', result.message);
                else SpreadsheetApp.getUi().alert('Error', result.message);
              })
              .withFailureHandler(function(err) {
                google.script.host.close();
                SpreadsheetApp.getUi().alert('Error', 'Failed to send: ' + err.message);
              })
              .sendFeedbackEmail(selectedPresenter, "${speechType}"); // Pass speechType
          }
        </script>
      </body>
    </html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(200);
  ui.showModalDialog(htmlOutput, `Send ${speechType} Feedback`);
}

/**
 * Shows a dialog to select a student and preview feedback for a specific speech type.
 * @param {string} speechType The speech type for which to preview feedback.
 */
function previewSelectedStudentFeedbackForType(speechType) {
  const ui = SpreadsheetApp.getUi();
  const presenters = getUniquePresenters(speechType); // Takes speechType

  if (presenters.length === 0) {
    ui.alert('No Presenters', `No evaluation data found for any presenters for the "${speechType}" speech.`, ui.ButtonSet.OK);
    return;
  }

  const optionsHtml = presenters.map(p => `<option value="${encodeURIComponent(p)}">${p}</option>`).join('');
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top">
         <style> body {font-family: Arial, sans-serif; margin: 20px;} select, button {width: 100%; padding: 10px; margin-bottom:10px; border-radius: 4px; border: 1px solid #ccc;} button { background-color: #4285f4; color: white; cursor: pointer;} button:hover{background-color:#3367d6;} </style>
      </head>
      <body>
        <h3>Select Student to Preview "${speechType}" Feedback</h3>
        <select id="presenterSelect">${optionsHtml}</select>
        <button onclick="processPreview()">Preview Feedback</button>
        <script>
          function processPreview() {
            const selectedPresenter = decodeURIComponent(document.getElementById('presenterSelect').value);
            // Close this small dialog before showing the larger preview dialog
            google.script.host.close(); 
            google.script.run
              .withFailureHandler(function(err) {
                SpreadsheetApp.getUi().alert('Error', 'Failed to generate preview: ' + err.message);
              })
              .showFeedbackPreview(selectedPresenter, "${speechType}"); // Pass speechType
          }
        </script>
      </body>
    </html>`;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(200);
  ui.showModalDialog(htmlOutput, `Preview ${speechType} Feedback`);
}


/**
 * Gets a list of unique presenter names from the evaluation data for a specific speech type.
 * @param {string} speechType The type of the speech.
 * @return {Array<string>} An array of unique presenter names, sorted.
 */
function getUniquePresenters(speechType) {
  const sheetName = getSheetNameForSpeechType(speechType); // From Code.gs
  if (!sheetName) {
    console.error(`Could not determine sheet name for speech type: ${speechType} in getUniquePresenters.`);
    return [];
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`Sheet "${sheetName}" not found for speech type "${speechType}" when getting unique presenters.`);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; // No data beyond header

  const headers = data[0].map(h => h.toString().trim());
  const presenterNameColIndex = headers.indexOf('PresenterName');

  if (presenterNameColIndex === -1) {
    console.error(`'PresenterName' column not found in sheet "${sheetName}" for speech type "${speechType}".`);
    return [];
  }

  const presenterNames = new Set();
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
    if (data[i][presenterNameColIndex] && data[i][presenterNameColIndex].toString().trim() !== "") {
      presenterNames.add(data[i][presenterNameColIndex].toString().trim());
    }
  }
  return Array.from(presenterNames).sort();
}

/**
 * Shows a dialog to configure email settings, specifically the teacher's CC email.
 */
function showEmailSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentTeacherEmail = getTeacherEmail(); // From FeedbackProcessor.gs (or could be Code.gs if preferred)

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top">
        <style>
          body {font-family:Arial,sans-serif; margin:20px;} label{display:block; margin-bottom:5px; font-weight:bold;}
          input[type='email']{width:calc(100% - 22px); padding:10px; margin-bottom:15px; border:1px solid #ccc; border-radius:4px;}
          .buttons{text-align:right;} button{padding:10px 15px; margin-left:10px; border-radius:4px; cursor:pointer; border:none;}
          .save{background-color:#4285f4; color:white;} .save:hover{background-color:#3367d6;}
          .cancel{background-color:#f1f1f1;} .cancel:hover{background-color:#e0e0e0;}
        </style>
      </head>
      <body>
        <h3>Email Settings</h3>
        <label for="teacherEmail">Teacher Email (CC on feedback summaries):</label>
        <input type="email" id="teacherEmail" value="${currentTeacherEmail || ''}">
        <div class="buttons">
          <button class="cancel" onclick="google.script.host.close()">Cancel</button>
          <button class="save" onclick="saveSettings()">Save Settings</button>
        </div>
        <script>
          function saveSettings() {
            const email = document.getElementById('teacherEmail').value;
            google.script.run
              .withSuccessHandler(function(result) {
                google.script.host.close();
                if(result.success) SpreadsheetApp.getUi().alert('Success', result.message);
                else SpreadsheetApp.getUi().alert('Error', result.message);
              })
              .withFailureHandler(function(err) {
                google.script.host.close();
                SpreadsheetApp.getUi().alert('Error', 'Failed to save settings: ' + err.message);
              })
              .saveTeacherEmailSetting(email); // New function name for clarity
          }
        </script>
      </body>
    </html>`;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(450).setHeight(230);
  ui.showModalDialog(htmlOutput, 'Configure Email Settings');
}

/**
 * Saves the teacher's email setting to the 'Index' sheet.
 * @param {string} teacherEmail The email address to save.
 * @return {Object} An object indicating success or failure.
 */
function saveTeacherEmailSetting(teacherEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = ss.getSheetByName('Index');
    if (!indexSheet) {
      // Optionally create the Index sheet if it doesn't exist, though it's fundamental
      // For now, assume it should exist if this setting is being changed.
      return { success: false, message: "'Index' sheet not found. Cannot save teacher email." };
    }
    // Teacher email is stored in C2 (Row 2, Column 3)
    indexSheet.getRange('C2').setValue(teacherEmail);
    
    // Update the global variable in Code.gs by reloading student data (which also loads teacherEmail)
    if (typeof loadStudentData === "function") {
        loadStudentData(); 
    } else {
        console.warn("loadStudentData function not found to refresh global teacherEmail variable.");
    }

    return { success: true, message: "Teacher email setting saved successfully." };
  } catch (error) {
    console.error("Error saving teacher email setting:", error);
    return { success: false, message: `Error saving setting: ${error.toString()}` };
  }
}

function showConfirmationDialog_SendAll(speechType) {
  const ui = SpreadsheetApp.getUi();
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top">
        <style>
          body {font-family: Arial, sans-serif; margin: 20px; padding: 10px; text-align: center;}
          h3 { margin-top: 0;}
          .buttons { margin-top: 20px;}
          button {padding: 10px 20px; margin: 0 10px; border-radius: 4px; cursor:pointer; border: none; font-size: 14px;}
          .confirm {background-color: #d9534f; color:white;} .confirm:hover{background-color:#c9302c;}
          .cancel {background-color: #f0f0f0;} .cancel:hover{background-color:#e0e0e0;}
        </style>
      </head>
      <body>
        <h3>Send All Feedback for "${speechType}"?</h3>
        <p>This will send summary emails to all evaluated students for the ${speechType} speech. This action cannot be undone.</p>
        <div class="buttons">
          <button class="cancel" onclick="google.script.host.close()">Cancel</button>
          <button class="confirm" onclick="executeSendAll()">Confirm & Send</button>
        </div>
        <script>
          function executeSendAll() {
            google.script.run
              .withSuccessHandler(function() { google.script.host.close(); /* Optionally show main UI success */ })
              .withFailureHandler(function(err) { 
                google.script.host.close(); 
                SpreadsheetApp.getUi().alert("Error during Send All", err.message || String(err));
              })
              .sendAllFeedbackEmailsForType("${speechType}"); // speechType is embedded here by server
          }
        </script>
      </body>
    </html>`;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(450).setHeight(230);
  ui.showModalDialog(htmlOutput, `Confirm: Send All ${speechType} Feedback`);
}

// Remove specific speech type functions like sendPersuasiveFeedbackToAll, etc.
// as they are now dynamically generated by onOpen and call generic handlers.
/*
function sendPersuasiveFeedbackToAll() { sendAllFeedbackEmailsForType('persuasive'); }
function sendCommencementFeedbackToAll() { sendAllFeedbackEmailsForType('commencement'); }
// etc. for selected and preview
*/