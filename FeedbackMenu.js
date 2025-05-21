// Update onOpen function to add speech type selection
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Get all available speech types
  const speechTypes = getAvailableSpeechTypes();
  
  // Create main menu
  const menu = ui.createMenu('Speech Feedback');
  
  // Add speech type selection submenu
  const activeSpeechType = getActiveSpeechType();
  
  const typeMenu = ui.createMenu('Set Active Speech Type');
  Object.keys(speechTypes).forEach(speechType => {
    // Show checkmark next to current active type
    const displayName = speechType.charAt(0).toUpperCase() + speechType.slice(1) + 
                       (speechType === activeSpeechType ? ' ✓' : '');
    
    typeMenu.addItem(displayName, `setActiveType_${speechType}`);
    
    // Create the function if it doesn't exist
    if (typeof this[`setActiveType_${speechType}`] !== 'function') {
      this[`setActiveType_${speechType}`] = function() {
        const result = setActiveSpeechType(speechType);
        if (result.success) {
          SpreadsheetApp.getUi().alert('Success', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
          // Refresh the menu to show the checkmark
          onOpen();
        } else {
          SpreadsheetApp.getUi().alert('Error', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
        }
      };
    }
  });
  
  menu.addSubMenu(typeMenu);
  menu.addSeparator();
  
  // Add feedback items for each speech type
  Object.keys(speechTypes).forEach(speechType => {
    const displayName = speechType.charAt(0).toUpperCase() + speechType.slice(1) + ' Speech';
    
    menu.addSubMenu(ui.createMenu(displayName)
        .addItem('Send Feedback to All Students', `sendFeedbackToAll_${speechType}`)
        .addItem('Send Feedback to Selected Student', `sendFeedbackToSelected_${speechType}`)
        .addItem('Preview Feedback for Selected Student', `previewFeedback_${speechType}`));
    
    // Create the respective functions if they don't exist
    if (typeof this[`sendFeedbackToAll_${speechType}`] !== 'function') {
      this[`sendFeedbackToAll_${speechType}`] = function() {
        sendAllFeedbackEmails(speechType);
      };
    }
    
    if (typeof this[`sendFeedbackToSelected_${speechType}`] !== 'function') {
      this[`sendFeedbackToSelected_${speechType}`] = function() {
        sendSelectedStudentFeedback(speechType);
      };
    }
    
    if (typeof this[`previewFeedback_${speechType}`] !== 'function') {
      this[`previewFeedback_${speechType}`] = function() {
        previewSelectedStudentFeedback(speechType);
      };
    }
  });
  
  menu.addSeparator()
    .addItem('Configure Email Settings', 'showEmailSettings')
    .addToUi();
}

// Send feedback to all students who have been evaluated
function sendAllFeedbackEmails(speechType) {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm action
  const response = ui.alert(
    'Send Feedback Emails',
    'Are you sure you want to send feedback emails to ALL students? This cannot be undone.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Get unique presenters
  const presenters = getUniquePresenters();
  
  if (presenters.length === 0) {
    ui.alert('No Presenters Found', 'No evaluation data was found.', ui.ButtonSet.OK);
    return;
  }
  
  // Track success/failure
  let successCount = 0;
  let failureCount = 0;
  const failures = [];
  
  // Process each presenter
  for (const presenter of presenters) {
    try {
      const result = sendFeedbackEmail(presenter);
      if (result.success) {
        successCount++;
      } else {
        failureCount++;
        failures.push(`${presenter}: ${result.message}`);
      }
    } catch (error) {
      console.error(`Error sending email to ${presenter}:`, error);
      failureCount++;
      failures.push(`${presenter}: ${error.toString()}`);
    }
  }
  
  // Show results
  if (failureCount === 0) {
    ui.alert(
      'Success',
      `Successfully sent ${successCount} feedback emails.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Completed with Errors',
      `Sent ${successCount} emails successfully.\n${failureCount} emails failed.\n\n${failures.join('\n')}`,
      ui.ButtonSet.OK
    );
  }
}

// Send feedback to a selected student
function sendSelectedStudentFeedback(speechType) {
  const ui = SpreadsheetApp.getUi();
  
  // Get unique presenters for the dropdown
  const presenters = getUniquePresenters();
  
  if (presenters.length === 0) {
    ui.alert('No Presenters Found', 'No evaluation data was found.', ui.ButtonSet.OK);
    return;
  }
  
  // Create HTML for the selection dialog
  const html = HtmlService.createHtmlOutput(
    `<!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
          }
          select {
            width: 100%;
            padding: 8px;
            margin-bottom: 20px;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
          }
          button {
            padding: 8px 16px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background-color: #3367d6;
          }
        </style>
      </head>
      <body>
        <h3>Select a Student</h3>
        <select id="presenterSelect">
          ${presenters.map(presenter => `<option value="${presenter}">${presenter}</option>`).join('')}
        </select>
        <div class="button-container">
          <button onclick="sendFeedback()">Send Feedback</button>
        </div>
        
        <script>
          function sendFeedback() {
            const presenter = document.getElementById('presenterSelect').value;
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  alert("Feedback email sent successfully!");
                } else {
                  alert("Error: " + result.message);
                }
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert("Error: " + error.message);
                google.script.host.close();
              })
              .sendFeedbackEmail(presenter);
          }
        </script>
      </body>
    </html>`
  )
  .setWidth(400)
  .setHeight(250);
  
  ui.showModalDialog(html, 'Send Feedback Email');
}

// Preview feedback for a selected student
function previewSelectedStudentFeedback(speechType) {
  const ui = SpreadsheetApp.getUi();
  
  // Get unique presenters for the dropdown
  const presenters = getUniquePresenters();
  
  if (presenters.length === 0) {
    ui.alert('No Presenters Found', 'No evaluation data was found.', ui.ButtonSet.OK);
    return;
  }
  
  // Create HTML for the selection dialog
  const html = HtmlService.createHtmlOutput(
    `<!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
          }
          select {
            width: 100%;
            padding: 8px;
            margin-bottom: 20px;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
          }
          button {
            padding: 8px 16px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background-color: #3367d6;
          }
        </style>
      </head>
      <body>
        <h3>Select a Student</h3>
        <select id="presenterSelect">
          ${presenters.map(presenter => `<option value="${presenter}">${presenter}</option>`).join('')}
        </select>
        <div class="button-container">
          <button onclick="previewFeedback()">Preview Feedback</button>
        </div>
        
        <script>
          function previewFeedback() {
            const presenter = document.getElementById('presenterSelect').value;
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert("Error: " + error.message);
                google.script.host.close();
              })
              .showFeedbackPreview(presenter);
          }
        </script>
      </body>
    </html>`
  )
  .setWidth(400)
  .setHeight(250);
  
  ui.showModalDialog(html, 'Preview Feedback Email');
}

// Get a list of unique presenters from the evaluation data
function getUniquePresenters(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Peer Evaluations');
  
  if (!sheet) {
    return [];
  }
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Skip the header row if it exists
  const headerRow = 0;
  
  // Extract unique presenter names (column C, index 2)
  const presenterNames = new Set();
  for (let i = headerRow + 1; i < data.length; i++) {
    if (data[i][2]) {
      presenterNames.add(data[i][2]);
    }
  }
  
  return Array.from(presenterNames).sort();
}

// Show email settings dialog
function showEmailSettings() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current teacher email
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName('Index');
  const teacherEmail = indexSheet.getRange('D2').getValue();
  
  // Create the settings dialog
  const html = HtmlService.createHtmlOutput(
    `<!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
          }
          input {
            width: 100%;
            padding: 8px;
            margin-bottom: 20px;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
          }
          button {
            padding: 8px 16px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background-color: #3367d6;
          }
          .cancel-button {
            background-color: #f1f1f1;
            color: #333;
          }
          .cancel-button:hover {
            background-color: #e4e4e4;
          }
        </style>
      </head>
      <body>
        <h3>Email Settings</h3>
        <label for="teacherEmail">Teacher Email (CC on all feedback emails):</label>
        <input type="email" id="teacherEmail" value="${teacherEmail || ''}">
        
        <div class="button-container">
          <button class="cancel-button" onclick="google.script.host.close()">Cancel</button>
          <button onclick="saveSettings()">Save Settings</button>
        </div>
        
        <script>
          function saveSettings() {
            const teacherEmail = document.getElementById('teacherEmail').value;
            
            google.script.run
              .withSuccessHandler(function() {
                alert("Settings saved successfully!");
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert("Error: " + error.message);
              })
              .saveEmailSettings(teacherEmail);
          }
        </script>
      </body>
    </html>`
  )
  .setWidth(400)
  .setHeight(250);
  
  ui.showModalDialog(html, 'Email Settings');
}

// Save email settings
function saveEmailSettings(teacherEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName('Index');
  
  // Update teacher email in cell D2
  indexSheet.getRange('C2').setValue(teacherEmail);
  
  // Update the global variable
  loadStudentData(); // This refreshes the teacherEmail variable
  
  return { success: true };
}

// Speech-specific functions
function sendPersuasiveFeedbackToAll() {
  sendAllFeedbackEmails('persuasive');
}

function sendCommencementFeedbackToAll() {
  sendAllFeedbackEmails('commencement');
}

function sendPersuasiveFeedbackToSelected() {
  sendSelectedStudentFeedback('persuasive');
}

function sendCommencementFeedbackToSelected() {
  sendSelectedStudentFeedback('commencement');
}

function previewPersuasiveFeedback() {
  previewSelectedStudentFeedback('persuasive');
}

function previewCommencementFeedback() {
  previewSelectedStudentFeedback('commencement');
}