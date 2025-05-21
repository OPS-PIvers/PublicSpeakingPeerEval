// Google Apps Script to create a web app for peer speech evaluations
// This reads from "Index" tab and writes to "Peer Evaluations" tab

// Global variables
let studentData = [];
let teacherEmail = ''; // This will be populated by loadStudentData

// Main function to serve the web app HTML
function doGet(e) {
  // Load student data early
  loadStudentData(); 
  
  let speechType = e.parameter.type;
  if (!speechType) {
    speechType = getActiveSpeechType();
  }
  
  const template = HtmlService.createTemplateFromFile('Index');
  template.speechType = speechType;
  // MODIFICATION: Add script URL to the template
  template.scriptUrl = ScriptApp.getService().getUrl(); 
  
  return template.evaluate()
      .setTitle('Speech Peer Evaluation')
      .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/forms_48dp.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Function to get active speech type from Settings sheet
function getActiveSpeechType() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      // Create settings sheet if it doesn't exist
      const newSheet = ss.insertSheet('Settings');
      newSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
      newSheet.getRange('A1:B1').setFontWeight('bold'); // Make header bold
      newSheet.getRange('A2:B2').setValues([['ActiveSpeechType', 'persuasive']]); // Default
      SpreadsheetApp.flush(); // Ensure sheet is created before trying to read again if needed
      return 'persuasive'; 
    }
    
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
      if (data[i][0] === 'ActiveSpeechType') {
        return data[i][1] || 'persuasive'; // Default if value is empty
      }
    }
    
    // If not found, add it with default value
    const lastRow = settingsSheet.getLastRow();
    settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['ActiveSpeechType', 'persuasive']]);
    return 'persuasive';
  } catch (error) {
    console.error("Error getting active speech type:", error);
    return 'persuasive'; // Default to persuasive in case of error
  }
}

// Function to set active speech type
function setActiveSpeechType(speechType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
      settingsSheet.getRange('A1:B1').setFontWeight('bold');
    }
    
    const data = settingsSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
      if (data[i][0] === 'ActiveSpeechType') {
        settingsSheet.getRange(i + 1, 2).setValue(speechType);
        found = true;
        break;
      }
    }
    
    if (!found) {
      const lastRow = settingsSheet.getLastRow();
      settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['ActiveSpeechType', speechType]]);
    }
    return { success: true, message: `Active speech type set to "${speechType}"` };
  } catch (error) {
    console.error("Error setting active speech type:", error);
    return { success: false, message: "Error setting active speech type: " + error.toString() };
  }
}

// Load student data from the Index tab
function loadStudentData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = ss.getSheetByName('Index');
    if (!indexSheet) {
      console.error("Index sheet not found. Student data and teacher email cannot be loaded.");
      studentData = [];
      teacherEmail = getTeacherEmail(); // Fallback to default if Index sheet is missing
      return;
    }
  
    const data = indexSheet.getDataRange().getValues();
    studentData = []; // Reset before loading

    // Start from row 1 (index 1) to skip header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') { // Full Name in Column A
        studentData.push({
          fullName: data[i][0].toString().trim(), 
          email: (data[i][1] || '').toString().trim() // Email in Column B
        });
      }
    }
  
    studentData.sort((a, b) => a.fullName.localeCompare(b.fullName));
  
    // Get teacher email from cell C2 (row 2, column 3)
    if (data.length > 1 && data[1].length > 2 && data[1][2] && data[1][2].toString().trim() !== '') {
      teacherEmail = data[1][2].toString().trim();
    } else {
      // Attempt to get from getTeacherEmail as a fallback, which itself has a default
      teacherEmail = getTeacherEmail(); // Ensures it uses the constant if C2 is empty
      console.warn("Teacher email not found in Index sheet C2. Using fallback/default.");
    }
    console.log("Student data loaded. Count:", studentData.length, "Teacher Email:", teacherEmail);

  } catch (error) {
    console.error("Error loading student data:", error);
    studentData = []; // Ensure it's empty on error
    teacherEmail = DEFAULT_TEACHER_EMAIL; // Use defined default on error
  }
}


// Return student data to the client-side JavaScript
function getStudentData() {
  if (!studentData || studentData.length === 0) { // Ensure data is loaded if not already
      loadStudentData();
  }
  return studentData;
}

function processForm(formData) {
  try {
    console.log("Process form (for saving data) started with data:", JSON.stringify(formData));
    
    // Basic validation - can be enhanced by checking against speechConfiguration if needed
    if (!formData.speechType || !formData.evaluatorName || !formData.presenterName) {
        // 'evaluatorName' and 'presenterName' might not be universally required for all form types
        // but speechType is essential.
        // If 'evaluatorName' and 'presenterName' are dynamic, this check might be too strict here.
        // For now, keeping it as it was often a base requirement.
      return { success: false, message: "Missing critical fields (e.g., speech type, evaluator, presenter)." };
    }
    
    const saveResult = saveToSheet(formData); // Using the improved saveToSheet
    if (!saveResult.success) {
        return { success: false, message: saveResult.message || "Failed to save data to sheet."};
    }
    
    // DO NOT SEND INDIVIDUAL EMAIL HERE.
    // Emails with summaries/averages are sent via FeedbackMenu.gs -> FeedbackProcessor.gs
    console.log("Form data saved successfully. Summary emails are sent via menu options.");
    
    return { success: true, message: "Your evaluation has been submitted successfully!" };
  } catch (error) {
    console.error("Process form error:", error.toString(), error.stack);
    return { success: false, message: "Error processing form: " + error.toString() };
  }
}

/*
/**
 * Sends an email notification for an individual form submission.
 * This is different from the summary emails sent via FeedbackProcessor.gs.
 * @param {Object} formData The data from the submitted form.
 * @return {Object} An object indicating success or failure of email sending.
 
function sendIndividualSubmissionEmail(formData) {
  try {
    loadStudentData(); // Ensure studentData and teacherEmail global are fresh

    const speechConfiguration = getSpeechConfiguration(formData.speechType);
    if (speechConfiguration.error) {
      console.error("Cannot send individual email, speech configuration error:", speechConfiguration.error);
      return { success: false, message: "Speech configuration error for email." };
    }

    const presenterName = formData.presenterName;
    const evaluatorName = formData.evaluatorName;
    
    const presenterEmail = findPresenterEmail(presenterName);
    const currentTeacherEmail = getTeacherEmail(); // Use the robust getter

    if (!presenterEmail && !currentTeacherEmail) {
      console.warn("No recipient for individual submission email (presenter or teacher). Email not sent.");
      return { success: false, message: "No recipient for email." };
    }

    const subject = `New ${speechConfiguration.title || formData.speechType} Evaluation: ${presenterName} by ${evaluatorName}`;
    const htmlBody = createDynamicIndividualEmailBody(formData, speechConfiguration, 'html');
    // const textBody = createDynamicIndividualEmailBody(formData, speechConfiguration, 'text'); // For plain text version

    const mailOptions = {
      subject: subject,
      htmlBody: htmlBody
      // body: textBody // Uncomment if you want plain text part for multipart
    };

    let recipients = [];
    if (presenterEmail) recipients.push(presenterEmail);
    if (currentTeacherEmail) recipients.push(currentTeacherEmail); // Send to teacher as well or CC
    
    // De-duplicate recipients (e.g., if teacher is evaluating themselves, or presenter IS teacher)
    recipients = [...new Set(recipients)]; 

    if (recipients.length === 0) {
        console.warn("No valid recipients for the individual submission email.");
        return { success: false, message: "No valid email recipients."};
    }

    mailOptions.to = recipients.join(',');

    MailApp.sendEmail(mailOptions);
    console.log(`Individual submission email sent for ${presenterName} to: ${mailOptions.to}`);
    return { success: true, message: "Notification email sent." };

  } catch (e) {
    console.error("Error sending individual submission email:", e.toString(), e.stack);
    return { success: false, message: `Error sending email: ${e.toString()}` };
  }
}

/**
 * Creates a dynamic email body (HTML or Text) for an individual submission.
 * @param {Object} formData The submitted form data.
 * @param {Object} speechConfig The speech configuration object.
 * @param {string} format 'html' or 'text'.
 * @return {string} The generated email body.

function createDynamicIndividualEmailBody(formData, speechConfig, format = 'html') {
  let body = '';
  const nl = format === 'html' ? '<br>' : '\n';
  const h1Open = format === 'html' ? '<h1 style="color: #1a73e8;">' : '== ';
  const h1Close = format === 'html' ? '</h1>' : ' ==';
  const h2Open = format === 'html' ? '<h2 style="color: #333; border-bottom: 1px solid #eee; padding-bottom: 5px;">' : '-- ';
  const h2Close = format === 'html' ? '</h2>' : ' --';
  const strongOpen = format === 'html' ? '<strong>' : '';
  const strongClose = format === 'html' ? '</strong>' : '';
  const pOpen = format === 'html' ? '<p>' : '';
  const pClose = format === 'html' ? '</p>' : '';
  const divOpen = format === 'html' ? '<div>' : '';
  const divClose = format === 'html' ? '</div>' : '';
  const commentOpen = format === 'html' ? '<div style="margin-left: 15px; padding: 5px; background-color: #f9f9f9; border-left: 2px solid #ccc;"><em>' : '> ';
  const commentClose = format === 'html' ? '</em></div>' : '';


  if (format === 'html') {
    body += '<div style="font-family: Arial, sans-serif; max-width: 700px; margin: auto; padding: 15px; border: 1px solid #ddd;">';
  }

  body += `${h1Open}New ${speechConfig.title || formData.speechType} Evaluation Submitted${h1Close}${nl}${nl}`;
  body += `${pOpen}${strongOpen}Speech/Presenter:${strongClose} ${formData.presenterName || 'N/A'}${pClose}`;
  body += `${pOpen}${strongOpen}Evaluated By:${strongClose} ${formData.evaluatorName || 'N/A'}${pClose}`;
  body += `${pOpen}${strongOpen}Timestamp:${strongClose} ${new Date().toLocaleString()}${pClose}${nl}`;

  speechConfig.sections.forEach(section => {
    // Skip the "Review Your Evaluation" section in the email body
    if (section.title && section.title.toLowerCase().includes('review')) {
      return; 
    }

    body += `${h2Open}${section.title}${h2Close}`;
    section.questions.forEach(question => {
      const questionId = question.id;
      const questionText = question.text;
      let value = formData[questionId] || "N/A";

      // Special formatting for certain types
      if (question.type.toLowerCase() === 'checkbox' && value !== "N/A") {
        try {
          const parsedValue = JSON.parse(value);
          value = Array.isArray(parsedValue) ? parsedValue.join(', ') : value;
          if (value === "") value = "None selected";
        } catch (e) { / Ignore parsing error, use raw value  }
      } else if (question.type.toLowerCase() === 'rubric' && value !== "N/A") {
         value = `${value} / ${question.maxScore || 5}`; // Assume 5 if maxScore not in config for this question
      } else if (question.type.toLowerCase() === 'comment' && (value === "N/A" || value.trim() === "" || value.trim() === "No comments provided")) {
         value = "No comments provided."; // Standardize empty comment
      }
      
      if (question.type.toLowerCase() === 'comment') {
        body += `${divOpen}${strongOpen}${questionText}:${strongClose}${divClose}`;
        body += `${commentOpen}${value}${commentClose}${nl}`;
      } else {
        body += `${divOpen}${strongOpen}${questionText}:${strongClose} ${value}${divClose}${nl}`;
      }
    });
    body += nl;
  });

  if (format === 'html') {
    body += '<p style="font-size:0.9em; color:#777;">This is an automated notification.</p>';
    body += '</div>'; // Close main container
  }
  return body;
}


// OLD sendEmails function and its helpers are now effectively replaced by sendIndividualSubmissionEmail
// and createDynamicIndividualEmailBody if immediate individual emails are desired.
// The summary emails are handled by FeedbackProcessor.gs
/*
// Updated sendEmails function 
function sendEmails(formData) { 
  // This function is problematic if it's trying to use the OLD hardcoded email bodies.
  // It should call a dynamic email body generator.
  // For now, commenting out its content and relying on the new sendIndividualSubmissionEmail
  console.warn("Old 'sendEmails' function called. Consider using 'sendIndividualSubmissionEmail'.");
}

// Create the HTML email body
function createHtmlEmailBody(formData) { // OLD AND HARDCODED
  console.error("Deprecated createHtmlEmailBody called. This function is hardcoded.");
  // ... old hardcoded HTML ...
  return "This email body is from a deprecated function and likely incorrect.";
}

// Create plain text email body (fallback)
function createEmailBody(formData) { // OLD AND HARDCODED
  console.error("Deprecated createEmailBody called. This function is hardcoded.");
  // ... old hardcoded text ...
  return "This email body is from a deprecated function and likely incorrect.";
}
*/


// Updated saveToSheet function to handle speech-specific tabs

function saveToSheet(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Could not access the active spreadsheet");

    const targetSheetName = getSheetNameForSpeechType(formData.speechType); // Ensure this is correct casing from config
    if (!targetSheetName) {
      console.error(`Could not determine sheet name for speech type: ${formData.speechType}. Data not saved.`);
      return { success: false, message: `Configuration error: Sheet name not found for ${formData.speechType}.` };
    }

    let sheet = ss.getSheetByName(targetSheetName);
    const timestamp = new Date();
    const safeGet = (value, defaultValue = '') => (value !== undefined && value !== null ? value : defaultValue);

    // Define the core fields and their expected keys in formData
    const coreFieldMapping = {
      "Timestamp": () => timestamp, // Special handler
      "EvaluatorName": () => safeGet(formData.evaluatorName), // formData key is camelCase
      "PresenterName": () => safeGet(formData.presenterName), // formData key is camelCase
      "SpeechType": () => safeGet(formData.speechType)      // formData key is camelCase
      // Add other known critical fields here if their sheet header case might differ from formData key case
    };

    let actualSheetHeaders = [];

    if (!sheet) {
      sheet = ss.insertSheet(targetSheetName);
      console.log(`Sheet "${targetSheetName}" did not exist and was created.`);
      // For a new sheet, create headers using the PascalCase keys from coreFieldMapping
      // and then any other keys from formData.
      actualSheetHeaders = [...Object.keys(coreFieldMapping)];
      Object.keys(formData).forEach(key => {
        // Add formData keys if they don't correspond to a coreFieldMapping sheet header
        // (e.g. "evaluatorName" key from formData won't be added again if "EvaluatorName" is already in actualSheetHeaders)
        // This logic needs to be careful not to add camelCase versions if PascalCase is already there.
        const pascalKey = key.charAt(0).toUpperCase() + key.slice(1);
        if (!actualSheetHeaders.includes(key) && !actualSheetHeaders.includes(pascalKey)) {
          // Default to adding the header as it appears in formData if not a mapped core field
          actualSheetHeaders.push(key); 
        }
      });
      sheet.appendRow(actualSheetHeaders);
      sheet.getRange(1, 1, 1, actualSheetHeaders.length).setFontWeight('bold');
    } else {
      if (sheet.getLastRow() === 0) { // Sheet exists but is empty
        // Behave like a new sheet regarding headers
        actualSheetHeaders = [...Object.keys(coreFieldMapping)];
        Object.keys(formData).forEach(key => {
            const pascalKey = key.charAt(0).toUpperCase() + key.slice(1);
             if (!actualSheetHeaders.find(h => h.toLowerCase() === key.toLowerCase())) {
                actualSheetHeaders.push(key); // Add if no case-insensitive match found
            }
        });
        sheet.appendRow(actualSheetHeaders);
        sheet.getRange(1, 1, 1, actualSheetHeaders.length).setFontWeight('bold');
      } else {
        actualSheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
      }
    }

    // Verify essential headers (case-insensitive check for presence)
    const lowerCaseSheetHeaders = actualSheetHeaders.map(h => h.toLowerCase());
    Object.keys(coreFieldMapping).forEach(expectedHeader => {
        if (!lowerCaseSheetHeaders.includes(expectedHeader.toLowerCase())) {
            console.warn(`Expected header (or case-variant) "${expectedHeader}" was missing from sheet "${targetSheetName}". This might cause issues.`);
            // For full robustness, you could add missing essential columns here, then re-fetch headers.
            // For now, this warning is important. If it's "SpeechType", it's a problem.
        }
    });

    const rowData = [];
    actualSheetHeaders.forEach(sheetHeader => {
      if (coreFieldMapping[sheetHeader]) { // Check if it's a core field with PascalCase sheet header
        rowData.push(coreFieldMapping[sheetHeader]());
      } else if (coreFieldMapping[sheetHeader.charAt(0).toUpperCase() + sheetHeader.slice(1)] && 
                 sheetHeader.toLowerCase() === (sheetHeader.charAt(0).toUpperCase() + sheetHeader.slice(1)).toLowerCase() ) {
        // This handles if sheetHeader is camelCase but coreFieldMapping uses PascalCase
        rowData.push(coreFieldMapping[sheetHeader.charAt(0).toUpperCase() + sheetHeader.slice(1)]());
      }
      else if (formData.hasOwnProperty(sheetHeader)) { // Exact match for other fields
        rowData.push(safeGet(formData[sheetHeader]));
      } else if (formData.hasOwnProperty(sheetHeader.toLowerCase())) { // Try lowercase key from formData
        rowData.push(safeGet(formData[sheetHeader.toLowerCase()]));
      } else {
        // If a header exists in sheet but not in formData and not a core field, push empty
        rowData.push(''); 
        console.log(`Header "${sheetHeader}" found in sheet but not in formData nor as a core mapped field. Pushing empty value.`);
      }
    });
    
    sheet.appendRow(rowData);
    console.log(`Row successfully appended to "${targetSheetName}" sheet for speech type "${formData.speechType}".`);
    return { success: true };

  } catch (error) {
    console.error("Save to sheet error:", error.toString(), error.stack);
    return { success: false, message: `Error saving to sheet: ${error.toString()}` };
  }
}
// Function to include HTML partials (like Stylesheet.css)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Find a presenter's email from their full name
function findPresenterEmail(fullName) {
  if (!studentData || studentData.length === 0) {
      loadStudentData(); // Ensure data is loaded
  }
  
  const student = studentData.find(s => s.fullName === fullName);
  if (student && student.email) {
    console.log(`Found email for ${fullName}: ${student.email}`);
    return student.email;
  }
  
  console.warn(`No email found for presenter: ${fullName}. Current student data count: ${studentData.length}`);
  return ''; // Return empty string if not found
}


// Get speech configuration from Templates sheet
function getSpeechConfiguration(speechType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Templates');
    
    if (!templateSheet) {
      return { error: "Templates sheet not found. Cannot build form." };
    }
    
    const data = templateSheet.getDataRange().getValues();
    if (data.length < 1) {
        return { error: "Templates sheet is empty."};
    }
    const headers = data[0].map(h => h.toString().trim());
    
    const colIdx = {};
    headers.forEach((header, index) => { colIdx[header] = index; });
    
    const requiredColumns = ['SpeechType', 'SectionID', 'SectionTitle', 'QuestionID', 'QuestionText', 'QuestionType'];
    for (const col of requiredColumns) {
      if (!(col in colIdx)) {
        return { error: `Required column '${col}' not found in Templates sheet. Present headers: ${headers.join(', ')}` };
      }
    }
    
    const configRows = data.slice(1).filter(row => row[colIdx.SpeechType] === speechType);
    
    if (configRows.length === 0) {
      return { error: `No configuration found for speech type: "${speechType}" in Templates sheet.` };
    }
    
    const sectionsMap = new Map(); // Use Map to maintain insertion order for sections if IDs are not purely numeric
    let formTitle = `${speechType.charAt(0).toUpperCase() + speechType.slice(1)} Speech Evaluation`; // Default title
    
    configRows.forEach(row => {
      const sectionId = row[colIdx.SectionID].toString(); // Ensure ID is a string for map keys
      
      if (!sectionsMap.has(sectionId)) {
        sectionsMap.set(sectionId, {
          id: sectionId, // Store original ID, could be numeric or string
          title: row[colIdx.SectionTitle],
          questions: []
        });
        // Check for a form-level title if provided in a specific way (e.g. first row for a speechType)
        if (row[colIdx.FormTitle] && sectionsMap.size === 1) { // Example: if FormTitle column exists
            formTitle = row[colIdx.FormTitle];
        }
      }
      
      const questionId = row[colIdx.QuestionID];
      if (!questionId) return; // Row might be for section definition only
      
      let options = [];
      if (colIdx.Options !== undefined && row[colIdx.Options]) {
        const optStr = row[colIdx.Options].toString();
        try { options = JSON.parse(optStr); } 
        catch (e) { options = optStr.split('|').map(opt => opt.trim()).filter(o => o); }
      }
      
      let scoreCriteria = [];
      if (colIdx.ScoreCriteria !== undefined && row[colIdx.ScoreCriteria]) {
        const critStr = row[colIdx.ScoreCriteria].toString();
        try { scoreCriteria = JSON.parse(critStr); } 
        catch (e) { scoreCriteria = critStr.split('|').map(crit => crit.trim()).filter(c => c); }
      }
      
      sectionsMap.get(sectionId).questions.push({
        id: questionId.toString(),
        type: row[colIdx.QuestionType].toString(),
        text: row[colIdx.QuestionText].toString(),
        options: options,
        required: colIdx.Required !== undefined ? /true/i.test(row[colIdx.Required].toString()) : false,
        defaultValue: colIdx.DefaultValue !== undefined ? row[colIdx.DefaultValue].toString() : '',
        minScore: colIdx.MinScore !== undefined ? row[colIdx.MinScore].toString() : '1',
        maxScore: colIdx.MaxScore !== undefined ? row[colIdx.MaxScore].toString() : '5',
        scoreCriteria: scoreCriteria
      });
    });
    
    // Convert sections map to array, preserving order if IDs are sortable
    // If section IDs are numbers, sort numerically. Otherwise, by insertion order (Map behavior) or string sort.
    const sectionsArray = Array.from(sectionsMap.values()).sort((a, b) => {
        const aIsNum = !isNaN(parseFloat(a.id)) && isFinite(a.id);
        const bIsNum = !isNaN(parseFloat(b.id)) && isFinite(b.id);
        if (aIsNum && bIsNum) return parseFloat(a.id) - parseFloat(b.id);
        if (aIsNum && !bIsNum) return -1;
        if (!aIsNum && bIsNum) return 1;
        return a.id.localeCompare(b.id); // Fallback to string locale compare for non-numeric IDs
    });
    
    return {
      speechType: speechType,
      title: formTitle,
      sections: sectionsArray
    };
  } catch (error) {
    console.error(`Error getting speech configuration for "${speechType}":`, error, error.stack);
    return { error: "Error getting speech configuration: " + error.toString() };
  }
}


// Get all available speech types and their sheet names from Templates
function getAvailableSpeechTypes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Templates');
    
    if (!templateSheet) {
      console.error("Templates sheet not found. Cannot get available speech types.");
      return {}; // Return empty object
    }
    
    const data = templateSheet.getDataRange().getValues();
    if (data.length < 2) return {}; // No data beyond headers

    const headers = data[0].map(h => h.toString().trim());
    const speechTypeColIndex = headers.indexOf('SpeechType');
    const sheetNameColIndex = headers.indexOf('SheetName'); // This column defines the target data sheet
    
    if (speechTypeColIndex === -1) { // SheetName is optional, can be derived
      console.error("'SpeechType' column not found in Templates sheet.");
      return {};
    }
    
    const speechTypesMap = {}; // Use a map to store speechType -> sheetName
    for (let i = 1; i < data.length; i++) {
      const speechType = data[i][speechTypeColIndex];
      if (speechType && speechType.toString().trim() !== "") {
        let sheetName = null;
        if (sheetNameColIndex !== -1 && data[i][sheetNameColIndex] && data[i][sheetNameColIndex].toString().trim() !== "") {
            sheetName = data[i][sheetNameColIndex].toString().trim();
        } else {
            // Default sheet name generation if not specified
            sheetName = speechType.toString().trim().charAt(0).toUpperCase() + speechType.toString().trim().slice(1) + " Evaluations";
        }
        speechTypesMap[speechType.toString().trim()] = sheetName;
      }
    }
    return speechTypesMap;
  } catch (error) {
      console.error("Error getting available speech types:", error);
      return {};
  }
}

// Get sheet name for a given speech type
function getSheetNameForSpeechType(speechType) {
  const speechTypes = getAvailableSpeechTypes(); // This now returns a map: { speechType: sheetName }
  if (speechTypes[speechType]) {
    return speechTypes[speechType];
  }
  // Fallback if not in Templates (should ideally not happen if config is well-maintained)
  console.warn(`Sheet name for speech type "${speechType}" not found in Templates. Generating default name.`);
  return speechType.charAt(0).toUpperCase() + speechType.slice(1) + " Evaluations"; // Consistent default
}

/*
// Diagnostic functions (keep commented out unless debugging)
function debugSheetAccess() {
  // ...
}
function debugStudentData() {
  // ...
}
*/