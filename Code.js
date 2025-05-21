// Google Apps Script to create a web app for peer speech evaluations
// This reads from "Index" tab and writes to "Peer Evaluations" tab

// Global variables
let studentData = [];
let teacherEmail = '';

// Main function to serve the web app HTML
function doGet(e) {
  // Load student data
  loadStudentData();
  
  // Get speech type from URL parameter or settings
  let speechType = e.parameter.type;
  
  // If no speech type in URL, get from settings
  if (!speechType) {
    speechType = getActiveSpeechType();
  }
  
  // Create and return the HTML content
  const template = HtmlService.createTemplateFromFile('Index');
  template.speechType = speechType;
  
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
      newSheet.getRange('A2:B2').setValues([['ActiveSpeechType', 'persuasive']]);
      newSheet.getRange('A:B').setFontWeight('bold');
      return 'persuasive'; // Default
    }
    
    // Find the ActiveSpeechType setting
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'ActiveSpeechType') {
        return data[i][1] || 'persuasive';
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
      // Create settings sheet if it doesn't exist
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
      settingsSheet.getRange('A1:B1').setFontWeight('bold');
      settingsSheet.getRange('A2:B2').setValues([['ActiveSpeechType', speechType]]);
      return { success: true, message: `Active speech type set to "${speechType}"` };
    }
    
    // Find the ActiveSpeechType setting
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'ActiveSpeechType') {
        // Update the value
        settingsSheet.getRange(i + 1, 2).setValue(speechType);
        return { success: true, message: `Active speech type updated to "${speechType}"` };
      }
    }
    
    // If not found, add it
    const lastRow = settingsSheet.getLastRow();
    settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['ActiveSpeechType', speechType]]);
    return { success: true, message: `Active speech type set to "${speechType}"` };
  } catch (error) {
    console.error("Error setting active speech type:", error);
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Load student data from the Index tab
function loadStudentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName('Index');
  
  // Get all data from the Index sheet
  const data = indexSheet.getDataRange().getValues();
  
  // Skip the header row (row 1)
  studentData = [];
  for (let i = 1; i < data.length; i++) {
    // Only add if there's a non-empty name in column A
    if (data[i][0] && data[i][0].toString().trim() !== '') {
      studentData.push({
        fullName: data[i][0],    // Column A for full name
        email: data[i][1]        // Column B for email
      });
    }
  }
  
  // Sort the student data alphabetically by fullName
  studentData.sort((a, b) => a.fullName.localeCompare(b.fullName));
  
  // Get teacher email from cell C2 (based on your spreadsheet structure)
  teacherEmail = data.length > 1 && data[1].length > 2 ? data[1][2] : '';
}

// Return student data to the client-side JavaScript
function getStudentData() {
  loadStudentData(); // Refresh data
  return studentData;
}

function processForm(formData) {
  try {
    console.log("Process form started with data:", JSON.stringify(formData));
    
    // Validate required fields
    if (!formData.evaluatorName || !formData.presenterName || !formData.speechType) {
      return { success: false, message: "Missing required fields: evaluator, presenter, or speech type" };
    }
    
    // Save to the appropriate sheet based on speech type
    saveToSheet(formData);
    
    return { success: true, message: "Your evaluation has been submitted successfully!" };
  } catch (error) {
    console.error("Process form error:", error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Updated sendEmails function 
function sendEmails(formData) {
  // Force reload student data
  loadStudentData();
  
  // Find presenter email
  const presenterEmail = findPresenterEmail(formData.presenterName);
  
  if (!presenterEmail) {
    console.error("Failed to find email for presenter: " + formData.presenterName);
  } else {
    console.log("Found presenter email: " + presenterEmail);
  }
  
  // Create email content
  const subject = 'Speech Evaluation: ' + formData.presenterName + ' (evaluated by ' + formData.evaluatorName + ')';
  const body = createEmailBody(formData);
  
  // Send email to teacher
  if (teacherEmail) {
    try {
      MailApp.sendEmail({
        to: teacherEmail,
        subject: subject,
        htmlBody: createHtmlEmailBody(formData)
      });
      console.log("Email sent to teacher: " + teacherEmail);
    } catch (e) {
      console.error("Error sending email to teacher: " + e.toString());
    }
  }
  
  // Send email to presenter
  if (presenterEmail) {
    try {
      MailApp.sendEmail({
        to: presenterEmail,
        cc: teacherEmail, // CC the teacher on presenter emails
        subject: subject,
        htmlBody: createHtmlEmailBody(formData)
      });
      console.log("Email sent to presenter: " + presenterEmail);
    } catch (e) {
      console.error("Error sending email to presenter: " + e.toString());
    }
  }
}

// Create the HTML email body
function createHtmlEmailBody(formData) {
  // Process the rhetorical devices array
  let rhetoricalDevices = 'None identified';
  if (Array.isArray(formData.rhetoricalDevices) && formData.rhetoricalDevices.length > 0) {
    rhetoricalDevices = formData.rhetoricalDevices.join(', ');
  } else if (typeof formData.rhetoricalDevices === 'string' && formData.rhetoricalDevices) {
    try {
      const parsed = JSON.parse(formData.rhetoricalDevices);
      if (Array.isArray(parsed) && parsed.length > 0) {
        rhetoricalDevices = parsed.join(', ');
      }
    } catch (e) {
      rhetoricalDevices = formData.rhetoricalDevices;
    }
  }

  // Create HTML email content  
  let html = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 5px;">
    <div style="background-color: #1a73e8; color: white; padding: 15px; border-radius: 5px 5px 0 0; margin: -20px -20px 20px;">
      <h1 style="margin: 0; font-size: 22px;">Speech Evaluation Feedback</h1>
    </div>
    
    <div style="color: #5f6368; margin-bottom: 25px; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px;">
      <p><strong>Speaker:</strong> ${formData.presenterName}</p>
      <p><strong>Evaluated by:</strong> ${formData.evaluatorName}</p>
      <p><strong>Initial Position:</strong> ${formData.initialPosition}</p>
    </div>
    
    <div style="margin-bottom: 25px;">
      <h2 style="color: #1a73e8; font-size: 18px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0;">Speech Content</h2>
      
      <p><strong>Body of Speech Score:</strong> ${formData.bodyScore}/4</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 10px;">
        <p><strong>Comments:</strong> ${formData.bodyComments}</p>
      </div>
    </div>
    
    <div style="margin-bottom: 25px;">
      <h2 style="color: #1a73e8; font-size: 18px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0;">Diction & Rhetoric</h2>
      
      <p><strong>Diction Score:</strong> ${formData.dictionScore}/4</p>
      <p><strong>Rhetorical Devices Used:</strong> ${rhetoricalDevices}</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 10px;">
        <p><strong>Comments:</strong> ${formData.dictionComments}</p>
      </div>
    </div>
    
    <div style="margin-bottom: 25px;">
      <h2 style="color: #1a73e8; font-size: 18px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0;">Delivery</h2>
      
      <p><strong>Eye Contact Score:</strong> ${formData.eyeContactScore}/4</p>
      <p><strong>Posture & Gestures Score:</strong> ${formData.postureScore}/4</p>
      <p><strong>Vocal Variety Score:</strong> ${formData.vocalScore}/4</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 10px;">
        <p><strong>Comments:</strong> ${formData.deliveryComments}</p>
      </div>
    </div>
    
    <div style="margin-bottom: 25px;">
      <h2 style="color: #1a73e8; font-size: 18px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0;">Impact & Feedback</h2>
      
      <p><strong>Position Change After Speech:</strong> ${formData.positionChange}</p>
      <p><strong>Most Convincing Element:</strong> ${formData.mostConvincing}</p>
      <p><strong>Least Convincing Element:</strong> ${formData.leastConvincing}</p>
      <p><strong>What Was Done Well:</strong> ${formData.didWell}</p>
      <p><strong>Area for Improvement:</strong> ${formData.improvement}</p>
    </div>
    
    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0; font-size: 12px; color: #5f6368; text-align: center;">
      <p>This is an automated message from the Speech Peer Evaluation System.</p>
    </div>
  </div>
  `;
  
  return html;
}

// Create plain text email body (fallback)
function createEmailBody(formData) {
  let body = 'Speech Evaluation Summary\n\n';
  
  body += 'Evaluator: ' + formData.evaluatorName + '\n';
  body += 'Presenter: ' + formData.presenterName + '\n\n';
  
  body += 'Initial Position: ' + formData.initialPosition + '\n\n';
  
  body += 'Body of Speech Score: ' + formData.bodyScore + '/4\n';
  body += 'Comments: ' + formData.bodyComments + '\n\n';
  
  body += 'Diction and Rhetoric Score: ' + formData.dictionScore + '/4\n';
  
  // Process rhetorical devices
  let rhetoricalDevices = 'None identified';
  if (Array.isArray(formData.rhetoricalDevices) && formData.rhetoricalDevices.length > 0) {
    rhetoricalDevices = formData.rhetoricalDevices.join(', ');
  } else if (typeof formData.rhetoricalDevices === 'string' && formData.rhetoricalDevices) {
    try {
      const parsed = JSON.parse(formData.rhetoricalDevices);
      if (Array.isArray(parsed) && parsed.length > 0) {
        rhetoricalDevices = parsed.join(', ');
      }
    } catch (e) {
      rhetoricalDevices = formData.rhetoricalDevices;
    }
  }
  
  body += 'Rhetorical Devices Used: ' + rhetoricalDevices + '\n';
  body += 'Comments: ' + formData.dictionComments + '\n\n';
  
  body += 'Eye Contact Score: ' + formData.eyeContactScore + '/4\n';
  body += 'Posture and Gestures Score: ' + formData.postureScore + '/4\n';
  body += 'Vocal Variety Score: ' + formData.vocalScore + '/4\n';
  body += 'Delivery Comments: ' + formData.deliveryComments + '\n\n';
  
  body += 'Position Change After Speech: ' + formData.positionChange + '\n';
  body += 'Most Convincing Element: ' + formData.mostConvincing + '\n';
  body += 'Least Convincing Element: ' + formData.leastConvincing + '\n\n';
  
  body += 'What the Presenter Did Well: ' + formData.didWell + '\n';
  body += 'Suggestion for Improvement: ' + formData.improvement + '\n';
  
  return body;
}

// Updated saveToSheet function to handle speech-specific tabs
function saveToSheet(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Could not access the active spreadsheet");
    }
    
    // Get the sheet name for this speech type
    const sheetName = getSheetNameForSpeechType(formData.speechType);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Could not find sheet named '${sheetName}'`);
    }
    
    // Create a timestamp
    const timestamp = new Date();
    
    // Handle undefined or null values safely
    const safeGet = (value, defaultValue = '') => {
      return value !== undefined && value !== null ? value : defaultValue;
    };
    
    // Create row data starting with standard fields
    const rowData = [
      timestamp,
      safeGet(formData.evaluatorName),
      safeGet(formData.presenterName)
    ];
    
    // Get headers from the sheet or create them if the sheet is empty
    let headers;
    if (sheet.getLastRow() === 0) {
      // Sheet is empty, create headers
      headers = ['Timestamp', 'EvaluatorName', 'PresenterName'];
      
      // Add all other form fields to headers
      Object.keys(formData).forEach(key => {
        if (!['evaluatorName', 'presenterName', 'speechType'].includes(key)) {
          headers.push(key);
        }
      });
      
      sheet.appendRow(headers);
    } else {
      // Get existing headers
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // Add other form data based on headers
    for (let i = 3; i < headers.length; i++) {
      const fieldName = headers[i];
      if (formData[fieldName] !== undefined) {
        rowData.push(formData[fieldName]);
      } else {
        rowData.push('');
      }
    }
    
    // Append the row to the sheet
    sheet.appendRow(rowData);
    console.log(`Row successfully appended to "${sheetName}" sheet`);
    
    return true;
  } catch (error) {
    console.error("Save to sheet error:", error.toString());
    throw error;
  }
}

// Add this function to your Code.gs file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Find a presenter's email from their full name
function findPresenterEmail(fullName) {
  console.log("Looking for email for presenter: " + fullName);
  console.log("Current student data: " + JSON.stringify(studentData));
  
  // Force reload student data to ensure it's fresh
  loadStudentData();
  console.log("Refreshed student data: " + JSON.stringify(studentData));
  
  for (let i = 0; i < studentData.length; i++) {
    if (studentData[i].fullName === fullName) {
      console.log("Found email for " + fullName + ": " + studentData[i].email);
      return studentData[i].email;
    }
  }
  
  console.log("No email found for: " + fullName);
  return '';
}

// Get speech configuration from Templates sheet
function getSpeechConfiguration(speechType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Templates');
    
    if (!templateSheet) {
      return { error: "Templates sheet not found" };
    }
    
    // Get all template data
    const data = templateSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Create column index map
    const colIdx = {};
    headers.forEach((header, index) => {
      colIdx[header] = index;
    });
    
    // Required columns
    const requiredColumns = ['SpeechType', 'SheetName', 'SectionID', 'SectionTitle'];
    for (const col of requiredColumns) {
      if (!(col in colIdx)) {
        return { error: `Required column '${col}' not found in Templates sheet` };
      }
    }
    
    // Filter for requested speech type
    const configRows = data.slice(1).filter(row => 
      row[colIdx.SpeechType] === speechType
    );
    
    if (configRows.length === 0) {
      return { error: `No configuration found for speech type: ${speechType}` };
    }
    
    // Organize by sections
    const sections = {};
    
    configRows.forEach(row => {
      const sectionId = row[colIdx.SectionID];
      const sectionTitle = row[colIdx.SectionTitle];
      const questionId = row[colIdx.QuestionID];
      
      // Initialize section if it doesn't exist
      if (!sections[sectionId]) {
        sections[sectionId] = {
          id: sectionId,
          title: sectionTitle,
          questions: []
        };
      }
      
      // Skip if this row is just defining a section without a question
      if (!questionId) return;
      
      // Parse options
      let options = [];
      if (colIdx.Options !== undefined && row[colIdx.Options]) {
        // Try to parse as JSON first
        try {
          options = JSON.parse(row[colIdx.Options]);
        } catch (e) {
          // If that fails, split by pipe
          options = row[colIdx.Options].split('|').map(opt => opt.trim());
        }
      }
      
      // Parse score criteria
      let scoreCriteria = [];
      if (colIdx.ScoreCriteria !== undefined && row[colIdx.ScoreCriteria]) {
        // Try to parse as JSON first
        try {
          scoreCriteria = JSON.parse(row[colIdx.ScoreCriteria]);
        } catch (e) {
          // If that fails, split by pipe
          scoreCriteria = row[colIdx.ScoreCriteria].split('|').map(crit => crit.trim());
        }
      }
      
      // Get question properties
      const questionType = colIdx.QuestionType !== undefined ? row[colIdx.QuestionType] : '';
      const questionText = colIdx.QuestionText !== undefined ? row[colIdx.QuestionText] : '';
      const required = colIdx.Required !== undefined ? 
        (row[colIdx.Required] === true || row[colIdx.Required] === 'TRUE' || row[colIdx.Required] === 'true') : false;
      const defaultValue = colIdx.DefaultValue !== undefined ? row[colIdx.DefaultValue] : '';
      const minScore = colIdx.MinScore !== undefined ? row[colIdx.MinScore] : '';
      const maxScore = colIdx.MaxScore !== undefined ? row[colIdx.MaxScore] : '';
      
      // Add question to section
      sections[sectionId].questions.push({
        id: questionId,
        type: questionType,
        text: questionText,
        options: options,
        required: required,
        defaultValue: defaultValue,
        minScore: minScore,
        maxScore: maxScore,
        scoreCriteria: scoreCriteria
      });
    });
    
    // Convert to array and sort by section ID
    const sectionsArray = Object.values(sections).sort((a, b) => {
      // Convert to numbers if possible, otherwise compare as strings
      const aId = !isNaN(a.id) ? Number(a.id) : a.id;
      const bId = !isNaN(b.id) ? Number(b.id) : b.id;
      
      if (typeof aId === 'number' && typeof bId === 'number') {
        return aId - bId;
      }
      return String(aId).localeCompare(String(bId));
    });
    
    return {
      speechType: speechType,
      title: `${speechType.charAt(0).toUpperCase() + speechType.slice(1)} Speech Evaluation`,
      sections: sectionsArray
    };
  } catch (error) {
    console.error("Error getting speech configuration:", error);
    return { error: "Error: " + error.toString() };
  }
}

// Get all available speech types and their sheet names from Templates
function getAvailableSpeechTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName('Templates');
  
  if (!templateSheet) {
    console.error("Templates sheet not found");
    return [];
  }
  
  // Get all template data
  const data = templateSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const speechTypeColIndex = headers.indexOf('SpeechType');
  const sheetNameColIndex = headers.indexOf('SheetName');
  
  if (speechTypeColIndex === -1 || sheetNameColIndex === -1) {
    console.error("Required columns not found in Templates sheet");
    return [];
  }
  
  // Get unique speech types and their sheet names
  const speechTypes = {};
  
  // Start from row 1 (skipping header)
  for (let i = 1; i < data.length; i++) {
    const speechType = data[i][speechTypeColIndex];
    const sheetName = data[i][sheetNameColIndex];
    
    if (speechType && sheetName) {
      speechTypes[speechType] = sheetName;
    }
  }
  
  return speechTypes;
}

// Get sheet name for a given speech type
function getSheetNameForSpeechType(speechType) {
  // Get all speech types and their sheet names
  const speechTypes = getAvailableSpeechTypes();
  
  // Return the sheet name if found
  if (speechTypes[speechType]) {
    return speechTypes[speechType];
  }
  
  // Fallback: Generate a sheet name based on the speech type
  return speechType.charAt(0).toUpperCase() + speechType.slice(1) + " Speech";
}

/*
// Diagnostic function to test if we can access the spreadsheet and sheet
function debugSheetAccess() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: "Could not access the active spreadsheet" };
    }
    
    const sheetNames = ss.getSheets().map(sheet => sheet.getName());
    
    const peerEvalSheet = ss.getSheetByName('Peer Evaluations');
    if (!peerEvalSheet) {
      return { 
        success: false, 
        message: "Could not find 'Peer Evaluations' sheet",
        availableSheets: sheetNames
      };
    }
    
    // Try to write a test row
    try {
      const testRow = ["TEST", "Diagnostic Test", new Date(), "Please delete this row"];
      peerEvalSheet.appendRow(testRow);
      return { 
        success: true, 
        message: "Successfully accessed spreadsheet and appended test row",
        sheetId: ss.getId(),
        sheetUrl: ss.getUrl(),
        sheetNames: sheetNames
      };
    } catch (writeError) {
      return {
        success: false,
        message: "Could access spreadsheet but failed to write: " + writeError.toString(),
        sheetNames: sheetNames
      };
    }
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Function to log student data for debugging
function debugStudentData() {
  loadStudentData();
  
  const studentInfo = studentData.map(student => ({
    name: student.fullName,
    email: student.email
  }));
  
  console.log("Student data loaded:");
  console.log(JSON.stringify(studentInfo, null, 2));
  
  return studentInfo;
}
*/