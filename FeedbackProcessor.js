// Google Apps Script: FeedbackProcessor.gs

// Default teacher email as a fallback
const DEFAULT_TEACHER_EMAIL = "notifications@orono.k12.mn.us"; // Or your preferred default

// Define inappropriate words list for comment sanitization
const INAPPROPRIATE_WORDS = [
  'damn', 'hell', 'crap', 'stupid', 'idiot', 'dumb', 'fool', 'jerk',
  'suck', 'hate', 'terrible', 'awful', 'worst', 'bad', 'horrible'
  // Add more words as needed, keep them lowercase
];

/**
 * Retrieves the teacher's email address.
 * Tries global variable, then 'Index' sheet, then a default.
 * @return {string} The teacher's email address.
 */
function getTeacherEmail() {
  // 1. Try the global variable `teacherEmail` set by `loadStudentData()` in Code.gs
  if (typeof teacherEmail !== 'undefined' && teacherEmail && teacherEmail.trim() !== '') {
    return teacherEmail;
  }

  // 2. Try to load it from the 'Index' sheet (cell C2)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = ss.getSheetByName('Index');
    if (indexSheet) {
      const emailFromSheet = indexSheet.getRange('C2').getValue(); // Consistent with loadStudentData
      if (emailFromSheet && typeof emailFromSheet === 'string' && emailFromSheet.trim() !== '') {
        // Update the global variable for this session if found
        this.teacherEmail = emailFromSheet.trim(); 
        return this.teacherEmail;
      }
    }
  } catch (error) {
    console.error("Error loading teacher email from Index sheet:", error);
  }

  // 3. Use the default as a last resort
  console.warn("Teacher email not found in global variable or Index sheet; using default.");
  return DEFAULT_TEACHER_EMAIL;
}

/**
 * Sanitizes a comment by replacing inappropriate words with asterisks.
 * @param {string} comment The comment to sanitize.
 * @return {string} The sanitized comment.
 */
function sanitizeComment(comment) {
  if (!comment || typeof comment !== 'string') {
    return "";
  }
  let sanitized = comment;
  INAPPROPRIATE_WORDS.forEach(word => {
    const regex = new RegExp('\\b' + word + '\\b', 'gi');
    sanitized = sanitized.replace(regex, '*'.repeat(word.length));
  });
  return sanitized;
}

/**
 * Generates HTML for a star rating.
 * @param {number} score The score to represent.
 * @param {number} maxScore The maximum possible score.
 * @return {string} HTML string for the star rating.
 */
function generateStarRating(score, maxScore) {
  const normalizedScore = Math.max(0, Math.min(score, maxScore)); // Clamp score
  const fullStars = Math.floor(normalizedScore);
  const halfStar = (normalizedScore % 1) >= 0.4 && (normalizedScore % 1) <= 0.6; // More generous half-star
  const demiStar = (normalizedScore % 1) > 0.6; // if more than half, count as full for display simplicity or adjust icon
  
  let starsHtml = '';
  let currentStars = 0;

  for (let i = 0; i < fullStars; i++) {
    starsHtml += '★';
    currentStars++;
  }
  if (halfStar) {
    starsHtml += '½'; // Or another icon for half star e.g. using Material Icons font if available in email
    currentStars++;
  } else if (demiStar && currentStars < maxScore) { // If it was rounded up essentially
     starsHtml += '★';
     currentStars++;
  }
  
  const emptyStars = maxScore - currentStars;
  for (let i = 0; i < emptyStars; i++) {
    starsHtml += '☆';
  }
  
  return starsHtml;
}

/**
 * Fetches all evaluations for a specific presenter and speech type.
 * @param {string} presenterName The name of the presenter.
 * @param {string} speechType The type of the speech.
 * @return {Array<Object>} An array of evaluation objects.
 */
function getPresenterEvaluations(presenterName, speechType) {
  const sheetName = getSheetNameForSpeechType(speechType); // From Code.gs
  if (!sheetName) {
    console.error(`Could not determine sheet name for speech type: ${speechType}`);
    return [];
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.error(`Sheet "${sheetName}" not found for speech type "${speechType}".`);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { // No data beyond headers
    console.log(`No data found in sheet "${sheetName}".`);
    return [];
  }

  const headers = data[0].map(header => header.toString().trim());
  const presenterNameColIndex = headers.indexOf('PresenterName'); // Standardized header from form

  if (presenterNameColIndex === -1) {
    console.error(`'PresenterName' column not found in sheet "${sheetName}". Headers: ${headers.join(', ')}`);
    return [];
  }

  const evaluations = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[presenterNameColIndex] === presenterName) {
      const evaluation = {};
      headers.forEach((header, index) => {
        if (header) { // Ensure header is not empty
          evaluation[header] = row[index];
        }
      });
      evaluation.speechType = speechType; // Add speechType for context if needed later
      evaluations.push(evaluation);
    }
  }

  console.log(`Found ${evaluations.length} evaluations for ${presenterName} in "${sheetName}" for speech type "${speechType}".`);
  return evaluations;
}


/**
 * Aggregates evaluation data based on the speech configuration.
 * @param {Array<Object>} evaluations An array of raw evaluation data objects for a presenter.
 * @param {Object} speechConfiguration The configuration object for the speech type.
 * @return {Object} An object containing aggregated data, structured by sections and questions.
 */
function aggregateEvaluationData(evaluations, speechConfiguration) {
  const aggregatedData = {
    speechConfigTitle: speechConfiguration.title,
    evaluationCount: evaluations.length,
    sections: []
  };

  if (evaluations.length === 0) {
    return aggregatedData; // Return basic structure if no evaluations
  }

  speechConfiguration.sections.forEach(sectionConfig => {
    const aggregatedSection = {
      id: sectionConfig.id,
      title: sectionConfig.title,
      questions: []
    };

    sectionConfig.questions.forEach(questionConfig => {
      const questionId = questionConfig.id;
      const questionType = questionConfig.type.toLowerCase();
      const aggregatedQuestion = {
        id: questionId,
        text: questionConfig.text,
        type: questionType,
        // Add specific aggregation results below
      };

      let allResponses = evaluations.map(ev => ev[questionId]).filter(val => val !== undefined && val !== null && val !== "");

      switch (questionType) {
        case 'rubric':
          const numericScores = allResponses
            .map(score => parseFloat(score))
            .filter(score => !isNaN(score));
          
          aggregatedQuestion.scores = numericScores;
          aggregatedQuestion.averageScore = numericScores.length > 0 
            ? (numericScores.reduce((sum, score) => sum + score, 0) / numericScores.length) 
            : 0;
          aggregatedQuestion.minScore = parseFloat(questionConfig.minScore) || 1; // From config
          aggregatedQuestion.maxScore = parseFloat(questionConfig.maxScore) || 5; // From config
          aggregatedQuestion.scoreCriteria = questionConfig.scoreCriteria || []; // From config
          break;

        case 'option':
        case 'dropdown':
          aggregatedQuestion.optionCounts = {};
          allResponses.forEach(response => {
            aggregatedQuestion.optionCounts[response] = (aggregatedQuestion.optionCounts[response] || 0) + 1;
          });
          aggregatedQuestion.options = questionConfig.options || []; // From config
          break;

        case 'checkbox':
          aggregatedQuestion.optionCounts = {};
          allResponses.forEach(responseStr => {
            try {
              const selectedOptions = JSON.parse(responseStr);
              if (Array.isArray(selectedOptions)) {
                selectedOptions.forEach(option => {
                  aggregatedQuestion.optionCounts[option] = (aggregatedQuestion.optionCounts[option] || 0) + 1;
                });
              }
            } catch (e) {
              console.warn(`Could not parse checkbox response for ${questionId}: ${responseStr}`, e);
            }
          });
          aggregatedQuestion.options = questionConfig.options || []; // From config
          break;
        
        case 'comment':
          aggregatedQuestion.comments = allResponses.map(comment => sanitizeComment(comment)).filter(c => c.trim() !== "" && c !== "No comments provided");
          break;
          
        default: // Handles 'text', or any other type as a list of text responses
          aggregatedQuestion.responses = allResponses.map(response => sanitizeComment(response)).filter(r => r.trim() !== "" && r !== "No comments provided");
          break;
      }
      aggregatedSection.questions.push(aggregatedQuestion);
    });
    aggregatedData.sections.push(aggregatedSection);
  });
  
  console.log("Aggregated Data Sample (first section, first question):", aggregatedData.sections[0]?.questions[0]);
  return aggregatedData;
}

/**
 * Generates dynamic HTML content for the feedback email.
 * @param {string} presenterName The name of the presenter.
 * @param {Object} aggregatedData The aggregated evaluation data.
 * @param {Object} speechConfiguration The configuration for this speech type.
 * @return {string} HTML string for the email body.
 */
function generateDynamicFeedbackEmailHtml(presenterName, aggregatedData, speechConfiguration) {
  let html = `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; }
        .header { background-color: #1a73e8; color: white; padding: 20px; text-align: center; border-radius: 5px 5px 0 0; }
        .header h1 { margin: 0; font-size: 24px;}
        .header h2 { margin: 5px 0 0; font-size: 20px; font-weight: normal;}
        .header p { margin: 5px 0 0; font-size: 14px; }
        .content { padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 5px 5px; }
        .section-card { background-color: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 5px; padding: 15px; margin-bottom: 20px; }
        .section-title { font-size: 20px; font-weight: bold; margin: 0 0 15px; padding-bottom: 10px; border-bottom: 2px solid #1a73e8; color: #1a73e8; }
        .question-block { margin-bottom: 15px; }
        .question-text { font-weight: bold; color: #444; margin-bottom: 5px; }
        .question-response { margin-left: 15px; }
        .score-display { font-size: 16px; }
        .star-rating { color: #fbbc04; letter-spacing: 2px; font-size: 18px; }
        .comment { font-style: italic; padding: 8px; background-color: #f1f8ff; border-left: 3px solid #1a73e8; margin: 5px 0 5px 15px; }
        .option-count { margin-left: 15px; }
        .option-name { font-weight: normal; }
        .option-bar-container { display: flex; align-items: center; margin-bottom: 3px; }
        .option-bar { background-color: #ddd; height: 10px; border-radius: 5px; min-width: 100px; margin-right: 5px; }
        .option-bar-fill { background-color: #4285f4; height: 100%; border-radius: 5px; }
        .footer { margin-top: 30px; text-align: center; font-size: 12px; color: #777; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>${speechConfiguration.title || 'Speech Evaluation Feedback'}</h1>
        <h2>${presenterName}</h2>
        <p>Based on feedback from ${aggregatedData.evaluationCount} peer evaluator(s)</p>
      </div>
      <div class="content">
  `;

  aggregatedData.sections.forEach(section => {
    // Only create a card if there are questions with data to show in this section
    let sectionHasContent = section.questions.some(q => {
        return (q.scores && q.scores.length > 0) || 
               (q.optionCounts && Object.keys(q.optionCounts).length > 0) ||
               (q.comments && q.comments.length > 0) ||
               (q.responses && q.responses.length > 0);
    });

    // Also consider if the section is the "Review" section which might be empty of configured questions but could have specific handling
    if (section.title.toLowerCase().includes('review')) {
        sectionHasContent = false; // Typically review sections are for client-side, not email, unless designed otherwise
    }


    if (!sectionHasContent) return; // Skip rendering empty sections (unless it's a Review section we decide to handle)

    html += `<div class="section-card">`;
    html += `<div class="section-title">${section.title}</div>`;

    section.questions.forEach(question => {
      html += `<div class="question-block">`;
      html += `<div class="question-text">${question.text}</div>`;
      html += `<div class="question-response">`;

      switch (question.type) {
        case 'rubric':
          if (question.scores && question.scores.length > 0) {
            html += `<div class="score-display">`;
            html += `<span class="star-rating">${generateStarRating(question.averageScore, question.maxScore)}</span> `;
            html += `(${question.averageScore.toFixed(1)} / ${question.maxScore} average from ${question.scores.length} ratings)`;
            html += `</div>`;
            // Optionally, display score criteria or individual scores if desired (can make email long)
            // question.scoreCriteria might be useful here
          } else {
            html += `<div>No rubric scores submitted.</div>`;
          }
          break;

        case 'option':
        case 'dropdown':
        case 'checkbox':
          if (question.optionCounts && Object.keys(question.optionCounts).length > 0) {
            const totalSelections = Object.values(question.optionCounts).reduce((sum, count) => sum + count, 0);
            question.options.forEach(opt => { // Iterate through configured options to maintain order
              const count = question.optionCounts[opt] || 0;
              if (count > 0 || (question.options.includes(opt) && totalSelections > 0)) { // Show if has votes or is a valid option and there were votes
                const percentage = totalSelections > 0 ? (count / totalSelections) * 100 : 0;
                html += `<div class="option-count">`;
                html += `<div class="option-bar-container">`;
                html +=   `<div class="option-bar"><div class="option-bar-fill" style="width: ${percentage.toFixed(0)}%;"></div></div>`;
                html +=   `<span class="option-name">${opt}</span>: ${count} (${percentage.toFixed(0)}%)`;
                html += `</div>`;
                html += `</div>`;
              }
            });
             // Show options that received votes but might not be in the original config (e.g. if options can be dynamic - less common)
            Object.keys(question.optionCounts).forEach(opt => {
              if (!question.options.includes(opt)) {
                 const count = question.optionCounts[opt] || 0;
                 const percentage = totalSelections > 0 ? (count / totalSelections) * 100 : 0;
                  html += `<div class="option-count">`;
                  html += `<div class="option-bar-container">`;
                  html +=   `<div class="option-bar"><div class="option-bar-fill" style="width: ${percentage.toFixed(0)}%;"></div></div>`;
                  html +=   `<span class="option-name">${opt}</span>: ${count} (${percentage.toFixed(0)}%)`;
                  html += `</div>`;
                  html += `</div>`;
              }
            });
          } else {
            html += `<div>No selections made.</div>`;
          }
          break;

        case 'comment':
          if (question.comments && question.comments.length > 0) {
            question.comments.forEach(comment => {
              html += `<div class="comment">${comment}</div>`;
            });
          } else {
            html += `<div>No comments provided.</div>`;
          }
          break;

        default: // 'text' or other types
          if (question.responses && question.responses.length > 0) {
            question.responses.forEach(response => {
              // For general text, could be a list or just concatenated if appropriate
              html += `<div class="comment">${response}</div>`; // Using 'comment' style for now
            });
          } else {
            html += `<div>No responses provided.</div>`;
          }
          break;
      }
      html += `</div></div>`; // Close question-response and question-block
    });
    html += `</div>`; // Close section-card
  });

  html += `
        <div class="footer">
          <p>This is an automated summary generated by the Speech Peer Evaluation System.</p>
          <p>Generated on ${new Date().toLocaleDateString()}</p>
        </div>
      </div>
    </body>
  </html>
  `;
  return html;
}


/**
 * Sends a feedback email to a specific presenter for a given speech type.
 * @param {string} presenterName The name of the presenter.
 * @param {string} speechType The type of the speech.
 * @return {Object} An object indicating success or failure and a message.
 */
function sendFeedbackEmail(presenterName, speechType) {
  try {
    const speechConfiguration = getSpeechConfiguration(speechType); // From Code.gs
    if (speechConfiguration.error) {
      console.error(`Error getting speech configuration for ${speechType}: ${speechConfiguration.error}`);
      return { success: false, message: `Configuration error for ${speechType}: ${speechConfiguration.error}` };
    }

    const evaluations = getPresenterEvaluations(presenterName, speechType);
    if (evaluations.length === 0) {
      return { success: false, message: `No evaluation data found for ${presenterName} for ${speechType}.` };
    }

    const presenterEmail = findPresenterEmail(presenterName); // From Code.gs
    if (!presenterEmail) {
      console.warn(`No email address found for presenter: ${presenterName}. Feedback not sent to presenter.`);
      // Decide if you still want to send to teacher or just return error
    }

    const teacherCcEmail = getTeacherEmail();
    const aggregatedData = aggregateEvaluationData(evaluations, speechConfiguration);
    const emailHtml = generateDynamicFeedbackEmailHtml(presenterName, aggregatedData, speechConfiguration);
    const emailSubject = `${speechConfiguration.title || speechType} - Feedback Summary for ${presenterName}`;

    const mailOptions = {
      cc: teacherCcEmail,
      subject: emailSubject,
      htmlBody: emailHtml
    };

    if (presenterEmail) {
      mailOptions.to = presenterEmail;
      MailApp.sendEmail(mailOptions);
      console.log(`Feedback email sent to ${presenterName} <${presenterEmail}>, CC: ${teacherCcEmail} for ${speechType}.`);
      return {
        success: true,
        message: `Feedback email sent to ${presenterName} (${presenterEmail}), CC: ${teacherCcEmail}.`
      };
    } else if (teacherCcEmail) { // Only send to teacher if presenter email not found
      mailOptions.to = teacherCcEmail; // Change CC to To
      delete mailOptions.cc;
      mailOptions.subject = `${emailSubject} (Presenter email not found)`;
      MailApp.sendEmail(mailOptions);
      console.log(`Feedback email sent ONLY to teacher <${teacherCcEmail}> for ${presenterName} (presenter email not found).`);
      return {
        success: true,
        message: `Feedback email sent to teacher ${teacherCcEmail} (presenter email for ${presenterName} not found).`
      };
    } else {
      console.error(`Neither presenter nor teacher email found for ${presenterName}. Email not sent.`);
      return { success: false, message: `No recipient (presenter or teacher) email found for ${presenterName}.`};
    }

  } catch (error) {
    console.error(`Error sending feedback to ${presenterName} for ${speechType}:`, error, error.stack);
    return { success: false, message: `Error sending email: ${error.toString()}` };
  }
}

/**
 * Shows a preview of the feedback email for a selected presenter and speech type.
 * @param {string} presenterName The name of the presenter.
 * @param {string} speechType The type of the speech.
 */
function showFeedbackPreview(presenterName, speechType) {
  const ui = SpreadsheetApp.getUi();
  try {
    const speechConfiguration = getSpeechConfiguration(speechType); // From Code.gs
    if (speechConfiguration.error) {
      ui.alert('Error', `Could not load speech configuration: ${speechConfiguration.error}`, ui.ButtonSet.OK);
      return;
    }

    const evaluations = getPresenterEvaluations(presenterName, speechType);
    if (evaluations.length === 0) {
      ui.alert('No Data', `No evaluation data found for ${presenterName} for speech type "${speechType}".`, ui.ButtonSet.OK);
      return;
    }

    const aggregatedData = aggregateEvaluationData(evaluations, speechConfiguration);
    const emailHtml = generateDynamicFeedbackEmailHtml(presenterName, aggregatedData, speechConfiguration);

    const htmlOutput = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .preview-container { border: 1px solid #ddd; padding: 10px; max-height: 70vh; overflow-y: auto; margin-bottom: 20px; }
            .button-container { display: flex; justify-content: flex-end; gap: 10px; }
            button { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; }
            .send-button { background-color: #4285f4; color: white; }
            .send-button:hover { background-color: #3367d6; }
            .cancel-button { background-color: #f1f1f1; color: #333; }
            .cancel-button:hover { background-color: #e0e0e0; }
          </style>
        </head>
        <body>
          <h2>Email Preview: ${presenterName} (${speechType})</h2>
          <div class="preview-container">${emailHtml}</div>
          <div class="button-container">
            <button class="cancel-button" onclick="google.script.host.close()">Close</button>
            <button class="send-button" onclick="confirmAndSend()">Send Email</button>
          </div>
          <script>
            function confirmAndSend() {
              if (confirm("Are you sure you want to send this feedback email?")) {
                google.script.run
                  .withSuccessHandler(function(result) {
                    if (result.success) {
                      SpreadsheetApp.getUi().alert("Success", result.message);
                    } else {
                      SpreadsheetApp.getUi().alert("Error", result.message);
                    }
                    google.script.host.close();
                  })
                  .withFailureHandler(function(error) {
                    SpreadsheetApp.getUi().alert("Error", "Failed to send email: " + error.message);
                    google.script.host.close();
                  })
                  .sendFeedbackEmail("${presenterName}", "${speechType}");
              }
            }
          </script>
        </body>
      </html>`)
      .setWidth(800)
      .setHeight(600); // Adjusted height for better preview

    ui.showModalDialog(htmlOutput, `Preview Feedback: ${presenterName}`);

  } catch (error) {
    console.error(`Error generating preview for ${presenterName}, ${speechType}:`, error, error.stack);
    ui.alert('Error', `Could not generate preview: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ==================================================================================
// OLD FUNCTIONS - To be reviewed and likely removed after new system is validated
// ==================================================================================

/*
// OLD: generateFeedbackEmailHtml - Replaced by generateDynamicFeedbackEmailHtml
function generateFeedbackEmailHtml(presenterName, evaluations) {
  // ... old hardcoded implementation ...
  const stats = calculateStatistics(evaluations); // Old function
  const feedbackByCategory = organizeFeedbackByCategory(evaluations); // Old function
  // ... rest of old hardcoded HTML generation ...
  return html;
}
*/

/*
// OLD: calculateStatistics - Replaced by aggregateEvaluationData
function calculateStatistics(evaluations) {
  // ... old hardcoded implementation ...
  return stats;
}
*/

/*
// OLD: organizeFeedbackByCategory - Replaced by aggregateEvaluationData
function organizeFeedbackByCategory(evaluations) {
  // ... old hardcoded implementation ...
  return { bodyComments, dictionComments, deliveryComments, didWell, improvement };
}
*/

/*
// OLD: generateCommentsList - Functionality integrated into generateDynamicFeedbackEmailHtml
function generateCommentsList(comments) {
  // ... old implementation ...
}
*/

/*
// OLD: generateFeedbackList - Functionality integrated into generateDynamicFeedbackEmailHtml
function generateFeedbackList(items) {
  // ... old implementation ...
}
*/

/*
// OLD: generatePositionChangeHtml - Functionality to be integrated into generateDynamicFeedbackEmailHtml based on config
function generatePositionChangeHtml(positionChanges) {
  // ... old implementation ...
}
*/

/*
// OLD: generateConvincingElementsHtml - Functionality to be integrated into generateDynamicFeedbackEmailHtml based on config
function generateConvincingElementsHtml(elements) {
  // ... old implementation ...
}
*/

/*
// OLD: sendSummaryEmail - Replaced by sendFeedbackEmail with dynamic content
function sendSummaryEmail(presenterName) {
  // ... old implementation using old helpers ...
}
*/

/*
// OLD: groupComments - Replaced by aggregateEvaluationData
function groupComments(evaluations) {
  // ... old implementation ...
}
*/

/*
// OLD: calculateAverages - Replaced by aggregateEvaluationData
function calculateAverages(evaluations) {
  // ... old implementation ...
}
*/