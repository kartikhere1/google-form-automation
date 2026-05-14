// ============================================================
// Code.gs — Excel to Google Form Generator
// Project: google-form-automation
//
// HOW IT WORKS:
//   1. You upload your Excel file to Google Drive
//   2. Google converts it to a Google Sheet automatically
//   3. This script reads your questions from that Sheet
//   4. It builds a Google Form with all your questions
//   5. Form responses are linked to a new Google Sheet
// ============================================================


// ─────────────────────────────────────────────────────────────
// CONFIGURATION — Update these before running
// ─────────────────────────────────────────────────────────────

var CONFIG = {
  // Paste the ID of your Google Sheet here (from its URL)
  // URL format: https://docs.google.com/spreadsheets/d/YOUR_ID/edit
  QUESTIONS_SHEET_ID: "YOUR_SPREADSHEET_ID_HERE",

  // The tab name inside your Sheet that has the questions
  // This should match exactly (case-sensitive)
  QUESTIONS_TAB_NAME: "Questions",

  // The title of the Google Form that will be created
  FORM_TITLE: "My Generated Form",

  // Optional: A description shown at the top of the form
  FORM_DESCRIPTION: "Please fill in this form. All fields marked * are required.",

  // If true, the script sends you an email with the form link when done
  SEND_EMAIL_ON_COMPLETION: true,
  NOTIFICATION_EMAIL: "your@email.com",
};


// ─────────────────────────────────────────────────────────────
// EXPECTED EXCEL / SHEET FORMAT
//
// Your Excel file must have these columns (Row 1 = headers):
//
// | Question         | Type            | Options              | Required | Help Text          |
// |------------------|-----------------|----------------------|----------|--------------------|
// | What is name?    | short_answer    |                      | TRUE     |                    |
// | Choose a color   | multiple_choice | Red, Blue, Green     | TRUE     | Pick just one      |
// | Select all tags  | checkbox        | Tech, Art, Sports    | FALSE    | Choose all that fit|
// | Pick department  | dropdown        | HR, Tech, Finance    | TRUE     |                    |
// | Tell us more     | paragraph       |                      | FALSE    | Optional details   |
// | Rate your exp.   | linear_scale    | 1, 5, Bad, Excellent | TRUE     | 1=Bad, 5=Excellent |
// | Date of event    | date            |                      | FALSE    |                    |
// | Time preference  | time            |                      | FALSE    |                    |
//
// SUPPORTED TYPES:
//   short_answer    → Single-line text box
//   paragraph       → Multi-line text box
//   multiple_choice → Radio buttons (pick one)
//   checkbox        → Checkboxes (pick many)
//   dropdown        → Dropdown menu (pick one)
//   linear_scale    → 1–5 star/rating scale
//                     Options format: "min, max, minLabel, maxLabel"
//                     Example: "1, 5, Bad, Excellent"
//   date            → Date picker
//   time            → Time picker
// ─────────────────────────────────────────────────────────────


// ─────────────────────────────────────────────────────────────
// MAIN FUNCTION: generateForm
// Run this from the Apps Script editor to create your form.
// ─────────────────────────────────────────────────────────────

/**
 * Reads questions from your Google Sheet and generates a Google Form.
 * Run this function once to create the form.
 */
function generateForm() {
  Logger.log("🚀 Starting form generation...");

  try {
    // ── Step 1: Read questions from the Sheet ──────────────
    var questions = readQuestionsFromSheet();

    if (questions.length === 0) {
      Logger.log("❌ No questions found in the sheet. Check your tab name and data.");
      return;
    }

    Logger.log("✅ Found " + questions.length + " questions. Building form...");

    // ── Step 2: Create a new Google Form ──────────────────
    var form = FormApp.create(CONFIG.FORM_TITLE);
    form.setDescription(CONFIG.FORM_DESCRIPTION);
    form.setCollectEmail(false); // Set to true if you want respondent emails

    Logger.log("✅ Form created: " + form.getEditUrl());

    // ── Step 3: Add each question to the form ─────────────
    for (var i = 0; i < questions.length; i++) {
      var q = questions[i];
      Logger.log("  Adding question " + (i + 1) + ": [" + q.type + "] " + q.question);
      addQuestionToForm(form, q);
    }

    Logger.log("✅ All " + questions.length + " questions added to form.");

    // ── Step 4: Link responses to a Google Sheet ──────────
    var responseSheet = linkResponsesToSheet(form);

    // ── Step 5: Output results ────────────────────────────
    var formUrl       = form.getPublishedUrl();
    var formEditUrl   = form.getEditUrl();
    var responseSheetUrl = responseSheet.getUrl();

    Logger.log("─────────────────────────────────────────");
    Logger.log("🎉 FORM GENERATION COMPLETE!");
    Logger.log("📋 Form Title:       " + CONFIG.FORM_TITLE);
    Logger.log("🔗 Share this link:  " + formUrl);
    Logger.log("✏️  Edit form here:   " + formEditUrl);
    Logger.log("📊 Responses sheet:  " + responseSheetUrl);
    Logger.log("─────────────────────────────────────────");

    // ── Step 6: Send completion email ─────────────────────
    if (CONFIG.SEND_EMAIL_ON_COMPLETION) {
      sendCompletionEmail(formUrl, formEditUrl, responseSheetUrl, questions.length);
    }

    // Show a popup in the editor
    SpreadsheetApp.getUi().alert(
      "✅ Form Created!\n\n" +
      "Share link: " + formUrl + "\n\n" +
      "Responses will be saved to: " + responseSheetUrl
    );

  } catch (error) {
    Logger.log("❌ Error during form generation: " + error.message);
    Logger.log("Stack: " + error.stack);
  }
}


// ─────────────────────────────────────────────────────────────
// FUNCTION: readQuestionsFromSheet
// Opens your Google Sheet and reads all question rows.
// Returns an array of question objects.
// ─────────────────────────────────────────────────────────────

/**
 * Reads question data from the configured Google Sheet.
 * @returns {Array} Array of question objects.
 */
function readQuestionsFromSheet() {
  // Open the spreadsheet
  var spreadsheet = SpreadsheetApp.openById(CONFIG.QUESTIONS_SHEET_ID);

  // Get the specified tab
  var sheet = spreadsheet.getSheetByName(CONFIG.QUESTIONS_TAB_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet tab "' + CONFIG.QUESTIONS_TAB_NAME + '" not found. ' +
      'Check CONFIG.QUESTIONS_TAB_NAME matches exactly.'
    );
  }

  // Get all data (skip row 1 which is the header)
  var data = sheet.getDataRange().getValues();

  var questions = [];

  // Loop from row index 1 (skip header row at index 0)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Skip completely empty rows
    if (!row[0] || String(row[0]).trim() === "") {
      continue;
    }

    // Map columns to named properties
    var questionObj = {
      question : String(row[0]).trim(),          // Column A: Question text
      type     : String(row[1]).trim().toLowerCase(), // Column B: Question type
      options  : row[2] ? String(row[2]).trim() : "", // Column C: Options (comma-separated)
      required : String(row[3]).toUpperCase() === "TRUE", // Column D: TRUE / FALSE
      helpText : row[4] ? String(row[4]).trim() : "",    // Column E: Help/description text
    };

    questions.push(questionObj);
  }

  return questions;
}


// ─────────────────────────────────────────────────────────────
// FUNCTION: addQuestionToForm
// Takes a question object and adds the correct item type
// to the Google Form.
// ─────────────────────────────────────────────────────────────

/**
 * Adds a single question to the given form based on its type.
 * @param {GoogleAppsScript.Forms.Form} form - The form to add the question to.
 * @param {Object} q - The question object with .question, .type, .options, etc.
 */
function addQuestionToForm(form, q) {
  var item;

  switch (q.type) {

    // ── Single-line text ──────────────────────────────
    case "short_answer":
      item = form.addTextItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      break;

    // ── Multi-line text ───────────────────────────────
    case "paragraph":
      item = form.addParagraphTextItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      break;

    // ── Radio buttons (pick one) ──────────────────────
    case "multiple_choice":
      item = form.addMultipleChoiceItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      if (q.options) {
        var mcChoices = parseOptions(q.options);
        item.setChoiceValues(mcChoices);
      }
      break;

    // ── Checkboxes (pick many) ────────────────────────
    case "checkbox":
      item = form.addCheckboxItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      if (q.options) {
        var cbChoices = parseOptions(q.options);
        item.setChoiceValues(cbChoices);
      }
      break;

    // ── Dropdown (pick one from list) ─────────────────
    case "dropdown":
      item = form.addListItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      if (q.options) {
        var ddChoices = parseOptions(q.options);
        item.setChoiceValues(ddChoices);
      }
      break;

    // ── Linear scale (e.g. 1–5 rating) ───────────────
    // Options format: "minValue, maxValue, minLabel, maxLabel"
    // Example:        "1, 5, Bad, Excellent"
    case "linear_scale":
      item = form.addScaleItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      if (q.options) {
        var scaleParts = parseOptions(q.options);
        var minVal   = parseInt(scaleParts[0]) || 1;
        var maxVal   = parseInt(scaleParts[1]) || 5;
        var minLabel = scaleParts[2] || "";
        var maxLabel = scaleParts[3] || "";
        item.setBounds(minVal, maxVal);
        item.setLabels(minLabel, maxLabel);
      }
      break;

    // ── Date picker ───────────────────────────────────
    case "date":
      item = form.addDateItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      break;

    // ── Time picker ───────────────────────────────────
    case "time":
      item = form.addTimeItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      break;

    // ── Unknown type ──────────────────────────────────
    default:
      Logger.log(
        "⚠️  Unknown type '" + q.type + "' for question: " + q.question +
        ". Defaulting to short_answer."
      );
      item = form.addTextItem();
      item.setTitle(q.question);
      item.setRequired(q.required);
      if (q.helpText) item.setHelpText(q.helpText);
      break;
  }
}


// ─────────────────────────────────────────────────────────────
// FUNCTION: linkResponsesToSheet
// Creates a new Google Sheet and links it to the form so
// all responses are saved there automatically.
// ─────────────────────────────────────────────────────────────

/**
 * Links the form's responses to a new Google Spreadsheet.
 * @param {GoogleAppsScript.Forms.Form} form - The form to link.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The response spreadsheet.
 */
function linkResponsesToSheet(form) {
  var responseSheetName = CONFIG.FORM_TITLE + " — Responses";
  var responseSpreadsheet = SpreadsheetApp.create(responseSheetName);

  // Link the form to this spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSpreadsheet.getId());

  Logger.log("✅ Responses linked to sheet: " + responseSpreadsheet.getUrl());
  return responseSpreadsheet;
}


// ─────────────────────────────────────────────────────────────
// HELPER: parseOptions
// Splits a comma-separated string of options into an array.
// Trims whitespace from each option.
// ─────────────────────────────────────────────────────────────

/**
 * Parses a comma-separated options string into a trimmed array.
 * @param {string} optionsStr - e.g. "Red, Blue, Green"
 * @returns {string[]} e.g. ["Red", "Blue", "Green"]
 */
function parseOptions(optionsStr) {
  return optionsStr.split(",").map(function (opt) {
    return opt.trim();
  }).filter(function (opt) {
    return opt !== "";
  });
}


// ─────────────────────────────────────────────────────────────
// FUNCTION: sendCompletionEmail
// Sends a summary email when the form is generated.
// ─────────────────────────────────────────────────────────────

/**
 * Sends an email notification with the form and response sheet links.
 */
function sendCompletionEmail(formUrl, formEditUrl, responseSheetUrl, questionCount) {
  var subject = "✅ Your Google Form is Ready — " + CONFIG.FORM_TITLE;
  var body =
    "Hello,\n\n" +
    "Your Google Form has been generated successfully!\n\n" +
    "─────────────────────────\n" +
    "📋 Form Title:      " + CONFIG.FORM_TITLE + "\n" +
    "❓ Questions Added: " + questionCount + "\n\n" +
    "🔗 Share this link with respondents:\n" + formUrl + "\n\n" +
    "✏️  Edit the form:\n" + formEditUrl + "\n\n" +
    "📊 View responses here:\n" + responseSheetUrl + "\n" +
    "─────────────────────────\n\n" +
    "— Automated by Google Apps Script";

  GmailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL, subject, body);
  Logger.log("📧 Completion email sent to: " + CONFIG.NOTIFICATION_EMAIL);
}


// ─────────────────────────────────────────────────────────────
// FUNCTION: testWithSampleData
// Creates a form using fake built-in data — no Excel needed.
// Run this first to verify everything works before using
// your real spreadsheet.
// ─────────────────────────────────────────────────────────────

/**
 * Generates a form using hard-coded sample questions.
 * Use this to test the script without needing your real Sheet.
 */
function testWithSampleData() {
  Logger.log("🧪 Running test with sample data...");

  // Fake question data (mimics what would be read from your Excel)
  var sampleQuestions = [
    { question: "What is your full name?",          type: "short_answer",    options: "",                        required: true,  helpText: "" },
    { question: "Your email address",               type: "short_answer",    options: "",                        required: true,  helpText: "We will not spam you" },
    { question: "Which department are you in?",     type: "dropdown",        options: "HR, Engineering, Sales, Marketing, Finance", required: true,  helpText: "" },
    { question: "How did you hear about us?",       type: "multiple_choice", options: "Social Media, Friend, Search Engine, Event", required: false, helpText: "" },
    { question: "Select your areas of interest",   type: "checkbox",        options: "Technology, Design, Business, Marketing",   required: false, helpText: "Pick all that apply" },
    { question: "Rate your overall experience",    type: "linear_scale",    options: "1, 5, Very Poor, Excellent",               required: true,  helpText: "" },
    { question: "Tell us about yourself",          type: "paragraph",       options: "",                        required: false, helpText: "Optional — 200 words max" },
    { question: "Date of your last visit",         type: "date",            options: "",                        required: false, helpText: "" },
  ];

  var form = FormApp.create("🧪 Test Form — Sample Data");
  form.setDescription("This is a test form generated by the Apps Script.");

  for (var i = 0; i < sampleQuestions.length; i++) {
    addQuestionToForm(form, sampleQuestions[i]);
  }

  var responseSheet = linkResponsesToSheet(form);

  Logger.log("─────────────────────────────────────────");
  Logger.log("✅ TEST FORM CREATED!");
  Logger.log("🔗 Share link: " + form.getPublishedUrl());
  Logger.log("✏️  Edit link:  " + form.getEditUrl());
  Logger.log("📊 Responses:  " + responseSheet.getUrl());
  Logger.log("─────────────────────────────────────────");
}
