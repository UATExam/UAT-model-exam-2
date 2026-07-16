/**
 * UAT MODEL EXAM — Google Apps Script (paste into Extensions > Apps Script
 * on your Google Sheet, then deploy as a Web App). See README.md for the
 * full step-by-step.
 *
 * This script receives a JSON POST from the exam website every time a
 * student finishes the Math section, and appends one row per attempt.
 */

const SHEET_NAME = "Results"; // change if you want a different tab name

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Add header row once
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Name", "City", "School",
        "English Score", "English Total",
        "Math Score", "Math Total",
        "Total Score", "Total Questions", "Percentage",
        "English Restarted?", "Math Restarted?", "Device"
      ]);
    }

    const data = JSON.parse(e.postData.contents);

    const total = Number(data.totalScore) || 0;
    const totalQ = Number(data.totalQuestions) || 100;
    const pct = totalQ > 0 ? Math.round((total / totalQ) * 100) + "%" : "";

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name || "",
      data.city || "",
      data.school || "",
      data.englishScore ?? "",
      data.englishTotal ?? 55,
      data.mathScore ?? "",
      data.mathTotal ?? 45,
      total,
      totalQ,
      pct,
      data.englishRestarted ? "Yes" : "No",
      data.mathRestarted ? "Yes" : "No",
      data.device || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: lets you sanity-check the deployment by visiting the Web App
// URL directly in a browser (GET request).
function doGet(e) {
  return ContentService
    .createTextOutput("UAT Model Exam results endpoint is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}
