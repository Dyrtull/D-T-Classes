// ============================================================
// Code.gs  -  Exam Auto-Grader Backend
// Paste this entire file into your Google Apps Script editor,
// replace PASTE_YOUR_KEY_HERE with your Gemini API key,
// then deploy as a new Web App version (Anyone can access).
// ============================================================

var GEMINI_API_KEY  = 'PASTE_YOUR_KEY_HERE';
var SUBMISSIONS_SHEET = 'Submissions';
var DETAILS_SHEET     = 'Detailed_Answers';
var QUESTIONS_SHEET   = 'Questions';
var CONFIG_SHEET      = 'Config';

var OPEN_MAX_PER_Q = 8;
var YGG_MAX_PER_Q  = 10;

// 1-based column indices for the Submissions sheet
var COL = {
  TIMESTAMP:    1,
  NAME:         2,
  CLASS:        3,
  SET:          4,
  MC_SCORE:     5,
  OPEN_SCORE:   6,
  YGG_SCORE:    7,
  PENALTY:      8,
  TAB_SWITCHES: 9,
  TOTAL_SCORE:  10,
  MAX_SCORE:    11,
  PERCENTAGE:   12,
  FINAL_GRADE:  13,
  BONUS:        14,
  OPEN_FB:      15,
  YGG_FB:       16,
  Q1_AI:        17,
  Q2_AI:        18,
  Q3_AI:        19,
  Q4_AI:        20,
  Q5_AI:        21,
  Q1_OVR:       22,
  Q2_OVR:       23,
  Q3_OVR:       24,
  Q4_OVR:       25,
  Q5_OVR:       26,
  YGG1_AI:      27,
  YGG2_AI:      28,
  YGG3_AI:      29,
  YGG1_OVR:     30,
  YGG2_OVR:     31,
  YGG3_OVR:     32
};

var GRADE_BOUNDARIES = {
  A: [
    {grade: '2.3', min: 85}, {grade: '2.0', min: 70},
    {grade: '1.7', min: 50}, {grade: '1.3', min: 35}, {grade: '1.0', min: 0}
  ],
  B: [
    {grade: '3.3', min: 85}, {grade: '3.0', min: 75}, {grade: '2.7', min: 65},
    {grade: '2.3', min: 55}, {grade: '2.0', min: 45}, {grade: '1.7', min: 35},
    {grade: '1.3', min: 20}, {grade: '1.0', min: 0}
  ],
  C: [
    {grade: '4.0', min: 90}, {grade: '3.7', min: 80}, {grade: '3.3', min: 70},
    {grade: '3.0', min: 60}, {grade: '2.7', min: 50}, {grade: '2.3', min: 40},
    {grade: '2.0', min: 30}, {grade: '1.7', min: 20}, {grade: '1.3', min: 10}, {grade: '1.0', min: 0}
  ]
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Exam Tools')
    .addItem('Setup: Create Config & Questions sheets', 'setup')
    .addSeparator()
    .addItem('Recalculate Grades with Overrides', 'recalculateWithOverrides')
    .addSeparator()
    .addItem('Backfill AI Scores (for old rows)', 'backfillAIScores')
    .addToUi();
}

// ============================================================
// GET HANDLER
// ============================================================
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action === 'submissions') { return getSubmissionsResponse(); }
  if (action === 'config')      { return getConfigResponse(); }
  return getQuestionsResponse();
}

// ============================================================
// GET: QUESTIONS + CONFIG
// ============================================================
function getQuestionsResponse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) { configSheet = createConfigSheet(ss); }
  var configRows = configSheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < configRows.length; i++) {
    var k = String(configRows[i][0]).trim();
    var v = String(configRows[i][1]).trim();
    if (k) { config[k] = v; }
  }

  var qSheet = ss.getSheetByName(QUESTIONS_SHEET);
  if (!qSheet) { qSheet = createQuestionsSheet(ss); }
  var qRows = qSheet.getDataRange().getValues();

  var questions = {A: [], B: [], C: [], Yggdrasil: []};

  for (var r = 1; r < qRows.length; r++) {
    var row = qRows[r];
    var set     = String(row[0]).trim();
    var id      = String(row[1]).trim();
    var section = String(row[2]).trim();
    var type    = String(row[3]).trim();
    var points  = Number(row[4]);
    var text    = String(row[5]).trim();
    var optA    = String(row[6]).trim();
    var optB    = String(row[7]).trim();
    var optC    = String(row[8]).trim();
    var optD    = String(row[9]).trim();
    var correctIndex = (row[10] !== '' && row[10] !== null && row[10] !== undefined) ? Number(row[10]) : null;
    var rubric  = String(row[11]).trim();

    if (!set || !id || !text) { continue; }

    var q = {id: id, section: section, type: type, points: points, text: text, rubric: rubric};

    if (type === 'mc') {
      var opts = [optA, optB, optC, optD];
      var indices = [0, 1, 2, 3];
      shuffleIndices(indices);
      q.options      = opts;
      q.correctIndex = correctIndex;
      q.shuffled     = indices;
    }

    if (set === 'Yggdrasil') {
      questions.Yggdrasil.push(q);
    } else if (questions[set] !== undefined) {
      questions[set].push(q);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({success: true, questions: questions, config: config}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// GET: CONFIG ONLY  (lightweight - for admin panel)
// ============================================================
function getConfigResponse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) {
    return ContentService
      .createTextOutput(JSON.stringify({config: {}}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var configRows = configSheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < configRows.length; i++) {
    var k = String(configRows[i][0]).trim();
    var v = String(configRows[i][1]).trim();
    if (k) { config[k] = v; }
  }
  return ContentService
    .createTextOutput(JSON.stringify({config: config}))
    .setMimeType(ContentService.MimeType.JSON);
}

function shuffleIndices(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
}

// ============================================================
// GET: SUBMISSIONS  (for admin.html)
// ============================================================
function getSubmissionsResponse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({submissions: []}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var submissions = [];

  var AI_COLS  = [COL.Q1_AI,   COL.Q2_AI,   COL.Q3_AI,   COL.Q4_AI,   COL.Q5_AI];
  var OVR_COLS = [COL.Q1_OVR,  COL.Q2_OVR,  COL.Q3_OVR,  COL.Q4_OVR,  COL.Q5_OVR];
  var YAI_COLS = [COL.YGG1_AI, COL.YGG2_AI, COL.YGG3_AI];
  var YOVR_COLS= [COL.YGG1_OVR,COL.YGG2_OVR,COL.YGG3_OVR];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var q1ai = row[COL.Q1_AI - 1];
    var hasPerQ = (q1ai !== '' && q1ai !== null && q1ai !== undefined);

    var qScores = [];
    for (var qi = 0; qi < 5; qi++) {
      var ai  = row[AI_COLS[qi] - 1];
      var ovr = row[OVR_COLS[qi] - 1];
      if (ai !== '' && ai !== null && ai !== undefined) {
        qScores.push({
          aiScore:  Number(ai),
          override: (ovr !== '' && ovr !== null && ovr !== undefined) ? Number(ovr) : null,
          maxScore: OPEN_MAX_PER_Q
        });
      }
    }

    var yScores = [];
    for (var yi = 0; yi < 3; yi++) {
      var yai  = row[YAI_COLS[yi] - 1];
      var yovr = row[YOVR_COLS[yi] - 1];
      yScores.push({
        aiScore:  (yai !== '' && yai !== null && yai !== undefined) ? Number(yai) : 0,
        override: (yovr !== '' && yovr !== null && yovr !== undefined) ? Number(yovr) : null,
        maxScore: YGG_MAX_PER_Q
      });
    }

    var maxScore = Number(row[COL.MAX_SCORE - 1]) || 0;
    var openMax  = qScores.length * OPEN_MAX_PER_Q;
    var mcMax    = maxScore - openMax;

    submissions.push({
      timestamp:         String(row[COL.TIMESTAMP - 1]),
      name:              String(row[COL.NAME - 1]),
      class:             String(row[COL.CLASS - 1]),
      set:               String(row[COL.SET - 1]),
      mcScore:           Number(row[COL.MC_SCORE - 1])    || 0,
      mcMax:             mcMax,
      openScore:         Number(row[COL.OPEN_SCORE - 1])  || 0,
      openMax:           openMax,
      penalty:           Number(row[COL.PENALTY - 1])     || 0,
      scoreAfterPenalty: Number(row[COL.TOTAL_SCORE - 1]) || 0,
      maxScore:          maxScore,
      percentage:        Number(row[COL.PERCENTAGE - 1])  || 0,
      finalGrade:        String(row[COL.FINAL_GRADE - 1]),
      bonus:             Number(row[COL.BONUS - 1])       || 0,
      hasPerQuestionData: hasPerQ,
      questionScores:    qScores,
      yggScores:         yScores
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({submissions: submissions}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// POST HANDLER
// ============================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'override')      { return handleOverride(data); }
    if (data.action === 'updateConfig')  { return handleUpdateConfig(data); }
    return handleSubmission(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// HANDLE EXAM SUBMISSION
// ============================================================
function handleSubmission(data) {
  var openEnded    = data.openEnded    || [];
  var yggAnswers   = data.yggAnswers   || [];
  var mcAnswers    = data.mcAnswers    || [];
  var set          = data.set          || 'A';
  var mcScore      = Number(data.mcScore)      || 0;
  var penaltyPoints= Number(data.penaltyPoints) || 0;
  var tabSwitches  = Number(data.tabSwitches)   || 0;

  // Grade open-ended questions
  var openFeedback     = [];
  var openAIScores     = [];
  var totalOpenScore   = 0;
  var openFeedbackParts= [];

  for (var i = 0; i < openEnded.length; i++) {
    var q      = openEnded[i];
    var result = gradeWithGemini(q.text, q.rubric, q.points);
    openFeedback.push({questionId: q.questionId, score: result.score, maxScore: q.points, feedback: result.feedback});
    openAIScores.push(result.score);
    totalOpenScore += result.score;
    openFeedbackParts.push(q.questionId + ': ' + result.score + '/' + q.points + ' - ' + result.feedback);
  }

  // Grade Yggdrasil
  var yggFeedback      = [];
  var yggAIScores      = [];
  var yggScore         = 0;
  var yggFeedbackParts = [];
  var yggBonus         = 0;

  for (var j = 0; j < yggAnswers.length; j++) {
    var yq     = yggAnswers[j];
    var yResult= {score: 0, feedback: 'Not attempted.'};
    if (yq.text && yq.text.trim().length > 2) {
      yResult = gradeWithGemini(yq.text, yq.rubric, yq.points);
    }
    yggFeedback.push({questionId: yq.questionId, score: yResult.score, maxScore: yq.points, feedback: yResult.feedback});
    yggAIScores.push(yResult.score);
    yggScore += yResult.score;
    yggFeedbackParts.push(yq.questionId + ': ' + yResult.score + '/' + yq.points + ' - ' + yResult.feedback);
    if (yResult.score >= 0.6 * yq.points) { yggBonus += 0.3; }
  }

  // Calculate totals
  var mcMax   = 0;
  for (var mi = 0; mi < mcAnswers.length; mi++)   { mcMax   += mcAnswers[mi].points; }
  var openMax = 0;
  for (var oi = 0; oi < openEnded.length; oi++)   { openMax += openEnded[oi].points; }
  var totalMax = mcMax + openMax;

  var rawScore        = mcScore + totalOpenScore;
  var scoreAfterPenalty = Math.max(0, rawScore - penaltyPoints);
  var percentage      = totalMax > 0 ? Math.round((scoreAfterPenalty / totalMax) * 100) : 0;
  var baseGrade       = calculateGrade(percentage, set);
  var bonus           = Math.round(yggBonus * 10) / 10;
  var finalGradeNum   = Math.min(5.0, parseFloat(baseGrade) + bonus);
  var finalGrade      = finalGradeNum.toFixed(1);

  var timestamp = new Date().toISOString();

  writeSummaryToSheet({
    timestamp:    timestamp,
    studentName:  data.studentName,
    studentClass: data.class,
    set:          set,
    mcScore:      mcScore,
    openScore:    totalOpenScore,
    yggScore:     yggScore,
    penalty:      penaltyPoints,
    tabSwitches:  tabSwitches,
    scoreAfterPenalty: scoreAfterPenalty,
    maxScore:     totalMax,
    percentage:   percentage,
    finalGrade:   finalGrade,
    bonus:        bonus,
    openFeedback: openFeedbackParts.join('\n'),
    yggFeedback:  yggFeedbackParts.join('\n'),
    openAIScores: openAIScores,
    yggAIScores:  yggAIScores
  });

  return ContentService
    .createTextOutput(JSON.stringify({
      success:      true,
      studentName:  data.studentName,
      set:          set,
      totalScore:   scoreAfterPenalty,
      maxScore:     totalMax,
      percentage:   percentage,
      finalGrade:   finalGrade,
      penalty:      penaltyPoints,
      bonus:        bonus,
      openFeedback: openFeedback,
      yggFeedback:  yggFeedback
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// GEMINI GRADING
// ============================================================
function gradeWithGemini(studentAnswer, rubric, maxPoints) {
  if (!studentAnswer || studentAnswer.trim().length < 2) {
    return {score: 0, feedback: 'No answer provided.'};
  }

  var prompt =
    'You are a strict but fair grading assistant for a high school Design & Technology class.\n\n' +
    'RUBRIC: ' + rubric + '\n' +
    'STUDENT ANSWER: ' + studentAnswer + '\n' +
    'MAX POINTS: ' + maxPoints + '\n\n' +
    'Respond in EXACTLY this format (nothing else):\n' +
    'SCORE: [number from 0 to ' + maxPoints + ']\n' +
    'FEEDBACK: [1-2 sentences]';

  var responseText = callGeminiAPI(prompt);

  var scoreMatch   = responseText.match(/SCORE:\s*(\d+)/);
  var feedbackMatch= responseText.match(/FEEDBACK:\s*([\s\S]+)/);

  var score    = scoreMatch    ? Math.min(maxPoints, Math.max(0, parseInt(scoreMatch[1]))) : 0;
  var feedback = feedbackMatch ? feedbackMatch[1].trim() : responseText.substring(0, 200);

  return {score: score, feedback: feedback};
}

function callGeminiAPI(prompt) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;
  var payload = JSON.stringify({
    contents: [{parts: [{text: prompt}]}],
    generationConfig: {temperature: 0.1, maxOutputTokens: 200}
  });
  var options = {method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true};
  var delays  = [2000, 4000, 8000];

  for (var attempt = 0; attempt <= delays.length; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      return json.candidates[0].content.parts[0].text;
    }

    if ((code === 429 || code === 503) && attempt < delays.length) {
      Utilities.sleep(delays[attempt]);
    } else {
      break;
    }
  }

  return 'SCORE: 0\nFEEDBACK: Grading service error.';
}

// ============================================================
// WRITE SUBMISSION TO SHEET
// ============================================================
function writeSummaryToSheet(d) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) { sheet = ss.insertSheet(SUBMISSIONS_SHEET); }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp','Name','Class','Set','MC Score','Open Score','Ygg Score',
      'Penalty','Tab Switches','Total Score','Max Score','Percentage','Final Grade','Bonus',
      'Open Feedback','Ygg Feedback',
      'Q1_AI','Q2_AI','Q3_AI','Q4_AI','Q5_AI',
      'Q1_Override','Q2_Override','Q3_Override','Q4_Override','Q5_Override',
      'YGG1_AI','YGG2_AI','YGG3_AI',
      'YGG1_Override','YGG2_Override','YGG3_Override'
    ]);
  }

  var row = new Array(32);
  for (var x = 0; x < 32; x++) { row[x] = ''; }

  row[COL.TIMESTAMP    - 1] = d.timestamp;
  row[COL.NAME         - 1] = d.studentName;
  row[COL.CLASS        - 1] = d.studentClass;
  row[COL.SET          - 1] = d.set;
  row[COL.MC_SCORE     - 1] = d.mcScore;
  row[COL.OPEN_SCORE   - 1] = d.openScore;
  row[COL.YGG_SCORE    - 1] = d.yggScore;
  row[COL.PENALTY      - 1] = d.penalty;
  row[COL.TAB_SWITCHES - 1] = d.tabSwitches;
  row[COL.TOTAL_SCORE  - 1] = d.scoreAfterPenalty;
  row[COL.MAX_SCORE    - 1] = d.maxScore;
  row[COL.PERCENTAGE   - 1] = d.percentage;
  row[COL.FINAL_GRADE  - 1] = d.finalGrade;
  row[COL.BONUS        - 1] = d.bonus;
  row[COL.OPEN_FB      - 1] = d.openFeedback;
  row[COL.YGG_FB       - 1] = d.yggFeedback;

  var AI_COLS  = [COL.Q1_AI,   COL.Q2_AI,   COL.Q3_AI,   COL.Q4_AI,   COL.Q5_AI];
  var YAI_COLS = [COL.YGG1_AI, COL.YGG2_AI, COL.YGG3_AI];

  var oScores = d.openAIScores || [];
  for (var i = 0; i < oScores.length && i < 5; i++) { row[AI_COLS[i] - 1] = oScores[i]; }

  var yScores = d.yggAIScores || [];
  for (var j = 0; j < yScores.length && j < 3; j++) { row[YAI_COLS[j] - 1] = yScores[j]; }

  sheet.appendRow(row);
}

// ============================================================
// HANDLE OVERRIDE (from admin.html)
// ============================================================
function handleOverride(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: 'Sheet not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var allData  = sheet.getDataRange().getValues();
  var targetRow= -1;

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][COL.TIMESTAMP - 1]) === String(data.timestamp)) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: 'Submission not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var row      = allData[targetRow - 1];
  var qScores  = data.questionScores || [];
  var yScores  = data.yggScores      || [];

  var OVR_COLS = [COL.Q1_OVR,   COL.Q2_OVR,   COL.Q3_OVR,   COL.Q4_OVR,   COL.Q5_OVR];
  var YOVR_COLS= [COL.YGG1_OVR, COL.YGG2_OVR, COL.YGG3_OVR];

  for (var qi = 0; qi < 5; qi++) {
    var val = (qScores[qi] !== null && qScores[qi] !== undefined) ? qScores[qi] : '';
    row[OVR_COLS[qi] - 1] = val;
    sheet.getRange(targetRow, OVR_COLS[qi]).setValue(val);
  }
  for (var yi = 0; yi < 3; yi++) {
    var yval = (yScores[yi] !== null && yScores[yi] !== undefined) ? yScores[yi] : '';
    row[YOVR_COLS[yi] - 1] = yval;
    sheet.getRange(targetRow, YOVR_COLS[yi]).setValue(yval);
  }

  var updated = recalcRow(row);

  sheet.getRange(targetRow, COL.OPEN_SCORE).setValue(updated.openScore);
  sheet.getRange(targetRow, COL.YGG_SCORE).setValue(updated.yggScore);
  sheet.getRange(targetRow, COL.TOTAL_SCORE).setValue(updated.totalScore);
  sheet.getRange(targetRow, COL.PERCENTAGE).setValue(updated.percentage);
  sheet.getRange(targetRow, COL.FINAL_GRADE).setValue(updated.finalGrade);
  sheet.getRange(targetRow, COL.BONUS).setValue(updated.bonus);

  return ContentService
    .createTextOutput(JSON.stringify({success: true, finalGrade: updated.finalGrade, percentage: updated.percentage}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HANDLE CONFIG UPDATE  (from admin.html)
// ============================================================
function handleUpdateConfig(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: 'Config sheet not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var rows = sheet.getDataRange().getValues();
  var updates = data.updates || {};

  for (var i = 0; i < rows.length; i++) {
    var key = String(rows[i][0]).trim();
    if (updates.hasOwnProperty(key)) {
      sheet.getRange(i + 1, 2).setValue(String(updates[key]));
      delete updates[key];
    }
  }
  // Append any keys that don't exist yet
  for (var k in updates) {
    sheet.appendRow([k, String(updates[k])]);
  }
  return ContentService
    .createTextOutput(JSON.stringify({success: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHARED RECALCULATION HELPER
// ============================================================
function recalcRow(row) {
  var set      = String(row[COL.SET - 1]);
  var mcScore  = Number(row[COL.MC_SCORE - 1])  || 0;
  var penalty  = Number(row[COL.PENALTY - 1])   || 0;
  var maxScore = Number(row[COL.MAX_SCORE - 1]) || 60;

  var AI_COLS  = [COL.Q1_AI,   COL.Q2_AI,   COL.Q3_AI,   COL.Q4_AI,   COL.Q5_AI];
  var OVR_COLS = [COL.Q1_OVR,  COL.Q2_OVR,  COL.Q3_OVR,  COL.Q4_OVR,  COL.Q5_OVR];
  var YAI_COLS = [COL.YGG1_AI, COL.YGG2_AI, COL.YGG3_AI];
  var YOVR_COLS= [COL.YGG1_OVR,COL.YGG2_OVR,COL.YGG3_OVR];

  var openScore = 0;
  for (var qi = 0; qi < 5; qi++) {
    var ai  = row[AI_COLS[qi] - 1];
    var ovr = row[OVR_COLS[qi] - 1];
    if (ai !== '' && ai !== null && ai !== undefined) {
      var effective = (ovr !== '' && ovr !== null && ovr !== undefined) ? Number(ovr) : Number(ai);
      openScore += effective;
    }
  }

  var yggScore = 0;
  var yggBonus = 0;
  for (var yi = 0; yi < 3; yi++) {
    var yai  = row[YAI_COLS[yi] - 1];
    var yovr = row[YOVR_COLS[yi] - 1];
    if (yai !== '' && yai !== null && yai !== undefined) {
      var yEff = (yovr !== '' && yovr !== null && yovr !== undefined) ? Number(yovr) : Number(yai);
      yggScore += yEff;
      if (yEff >= 0.6 * YGG_MAX_PER_Q) { yggBonus += 0.3; }
    }
  }

  var bonus         = Math.round(yggBonus * 10) / 10;
  var totalScore    = Math.max(0, mcScore + openScore - penalty);
  var percentage    = maxScore > 0 ? Math.round((totalScore / maxScore) * 100) : 0;
  var baseGrade     = calculateGrade(percentage, set);
  var finalGradeNum = Math.min(5.0, parseFloat(baseGrade) + bonus);

  return {
    openScore:  openScore,
    yggScore:   yggScore,
    totalScore: totalScore,
    percentage: percentage,
    finalGrade: finalGradeNum.toFixed(1),
    bonus:      bonus
  };
}

// ============================================================
// GRADE LOOKUP
// ============================================================
function calculateGrade(percentage, set) {
  var boundaries = GRADE_BOUNDARIES[set] || GRADE_BOUNDARIES['A'];
  for (var i = 0; i < boundaries.length; i++) {
    if (percentage >= boundaries[i].min) { return boundaries[i].grade; }
  }
  return '1.0';
}

// ============================================================
// MENU ACTION: RECALCULATE ALL ROWS
// ============================================================
function recalculateWithOverrides() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('No Submissions sheet found.'); return; }

  var data    = sheet.getDataRange().getValues();
  var changed = 0;

  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var q1ai = row[COL.Q1_AI - 1];
    if (q1ai === '' || q1ai === null || q1ai === undefined) { continue; }

    var updated  = recalcRow(row);
    var sheetRow = i + 1;

    sheet.getRange(sheetRow, COL.OPEN_SCORE).setValue(updated.openScore);
    sheet.getRange(sheetRow, COL.YGG_SCORE).setValue(updated.yggScore);
    sheet.getRange(sheetRow, COL.TOTAL_SCORE).setValue(updated.totalScore);
    sheet.getRange(sheetRow, COL.PERCENTAGE).setValue(updated.percentage);
    sheet.getRange(sheetRow, COL.FINAL_GRADE).setValue(updated.finalGrade);
    sheet.getRange(sheetRow, COL.BONUS).setValue(updated.bonus);
    changed++;
  }

  SpreadsheetApp.getUi().alert('Done. Recalculated ' + changed + ' submission(s).');
}

// ============================================================
// MENU ACTION: BACKFILL AI SCORES FOR OLD ROWS
// ============================================================
function backfillAIScores() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('No Submissions sheet found.'); return; }

  var data   = sheet.getDataRange().getValues();
  var filled = 0;
  var regex  = /([A-Z]+_\d+):\s*(\d+)\/(\d+)/g;

  var AI_COLS  = [COL.Q1_AI,   COL.Q2_AI,   COL.Q3_AI,   COL.Q4_AI,   COL.Q5_AI];
  var YAI_COLS = [COL.YGG1_AI, COL.YGG2_AI, COL.YGG3_AI];

  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var q1ai = row[COL.Q1_AI - 1];
    if (q1ai !== '' && q1ai !== null && q1ai !== undefined) { continue; }

    var openFb = String(row[COL.OPEN_FB - 1] || '');
    var yggFb  = String(row[COL.YGG_FB  - 1] || '');
    if (!openFb) { continue; }

    var sheetRow = i + 1;
    var qIdx = 0;
    var yIdx = 0;
    var match;

    regex.lastIndex = 0;
    while ((match = regex.exec(openFb)) !== null) {
      var id = match[1];
      if (id.indexOf('YGG') !== -1) { continue; }
      if (qIdx < 5) {
        sheet.getRange(sheetRow, AI_COLS[qIdx]).setValue(parseInt(match[2]));
        qIdx++;
      }
    }

    regex.lastIndex = 0;
    while ((match = regex.exec(yggFb)) !== null) {
      if (yIdx < 3) {
        sheet.getRange(sheetRow, YAI_COLS[yIdx]).setValue(parseInt(match[2]));
        yIdx++;
      }
    }

    if (qIdx > 0) { filled++; }
  }

  SpreadsheetApp.getUi().alert('Done. Backfilled AI scores for ' + filled + ' submission(s).');
}

// ============================================================
// MANUAL SETUP  -  run this from the Apps Script editor
// to create the Config and Questions tabs immediately.
// ============================================================
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG_SHEET)) {
    createConfigSheet(ss);
    console.log('Config sheet created.');
  } else {
    console.log('Config sheet already exists - skipped.');
  }
  if (!ss.getSheetByName(QUESTIONS_SHEET)) {
    createQuestionsSheet(ss);
    console.log('Questions sheet created with 33 questions.');
  } else {
    console.log('Questions sheet already exists - skipped.');
  }
  console.log('Setup complete.');
}

// ============================================================
// AUTO-CREATE: CONFIG SHEET
// ============================================================
function createConfigSheet(ss) {
  if (!ss) { ss = SpreadsheetApp.getActiveSpreadsheet(); }
  var sheet = ss.insertSheet(CONFIG_SHEET);
  sheet.appendRow(['exam_title',            'Smart Vertical Garden - Session 5']);
  sheet.appendRow(['exam_subtitle',         'Select your question set to begin.']);
  sheet.appendRow(['class_options',         '10A,10B,10C']);
  sheet.appendRow(['exam_duration_minutes', '60']);
  sheet.appendRow(['exam_active',           'true']);
  return sheet;
}

// ============================================================
// AUTO-CREATE: QUESTIONS SHEET  (33 questions pre-loaded)
// ============================================================
function createQuestionsSheet(ss) {
  if (!ss) { ss = SpreadsheetApp.getActiveSpreadsheet(); }
  var sheet = ss.insertSheet(QUESTIONS_SHEET);
  sheet.appendRow(['Set','ID','Section','Type','Points','Text','Option_A','Option_B','Option_C','Option_D','CorrectIndex','Rubric']);

  var rows = [
    // --- SET A ---
    ['A','A_1','A','openEnded',8,
     '1. What component acts like a "kink in a hose" to limit the flow of electricity?',
     '','','','','',
     'Award 8 points for identifying "Resistor".'],
    ['A','A_2','A','openEnded',8,
     '2. The battery pack in your Smart Garden supplies DC power. What does DC stand for?',
     '','','','','',
     'Award 8 points for identifying "Direct Current".'],
    ['A','A_3','A','openEnded',8,
     '3. What unit of measurement is used to measure electrical pressure (Voltage)?',
     '','','','','',
     'Award 8 points for identifying "Volts" or "V".'],
    ['A','A_4','A','openEnded',8,
     '4. What happens if you install an LED into your circuit backwards?',
     '','','','','',
     'Award 8 points for stating it will not turn on or light up, because LEDs only allow current in one direction.'],
    ['A','A_5','A','openEnded',8,
     '5. Write the standard mathematical formula for Ohm\'s Law.',
     '','','','','',
     'Award 8 points for V = I * R (or V = IR).'],
    ['A','A_6','B','mc',4,
     '6. Which circuit symbol is drawn as a plain rectangle?',
     'Battery','Switch','Resistor','LED',2,
     'Award 4 points for Resistor (index 2).'],
    ['A','A_7','B','mc',4,
     '7. What do the colored bands painted on a fixed resistor tell you?',
     'Its resistance value and tolerance','How much voltage it produces','Which way to plug it in','If it uses AC or DC power',0,
     'Award 4 points for "Its resistance value and tolerance".'],
    ['A','A_8','B','mc',4,
     '8. To measure voltage, the multimeter probes must be connected:',
     'In parallel (across the component)','In series (breaking the circuit)','Only when the power is OFF','Directly to the wall socket',0,
     'Award 4 points for "In parallel (across the component)".'],
    ['A','A_9','B','mc',4,
     '9. Which component acts as the main power source for the Smart Garden?',
     'The Resistor','The LED','The 7.4V Battery Pack','The Multimeter',2,
     'Award 4 points for "The 7.4V Battery Pack".'],
    ['A','A_10','B','mc',4,
     '10. What is the function of a switch in a circuit?',
     'To create electrons','To change DC to AC','To break or complete the path of electricity','To change colors of the LED',2,
     'Award 4 points for "To break or complete the path of electricity".'],

    // --- SET B ---
    ['B','B_1','A','openEnded',8,
     '1. Your plant monitoring system is powered by a 7.4V battery pack. Total resistance is 330 ohms. Calculate the exact current flowing through the system.',
     '','','','','',
     'Award 4 points for formula (I=V/R) and values. Award 4 points for answer: 0.022A or 22mA.'],
    ['B','B_2','A','openEnded',8,
     '2. You find a resistor painted Brown, Black, Red, Gold. Calculate its nominal value and state the acceptable minimum and maximum resistance based on tolerance.',
     '','','','','',
     'Award 4 pts for nominal value (1000 ohms / 1k ohm). Award 4 pts for tolerance range: 950 ohms to 1050 ohms.'],
    ['B','B_3','A','openEnded',8,
     '3. A student reads 0.00A on a circuit that is switched ON. Write a checklist of three physical connections they should check.',
     '','','','','',
     'Award 8 pts for 3 valid physical checks (LED backwards, battery disconnected, switch seated wrong, meter not in series). 4 pts for 1-2 checks.'],
    ['B','B_4','A','openEnded',8,
     '4. Match each scenario to Voltage, Current, or Resistance:\na) Water pump blocked by debris.\nb) Connecting multimeter without breaking circuit.\nc) Removing component completely to test safely.',
     '','','','','',
     'Award 8 pts total. a) Resistance (3 pts). b) Voltage (3 pts). c) Resistance (2 pts).'],
    ['B','B_5','A','openEnded',8,
     '5. Explain, using the water hose analogy, why adding a second resistor in series makes the LED dimmer.',
     '','','','','',
     'Award 8 points for explaining a resistor is a kink in the hose, and a second kink further restricts flow (current), leaving less to power the LED.'],
    ['B','B_6','B','mc',4,
     '6. The switch is closed, but the LED does not turn on. Most likely cause?',
     'The LED long leg (anode) is connected to 0V.','Resistor is backwards.','Battery is supplying DC.','Switch is bridging the gap.',0,
     'Award 4 points for "The LED long leg (anode) is connected to 0V."'],
    ['B','B_7','B','mc',4,
     '7. To measure voltage across a single resistor in a functioning circuit, you must:',
     'Connect probes across the resistor while power is ON.','Break circuit and insert in series.','Turn power OFF and remove resistor.','Switch to A setting.',0,
     'Award 4 points for "Connect probes across the resistor while power is ON."'],
    ['B','B_8','B','mc',4,
     '8. A design calls for 150 ohms, but you only have standard E12 values. Which component do you select?',
     '150 ohms','120 ohms','180 ohms','100 ohms',0,
     'Award 4 points for 150 ohms (it IS a standard E12 value).'],
    ['B','B_9','B','mc',4,
     '9. Why do we strictly use DC power in the garden projects instead of AC?',
     'DC flows in one direction and is a safe voltage.','AC does not work with resistors.','DC reverses direction, safer for plants.','AC changes LED colors.',0,
     'Award 4 points for "DC flows in one direction and is a safe voltage."'],
    ['B','B_10','B','mc',4,
     '10. You want to adjust light sensitivity daily without tools. Which component is correct?',
     'Potentiometer (Variable Resistor)','Fixed Resistor','Pre-set Resistor','470 ohm Resistor',0,
     'Award 4 points for "Potentiometer (Variable Resistor)".'],

    // --- SET C ---
    ['C','C_1','A','openEnded',8,
     '1. You need an LED to drop 2.0V from a 7.4V battery pack. Current must be 20mA (0.020A). Calculate the exact resistance needed, then state the E12 resistor to use.',
     '','','','','',
     'Award 4 pts for voltage drop (5.4V). Award 2 pts for ideal resistance (270 ohms). Award 2 pts for noting 270 ohms IS an E12 value.'],
    ['C','C_2','A','openEnded',8,
     '2. You measure a Red-Red-Brown-Gold resistor. The meter reads 215 ohms. Prove mathematically if it is safe to use.',
     '','','','','',
     'Award 4 pts for nominal value (220 ohms +/- 5%). Award 4 pts for calculating range (209 to 231 ohms) and concluding 215 ohms is safe.'],
    ['C','C_3','A','openEnded',8,
     '3. Draw or describe the exact probe placement to measure current flowing through the LED. What must be done to the physical wires?',
     '','','','','',
     'Award 4 pts for stating the circuit must be broken. Award 4 pts for describing the multimeter in SERIES (flow goes into one probe and out the other).'],
    ['C','C_4','A','openEnded',8,
     '4. Match each incorrect color code translation to the mistake made:\na) "Brown-Black-Red" read as 102 ohms.\nb) "Red-Red-Brown" read as 2200 ohms.\nc) 470 ohm resistor read backwards as Brown-Violet-Yellow.',
     '','','','','',
     'Award 8 pts total. a) Band 3 read as a digit instead of a multiplier (3 pts). b) Multiplied by 100 instead of 10 (3 pts). c) Read from the tolerance end (2 pts).'],
    ['C','C_5','A','openEnded',8,
     '5. A classmate sets the meter to ohms and probes a resistor while the 7.4V battery is still ON. Describe two consequences.',
     '','','','','',
     'Award 4 pts for stating the reading will be inaccurate or false. Award 4 pts for stating it will likely blow the meter fuse.'],
    ['C','C_6','B','mc',4,
     '6. A circuit requires 315 ohms. You use an E12 330 ohm resistor. How will this affect the current?',
     'Current will decrease slightly.','Current will increase slightly.','Voltage drops to zero.','Tolerance changes to +/-10%.',0,
     'Award 4 points for "Current will decrease slightly."'],
    ['C','C_7','B','mc',4,
     '7. You measure 7.4V across an OPEN switch and 0V when the switch is CLOSED. This means:',
     'Normal electrical behavior.','Switch is broken.','Battery is wired backwards.','Meter blew a fuse.',0,
     'Award 4 points for "Normal electrical behavior."'],
    ['C','C_8','B','mc',4,
     '8. A student\'s LED flashes briefly then permanently dies. Which resistor color code was most likely used?',
     'Brown-Black-Black-Gold (10 ohms)','Yellow-Violet-Brown-Gold (470 ohms)','Red-Red-Red-Gold (2.2k ohms)','Brown-Black-Orange-Gold (10k ohms)',0,
     'Award 4 points for "Brown-Black-Black-Gold (10 ohms)" - too low a resistance caused overcurrent.'],
    ['C','C_9','B','mc',4,
     '9. If the multiplier band (Band 3) is Orange, the resistor value is in the:',
     'Thousands of Ohms (kOhm)','Tens of Ohms','Hundreds of Ohms','Millions of Ohms (MOhm)',0,
     'Award 4 points for "Thousands of Ohms (kOhm)".'],
    ['C','C_10','B','mc',4,
     '10. In a series circuit, if you swap the position of the resistor and the LED:',
     'Circuit functions exactly the same.','LED burns out immediately.','Current flows in reverse.','Switch stops controlling the LED.',0,
     'Award 4 points for "Circuit functions exactly the same."'],

    // --- YGGDRASIL (The Bonus Challenge) ---
    ['Yggdrasil','YGG_1','Y','openEnded',10,
     'A student receives a failing grade because their 12mA soil moisture sensor burned out. They powered it with 7.4V using a Brown-Black-Brown-Gold resistor. Prove mathematically why it failed.',
     '','','','','',
     'Decoding resistor: 100 ohms. Current: I = 7.4 / 100 = 74mA. Conclude 74mA destroys a 12mA sensor. Full marks for all three steps.'],
    ['Yggdrasil','YGG_2','Y','openEnded',10,
     'Reverse-engineer: you need to drop 5V across a resistor at exactly 10mA. Calculate the required resistance, select the best E12 value, and write the exact four color bands.',
     '','','','','',
     'R = V/I = 5 / 0.010 = 500 ohms. Nearest E12 value: 470 ohms. Color bands: Yellow, Violet, Brown, Gold. Full marks for all three steps.'],
    ['Yggdrasil','YGG_3','Y','openEnded',10,
     'A fully charged 8.4V battery powers a 2.0V LED. Calculate the minimum E12 resistor value to ensure the current NEVER exceeds 25mA.',
     '','','','','',
     'Voltage drop: 8.4 - 2.0 = 6.4V. Resistance: 6.4 / 0.025 = 256 ohms. Nearest HIGHER E12 value (for safety): 270 ohms. Full marks for all three steps.']
  ];

  for (var i = 0; i < rows.length; i++) {
    sheet.appendRow(rows[i]);
  }

  return sheet;
}
