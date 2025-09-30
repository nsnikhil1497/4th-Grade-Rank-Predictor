const SHEET_NAME = 'Final'; // Ensure your Google Sheet is named "ScoreData"

/**
 * Handles GET requests and serves the HTML file.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('4th Grade Rank Predictor')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Helper function to calculate Raw Score using the agreed-upon scheme.
 * Scheme: Correct (+1.66), Wrong (-0.55)
 */
function calculateRawScore(correct, wrong) {
  const positiveMark = 1.66;
  const negativeMark = 0.55; 

  const correctCount = parseInt(correct);
  const wrongCount = parseInt(wrong);

  let rawScore = (correctCount * positiveMark) - (wrongCount * negativeMark);

  // Rounding to two decimal places
  return Math.round(rawScore * 100) / 100;
}


/**
 * Processes the score submission from the form.
 */
function processSubmission(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return { success: false, message: 'Spreadsheet Error: Sheet "ScoreData" not found. Please create a sheet named "ScoreData".' };
  }

  const rollNo = formData.rollNo.trim();
  const rawScore = calculateRawScore(formData.correct, formData.wrong);

  // Check for duplicate submission
  const data = sheet.getDataRange().getValues();
  const existingEntry = data.find(row => row[1] == rollNo); // RollNo is at index 1
  if (existingEntry) {
    return { success: false, message: 'Roll No. already submitted. You cannot submit twice.' };
  }

  // Prepare new row data
  const newRow = [
    new Date(),
    rollNo,
    formData.name.trim(),
    formData.category,
    formData.shift,
    parseInt(formData.attempted),
    parseInt(formData.correct),
    parseInt(formData.wrong),
    rawScore
  ];
  
  // Append new data to the sheet
  sheet.appendRow(newRow);

  // Calculate ranks for the newly added entry
  const { overallRank, categoryRank, shiftRank } = calculateRanks(sheet, formData.category, formData.shift, rawScore);
  
  return {
    success: true,
    rawScore: rawScore,
    overallRank: overallRank,
    categoryRank: categoryRank,
    shiftRank: shiftRank
  };
}


/**
 * Searches for a candidate's rank by Roll No. (Used by the 'Check Current Rank' tab).
 */
function searchRankByRollNo(rollNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return { success: false, message: 'Spreadsheet Error: Sheet "ScoreData" not found.' };
  }

  const data = sheet.getDataRange().getValues();
  const allEntries = data.slice(1); 
  
  // Find the candidate
  const candidateEntry = allEntries.find(row => row[1] == rollNo); 

  if (!candidateEntry) {
    return { success: false, message: `No submission found for Roll No.: ${rollNo}` };
  }
  
  // Extract candidate data (data indices match the appended array in processSubmission)
  const name = candidateEntry[2];
  const category = candidateEntry[3];
  const shift = candidateEntry[4];
  const rawScore = candidateEntry[8];

  // Calculate ranks for the found entry
  const { overallRank, categoryRank, shiftRank } = calculateRanks(sheet, category, shift, rawScore);
  
  return {
    success: true,
    name: name,
    rawScore: rawScore,
    overallRank: overallRank,
    categoryRank: categoryRank,
    shiftRank: shiftRank
  };
}


/**
 * Helper function to calculate Overall, Category, and Shift ranks.
 * This ensures the rank is updated based on ALL current data in the sheet.
 */
function calculateRanks(sheet, category, shift, rawScore) {
  const data = sheet.getDataRange().getValues();
  const allEntries = data.slice(1); // Exclude header row

  // 1. Overall Rank
  const overallRank = getRank(allEntries, rawScore);

  // 2. Category Rank (Category is at index 3 in the Sheet row array)
  const categoryEntries = allEntries.filter(row => row[3] === category); 
  const categoryRank = getRank(categoryEntries, rawScore);

  // 3. Shift Rank (Shift is at index 4 in the Sheet row array)
  const shiftEntries = allEntries.filter(row => row[4] === shift); 
  const shiftRank = getRank(shiftEntries, rawScore);
  
  return {
    overallRank: overallRank,
    categoryRank: categoryRank,
    shiftRank: shiftRank
  };
}


/**
 * Calculates the rank ensuring same score gets the same rank (e.g., 1, 1, 3 for scores 10, 10, 9).
 * @param {Array<Array<any>>} entries - The filtered list of candidate entries.
 * @param {number} targetScore - The score of the candidate whose rank is being calculated.
 * @return {number} The calculated rank.
 */
function getRank(entries, targetScore) {
  // Sort entries by RawScore (index 8) in descending order.
  entries.sort((a, b) => b[8] - a[8]); 

  let rank = 1;
  
  for (let i = 0; i < entries.length; i++) {
    const currentScore = entries[i][8];
    
    // Check if the current score is less than the score before the current rank block
    if (i > 0 && currentScore < entries[i-1][8]) {
      // New distinct score, rank is current position + 1
      rank = i + 1;
    }
    
    // If we find the target score, return the calculated rank
    if (currentScore === targetScore) {
      return rank;
    }
  }
  
  // Should not happen if data is correct
  return entries.length + 1; 
}
