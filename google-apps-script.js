/**
 * Google Apps Script for Sacred Pathways Quiz
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1mRChINgrxaiudooydofW1m9Yy4FQv9b_3xaD9hdUUFY
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click "Deploy" > "New deployment"
 * 5. Select type: "Web app"
 * 6. Set "Execute as": "Me"
 * 7. Set "Who has access": "Anyone"
 * 8. Click "Deploy" and authorize when prompted
 * 9. Copy the Web app URL - you'll need it for the quiz
 */

// Handle POST requests from the quiz
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // Ensure headers exist
    ensureHeaders(sheet);

    // Build the row data
    const row = [
      new Date().toISOString(),                    // Timestamp
      data.name || 'Anonymous',                     // Name
      // Top 5 pathways (name and score)
      data.top5[0]?.name || '', data.top5[0]?.score || '',
      data.top5[1]?.name || '', data.top5[1]?.score || '',
      data.top5[2]?.name || '', data.top5[2]?.score || '',
      data.top5[3]?.name || '', data.top5[3]?.score || '',
      data.top5[4]?.name || '', data.top5[4]?.score || '',
      // All pathway scores (for analysis)
      data.allScores.naturalist || 0,
      data.allScores.sensate || 0,
      data.allScores.traditionalist || 0,
      data.allScores.activist || 0,
      data.allScores.ascetic || 0,
      data.allScores.caregiver || 0,
      data.allScores.enthusiast || 0,
      data.allScores.contemplative || 0,
      data.allScores.intellectual || 0
    ];

    // Append the row
    sheet.appendRow(row);

    // Return success
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Results submitted successfully!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET requests (for testing)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Sacred Pathways Quiz API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ensure the sheet has proper headers
function ensureHeaders(sheet) {
  const firstCell = sheet.getRange('A1').getValue();
  if (firstCell !== 'Timestamp') {
    const headers = [
      'Timestamp',
      'Name',
      'Rank1_Pathway', 'Rank1_Score',
      'Rank2_Pathway', 'Rank2_Score',
      'Rank3_Pathway', 'Rank3_Score',
      'Rank4_Pathway', 'Rank4_Score',
      'Rank5_Pathway', 'Rank5_Score',
      'Naturalist', 'Sensate', 'Traditionalist', 'Activist', 'Ascetic',
      'Caregiver', 'Enthusiast', 'Contemplative', 'Intellectual'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

/**
 * BONUS: Helper function to get congregation statistics
 * Run this from the Apps Script editor to see aggregate data
 */
function getCongregationStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('No submissions yet');
    return;
  }

  const pathways = ['Naturalist', 'Sensate', 'Traditionalist', 'Activist', 'Ascetic',
                    'Caregiver', 'Enthusiast', 'Contemplative', 'Intellectual'];

  // Column indices for pathway scores (0-indexed, starting after Top5 columns)
  const startCol = 12; // Column M (index 12)

  const stats = {};
  pathways.forEach((p, i) => {
    const scores = data.slice(1).map(row => row[startCol + i] || 0);
    const validScores = scores.filter(s => s > 0);
    stats[p] = {
      count: validScores.length,
      total: validScores.reduce((a, b) => a + b, 0),
      average: validScores.length > 0 ? (validScores.reduce((a, b) => a + b, 0) / validScores.length).toFixed(1) : 0,
      timesRanked1: data.slice(1).filter(row => row[2] === p).length
    };
  });

  Logger.log('=== CONGREGATION PATHWAY STATISTICS ===');
  Logger.log('Total submissions: ' + (data.length - 1));
  Logger.log('');

  // Sort by average score
  const sorted = Object.entries(stats).sort((a, b) => b[1].average - a[1].average);
  sorted.forEach(([pathway, s]) => {
    Logger.log(pathway + ':');
    Logger.log('  Average score: ' + s.average + '/30');
    Logger.log('  Times ranked #1: ' + s.timesRanked1);
    Logger.log('');
  });

  return stats;
}
