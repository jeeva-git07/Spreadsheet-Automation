/**
 * Main function to run the full CSAT report, creating four pivot tables 
 * (Asia, China, Europe and America sites) and generating a dynamic summary 
 * table next to each pivot.
 *
 * NOTE: Assumes the raw data is on a sheet named "Sheet1".
 */
function runCSATReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert("Error: Data sheet 'Sheet1' not found.");
    return;
  }
  
  // Set up the output sheet
  let pivotSheet = ss.getSheetByName("Pivot Summary");
  if (!pivotSheet) pivotSheet = ss.insertSheet("Pivot Summary");
  pivotSheet.clear();
  
  // --- Configuration ---
  const lr=dataSheet.getLastRow()
  const lc=dataSheet.getLastColumn()
  const sourceRange = dataSheet.getRange(3,1,lr-2,lc);
  const keepTeams = ['Endpoint', 'Network Services', 'Server And Datacenter'];
  
  // Column indices for the source data (assuming 1-based indexing for Apps Script methods)
  const COL_TEAM = 9; // Team
  const COL_SURVEY_RATING = 12; // Survey RatingRest (Pivot Column Group / Value)
  const COL_IT_CENTER = 14; // IT Center (Filter)
  
  // Variable to track the starting row for the next pivot table. Starts at row 1 (A1).
  let nextStartRow = 1;
  
  // --- 1. Generate ASIA Report (Starts at A1) ---
  const asiaSites = ['IN', 'JP', 'KR', 'TH'];
  nextStartRow = createPivotAndSummary(
    asiaSites, 
    `A${nextStartRow}`, 
    pivotSheet, 
    sourceRange, 
    COL_TEAM, 
    COL_SURVEY_RATING, 
    COL_IT_CENTER, 
    keepTeams
  ) + 2;
  
  // --- 2. Generate CHINA Report ---
  const chinaSites = ['CN'];
  nextStartRow = createPivotAndSummary(
    chinaSites, 
    `A${nextStartRow}`, 
    pivotSheet, 
    sourceRange, 
    COL_TEAM, 
    COL_SURVEY_RATING, 
    COL_IT_CENTER, 
    keepTeams
  ) + 2;

  // --- 3. Generate AMERICA Report ---
  const americaSites = ['BR','MX','US'];
  nextStartRow = createPivotAndSummary(
    americaSites, 
    `A${nextStartRow}`, 
    pivotSheet, 
    sourceRange, 
    COL_TEAM, 
    COL_SURVEY_RATING, 
    COL_IT_CENTER, 
    keepTeams
  ) + 2;

  // --- 4. Generate EUROPE Report ---
  const europeSites = ['CZ','DE','PL'];
  nextStartRow = createPivotAndSummary(
    europeSites, 
    `A${nextStartRow}`, 
    pivotSheet, 
    sourceRange, 
    COL_TEAM, 
    COL_SURVEY_RATING, 
    COL_IT_CENTER, 
    keepTeams
  ) + 2;

  // Final cleanup and formatting for the sheet
  pivotSheet.autoResizeColumns(1, pivotSheet.getMaxColumns());
}

/**
 * Creates a single pivot table and generates its corresponding summary table immediately next to it.
 *
 * @param {string[]} sites The list of IT Center values to filter (e.g., ['IN', 'JP']).
 * @param {string} startCellA1 The A1 notation for where the pivot table should start (e.g., "A1").
 * @param {GoogleAppsScript.Spreadsheet.Sheet} pivotSheet The sheet to write the pivot table to.
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRange The range of the source data.
 * @param {number} colTeam Index of the Team column (e.g., 9).
 * @param {number} colSurveyRating Index of the Survey Rating column (e.g., 12).
 * @param {number} colItCenter Index of the IT Center column (e.g., 14).
 * @param {string[]} keepTeams The list of team names to filter (e.g., ['Endpoint', ...]).
 * @returns {number} The row number immediately following the generated pivot table.
 */
function createPivotAndSummary(sites, startCellA1, pivotSheet, sourceRange, colTeam, colSurveyRating, colItCenter, keepTeams) {
  
  // 1. CREATE PIVOT TABLE
  
  const pivotTable = pivotSheet.getRange(startCellA1).createPivotTable(sourceRange);

  // Set Row Groups (Team)
  pivotTable.addRowGroup(colTeam);
  // Set Column Groups (Survey RatingRest)
  pivotTable.addColumnGroup(colSurveyRating);
  // Set Value (Count of Survey RatingRest)
  pivotTable.addPivotValue(colSurveyRating, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);

  // Filter 1: IT Center
  const siteCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(sites)
    .build();
  pivotTable.addFilter(colItCenter, siteCriteria);

  // Filter 2: Specific Teams
  const teamCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(keepTeams)
    .build();
  pivotTable.addFilter(colTeam, teamCriteria);
  
  // The pivot table is now created on the sheet. We need to read its actual data.
  // Note: App Script pivot tables often place the result in the second or third row,
  // but getDataRegion() handles finding the full, compiled table size.
  
  // Wait briefly (optional, but sometimes helps ensure the pivot data is fully rendered)
  SpreadsheetApp.flush(); 

  // Get the range containing the generated pivot data
  const pivotRange = pivotSheet.getRange(startCellA1).getDataRegion();
  const pivotData = pivotRange.getValues();
  
  // If the pivot table is empty or malformed, exit
  if (pivotData.length < 3 || pivotData[0].length < 2) {
    Logger.log(`Pivot table starting at ${startCellA1} is empty or too small.`);
    return pivotSheet.getLastRow(); // Return current last row
  }
  
  // 2. GENERATE DYNAMIC SUMMARY
  
  // The header row containing the rating names (3, 4, 5, Null, Grand Total) is usually the second row.
  const headers = pivotData[1]; 
  
  // Dynamically Find Column Indices
  const ratingCols = [];
  let grandTotalCol = -1;
  const numericRegex = /^\d+$/;

  for (let i = 1; i < headers.length; i++) { // Start from index 1 to skip the 'Team' column label
    const headerStr = String(headers[i]).trim();
    
    if (numericRegex.test(headerStr)) {
      ratingCols.push({
        index: i,
        rating: parseInt(headerStr)
      });
    } else if (headerStr.toLowerCase().includes('grand total')) {
      grandTotalCol = i;
    }
  }

  if (ratingCols.length === 0 || grandTotalCol === -1) {
    Logger.log(`Rating columns or 'Grand Total' column not found in headers for table at ${startCellA1}.`);
    return pivotRange.getLastRow();
  }
  
  // Pivot table body data (excluding header rows (0 and 1) and the final Grand Total row)
  const bodyData = pivotData.slice(2, pivotData.length - 1);
  
  const summaryHeaders = [
    "SUM", 
    "Response %", 
    "Rating", 
    "Low Rating", 
    "Low Response", 
    "Feedback Requested", 
    "Feedback Received"
  ];
  
  const summaryBody = [];
  
  // Helper function for weighted average calculation
  function calculateWeightedAvg(rowData, ratingCols) {
    let total = 0;
    let count = 0;
    
    ratingCols.forEach(col => {
      // Ensure the value is treated as a number (use 0 for null/empty)
      const val = parseFloat(rowData[col.index]) || 0; 
      total += val * col.rating;
      count += val;
    });
    
    return count > 0 ? (total / count) : 0;
  }

  // Process Team rows
  let totalSum = 0;
  let totalFeedbackReceived = 0;
  let totalFeedbackRequested = 0;
  
  bodyData.forEach(row => {
    // Feedback Received = sum of all numeric rating counts
    const feedbackReceived = ratingCols.reduce((sum, col) => sum + (parseFloat(row[col.index]) || 0), 0);
    
    // SUM = weighted total (count × rating value)
    const sum = ratingCols.reduce((sum, col) => sum + ((parseFloat(row[col.index]) || 0) * col.rating), 0);
    
    // Feedback Requested = Grand Total column value (Rated + Null)
    const feedbackRequested = parseFloat(row[grandTotalCol]) || 0; 
    
    // Response % calculation: Feedback Received / Grand Total
    const responsePercent = (feedbackRequested > 0) ? (feedbackReceived / feedbackRequested) : 0;
    
    // Rating (Weighted Average)
    const rating = calculateWeightedAvg(row, ratingCols);
    
    // Accumulate totals for the final Grand Total row
    totalSum += sum;
    totalFeedbackReceived += feedbackReceived;
    totalFeedbackRequested += feedbackRequested;
    
    const summaryRow = [
      sum,
      responsePercent, 
      rating, 
      4, // Low Rating constant
      15, // Low Response constant
      feedbackRequested,
      feedbackReceived
    ];
    
    summaryBody.push(summaryRow);
  });
  
  // Calculate and add the Grand Total row for the summary table
  const grandTotalResponsePercent = (totalFeedbackRequested > 0) ? 
    (totalFeedbackReceived / totalFeedbackRequested) : 0;
    
  const grandTotalRating = (totalFeedbackReceived > 0) ? 
    (totalSum / totalFeedbackReceived) : 0;
    
  const summaryGrandTotalRow = [
    totalSum,
    grandTotalResponsePercent, 
    grandTotalRating, 
    4,
    15,
    totalFeedbackRequested,
    totalFeedbackReceived
  ];
  
  // 3. WRITE SUMMARY TABLE TO SHEET

  const finalSummaryData = [summaryHeaders].concat(summaryBody).concat([summaryGrandTotalRow]);
  
  // Determine the starting column for the summary table. User requested column J (index 10).
  const requestedStartCol = 10; 
  
  // Ensure the summary does not overlap the pivot table if the pivot table is wider than column I (9).
  const pivotEndCol = pivotRange.getLastColumn();
  const dynamicStartCol = pivotEndCol + 2; 
  
  // Use column J (10) unless the pivot table is too wide, in which case use the dynamic placement.
  const startCol = Math.max(requestedStartCol, dynamicStartCol);
  
  // Determine the starting row for the summary table (same row as the pivot table start)
  const startRow = pivotRange.getRow();

  const outputRange = pivotSheet.getRange(startRow, startCol, finalSummaryData.length, finalSummaryData[0].length);
  outputRange.setValues(finalSummaryData);

  // 4. APPLY NUMBER FORMATTING & STYLING 
  
  // Header styling (Pivot Table ONLY)
  // We use pivotSheet.getRange() and limit the column span to the pivot's actual width.
  const pivotHeaderRange = pivotSheet.getRange(startRow, 1, 2, pivotRange.getLastColumn());
  pivotHeaderRange.setBackground('#4A708B').setFontColor('white').setFontWeight('bold');
  
  // Data range starts at row 3 of the pivot (index 2 in pivotData)
  const dataStartRow = startRow + 2;
  const dataNumRows = finalSummaryData.length - 1; // Exclude header row

  // Formatting for 'Response %' column (column index 1 of the summary data)
  const responsePercentCol = outputRange.offset(1, 1, dataNumRows, 1);
  responsePercentCol.setNumberFormat('0.0%'); 
  
  // Formatting for 'Rating' column (column index 2 of the summary data)
  const ratingCol = outputRange.offset(1, 2, dataNumRows, 1);
  ratingCol.setNumberFormat('0.0'); 
  
  // Formatting for integer columns
  const intCols = [0, 5, 6]; 
  intCols.forEach(colIndex => {
    const intColRange = outputRange.offset(1, colIndex, dataNumRows, 1);
    intColRange.setNumberFormat('0');
  });

  // Highlight the Grand Total row (last row of both tables)
  // This highlight still spans both tables since the Grand Total row is shared.
  pivotSheet.getRange(pivotRange.getLastRow(), 1, 1, outputRange.getLastColumn()).setFontWeight('bold');

  return pivotRange.getLastRow();
}
