function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index');
  html.startupPage = e.parameter.page || 'entry'; 
  return html.evaluate().setTitle('Inbound Stats @YIC');
}

function getAllCountryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countriesSheet = ss.getSheetByName("Countries");
  if (!countriesSheet) {
    throw new Error("The 'Countries' sheet was not found. Please ensure it exists.");
  }
  
  const data = countriesSheet.getDataRange().getValues();
  const result = [];
  
  // Assuming headers are in row 1, so we start reading from row 2 (index 1)
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const area = data[i][1];
    const isFrequent = data[i][2] === true; // Ensure it's a true boolean
    
    if (name && area) {
      result.push({ name: name, area: area, isFrequent: isFrequent });
    }
  }
  return result;
}

// --- HELPER FUNCTIONS for Fiscal Year Management ---

/**
 * Determines the fiscal year sheet name for a given date (April-March).
 * @param {Date} date The date to check.
 * @return {string} The name of the sheet, e.g., "".
 */
function getFiscalYearSheetName(date) {
  const year = date.getFullYear();
  const month = date.getMonth(); // 0 = January, 3 = April
  const fiscalYear = (month < 3) ? year - 1 : year;
  return `${fiscalYear}`;
}

/**
 * Checks if a sheet exists, creates it with headers if not.
 * @param {string} sheetName The name of the sheet to ensure exists.
 * @return {Sheet} The Google Sheet object.
 */
function ensureSheetExists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, 0); // Insert as first sheet for easy access
    const headers = [["Timestamp", "Region", "Country", "Number of Visitors", "Inquiry Details", "Accommodation"]];
    sheet.getRange("A2:F2").setValues(headers).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat('@'); // Set Timestamp column to Plain Text for ISO strings
  }
  return sheet;
}

// --- DATA ENTRY AND DAILY LOG FUNCTIONS ---

/**
 * Appends a new visitor entry to the correct fiscal year sheet.
 */
function submitVisitorData(area, country, numberOfVisitors, inquiryDetails, accommodation) {
  const timestamp = new Date();
  const sheetName = getFiscalYearSheetName(timestamp);
  const sheet = ensureSheetExists(sheetName);
  
  sheet.appendRow([
    timestamp.toISOString(), area, country,
    numberOfVisitors, accommodation || "", inquiryDetails || "",
  ]);
}

/**
 * Retrieves all entries from the past 7 days.
 * This function can handle date ranges that span across two fiscal year sheets.
 */
function getRecentEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allResults = [];
  
  // 1. Define the 14-day date range
  const now = new Date();
  const fourteenDaysAgo = new Date();
  fourteenDaysAgo.setDate(now.getDate() - 7);
  fourteenDaysAgo.setHours(0, 0, 0, 0); // Start of the 7-day window

  // 2. Determine which sheet(s) to search
  const currentFYSheetName = getFiscalYearSheetName(now);
  const previousFYSheetName = getFiscalYearSheetName(fourteenDaysAgo);

  const sheetNamesToSearch = [currentFYSheetName];
  // If the 7-day window crosses the fiscal year boundary (April 1st), add the previous year's sheet
  if (currentFYSheetName !== previousFYSheetName) {
    sheetNamesToSearch.push(previousFYSheetName);
  }
  
  // 3. Loop through the necessary sheets and gather matching records
  for (const sheetName of sheetNamesToSearch) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue; // Skip if a sheet for a past fiscal year doesn't exist

    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) { // Assumes header on row 2
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows
      
      const recordTimestamp = new Date(row[0]);
      
      // Check if the record's date falls within our 14-day window
      if (recordTimestamp >= fourteenDaysAgo && recordTimestamp <= now) {
        allResults.push({
          timestamp: recordTimestamp.toISOString(),
          area: row[1],
          country: row[2],
          visitors: row[3],
          accommodation: row[4] || "",
          inquiryDetails: row[5] || "",
        });
      }
    }
  }
  
  // 4. Sort all combined results by newest first before returning
  allResults.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
  
  return allResults;
}

/**
 * Finds and edits an entry within the entry's corresponding fiscal year sheet.
 */
function editEntry(timestampStr, updatedEntry) {
  const timestamp = new Date(timestampStr);
  const sheetName = getFiscalYearSheetName(timestamp);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Data sheet for the corresponding year (${sheetName}) not found.`);

  const textFinder = sheet.getRange("A:A").createTextFinder(timestampStr).matchEntireCell(true);
  const foundCell = textFinder.findNext();

  if (foundCell) {
    const row = foundCell.getRow();
    sheet.getRange(row, 2, 1, 5).setValues([[
      updatedEntry.area, 
      updatedEntry.country, 
      updatedEntry.visitors,
      updatedEntry.accommodation || "", 
      updatedEntry.inquiryDetails || "",
    ]]);
    return true;
  }
  throw new Error("Could not find the entry to update.");
}

/**
 * Finds and deletes an entry within the entry's corresponding fiscal year sheet.
 */
function deleteEntry(timestampStr) {
  const timestamp = new Date(timestampStr);
  const sheetName = getFiscalYearSheetName(timestamp);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Data sheet for the corresponding year (${sheetName}) not found.`);

  const textFinder = sheet.getRange("A:A").createTextFinder(timestampStr).matchEntireCell(true);
  const foundCell = textFinder.findNext();

  if (foundCell) {
    sheet.deleteRow(foundCell.getRow());
    return true;
  }
  throw new Error("Entry not found.");
}

// --- REPORTING AND ANALYTICS FUNCTIONS ---

/**
 * The main calculation engine for monthly summaries for any given month.
 */
function calculateMonthlySummary(year, month_0_based) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countriesSheet = ss.getSheetByName("Countries");
  
  const reportDate = new Date(year, month_0_based, 15);
  const dataSheetName = getFiscalYearSheetName(reportDate);
  const dataSheet = ss.getSheetByName(dataSheetName);
  
  const countryInfo = {};
  const countryData = countriesSheet.getDataRange().getValues();
  for (let i = 1; i < countryData.length; i++) { // Assumes headers on row 1 of Countries sheet
    const name = countryData[i][0];
    if (name) {
      countryInfo[name] = {
        region: countryData[i][1],
        isFrequent: countryData[i][2] === true,
        order: countryData[i][3] || 999
      };
    }
  }

  const monthlyTotals = {};
  if (dataSheet) {
    const allData = dataSheet.getDataRange().getValues();
    if (allData.length > 2) {
      const header = allData[1];
      const dataRows = allData.slice(2);
      const timestampIdx = header.indexOf("Timestamp");
      const countryIdx = header.indexOf("Country");
      const visitorsIdx = header.indexOf("Number of Visitors");
      if(timestampIdx === -1 || countryIdx === -1 || visitorsIdx === -1) { throw new Error(`Required headers not found in ${dataSheetName}`); }

      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (!row[timestampIdx]) continue;
        const timestamp = new Date(row[timestampIdx]);
        if (timestamp.getFullYear() === year && timestamp.getMonth() === month_0_based) {
          const countryName = row[countryIdx];
          const visitors = Number(row[visitorsIdx]) || 0;
          monthlyTotals[countryName] = (monthlyTotals[countryName] || 0) + visitors;
        }
      }
    }
  }

  const summaryByRegion = {};
  for (const countryName in countryInfo) {
    const info = countryInfo[countryName];
    const region = info.region;
    const visitors = monthlyTotals[countryName] || 0;

    if (!summaryByRegion[region]) {
      summaryByRegion[region] = {};
      if (region !== "Unknown") {
        const restKey = `Rest of ${region}`;
        summaryByRegion[region][restKey] = { country: restKey, visitors: 0, isFrequent: false, order: 9999 };
      }
    }
    
    if (info.isFrequent) {
      summaryByRegion[region][countryName] = { country: countryName, visitors: visitors, isFrequent: true, order: info.order };
    } else {
      if (region !== "Unknown") {
        const restKey = `Rest of ${region}`;
        summaryByRegion[region][restKey].visitors += visitors;
      }
    }
  }

  const finalOutput = {};
  for (const region in summaryByRegion) {
    finalOutput[region] = Object.values(summaryByRegion[region]).sort((a, b) => a.order - b.order);
  }
  
  return finalOutput;
}

/**
 * Gets the summary for any month specified by the frontend.
 */
function getSummaryForMonth(year, month_1_based) {
  return calculateMonthlySummary(year, month_1_based - 1);
}

/**
 * Searches through inquiry records for a specific keyword.
 */
function searchInquiries(keyword, fiscalYear) {
  if (!keyword) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToSearch = [];

  if (fiscalYear === "All") {
    const allSheets = ss.getSheets();
    for (const sheet of allSheets) {
    const sheetName = `${fiscalYear}`;
    const specificSheet = ss.getSheetByName(sheetName);
    if (specificSheet) sheetsToSearch.push(specificSheet);
  if (sheetsToSearch.length === 0) return [];

  const allResults = [];
  const INQUIRY_COLUMN_LETTER = 'E';

  for (const sheet of sheetsToSearch) {
    const textFinder = sheet.getRange(INQUIRY_COLUMN_LETTER + ":" + INQUIRY_COLUMN_LETTER)
                           .createTextFinder(keyword)
                           .ignoreCase(true)
                           .findAll();
    
    for (const range of textFinder) {
      const rowNum = range.getRow();
      const rowData = sheet.getRange(rowNum, 1, 1, 6).getValues()[0];
      allResults.push({
        timestamp: rowData[0], // Column A
        country: rowData[2],   // Column C
        inquiryDetails: rowData[4], // Column E
        accommodation: rowData[5] // Column F
      });
    }
  }
  
  allResults.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
  return allResults;
}}
}
