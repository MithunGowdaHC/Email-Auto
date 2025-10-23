// ============================================
// AMAZON ADS REPORT AUTOMATION SCRIPT
// With Excel (.xlsx) and CSV Support
// ============================================

// ⚠️ IMPORTANT SETUP STEPS:
// 1. Enable Drive API: In the Apps Script editor, click "+" next to Services → Add "Drive API" (select the v3 version if asked).
// 2. Share the mapping sheet with your email (the one running this script) with at least 'Viewer' access, 'Editor' is safer.
// 3. Run setupDailyTrigger() function manually *once* to schedule the daily execution.

// CONFIGURATION - UPDATE THESE VALUES
const CONFIG = {
  MAPPING_SHEET_URL: 'https://docs.google.com/spreadsheets/d/1Ui1i0oikCy7MHhJ_obRlwZ6OZZf_oR1ugDaJV6QaNrc/edit?gid=0#gid=0',
  SENDER_EMAIL: 'mithun.gowda@hiveminds.in',
  RECEIVER_EMAIL: 'mithungowdahc12a@gmail.com',
  CLIENT_EMAIL: 'mgowda8431@gmail.com',
  YOUR_EMAIL: 'mithungowdahc12a@gmail.com', // For error/success notifications
  DRIVE_FOLDER_ID: '1gQKp7tY5aKRx0LG5fmVQhWc5Qdl_2qks', // Optional: ID of a Google Drive folder to save collated reports. e.g. '1a2b3c4d5e6f_g7h8i9j'
  EMAIL_SUBJECT: 'Amazon Ads Performance Report - [DATE]'
};

// ============================================
// MAIN FUNCTION - RUN THIS DAILY VIA TRIGGER
// ============================================
function automateAmazonAdsReport() {
  try {
    Logger.log('🚀 Starting Amazon Ads Automation...');
    
    const reports = fetchReportsFromGmail();
    if (!reports || (reports.sp === null && reports.sb === null && reports.sd === null)) {
      Logger.log('⚠️ No valid reports found. Notifying user.');
      notifyNoReportsFound();
      return;
    }
    
    const collatedData = collateReports(reports);
    if (collatedData.length === 0) {
      Logger.log('⚠️ Reports found but collation resulted in 0 rows. Check column mappings or report content.');
      notifyError(new Error('Reports were found, but no data could be collated. Please check report formats and column mappings.'));
      return;
    }
    
    const enrichedData = applyMapping(collatedData);
    const pivotData = createCategoryPivot(enrichedData);
    
    saveCollatedFile(enrichedData);
    createDraftEmail(pivotData);
    notifyReportReady();
    
    Logger.log('✅ Automation completed successfully!');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    notifyError(error);
  }
}

// ============================================
// STEP 1: FETCH REPORTS FROM GMAIL
// ============================================
function fetchReportsFromGmail() {
  const searchQuery = `from:${CONFIG.SENDER_EMAIL} to:${CONFIG.RECEIVER_EMAIL} newer_than:1d has:attachment`;
  
  Logger.log('🔍 Searching emails: ' + searchQuery);
  
  // Get the first thread (most recent)
  const threads = GmailApp.search(searchQuery, 0, 1);
  
  if (threads.length === 0) {
    Logger.log('⚠️ No emails found from ' + CONFIG.SENDER_EMAIL);
    return null;
  }
  
  const messages = threads[0].getMessages();
  const latestMessage = messages[messages.length - 1]; // Get the last message in the thread
  const attachments = latestMessage.getAttachments();
  
  Logger.log('📧 Found email with subject: "' + latestMessage.getSubject() + '" from: ' + latestMessage.getFrom());
  Logger.log('📎 Number of attachments: ' + attachments.length);
  
  const reports = {
    sp: null,
    sb: null,
    sd: null
  };
  
  attachments.forEach(attachment => {
    const name = attachment.getName();
    Logger.log('📄 Processing attachment: ' + name);
    
    try {
      let data;
      
      // Check file type and parse accordingly
      if (name.endsWith('.csv')) {
        // Handle CSV files
        Logger.log('   → Detected CSV file');
        const content = attachment.getDataAsString();
        data = Utilities.parseCsv(content);
        
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        // Handle Excel files using Drive API v3
        Logger.log('   → Detected Excel file');
        const blob = attachment.copyBlob();
        
        // Create temp file in Drive
        // This creates a file in the root folder by default
        const tempFile = DriveApp.createFile(blob);
        const tempFileId = tempFile.getId();
        Logger.log('   → Created temp file: ' + tempFileId);
        
        // Convert to Google Sheets using Drive API v3
        const resource = {
          name: name + '_converted_' + new Date().getTime(),
          mimeType: MimeType.GOOGLE_SHEETS
        };
        
        // The Drive.Files.copy makes a *copy* of the file and converts it
        const fileJson = Drive.Files.copy(resource, tempFileId);
        Logger.log('   → Converted to GSheet: ' + fileJson.id);
        
        // Read data from converted sheet
        const ss = SpreadsheetApp.openById(fileJson.id);
        data = ss.getSheets()[0].getDataRange().getValues();
        
        // Clean up both files (original Excel and converted GSheet)
        Drive.Files.remove(tempFileId);
        Drive.Files.remove(fileJson.id);
        Logger.log('   → Cleaned up temp files');
        
      } else {
        Logger.log('   ⚠️ Unsupported file type: ' + name);
        return;
      }
      
      // Identify report type
      if (name.includes('Sponsored_Products') || name.includes('Sponsored Products')) {
        reports.sp = data;
        Logger.log('✅ Loaded Sponsored Products: ' + data.length + ' rows');
      } else if (name.includes('Sponsored_Brands') || name.includes('Sponsored Brands')) {
        reports.sb = data;
        Logger.log('✅ Loaded Sponsored Brands: ' + data.length + ' rows');
      } else if (name.includes('Sponsored_Display') || name.includes('Sponsored Display')) {
        reports.sd = data;
        Logger.log('✅ Loaded Sponsored Display: ' + data.length + ' rows');
      }
      
    } catch (e) {
      Logger.log('⚠️ Error processing ' + name + ': ' + e.toString());
      // Attempt to clean up temp files if they exist
      try {
        if (typeof tempFileId !== 'undefined') Drive.Files.remove(tempFileId);
        if (typeof fileJson !== 'undefined') Drive.Files.remove(fileJson.id);
      } catch (cleanupError) {
        Logger.log('⚠️ Error during cleanup: ' + cleanupError.toString());
      }
    }
  });
  
  Logger.log('✅ Reports loaded: SP=' + !!reports.sp + ', SB=' + !!reports.sb + ', SD=' + !!reports.sd);
  return reports;
}

// ============================================
// STEP 2: COLLATE REPORTS
// ============================================
function collateReports(reports) {
  const collated = [];
  
  // Column mappings - Adjust these 0-based indexes based on your report structure
  // Example: 'date: 0' means 'Date' is in column A
  const columnMaps = {
    sp: {
      date: 0,        // Col A
      portfolioName: 1, // Col B
      campaignName: 3,  // Col D
      impressions: 9,   // Col J
      clicks: 11,       // Col L
      spend: 14,        // Col O
      orders: 17,       // Col R
      sales: 20         // Col U
    },
    sb: {
      date: 0,        // Col A
      portfolioName: 1, // Col B
      campaignName: 3,  // Col D
      impressions: 6,   // Col G
      clicks: 7,        // Col H
      spend: 10,        // Col K
      orders: 14,       // Col O
      sales: 13         // Col N
    },
    sd: {
      date: 0,        // Col A
      portfolioName: 5, // Col F
      campaignName: 4,  // Col E
      impressions: 7,   // Col H
      clicks: 9,        // Col J
      spend: 11,        // Col L
      orders: 18,       // Col S
      sales: 20         // Col U
    }
  };
  
  Object.keys(reports).forEach(reportType => {
    const report = reports[reportType];
    if (!report || report.length < 2) {
      Logger.log(`ℹ️ Skipping ${reportType}: no data found.`);
      return;
    }
    
    const colMap = columnMaps[reportType];
    
    // Log headers for verification
    Logger.log(`📊 ${reportType.toUpperCase()} headers (first 5): ${report[0].slice(0, 5).join(', ')}`);
    
    // Skip header row (start at i = 1)
    for (let i = 1; i < report.length; i++) {
      const row = report[i];
      
      // Skip empty/total rows - check for a campaign name
      if (!row[colMap.campaignName] || row[colMap.campaignName] === 'Total') continue;
      
      collated.push({
        date: row[colMap.date],
        campaignName: row[colMap.campaignName],
        portfolioName: row[colMap.portfolioName],
        impressions: parseNumber(row[colMap.impressions]),
        clicks: parseNumber(row[colMap.clicks]),
        spend: parseNumber(row[colMap.spend]),
        sales: parseNumber(row[colMap.sales]),
        orders: parseNumber(row[colMap.orders])
      });
    }
    Logger.log(`📈 Collated ${collated.length} rows from ${reportType.toUpperCase()}`);
  });
  
  Logger.log('✅ Total collated rows: ' + collated.length);
  return collated;
}

// ============================================
// STEP 3: APPLY MAPPING (VLOOKUP)
// ============================================
function applyMapping(data) {
  try {
    const mappingSheet = SpreadsheetApp.openByUrl(CONFIG.MAPPING_SHEET_URL).getSheets()[0];
    // Get data, skip header row (A2:D)
    const mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 4).getValues(); 
    
    // Create lookup dictionary { CampaignName: { category: 'X', intentType: 'Y' } }
    const mapping = {};
    for (let i = 0; i < mappingData.length; i++) {
      const campaignName = mappingData[i][0]; // Column A
      if (campaignName) {
        mapping[campaignName] = {
          category: mappingData[i][2] || 'Unmapped',   // Column C
          intentType: mappingData[i][3] || 'Unknown' // Column D
        };
      }
    }
    
    Logger.log('✅ Loaded ' + Object.keys(mapping).length + ' campaign mappings');
    
    let unmappedCount = 0;
    // Apply mapping to each row
    data.forEach(row => {
      const mapped = mapping[row.campaignName];
      if (mapped) {
        row.category = mapped.category;
        row.intentType = mapped.intentType;
      } else {
        row.category = 'Unmapped';
        row.intentType = 'Unknown';
        unmappedCount++;
      }
    });
    
    if (unmappedCount > 0) {
      Logger.log(`⚠️ Found ${unmappedCount} rows with unmapped campaigns. Please update the mapping sheet.`);
    }
    Logger.log('✅ Applied mapping');
    return data;
    
  } catch (e) {
    Logger.log('❌ MAPPING SHEET ERROR: ' + e.toString());
    Logger.log('⚠️ Could not open or read mapping sheet: ' + CONFIG.MAPPING_SHEET_URL);
    Logger.log('⚠️ Continuing without mapping - all campaigns will be marked as "Unmapped"');
    notifyError(new Error('CRITICAL: Could not read mapping sheet. ' + e.toString() + '. Report generated, but all data is "Unmapped".'));
    
    // Continue without mapping
    data.forEach(row => {
      row.category = 'Unmapped';
      row.intentType = 'Unknown';
    });
    
    return data;
  }
}

// ============================================
// STEP 4: CREATE CATEGORY PIVOT
// ============================================
function createCategoryPivot(data) {
  const pivot = {};
  
  data.forEach(row => {
    const cat = row.category || 'Unmapped';
    
    if (!pivot[cat]) {
      pivot[cat] = {
        impressions: 0,
        clicks: 0,
        spend: 0,
        orders: 0,
        sales: 0
      };
    }
    
    pivot[cat].impressions += row.impressions;
    pivot[cat].clicks += row.clicks;
    pivot[cat].spend += row.spend;
    pivot[cat].orders += row.orders;
    pivot[cat].sales += row.sales;
  });
  
  // Convert to array and calculate metrics
  const result = Object.keys(pivot).map(category => {
    const p = pivot[category];
    return {
      category: category,
      impressions: p.impressions,
      clicks: p.clicks,
      spend: p.spend,
      orders: p.orders,
      sales: p.sales,
      ctr: p.impressions > 0 ? (p.clicks / p.impressions * 100).toFixed(2) + '%' : '0.00%',
      cpc: p.clicks > 0 ? (p.spend / p.clicks).toFixed(1) : '0.0',
      convRate: p.clicks > 0 ? (p.orders / p.clicks * 100).toFixed(2) + '%' : '0.00%',
      roi: p.spend > 0 ? (p.sales / p.spend).toFixed(1) : '0.0',
      aov: p.orders > 0 ? Math.round(p.sales / p.orders) : 0,
      cpo: p.orders > 0 ? Math.round(p.spend / p.orders) : 0
    };
  });
  
  // Sort by sales descending
  result.sort((a, b) => b.sales - a.sales);
  
  // Add total row
  const total = result.reduce((acc, row) => {
    acc.impressions += row.impressions;
    acc.clicks += row.clicks;
    acc.spend += row.spend;
    acc.orders += row.orders;
    acc.sales += row.sales;
    return acc;
  }, {impressions: 0, clicks: 0, spend: 0, orders: 0, sales: 0});
  
  total.category = 'Total';
  total.ctr = total.impressions > 0 ? (total.clicks / total.impressions * 100).toFixed(2) + '%' : '0.00%';
  total.cpc = total.clicks > 0 ? (total.spend / total.clicks).toFixed(1) : '0.0';
  total.convRate = total.clicks > 0 ? (total.orders / total.clicks * 100).toFixed(2) + '%' : '0.00%';
  total.roi = total.spend > 0 ? (total.sales / total.spend).toFixed(1) : '0.0';
  total.aov = total.orders > 0 ? Math.round(total.sales / total.orders) : 0;
  total.cpo = total.orders > 0 ? Math.round(total.spend / total.orders) : 0;
  
  result.push(total);
  
  Logger.log('✅ Created pivot with ' + (result.length - 1) + ' categories + Total');
  return result;
}

// ============================================
// STEP 5: SAVE COLLATED FILE
// ============================================
function saveCollatedFile(data) {
  const fileName = 'Amazon_Ads_Collated_' + getDateString();
  const ss = SpreadsheetApp.create(fileName);
  const sheet = ss.getSheets()[0];
  
  // Headers
  const headers = ['Date', 'Campaign Name', 'Portfolio', 'Category', 'Intent Type', 
                   'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Data
  const rows = data.map(row => [
    row.date,
    row.campaignName,
    row.portfolioName,
    row.category,
    row.intentType,
    row.impressions,
    row.clicks,
    row.spend,
    row.sales,
    row.orders
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Format
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
  sheet.autoResizeColumns(1, headers.length);
  
  // Move to Drive folder if specified
  if (CONFIG.DRIVE_FOLDER_ID) {
    try {
      const file = DriveApp.getFileById(ss.getId());
      const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file); // Remove from root
      Logger.log('✅ Moved collated file to Drive folder: ' + folder.getName());
    } catch (e) {
      Logger.log('⚠️ Could not move file to Drive folder. File remains in root. Error: ' + e.toString());
      notifyError(new Error('Could not move file to Drive Folder ' + CONFIG.DRIVE_FOLDER_ID + '. ' + e.toString()));
    }
  }
  
  Logger.log('✅ Saved collated file: ' + ss.getUrl());
  return ss.getUrl();
}

// ============================================
// STEP 6: CREATE DRAFT EMAIL
// ============================================
function createDraftEmail(pivotData) {
  const subject = CONFIG.EMAIL_SUBJECT.replace('[DATE]', getDisplayDateString());
  const htmlTable = createHtmlTable(pivotData);
  
  const totalRow = pivotData[pivotData.length - 1];
  
  const body = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
      <p>Hi Team,</p>
      <p>Please find today's Amazon Advertising performance report for ${getDisplayDateString()}:</p>
      
      <h3 style="color: #4472C4;">Category-wise Performance</h3>
      ${htmlTable}
      
      <h3 style="color: #4472C4; margin-top: 25px;">Key Highlights:</h3>
      <ul style="line-height: 1.8;">
        <li><strong>Total Spend:</strong> ₹${formatNumber(totalRow.spend)}</li>
        <li><strong>Total Sales:</strong> ₹${formatNumber(totalRow.sales)}</li>
        <li><strong>Total Orders:</strong> ${formatNumber(totalRow.orders)}</li>
t      <li><strong>Overall ROAS:</strong> ${totalRow.roi}x</li>
        <li><strong>Average CPC:</strong> ₹${totalRow.cpc}</li>
        <li><strong>Conversion Rate:</strong> ${totalRow.convRate}</li>
      </ul>
      
      <p style="margin-top: 20px;">Best regards,<br>Amazon Ads Team</p>
    </div>
  `;
  
  GmailApp.createDraft(
    CONFIG.CLIENT_EMAIL,
    subject,
    '', // Plain text body (Gmail will use HTML body)
    {
      htmlBody: body,
      name: 'Amazon Ads Report' // Name in Gmail drafts list
    }
  );
  
  Logger.log('✅ Created draft email for client: ' + CONFIG.CLIENT_EMAIL);
}

// ============================================
// CREATE HTML TABLE (HELPER)
// ============================================
function createHtmlTable(data) {
  let html = `
    <table style="border-collapse: collapse; width: 100%; font-size: 12px; margin-top: 15px; border: 1px solid #ddd;">
      <thead>
        <tr style="background-color: #4472C4; color: white;">
          <th style="border: 1px solid #ddd; padding: 10px; text-align: left;">Category</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Impressions</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Clicks</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Spends (₹)</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Orders</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Sales (₹)</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">CTR</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">CPC (₹)</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Conv Rate</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">ROAS</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">AOV (₹)</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">CPO (₹)</th>
        </tr>
      </thead>
      <tbody>
  `;
  
  data.forEach((row, index) => {
    const isTotal = row.category === 'Total';
    const bgColor = isTotal ? '#D9E1F2' : (index % 2 === 0 ? '#F2F2F2' : 'white');
    const fontWeight = isTotal ? 'bold' : 'normal';
    
    html += `
      <tr style="background-color: ${bgColor}; font-weight: ${fontWeight};">
        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${row.category}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.impressions)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.clicks)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.spend)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.orders)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.sales)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${row.ctr}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${row.cpc}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${row.convRate}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${row.roi}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.aov)}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${formatNumber(row.cpo)}</td>
      </tr>
    `;
  });
  
  html += '</tbody></table>';
  return html;
}

// ============================================
// STEP 7: NOTIFICATIONS (to you)
// ============================================
function notifyReportReady() {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '✅ Amazon Ads Report Ready - Check Gmail Drafts',
    body: '🎉 Your Amazon Ads daily report has been processed successfully!\n\n' +
          '📧 A draft email is ready in your Gmail Drafts folder.\n' +
          '👀 Please review the report and click Send to deliver to the client.\n\n' +
          '📊 Client Email: ' + CONFIG.CLIENT_EMAIL + '\n' +
          '⏰ Timestamp: ' + new Date().toString() + '\n\n' +
          'Action Required: Open Gmail > Drafts > Review > Send'
  });
}

function notifyNoReportsFound() {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '⚠️ No Amazon Ads Reports Found Today',
    body: 'The automation ran but could not find any emails matching the search criteria.\n\n' +
          'Search Query:\n' +
          `from:${CONFIG.SENDER_EMAIL} to:${CONFIG.RECEIVER_EMAIL} newer_than:1d has:attachment` + '\n\n' +
          'Please check if:\n' +
          '1. Reports were sent by ' + CONFIG.SENDER_EMAIL + '\n' +
          '2. Reports were received at ' + CONFIG.RECEIVER_EMAIL + '\n' +
          '3. The email has attachments (.csv, .xls, .xlsx)\n\n' +
          'Timestamp: ' + new Date().toString()
  });
}

function notifyError(error) {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '❌ CRITICAL: Amazon Ads Automation Error',
    body: 'An error occurred during the Amazon Ads report automation:\n\n' +
          'Error: ' + error.toString() + '\n\n' +
          'Stack trace:\n' + error.stack + '\n\n' +
          'Timestamp: ' + new Date().toString()
  });
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Parses a string (which may contain '₹' or ',') into a number.
 * Returns 0 if input is invalid.
 */
function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  
  // Remove currency symbols, commas, and trim whitespace
  const cleanValue = value.toString().replace(/[₹,]/g, '').trim();
  
  const num = parseFloat(cleanValue);
  return isNaN(num) ? 0 : num;
}

/**
 * Formats a number as a string with Indian comma style (e.g., 1,23,456)
 * and rounds it to the nearest integer.
 */
function formatNumber(num) {
  if (typeof num !== 'number' || isNaN(num)) return '0';
  // 'en-IN' locale handles the Indian numbering system (lakhs, crores)
  return new Intl.NumberFormat('en-IN').format(Math.round(num));
}

/**
 * Returns today's date as 'YYYY-MM-DD'
 */
function getDateString() {
  const date = new Date();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Returns today's date as 'dd-MMM-yyyy' (e.g., 25-Oct-2023)
 */
function getDisplayDateString() {
  const date = new Date();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
}


// ============================================
// SETUP TRIGGER (RUN ONCE MANUALLY)
// ============================================
function setupDailyTrigger() {
  // Delete any existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automateAmazonAdsReport') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing trigger.');
    }
  });
  
  // Create a new daily trigger to run at 10 AM in your script's timezone
  ScriptApp.newTrigger('automateAmazonAdsReport')
    .timeBased()
    .everyDays(1)
    .atHour(10) // Runs between 10:00 and 11:00 AM
    .create();
  
  Logger.log('✅ Daily trigger set up successfully for 10 AM');
  SpreadsheetApp.getUi().alert('✅ Daily trigger set up successfully for 10 AM');
}
