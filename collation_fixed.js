// ===============================
// AMAZON ADS REPORT AUTOMATION SCRIPT — Website Collator Logic
// ===============================

// CONFIGURATION 
const CONFIG = {
  MAPPING_SHEET_URL: 'https://docs.google.com/spreadsheets/d/1Ui1i0oikCy7MHhJ_obRlwZ6OZZf_oR1ugDaJV6QaNrc/edit?gid=0#gid=0',
  SENDER_EMAIL: 'mithun.gowda@hiveminds.in',
  RECEIVER_EMAIL: 'mithungowdahc12a@gmail.com',
  CLIENT_EMAIL: 'mgowda8431@gmail.com',
  YOUR_EMAIL: 'mithungowdahc12a@gmail.com',
  DRIVE_FOLDER_ID: '1gQKp7tY5aKRx0LG5fmVQhWc5Qdl_2qks',
  EMAIL_SUBJECT: 'Amazon Ads Performance Report - [DATE]'
};

// All possible names (case-insensitive) — ADD more as seen in your exports
const COLUMN_MAP = {
  "Date": ["Date", "Report Date", "created date"],
  "Campaign Name": ["Campaign Name", "Campaign", "Campaign Title", "Ad campaign"],
  "Portfolio name": ["Portfolio", "Portfolio Name"],
  "Impressions": ["Impressions", "Views", "Impr", "View Count"],
  "Clicks": ["Clicks", "Click Count", "no_of_clicks"],
  "Spend": ["Spend", "Spend (INR)", "Cost", "Ad Spend", "Total Spend"],
  "14 Day Total Sales (₹)": ["14 Day Total Sales (₹)", "Sales", "Total Sales", "Revenue"],
  "14 Day Total Orders (#)": ["14 Day Total Orders (#)", "Orders", "Order Count"]
};
const OUTPUT_COLS = Object.keys(COLUMN_MAP);

// ========== SANITIZATION HELPERS ==========

const NBSP = /\u00A0/g, ZWSP = /\u200B/g;
function cleanString(v) {
  return String(v ?? "")
    .replace(NBSP, " ")
    .replace(ZWSP, "")
    .trim();
}
function cleanNumber(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  let s = String(v)
    .replace(/â‚¹|₹|\$|€|£/g, "")
    .replace(/,/g, "")
    .replace(NBSP, "")
    .replace(/\(([^)]+)\)/g, "-$1")
    .replace(/[^0-9.\-]/g, "")
    .trim();
  const n = Number.parseFloat(s);
  return isFinite(n) ? n : 0;
}
function sanitizeRow(row) {
  let out = Object.assign({}, row);
  if ("Campaign Name" in out) out["Campaign Name"] = cleanString(out["Campaign Name"]);
  if ("Portfolio name" in out) out["Portfolio name"] = cleanString(out["Portfolio name"]);
  [
    "Impressions", "Clicks", "Spend", "14 Day Total Sales (₹)", "14 Day Total Orders (#)"
  ].forEach(col => { if (col in out) out[col] = cleanNumber(out[col]); });
  return out;
}
function formatNumber(num) {
  if (typeof num !== 'number' || isNaN(num)) return '0';
  return new Intl.NumberFormat('en-IN').format(Math.round(num));
}

// ========== MAIN FUNCTION (for trigger) ==========

function automateAmazonAdsReport() {
  try {
    Logger.log('🚀 Starting Amazon Ads Automation...');
    const reports = fetchReportsFromGmail();
    if (!reports || (reports.sp == null && reports.sb==null && reports.sd == null)) {
      Logger.log('⚠️ No valid reports found. Notifying user.');
      notifyNoReportsFound();
      return;
    }
    const collated = collateReports(reports, COLUMN_MAP);
    if (collated.length === 0) {
      Logger.log('⚠️ Reports found but collation resulted in 0 rows. Check column mappings or report content.');
      notifyError(new Error('Reports found, but collation failed. Please check report formats.'));
      return;
    }
    const enriched = applyMapping(collated);
    const pivot = createCategoryPivot(enriched);
    saveCollatedFile(enriched);
    createDraftEmail(pivot);
    notifyReportReady();
    Logger.log('✅ Automation completed successfully!');
  } catch (e) {
    Logger.log('❌ Error: ' + (e && e.message || e));
    Logger.log('Stack: ' + (e && e.stack || e));
    notifyError(e);
  }
}

// ========== FETCH REPORTS FROM GMAIL ==========

function fetchReportsFromGmail() {
  const searchQuery = `from:${CONFIG.SENDER_EMAIL} to:${CONFIG.RECEIVER_EMAIL} newer_than:1d has:attachment`;
  Logger.log('🔍 Searching emails: ' + searchQuery);
  const threads = GmailApp.search(searchQuery, 0, 1);
  if (!threads.length) return null;
  const attachments = threads[0].getMessages().pop().getAttachments();
  const reports = { sp: null, sb: null, sd: null };
  attachments.forEach(attachment => {
    const name = attachment.getName();
    try {
      let data;
      if (name.endsWith('.csv')) {
        data = Utilities.parseCsv(attachment.getDataAsString());
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        const tempFile = DriveApp.createFile(attachment.copyBlob());
        const tempFileId = tempFile.getId();
        const fileJson = Drive.Files.copy({name: name + '_converted_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS}, tempFileId);
        data = SpreadsheetApp.openById(fileJson.id).getSheets()[0].getDataRange().getValues();
        Drive.Files.remove(tempFileId);
        Drive.Files.remove(fileJson.id);
      }
      if (name.toLowerCase().includes('sponsored_products')) reports.sp = data;
      else if (name.toLowerCase().includes('sponsored_brands')) reports.sb = data;
      else if (name.toLowerCase().includes('sponsored_display')) reports.sd = data;
    } catch(e) { Logger.log('⚠️ Error processing ' + name + ':' + e); }
  });
  return reports;
}

// ========== COLLATE LOGIC (Your Website Logic) ==========

function collateReports(reportsObj, columnMap) {
  var collated = [];
  var headerNorm = header => cleanString(header).toLowerCase();
  var duplicateKeys = {};
  Object.keys(reportsObj).forEach(reportType => {
    var data = reportsObj[reportType];
    if (!data || data.length < 2) return;
    var headers = data[0].map(h => cleanString(h));
    var headersNormMap = {};
    headers.forEach((h, i) => { headersNormMap[headerNorm(h)] = i; });
    var missingCols = [];
    Object.keys(columnMap).forEach(canonical => {
      const variations = columnMap[canonical].map(headerNorm);
      const found = Object.keys(headersNormMap).some(h => variations.indexOf(h) !== -1);
      if (!found) missingCols.push(canonical);
    });
    if (missingCols.length > 0) Logger.log("⚠️ " + reportType + " missing: " + missingCols.join(", "));

    for (var i = 1; i < data.length; ++i) {
      var row = data[i];
      var hasAny = row.some(x => (!!x && cleanString(x)!==""));
      if (!hasAny) continue;
      var outputRow = {};
      Object.keys(columnMap).forEach(function (canonical) {
        var variations = columnMap[canonical].map(headerNorm);
        var colIdx = -1;
        for (var j = 0; j < headers.length; j++) {
          if (variations.indexOf(headerNorm(headers[j])) !== -1) {colIdx = j; break;}
        }
        outputRow[canonical] = colIdx !== -1 ? row[colIdx] : "";
      });
      if ("Campaign Name" in outputRow && cleanString(outputRow["Campaign Name"]).toLowerCase()==="total") continue;
      outputRow = sanitizeRow(outputRow);
      var dedupKey = ["Date","Campaign Name","Portfolio name"].map(k => cleanString(outputRow[k])).join("|") + "|" + reportType;
      if (!duplicateKeys[dedupKey]) { collated.push(outputRow); duplicateKeys[dedupKey] = true; }
    }
  });
  return collated;
}

// ========== MAPPING VLOOKUP (Category/Intent/Type) ==========

function applyMapping(data) {
  try {
    const mappingSheet = SpreadsheetApp.openByUrl(CONFIG.MAPPING_SHEET_URL).getSheets()[0];
    const mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow()-1, 4).getValues();
    const mapping = {};
    for (let i = 0; i < mappingData.length; i++) {
      const campaignName = cleanString(mappingData[i][0]);
      if (campaignName) {
        mapping[campaignName] = {
          category: cleanString(mappingData[i][2]) || 'Unmapped',
          intentType: cleanString(mappingData[i][3]) || 'Unknown'
        };
      }
    }
    let unmappedCount = 0;
    data.forEach(row => {
      const mapped = mapping[cleanString(row["Campaign Name"])];
      if (mapped) {row.category = mapped.category; row.intentType = mapped.intentType;}
      else {row.category = 'Unmapped';row.intentType = 'Unknown';unmappedCount++;}
    });
    if (unmappedCount>0) Logger.log(`⚠️ Found ${unmappedCount} rows with unmapped campaigns.`);
    return data;
  } catch(e) {
    Logger.log('❌ MAPPING SHEET ERROR: ' + e);
    data.forEach(row => {row.category='Unmapped';row.intentType='Unknown';});
    return data;
  }
}

// ========== PIVOT ========================

function createCategoryPivot(data) {
  const metrics = ["Impressions","Clicks","Spend","14 Day Total Sales (₹)","14 Day Total Orders (#)"];
  const pivot = {};
  data.forEach(row=>{
    const cat = row.category || "Unmapped";
    if (!pivot[cat]) pivot[cat]={};
    metrics.forEach(key => pivot[cat][key] = (pivot[cat][key]||0) + cleanNumber(row[key]));
  });
  const result = Object.keys(pivot).map(category=>{
    const p = pivot[category];
    return {
      category: category,
      impressions: p["Impressions"], clicks: p["Clicks"], spend: p["Spend"],
      orders: p["14 Day Total Orders (#)"], sales: p["14 Day Total Sales (₹)"],
      ctr: p["Impressions"]>0 ? (p["Clicks"]/p["Impressions"]*100).toFixed(2)+"%" : "0.00%",
      cpc: p["Clicks"]>0 ? (p["Spend"]/p["Clicks"]).toFixed(1) : "0.0",
      convRate: p["Clicks"]>0 ? (p["14 Day Total Orders (#)"]/p["Clicks"]*100).toFixed(2)+"%" : "0.00%",
      roi: p["Spend"]>0 ? (p["14 Day Total Sales (₹)"]/p["Spend"]).toFixed(1) : "0.0",
      aov: p["14 Day Total Orders (#)"]>0 ? Math.round(p["14 Day Total Sales (₹)"]/p["14 Day Total Orders (#)"]) : 0,
      cpo: p["14 Day Total Orders (#)"]>0 ? Math.round(p["Spend"]/p["14 Day Total Orders (#)"]) : 0
    };
  });
  result.sort((a,b)=>b.sales-a.sales);
  // Total row
  const total = result.reduce((acc,row)=>
    {acc.impressions+=row.impressions;acc.clicks+=row.clicks;acc.spend+=row.spend;acc.orders+=row.orders;acc.sales+=row.sales;return acc;},
    {impressions:0,clicks:0,spend:0,orders:0,sales:0});
  total.category="Total";
  total.ctr = total.impressions>0 ? (total.clicks/total.impressions*100).toFixed(2)+"%":"0.00%";
  total.cpc = total.clicks>0 ? (total.spend/total.clicks).toFixed(1):"0.0";
  total.convRate = total.clicks>0 ? (total.orders/total.clicks*100).toFixed(2)+"%":"0.00%";
  total.roi = total.spend>0 ? (total.sales/total.spend).toFixed(1) : "0.0";
  total.aov = total.orders > 0 ? Math.round(total.sales/total.orders):0;
  total.cpo = total.orders > 0 ? Math.round(total.spend/total.orders):0;
  result.push(total);
  Logger.log('✅ Created pivot table.');
  return result;
}

// ========== OUTPUT/EMAIL ==========

function saveCollatedFile(data) {
  const fileName = 'Amazon_Ads_Collated_' + getDateString();
  const ss = SpreadsheetApp.create(fileName);
  const sheet = ss.getSheets()[0];
  const headers = OUTPUT_COLS.concat(["Category","Intent Type"]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const rows = data.map(row => OUTPUT_COLS.map(col => row[col])
    .concat([row.category,row.intentType]));
  if (rows.length > 0)
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
  sheet.autoResizeColumns(1, headers.length);
  if (CONFIG.DRIVE_FOLDER_ID) {
    try {
      const file = DriveApp.getFileById(ss.getId());
      const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
      Logger.log('✅ Moved collated file to Drive folder: ' + folder.getName());
    } catch(e) {
      Logger.log('⚠️ Could not move file to Drive folder.' + e);
      notifyError(new Error('Could not move file to Drive Folder: ' + e));
    }
  }
  Logger.log('✅ Saved collated file: ' + ss.getUrl());
  return ss.getUrl();
}

function createDraftEmail(pivotData) {
  const subject = CONFIG.EMAIL_SUBJECT.replace('[DATE]', getDisplayDateString());
  const htmlTable = createHtmlTable(pivotData);
  const totalRow = pivotData[pivotData.length-1];
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
        <li><strong>Overall ROAS:</strong> ${totalRow.roi}x</li>
        <li><strong>Average CPC:</strong> ₹${totalRow.cpc}</li>
        <li><strong>Conversion Rate:</strong> ${totalRow.convRate}</li>
      </ul>
      <p style="margin-top: 20px;">Best regards,<br>Amazon Ads Team</p>
    </div>`;
  GmailApp.createDraft(CONFIG.CLIENT_EMAIL, subject, '', {htmlBody: body, name: 'Amazon Ads Report'});
  Logger.log('✅ Created draft email for client: ' + CONFIG.CLIENT_EMAIL);
}

function createHtmlTable(data) {
  if (!Array.isArray(data)) {
    Logger.log('❌ HTML Table: pivotData not an array');
    return '<div style="color:red;">No data to display</div>';
  }
  let html = `<table style="border-collapse: collapse; width: 100%; font-size: 12px; margin-top: 15px; border: 1px solid #ddd;"><thead><tr style="background-color: #4472C4; color: white;">`;
  [
    "Category","Impressions","Clicks","Spend (₹)","Orders","Sales (₹)","CTR","CPC (₹)","Conv Rate","ROAS","AOV (₹)","CPO (₹)"
  ].forEach(col => { html += `<th style="border: 1px solid #ddd; padding: 10px; text-align: right;">${col}</th>`; });
  html += '</tr></thead><tbody>';
  data.forEach((row, idx) => {
    const isTotal = row.category==="Total";
    const bgColor = isTotal ? '#D9E1F2': (idx % 2 === 0 ? '#F2F2F2' : 'white');
    const fontWeight = isTotal ? 'bold':'normal';
    html += `<tr style="background-color: ${bgColor}; font-weight: ${fontWeight};">`
    [
      row.category,formatNumber(row.impressions),formatNumber(row.clicks),formatNumber(row.spend),
      formatNumber(row.orders),formatNumber(row.sales),row.ctr,row.cpc,row.convRate,row.roi,
      formatNumber(row.aov),formatNumber(row.cpo)
    ].forEach(cell => { html += `<td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${cell}</td>`; });
    html+='</tr>';
  });
  html+='</tbody></table>';
  return html;
}

// ========== NOTIFICATIONS/ERRORS ==========

function notifyReportReady() {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '✅ Amazon Ads Report Ready - Check Gmail Drafts',
    body: '🎉 Report processed successfully! Draft email is ready in Gmail Drafts.\nReview and send to client.'
  });
}
function notifyNoReportsFound() {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '⚠️ No Amazon Ads Reports Found Today',
    body: 'Automation ran but could not find any matching emails.\nPlease verify report delivery and try again.'
  });
}
function notifyError(error) {
  MailApp.sendEmail({
    to: CONFIG.YOUR_EMAIL,
    subject: '❌ CRITICAL: Amazon Ads Automation Error',
    body: 'An error occurred in the Amazon Ads automation:\n\n' +
      'Error: ' + (error && error.message ? error.message : error) + '\n\n' +
      'Stack trace:\n' + (error && error.stack ? error.stack : "")
  });
}

// ========== UTILITIES ==========

function getDateString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function getDisplayDateString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
}
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automateAmazonAdsReport') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('automateAmazonAdsReport').timeBased().everyDays(1).atHour(10).create();
  Logger.log('✅ Daily trigger set up for 10 AM.');
  SpreadsheetApp.getUi().alert('✅ Daily trigger set up successfully for 10 AM');
}
