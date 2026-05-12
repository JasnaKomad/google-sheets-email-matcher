/**
 * @OnlyCurrentDoc
 * 
 * Google Sheets Email Matcher
 * Automatically populates emails in Teams tab by matching against Database tab
 * 
 * Features:
 * - Real-time auto-fill on edit
 * - 3 execution modes (all, empty only, recheck)
 * - Conditional formatting ready
 * - Clickable email hyperlinks
 * - Audit notes with timestamps
 */

// ========== MENU ==========
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Email Generator')
    .addItem('Fill All Emails', 'fillAllEmails')
    .addItem('Fill Only Empty Emails', 'fillOnlyEmptyEmails')
    .addItem('Recheck All Emails', 'recheckAllEmails')  
    .addSeparator()
    .addItem('Delete Trigger', 'deleteTrigger')
    .addItem('Setup Trigger', 'setupTrigger')
    .addToUi();
  
  createEditTrigger();
}

// ========== TRIGGER MANAGEMENT ==========
function createEditTrigger() {
  deleteTrigger();
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditCheckAndFill')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  console.log('Trigger Created');
}

function deleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditCheckAndFill') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log('Trigger Deleted');
}

// ========== AUTO-FILL ON EDIT ==========
function onEditCheckAndFill(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  
  if (sheet.getName() !== 'Teams') return;
  if (col < 1 || col > 3) return;
  if (row < 2) return;
  
  Utilities.sleep(200);
  
  const firstName = sheet.getRange(row, 1).getValue()?.toString().trim().toLowerCase() || '';
  const lastName = sheet.getRange(row, 2).getValue()?.toString().trim().toLowerCase() || '';
  const company = sheet.getRange(row, 3).getValue()?.toString().trim().toLowerCase() || '';
  
  if (!firstName || !lastName || !company) return;
  
  fillSingleEmail(row, firstName, lastName, company);
}

function fillSingleEmail(row, firstName, lastName, company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const dbSheet = ss.getSheetByName('Database');
  
  const dbData = dbSheet.getDataRange().getValues();
  const emailMap = new Map();
  
  for (let i = 1; i < dbData.length; i++) {
    const dbFirstName = dbData[i][0]?.toString().trim().toLowerCase() || '';
    const dbLastName = dbData[i][1]?.toString().trim().toLowerCase() || '';
    const dbCompany = dbData[i][2]?.toString().trim().toLowerCase() || '';
    const dbEmail = dbData[i][3];
    
    if (dbFirstName && dbLastName && dbCompany && dbEmail) {
      const key = `${dbFirstName}|${dbLastName}|${dbCompany}`;
      emailMap.set(key, dbEmail);
    }
  }
  
  const key = `${firstName}|${lastName}|${company}`;
  const email = emailMap.get(key);
  
  if (email) {
    teamsSheet.getRange(row, 4).setValue(`=HYPERLINK("mailto:${email}", "${email}")`);
    teamsSheet.getRange(row, 4).setNote(`✅ FOUND: ${new Date().toLocaleString()}`);
  } else {
    teamsSheet.getRange(row, 4).setValue('NOT_FOUND_IN_DATABASE');
    teamsSheet.getRange(row, 4).setNote(`❌ NOT FOUND: ${new Date().toLocaleString()}`);
  }
}

// ========== FILL ALL EMAILS ==========
function fillAllEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const dbSheet = ss.getSheetByName('Database');
  
  const teamsData = teamsSheet.getDataRange().getValues();
  const dbData = dbSheet.getDataRange().getValues();
  
  const emailMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const firstName = dbData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = dbData[i][1]?.toString().trim().toLowerCase() || '';
    const company = dbData[i][2]?.toString().trim().toLowerCase() || '';
    const email = dbData[i][3];
    
    if (firstName && lastName && company && email) {
      const key = `${firstName}|${lastName}|${company}`;
      emailMap.set(key, email);
    }
  }
  
  const results = [];
  let foundCount = 0;
  let notFoundCount = 0;
  
  for (let i = 1; i < teamsData.length; i++) {
    const firstName = teamsData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = teamsData[i][1]?.toString().trim().toLowerCase() || '';
    const company = teamsData[i][2]?.toString().trim().toLowerCase() || '';
    
    if (firstName && lastName && company) {
      const key = `${firstName}|${lastName}|${company}`;
      const email = emailMap.get(key);
      
      if (email) {
        results.push([`=HYPERLINK("mailto:${email}", "${email}")`]);
        foundCount++;
      } else {
        results.push(['NOT_FOUND_IN_DATABASE']);
        notFoundCount++;
      }
    } else {
      results.push(['']);
      notFoundCount++;
    }
  }
  
  if (results.length > 0) {
    teamsSheet.getRange(2, 4, results.length, 1).setValues(results);
  }
  
  SpreadsheetApp.getUi().alert(`✅ Done!\n\nFound: ${foundCount}\nNot Found: ${notFoundCount}`);
}

// ========== FILL ONLY EMPTY EMAILS ==========
function fillOnlyEmptyEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const dbSheet = ss.getSheetByName('Database');
  
  const teamsData = teamsSheet.getDataRange().getValues();
  const dbData = dbSheet.getDataRange().getValues();
  
  const emailMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const firstName = dbData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = dbData[i][1]?.toString().trim().toLowerCase() || '';
    const company = dbData[i][2]?.toString().trim().toLowerCase() || '';
    const email = dbData[i][3];
    
    if (firstName && lastName && company && email) {
      const key = `${firstName}|${lastName}|${company}`;
      emailMap.set(key, email);
    }
  }
  
  let foundCount = 0;
  let notFoundCount = 0;
  let skippedCount = 0;
  
  for (let i = 1; i < teamsData.length; i++) {
    const currentEmail = teamsData[i][3];
    const firstName = teamsData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = teamsData[i][1]?.toString().trim().toLowerCase() || '';
    const company = teamsData[i][2]?.toString().trim().toLowerCase() || '';
    
    const isEmpty = !currentEmail || currentEmail.toString().trim() === '';
    
    if (isEmpty && firstName && lastName && company) {
      const key = `${firstName}|${lastName}|${company}`;
      const email = emailMap.get(key);
      
      if (email) {
        teamsSheet.getRange(i+1, 4).setValue(`=HYPERLINK("mailto:${email}", "${email}")`);
        foundCount++;
      } else {
        teamsSheet.getRange(i+1, 4).setValue('NOT_FOUND_IN_DATABASE');
        notFoundCount++;
      }
    } else if (!isEmpty) {
      skippedCount++;
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Done!\n\nFound: ${foundCount}\nNot Found: ${notFoundCount}\nSkipped: ${skippedCount}`);
}

// ========== RECHECK ALL ==========
function recheckAllEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const dbSheet = ss.getSheetByName('Database');
  
  const teamsData = teamsSheet.getDataRange().getValues();
  const dbData = dbSheet.getDataRange().getValues();
  
  const emailMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const firstName = dbData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = dbData[i][1]?.toString().trim().toLowerCase() || '';
    const company = dbData[i][2]?.toString().trim().toLowerCase() || '';
    const email = dbData[i][3];
    
    if (firstName && lastName && company && email) {
      const key = `${firstName}|${lastName}|${company}`;
      emailMap.set(key, email);
    }
  }
  
  let foundCount = 0;
  let notFoundCount = 0;
  let updatedCount = 0;
  
  for (let i = 1; i < teamsData.length; i++) {
    const firstName = teamsData[i][0]?.toString().trim().toLowerCase() || '';
    const lastName = teamsData[i][1]?.toString().trim().toLowerCase() || '';
    const company = teamsData[i][2]?.toString().trim().toLowerCase() || '';
    const currentEmail = teamsData[i][3]?.toString() || '';
    
    if (firstName && lastName && company) {
      const key = `${firstName}|${lastName}|${company}`;
      const email = emailMap.get(key);
      
      if (email) {
        const newEmailFormula = `=HYPERLINK("mailto:${email}", "${email}")`;
        if (currentEmail !== newEmailFormula && currentEmail !== email) {
          teamsSheet.getRange(i+1, 4).setValue(newEmailFormula);
          updatedCount++;
        }
        foundCount++;
      } else {
        if (currentEmail !== 'NOT_FOUND_IN_DATABASE') {
          teamsSheet.getRange(i+1, 4).setValue('NOT_FOUND_IN_DATABASE');
          updatedCount++;
        }
        notFoundCount++;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Done!\n\nFound: ${foundCount}\nNot Found: ${notFoundCount}\nUpdated: ${updatedCount}`);
}

function setupTrigger() {
  createEditTrigger();
  SpreadsheetApp.getUi().alert('✅ Trigger is on!');
}
