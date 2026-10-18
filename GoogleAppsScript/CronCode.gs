/**
 * RuPay Select Vault — Google Apps Script Engine
 * 100% Free Cloud Backend (Gmail -> Sheet -> Swift App REST API)
 * All timestamps strictly formatted in IST (Indian Standard Time: Asia/Kolkata)
 */

const SENDER_EMAIL = "rupay@truztee.com";
const IST_TIMEZONE = "Asia/Kolkata";
const VAULT_AUTH_KEY = "rupay_vault_auth_sec_8f93e2a1c0d4";

// Standard Unified Column Schemas
const UNIFIED_BRAND_HEADERS = [
  "#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI"
];

const UNIFIED_REDEEMED_HEADERS = [
  "#", "Brand", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Redeemed Timestamp (IST)", "Original Tab", "Appended Timestamp (IST)", "Display in UI"
];

const UNIFIED_CARD_HEADERS = [
  "Card ID", "Bank", "Card Type", "BIN", "Last 4", "Nickname", "Status", "Claimed Quota", "Total Quota", "Blocked Date", "Notes", "Last Updated (IST)"
];

const UNIFIED_SHARED_HEADERS = [
  "#", "Code", "Brand", "Amount", "PIN", "Shared Timestamp (IST)", "Status", "Platform", "Display in UI"
];

const UNIFIED_COOLDOWN_HEADERS = [
  "Account ID", "Last Redemption Time (IST)", "Status", "Last Updated (IST)"
];

/**
 * 1. AUTOMATIC EMAIL INGESTION (Triggered via cron or instant on-demand sync)
 * Two-Stage RUPS Lifecycle:
 *  - Stage 1 (Ack Email): Detects RUPS booking ID, delay/pending message -> records in Brand Tab as "PENDING" & starts 24h cooldown
 *  - Stage 2 (Delivery Email): Detects RUPS booking ID & alphanumeric code -> finds row by RUPS ID and populates Code/PIN as "ACTIVE" in SAME row!
 */
function processRupayEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    console.warn("processRupayEmails: script is currently running in another thread, skipping.");
    return;
  }

  try {
    // Search unread emails from Truztee or containing RUPS booking references
    const query = 'is:unread (from:rupay@truztee.com OR "rupay@truztee.com" OR "RUPS" OR subject:"Fwd:" OR subject:"Fw:")';
    const threads = GmailApp.search(query, 0, 25);
    if (!threads.length) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const msg of messages) {
        if (!msg.isUnread()) continue;

        const subject = msg.getSubject() || "";
        const body = msg.getPlainBody() || "";
        const toHeader = msg.getTo() || "";
        const fromHeader = msg.getFrom() || "";

        // 1. Extract Booking ID (prioritizing RUPS prefix)
        const bookingIdMatch = subject.match(/\b(RUPS[A-Za-z0-9\-]+)\b/i) || 
                               body.match(/\b(RUPS[A-Za-z0-9\-]+)\b/i) ||
                               subject.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i) || 
                               body.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i);
        const bookingId = bookingIdMatch ? bookingIdMatch[1].trim() : "";
        if (!bookingId) {
          // No booking ID found, skip non-redemption emails
          continue;
        }

        // 2. Extract original recipient account
        let account = extractAccountName(toHeader, body, fromHeader);
        if (!account) account = "vibhavoswal9";

        // 3. Extract Brand, Amount, Date
        const brand = extractBrand(subject, body);
        const amount = extractAmount(body);
        const dateOfIssue = extractField(body, "Date of Issue") || formatISTDate(new Date());

        // 4. Extract Voucher Code candidate
        const rawCode = extractField(body, "Coupon Code") || 
                        extractField(body, "Voucher Code") || 
                        extractField(body, "Code") || "";
        const cleanCode = rawCode.replace(/['"]/g, "").trim();

        // 5. Binary Real-Code Check:
        // A genuine code is a single clean alphanumeric/hexanumeric string (6-25 chars) with NO spaces,
        // and is not the booking ID itself. Any delay message ("will be sent...", "delay...", "NA", "-") is rejected as a code.
        const isDeliveredCode = /^[A-Za-z0-9]{6,25}$/.test(cleanCode) && 
                                !cleanCode.toUpperCase().startsWith("RUPS");

        // 6. Extract PIN if present
        const pin = extractField(body, "Coupon Pin") || 
                    extractField(body, "PIN") || 
                    extractField(body, "Voucher Pin") || 
                    extractField(body, "Card Pin") || "";
        const cleanPin = pin ? pin.replace(/['"]/g, "").trim() : "";

        // 7. Route based on Code Presence:
        if (!isDeliveredCode) {
          // ==========================================
          // STAGE 1: Acknowledgement Email
          // ==========================================
          // A. Start 24h cooldown timer on this account immediately
          updateCooldownTab(ss, account);

          // B. Record directly into Brand Tab as "PENDING"
          recordPendingRedemption(ss, {
            bookingId: bookingId,
            brand: brand,
            amount: amount,
            account: account,
            dateOfIssue: dateOfIssue
          });
        } else {
          // ==========================================
          // STAGE 2: Voucher Code Delivery Email
          // ==========================================
          // Search Google Sheet for Booking ID (RUPS...) and fulfill in the SAME row!
          fulfillVoucherCode(ss, {
            bookingId: bookingId,
            brand: brand,
            code: cleanCode,
            pin: cleanPin,
            amount: amount,
            dateOfIssue: dateOfIssue,
            account: account
          });
        }

        msg.markRead();
      }
    }

    SpreadsheetApp.flush();
  } catch(err) {
    console.error("Error in processRupayEmails:", err);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * Extracts account username (e.g. "vibhavoswal9") handling forwarded headers with display names
 */
function extractAccountName(toHeader, body, fromHeader) {
  // 1. Nested "To: <email>" inside forwarded block (handles names before email)
  const fwdToMatch = (body || "").match(/To:\s*(?:[^<@\r\n]+<)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>?/i);
  if (fwdToMatch && !fwdToMatch[1].toLowerCase().includes("rupayvouchers@gmail.com")) {
    return fwdToMatch[1].toLowerCase().replace(/@.*$/, "").trim();
  }

  // 2. From header if forwarded from personal account
  if (fromHeader && !fromHeader.toLowerCase().includes("rupay@truztee.com") && !fromHeader.toLowerCase().includes("google.com")) {
    const fromMatch = fromHeader.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (fromMatch && !fromMatch[1].toLowerCase().includes("rupayvouchers@gmail.com")) {
      return fromMatch[1].toLowerCase().replace(/@.*$/, "").trim();
    }
  }

  // 3. Direct To header
  if (toHeader) {
    const toMatch = toHeader.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (toMatch && !toMatch[1].toLowerCase().includes("rupayvouchers@gmail.com")) {
      return toMatch[1].toLowerCase().replace(/@.*$/, "").trim();
    }
  }

  return "vibhavoswal9";
}

function extractField(body, fieldName) {
  const re = new RegExp("\\b" + fieldName + "\\s*[:#-]?\\s*([^\n\r<]+)", "i");
  const m = body.match(re);
  return m ? m[1].trim() : null;
}

function extractBrand(subject, body) {
  const fullText = subject + "\n" + body;
  const provider = extractField(body, "Service Provider") || "";
  const combined = (provider + " " + fullText).toLowerCase();

  if (combined.includes("four fountains")) return "FourFountains";
  if (combined.includes("lakme") || combined.includes("lakmé")) return "Lakme";
  if (combined.includes("uber")) return "Uber";
  if (combined.includes("kalyan")) return "Kalyan";
  if (combined.includes("reliance")) return "Reliance";
  if (combined.includes("myntra")) return "Myntra";
  if (combined.includes("apollo")) return "Apollo";
  if (combined.includes("big basket") || combined.includes("bigbasket")) return "BigBasket";
  if (combined.includes("thyrocare") || combined.includes("health check")) return "Thyrocare";
  if (combined.includes("cult")) return "CultFit";
  if (combined.includes("amazon")) return "AmazonPrime";
  if (combined.includes("taxspanner")) return "TaxSpanner";
  if (combined.includes("decathlon")) return "Decathlon";
  if (combined.includes("tattva")) return "TattvaSpa";
  if (combined.includes("times prime") || combined.includes("timesprime") || combined.includes("tpace")) return "TimesPrime";

  return "Other";
}

function extractAmount(body) {
  const m = body.match(/INR\s*([0-9,]+)/i) || body.match(/Rs\.?\s*([0-9,]+)/i);
  return m ? parseInt(m[1].replace(/,/g, "")) : 500;
}

/**
 * 2. STAGE 1: RECORD PENDING REDEMPTION DIRECTLY IN BRAND TAB
 */
function recordPendingRedemption(ss, data) {
  const brand = data.brand;
  let sheet = ss.getSheetByName(brand);
  if (!sheet) {
    sheet = ss.insertSheet(brand);
    sheet.appendRow(UNIFIED_BRAND_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }

  // Idempotency: Does this Booking ID already exist in this brand sheet?
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === data.bookingId) {
      // Already recorded, don't duplicate
      return;
    }
  }

  const nextNum = rows.length;
  const rowData = [
    nextNum,                               // Col 1: #
    data.bookingId,                        // Col 2: Booking ID (RUPS...)
    "PENDING",                             // Col 3: Code
    "",                                    // Col 4: PIN
    data.amount || 500,                    // Col 5: Value
    data.dateOfIssue,                      // Col 6: Date of Issue
    data.account,                          // Col 7: Account
    "Unassigned",                          // Col 8: Card
    "PENDING",                             // Col 9: Status
    formatISTTimestamp(new Date()),        // Col 10: Appended Timestamp (IST)
    true                                   // Col 11: Display in UI
  ];
  sheet.appendRow(rowData);
}

/**
 * 3. STAGE 2: BRAND TAB FULFILLMENT BY BOOKING ID (IN SAME ROW)
 */
function fulfillVoucherCode(ss, data) {
  const brand = data.brand;
  let sheet = ss.getSheetByName(brand);
  if (!sheet) {
    sheet = ss.insertSheet(brand);
    sheet.appendRow(UNIFIED_BRAND_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }

  // 1. Search the target brand sheet first by Booking ID (RUPS...)
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === data.bookingId) {
      // MATCH FOUND: Populate Code & PIN and flip Status to ACTIVE in the EXACT SAME ROW!
      sheet.getRange(i + 1, 3).setValue(`'${data.code}`);
      if (data.pin) {
        sheet.getRange(i + 1, 4).setValue(`'${data.pin}`);
      }
      sheet.getRange(i + 1, 9).setValue("ACTIVE");
      sheet.getRange(i + 1, 10).setValue(formatISTTimestamp(new Date()));
      sheet.getRange(i + 1, 11).setValue(true);
      return;
    }
  }

  // 2. Global fallback search: What if brand detection differed slightly (e.g. Other vs Uber)?
  const allSheets = ss.getSheets();
  const systemTabs = new Set(["Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns", "⚡ Unassigned"]);
  for (const s of allSheets) {
    if (systemTabs.has(s.getName()) || s.getName() === brand) continue;
    const sRows = s.getDataRange().getValues();
    for (let i = 1; i < sRows.length; i++) {
      if (String(sRows[i][1]).trim() === data.bookingId) {
        s.getRange(i + 1, 3).setValue(`'${data.code}`);
        if (data.pin) {
          s.getRange(i + 1, 4).setValue(`'${data.pin}`);
        }
        s.getRange(i + 1, 9).setValue("ACTIVE");
        s.getRange(i + 1, 10).setValue(formatISTTimestamp(new Date()));
        s.getRange(i + 1, 11).setValue(true);
        return;
      }
    }
  }

  // 3. Fallback: If not found in any sheet (e.g. ack email was missed or instant code delivery)
  // Append as a new ACTIVE row in the Brand tab
  const nextNum = rows.length;
  const rowData = [
    nextNum,                               // Col 1: #
    data.bookingId,                        // Col 2: Booking ID
    `'${data.code}`,                       // Col 3: Code
    data.pin ? `'${data.pin}` : "",        // Col 4: PIN
    data.amount || 500,                    // Col 5: Value
    data.dateOfIssue,                      // Col 6: Date of Issue
    data.account,                          // Col 7: Account
    "Unassigned",                          // Col 8: Card
    "ACTIVE",                              // Col 9: Status
    formatISTTimestamp(new Date()),        // Col 10: Appended Timestamp (IST)
    true                                   // Col 11: Display in UI
  ];
  sheet.appendRow(rowData);
}

/**
 * 4. ASSIGN CARD IN BULK (Called by App)
 */
function assignCardToRedemptions(ss, assignments) {
  const unassignedSheet = ss.getSheetByName("⚡ Unassigned");
  if (!unassignedSheet) return { success: true };

  const uRows = unassignedSheet.getDataRange().getValues();
  const toDelete = [];

  for (const assign of assignments) {
    for (let i = 1; i < uRows.length; i++) {
      if (uRows[i][0] === assign.bookingId) {
        const brand = uRows[i][1];
        const amount = uRows[i][2];
        const account = uRows[i][3];
        const dateOfIssue = uRows[i][4];
        const cardName = assign.card;

        let bSheet = ss.getSheetByName(brand);
        if (!bSheet) {
          bSheet = ss.insertSheet(brand);
          bSheet.appendRow(UNIFIED_BRAND_HEADERS);
          bSheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
          bSheet.setFrozenRows(1);
        }

        const bRows = bSheet.getDataRange().getValues();
        const nextNum = bRows.length;
        const rowData = [
          nextNum,
          assign.bookingId,
          "PENDING_DELIVERY",
          "",
          amount,
          dateOfIssue,
          account,
          cardName,
          "PENDING",
          formatISTTimestamp(new Date())
        ];
        bSheet.appendRow(rowData);

        toDelete.push(i + 1);
        break;
      }
    }
  }

  toDelete.sort((a,b) => b - a).forEach(rowIdx => unassignedSheet.deleteRow(rowIdx));
  return { success: true, count: assignments.length };
}

/**
 * 5. PERSISTENT SYSTEM SHEETS INITIALIZATION & COOLDOWN ENGINE
 */
function getOrInitCooldownsSheet(ss) {
  let sheet = ss.getSheetByName("⏱ Cooldowns");
  if (!sheet) {
    sheet = ss.insertSheet("⏱ Cooldowns", 0);
    sheet.appendRow(UNIFIED_COOLDOWN_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_COOLDOWN_HEADERS.length).setFontWeight("bold").setBackground("#ECFDF5");
    sheet.setFrozenRows(1);
  }
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    const defaultAccounts = [
      "vibhavoswal9", "karthikrathod2067", "vibhavoswal2", "nikhildsouza369", "arunaoswal3", "madhuribhalekar369"
    ];
    defaultAccounts.forEach(acc => {
      sheet.appendRow([acc, formatISTTimestamp(new Date(Date.now() - 25 * 3600 * 1000)), "READY", formatISTTimestamp(new Date())]);
    });
  }
  return sheet;
}

function updateCooldownTab(ss, account) {
  const sheet = getOrInitCooldownsSheet(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === account) {
      sheet.getRange(i + 1, 2).setValue(formatISTTimestamp(new Date()));
      sheet.getRange(i + 1, 3).setValue("RUNNING");
      sheet.getRange(i + 1, 4).setValue(formatISTTimestamp(new Date()));
      return;
    }
  }
  sheet.appendRow([account, formatISTTimestamp(new Date()), "RUNNING", formatISTTimestamp(new Date())]);
}

function getOrInitCardsSheet(ss) {
  let sheet = ss.getSheetByName("Cards");
  if (!sheet) {
    sheet = ss.insertSheet("Cards");
    sheet.appendRow(UNIFIED_CARD_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_CARD_HEADERS.length).setFontWeight("bold").setBackground("#DBEAFE");
    sheet.setFrozenRows(1);
  }
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    const defaultCards = [
      ["c1", "PNB", "PNB Imperial RSDC", "'65249710", "'9102", "PNB Imperial Primary", "active", 3, 6, "", "", formatISTTimestamp(new Date())],
      ["c2", "Bank of Baroda", "BoB Bhoomi RSDC", "'652505", "'3314", "Vibhav BoB DC", "active", 2, 5, "", "", formatISTTimestamp(new Date())],
      ["c3", "Canara Bank", "Canara RSDC", "'652550", "'7828", "Canara DC 7828", "active", 2, 5, "", "", formatISTTimestamp(new Date())],
      ["c4", "Jana SFB", "Jana SFB RSDC", "'652520", "'7891", "Jana SFB DC", "active", 3, 5, "", "", formatISTTimestamp(new Date())],
      ["c5", "Canara Bank", "Canara Select CC", "'653020", "'1109", "Canara CC Primary", "active", 1, 2, "", "", formatISTTimestamp(new Date())],
      ["cx1", "PNB", "PNB Imperial RSDC", "'65249710", "'4821", "PNB Imperial (Old)", "blocked", 0, 0, "2026-10-14", "Q3 quotas exhausted", formatISTTimestamp(new Date())],
      ["cx2", "Canara Bank", "Canara RSDC", "'652550", "'1120", "Canara DC 1 (Old)", "blocked", 0, 0, "2026-09-28", "Blocked via mobile app", formatISTTimestamp(new Date())],
      ["cx3", "Bank of Baroda", "BoB Bhoomi RSDC", "'652505", "'2091", "Smiti BoB DC", "blocked", 0, 0, "2026-08-15", "Reissued new card number", formatISTTimestamp(new Date())]
    ];
    defaultCards.forEach(c => sheet.appendRow(c));
  }
  return sheet;
}

function getOrInitSharedActivitySheet(ss) {
  let sheet = ss.getSheetByName("Shared_Activity");
  if (!sheet) {
    sheet = ss.insertSheet("Shared_Activity");
    sheet.appendRow(UNIFIED_SHARED_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_SHARED_HEADERS.length).setFontWeight("bold").setBackground("#D1FAE5");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 6. MIGRATION TO UNIFIED 10-COLUMN SCHEMA ACROSS ALL TABS
 */
function migrateSheetsToUnifiedSchema(ss) {
  const sheets = ss.getSheets();
  const systemSheets = new Set(["⚡ Unassigned", "⏱ Cooldowns", "Cards", "Shared_Activity", "Redeemed"]);
  
  // 1. Ensure master "Redeemed" tab exists with UNIFIED_REDEEMED_HEADERS (14 cols)
  let redeemedSheet = ss.getSheetByName("Redeemed");
  if (!redeemedSheet) {
    redeemedSheet = ss.insertSheet("Redeemed");
  }
  redeemedSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setValues([UNIFIED_REDEEMED_HEADERS]);
  redeemedSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
  redeemedSheet.setFrozenRows(1);

  // 2. Ensure "Shared_Activity" tab has UNIFIED_SHARED_HEADERS (9 cols)
  let sharedSheet = ss.getSheetByName("Shared_Activity");
  if (sharedSheet) {
    sharedSheet.getRange(1, 1, 1, UNIFIED_SHARED_HEADERS.length).setValues([UNIFIED_SHARED_HEADERS]);
    sharedSheet.getRange(1, 1, 1, UNIFIED_SHARED_HEADERS.length).setFontWeight("bold").setBackground("#D1FAE5");
    sharedSheet.setFrozenRows(1);
  }

  const report = [];
  let totalMigratedVouchers = 0;

  for (const sheet of sheets) {
    const name = sheet.getName();
    if (systemSheets.has(name) || name === "Redeemed") continue;

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) continue;

    const rawHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    const headers = rawHeaders.map(h => String(h).trim().toLowerCase());
    
    // Check if "PIN" column exists
    const pinColIndex = headers.findIndex(h => h === "pin" || h === "coupon pin" || h === "voucher pin");

    if (pinColIndex === -1) {
      sheet.insertColumnBefore(4);
      sheet.getRange(1, 4).setValue("PIN").setFontWeight("bold").setBackground("#F3F4F6");
      report.push({ brand: name, status: "INSERTED_PIN_COLUMN", rows: lastRow - 1 });
    } else {
      report.push({ brand: name, status: "ALREADY_HAD_PIN_COLUMN", rows: lastRow - 1 });
    }

    totalMigratedVouchers += Math.max(0, lastRow - 1);

    // Standardize row 1 headers to UNIFIED_BRAND_HEADERS (11 cols)
    sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setValues([UNIFIED_BRAND_HEADERS]);
    sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }

  return {
    success: true,
    totalVouchersInSheet: totalMigratedVouchers,
    report: report,
    migratedAt: formatISTTimestamp(new Date())
  };
}

/**
 * 6B. DEDUPLICATE BRAND SHEETS (Removes duplicate rows where Code is identical)
 */
function deduplicateBrandSheets(ss) {
  const allSheets = ss.getSheets();
  const systemTabs = new Set(["Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns", "⚡ Unassigned"]);
  const report = {};

  for (const s of allSheets) {
    const sName = s.getName();
    if (systemTabs.has(sName)) continue;
    const vals = s.getDataRange().getValues();
    const seenCodes = new Set();
    const toDeleteRows = [];

    for (let i = 1; i < vals.length; i++) {
      const code = String(vals[i][2] || '').trim();
      if (!code || code === '-') continue;
      if (seenCodes.has(code)) {
        toDeleteRows.push(i + 1);
      } else {
        seenCodes.add(code);
      }
    }

    // Delete in descending order so earlier row indices don't shift
    toDeleteRows.sort((a, b) => b - a).forEach(r => s.deleteRow(r));
    if (toDeleteRows.length > 0) {
      report[sName] = toDeleteRows.length;
    }
  }

  SpreadsheetApp.flush();
  return report;
}

/**
 * 7. REST API (GET / POST)
 */
function doGet(e) {
  // SECURITY CHECK: Enforce Vault Auth Token
  const providedAuth = (e && e.parameter) ? (e.parameter.auth || e.parameter.key || e.parameter.token || "") : "";
  if (providedAuth !== VAULT_AUTH_KEY) {
    return ContentService.createTextOutput(JSON.stringify({
      error: "Unauthorized: Invalid or missing vault authentication token.",
      vouchers: [],
      cooldowns: [],
      unassigned: [],
      redeemed: []
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // One-off migration trigger via GET if requested
    if (e && e.parameter && e.parameter.migrate === "true") {
      try {
        const migRes = migrateSheetsToUnifiedSchema(ss);
        return ContentService.createTextOutput(JSON.stringify(migRes)).setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Migration error: " + err.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Clear test timestamps on historical vouchers (e.g. from undo testing)
    if (e && e.parameter && e.parameter.clean_test_timestamps === "true") {
      const brandsToClean = ["CultFit", "TattvaSpa"];
      const cleaned = [];
      brandsToClean.forEach(bName => {
        const sh = ss.getSheetByName(bName);
        if (sh) {
          const data = sh.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            if (data[i][9]) { // Column J
              sh.getRange(i + 1, 10).setValue(""); // clear cell
              cleaned.push({ brand: bName, code: data[i][2], clearedTimestamp: data[i][9] });
            }
          }
        }
      });
      return ContentService.createTextOutput(JSON.stringify({ success: true, cleaned: cleaned })).setMimeType(ContentService.MimeType.JSON);
    }

    // Deduplicate brand sheets on-demand via GET
    if (e && e.parameter && e.parameter.dedup === "true") {
      const dedupReport = deduplicateBrandSheets(ss);
      return ContentService.createTextOutput(JSON.stringify({ success: true, dedupReport: dedupReport })).setMimeType(ContentService.MimeType.JSON);
    }

    // Inspect recent inbox messages for debugging email forwarding and ingestion
    if (e && e.parameter && e.parameter.debug_inbox === "true") {
      const q = (e && e.parameter && e.parameter.q) ? e.parameter.q : "";
      const threads = q ? GmailApp.search(q, 0, 5) : GmailApp.getInboxThreads(0, 5);
      const inboxList = [];
      threads.forEach(t => {
        t.getMessages().forEach(m => {
          inboxList.push({
            id: m.getId(),
            subject: m.getSubject(),
            from: m.getFrom(),
            to: m.getTo(),
            date: m.getDate(),
            isUnread: m.isUnread(),
            snippet: m.getPlainBody().substring(0, 300)
          });
        });
      });
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        query: q,
        count: inboxList.length,
        messages: inboxList
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ON-DEMAND REAL-TIME EMAIL INGESTION
    if (!e || !e.parameter || e.parameter.ingest !== "false") {
      try {
        processRupayEmails();
      } catch (err) {
        console.warn("On-demand email ingestion error:", err.message);
      }
    }

    // Ensure persistent system tabs exist if missing
    if (!ss.getSheetByName("Cards")) {
      getOrInitCardsSheet(ss);
    }
    if (!ss.getSheetByName("⏱ Cooldowns")) {
      getOrInitCooldownsSheet(ss);
    }
    if (!ss.getSheetByName("Shared_Activity")) {
      getOrInitSharedActivitySheet(ss);
    }

    const sheets = ss.getSheets();
    const result = {
      sheetUrl: ss.getUrl(),
      unassigned: [],
      cooldowns: [],
      vouchers: [],
      redeemed: [],
      cards: [],
      shared: []
    };

    for (const sheet of sheets) {
      const name = sheet.getName();
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) continue;

      if (name === "⚡ Unassigned") {
        for (let i = 1; i < rows.length; i++) {
          result.unassigned.push({
            bookingId: rows[i][0],
            brand: rows[i][1],
            amount: rows[i][2],
            account: rows[i][3],
            dateOfIssue: rows[i][4]
          });
        }
      } else if (name === "⏱ Cooldowns") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] || String(rows[i][0]).trim() === "") continue;
          result.cooldowns.push({
            account: String(rows[i][0] || ''),
            lastRedeemed: rows[i][1] ? String(rows[i][1]) : '',
            status: rows[i][2] ? String(rows[i][2]) : 'READY',
            lastUpdated: rows[i][3] ? String(rows[i][3]) : ''
          });
        }
      } else if (name === "Cards") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] || String(rows[i][0]).trim() === "") continue;
          result.cards.push({
            id: String(rows[i][0] || ''),
            bank: String(rows[i][1] || ''),
            type: String(rows[i][2] || ''),
            bin: String(rows[i][3] || ''),
            last4: String(rows[i][4] || ''),
            nick: String(rows[i][5] || ''),
            status: String(rows[i][6] || 'active').toLowerCase(),
            claimed: Number(rows[i][7]) || 0,
            total: Number(rows[i][8]) || 5,
            blockedDate: rows[i][9] ? String(rows[i][9]) : '',
            notes: rows[i][10] ? String(rows[i][10]) : '',
            lastUpdated: rows[i][11] ? String(rows[i][11]) : ''
          });
        }
      } else if (name === "Shared_Activity") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][1] || String(rows[i][1]).trim() === "") continue;
          const status = String(rows[i][6] || 'SHARED').trim().toUpperCase();
          if (status === "DISMISSED") continue;
          const isDisp = rows[i].length > 8 ? (rows[i][8] !== false && String(rows[i][8]).toUpperCase() !== 'FALSE') : true;
          result.shared.push({
            code: String(rows[i][1] || ''),
            brand: String(rows[i][2] || ''),
            amount: Number(rows[i][3]) || 0,
            pin: rows[i][4] && String(rows[i][4]).trim() !== '' && String(rows[i][4]).trim() !== 'null' ? String(rows[i][4]).trim() : null,
            sharedAt: String(rows[i][5] || ''),
            status: status,
            platform: String(rows[i][7] || 'WhatsApp'),
            displayInUI: isDisp
          });
        }
      } else if (name === "Redeemed") {
        // Unified Redeemed Archive (Col 14: Display in UI)
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][3] || String(rows[i][3]).trim() === "") continue;
          const isDisp = rows[i].length > 13 ? (rows[i][13] !== false && String(rows[i][13]).toUpperCase() !== 'FALSE') : true;
          result.redeemed.push({
            brand: String(rows[i][1] || ''),
            bookingId: String(rows[i][2] || ''),
            code: String(rows[i][3] || ''),
            pin: rows[i][4] && String(rows[i][4]).trim() !== '' && String(rows[i][4]).trim() !== 'null' ? String(rows[i][4]).trim() : null,
            amount: Number(rows[i][5]) || 0,
            dateOfIssue: rows[i][6] ? String(rows[i][6]) : '',
            account: rows[i][7] ? String(rows[i][7]) : '',
            card: rows[i][8] ? String(rows[i][8]) : '',
            status: "REDEEMED",
            redeemedTimestampIST: rows[i][10] ? String(rows[i][10]) : '',
            originalTab: String(rows[i][11] || rows[i][1] || ''),
            displayInUI: isDisp
          });
        }
      } else {
        // Uniform Brand Tabs (Col 11: Display in UI)
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][2] || String(rows[i][2]).trim() === "") continue;
          const isDisp = rows[i].length > 10 ? (rows[i][10] !== false && String(rows[i][10]).toUpperCase() !== 'FALSE') : true;
          result.vouchers.push({
            brand: name,
            bookingId: String(rows[i][1] || ''),
            code: String(rows[i][2] || ''),
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            amount: Number(rows[i][4]) || 0,
            dateOfIssue: rows[i][5] ? String(rows[i][5]) : '',
            account: rows[i][6] ? String(rows[i][6]) : '',
            card: rows[i][7] ? String(rows[i][7]) : '',
            status: rows[i][8] ? String(rows[i][8]) : 'ACTIVE',
            timestampIST: rows[i][9] ? String(rows[i][9]) : null,
            displayInUI: isDisp
          });
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: "doGet error: " + err.message,
      vouchers: [],
      cooldowns: [],
      unassigned: [],
      redeemed: [],
      cards: [],
      shared: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const data = JSON.parse(e.postData.contents);
    const providedAuth = data.auth || data.token || (e.parameter ? (e.parameter.auth || e.parameter.key) : "") || "";
    if (providedAuth !== VAULT_AUTH_KEY) {
      return ContentService.createTextOutput(JSON.stringify({
        error: "Unauthorized: Invalid or missing vault authentication token."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 1. Run Migration on Demand
    if (data.action === "MIGRATE_UNIFIED_SCHEMA") {
      const migRes = migrateSheetsToUnifiedSchema(ss);
      return ContentService.createTextOutput(JSON.stringify(migRes)).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Assign Cards to Redemptions
    if (data.action === "ASSIGN_CARDS") {
      const res = assignCardToRedemptions(ss, data.assignments);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Move Voucher to "Redeemed" Tab (Approach A - True Cut & Paste with Deduplication Lock)
    if (data.action === "MOVE_TO_REDEEMED" || data.action === "MARK_REDEEMED") {
      let rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        rSheet = ss.insertSheet("Redeemed");
        rSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
        rSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
        rSheet.setFrozenRows(1);
      }

      const rawCodes = data.codes || (data.code ? [data.code] : []);
      // Unique codes only
      const codes = Array.from(new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []));
      if (!codes.length && data.bookingId) {
        codes.push(String(data.bookingId).trim());
      }

      const moved = [];
      const systemTabs = new Set(["Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns", "⚡ Unassigned"]);

      // Pre-fetch Redeemed sheet codes to prevent duplicate entries in Redeemed tab
      const rVals = rSheet.getDataRange().getValues();
      const existingRedeemedCodes = new Set();
      for (let r = 1; r < rVals.length; r++) {
        const c = String(rVals[r][3] || '').trim();
        if (c) existingRedeemedCodes.add(c);
      }

      codes.forEach(codeVal => {
        let targetSheet = null;
        let matchingRowIndices = [];
        let rowData = null;

        // 1. Try provided brand if present
        if (data.brand) {
          const directSheet = ss.getSheetByName(data.brand);
          if (directSheet && directSheet.getName() !== "Redeemed") {
            const vals = directSheet.getDataRange().getValues();
            for (let i = 1; i < vals.length; i++) {
              if (String(vals[i][2]).trim() === codeVal || String(vals[i][1]).trim() === codeVal) {
                targetSheet = directSheet;
                matchingRowIndices.push(i + 1);
                if (!rowData) rowData = vals[i];
              }
            }
          }
        }

        // 2. If not found in provided brand, scan ALL brand sheets (excluding system tabs)
        if (!targetSheet) {
          const allSheets = ss.getSheets();
          for (const s of allSheets) {
            const sName = s.getName();
            if (systemTabs.has(sName)) continue;
            const vals = s.getDataRange().getValues();
            for (let i = 1; i < vals.length; i++) {
              if (String(vals[i][2]).trim() === codeVal || String(vals[i][1]).trim() === codeVal) {
                targetSheet = s;
                matchingRowIndices.push(i + 1);
                if (!rowData) rowData = vals[i];
              }
            }
            if (targetSheet) break;
          }
        }

        if (targetSheet && rowData) {
          const origBrand = targetSheet.getName();

          // DEDUPLICATION LOCK: Only append to Redeemed if not already present in Redeemed!
          if (!existingRedeemedCodes.has(codeVal)) {
            const redeemedRow = [
              rSheet.getLastRow(),                    // A: #
              origBrand,                              // B: Brand
              rowData[1],                             // C: Booking ID
              `'${rowData[2]}`,                       // D: Code
              rowData[3] ? `'${rowData[3]}` : "",     // E: PIN
              rowData[4],                             // F: Value
              rowData[5],                             // G: Date of Issue
              rowData[6],                             // H: Account
              rowData[7],                             // I: Card
              "REDEEMED",                             // J: Status
              formatISTTimestamp(new Date()),         // K: Redeemed Timestamp (IST)
              origBrand,                              // L: Original Tab
              rowData[9] ? String(rowData[9]) : "",   // M: Original Ingested Timestamp (if any)
              true                                    // N: Display in UI
            ];
            rSheet.appendRow(redeemedRow);
            existingRedeemedCodes.add(codeVal);
          }

          // ATOMIC CUT: Delete all matching instances from target brand sheet in reverse (descending) order
          matchingRowIndices.sort((a, b) => b - a).forEach(idx => {
            targetSheet.deleteRow(idx);
          });

          moved.push({ code: codeVal, brand: origBrand });
        }
      });

      // Force atomic persistence
      SpreadsheetApp.flush();

      // Also update Shared_Activity status to REDEEMED for these codes if present
      const sSheet = ss.getSheetByName("Shared_Activity");
      if (sSheet && moved.length > 0) {
        const movedCodesSet = new Set(moved.map(m => m.code));
        const sRows = sSheet.getDataRange().getValues();
        for (let i = 1; i < sRows.length; i++) {
          if (movedCodesSet.has(String(sRows[i][1]).trim())) {
            sSheet.getRange(i + 1, 7).setValue("REDEEMED");
          }
        }
      }

      if (moved.length > 0) {
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          action: "MOVE_TO_REDEEMED",
          movedCount: moved.length,
          moved: moved
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        error: "Voucher(s) not found in any brand tab: " + codes.join(", ")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Undo Redemption (Restore to Brand Tab - Cut from Redeemed, Paste into Brand with Deduplication Lock)
    if (data.action === "UNDO_REDEEM") {
      const rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Redeemed sheet not found"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = Array.from(new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []));
      if (!codes.length && data.bookingId) {
        codes.push(String(data.bookingId).trim());
      }

      const restored = [];
      const rRows = rSheet.getDataRange().getValues();
      const rowsToDeleteFromRedeemed = [];

      // Search backwards so deleting rows does not alter row index of earlier rows
      for (let i = rRows.length - 1; i >= 1; i--) {
        const rowCode = String(rRows[i][3]).trim();
        const rowBooking = String(rRows[i][2]).trim();

        if (codes.includes(rowCode) || codes.includes(rowBooking)) {
          const origBrand = String(rRows[i][11] || rRows[i][1]);
          let bSheet = ss.getSheetByName(origBrand);
          if (!bSheet) {
            bSheet = ss.insertSheet(origBrand);
            bSheet.appendRow(UNIFIED_BRAND_HEADERS);
            bSheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
            bSheet.setFrozenRows(1);
          }

          // DEDUPLICATION LOCK: Inspect brand sheet first! Does this code already exist in brand sheet?
          const bVals = bSheet.getDataRange().getValues();
          let existingBrandRowIndex = -1;
          for (let b = 1; b < bVals.length; b++) {
            if (String(bVals[b][2]).trim() === rowCode) {
              existingBrandRowIndex = b + 1;
              break;
            }
          }

          if (existingBrandRowIndex > 0) {
            // ALREADY EXISTS in brand tab! Do NOT append a duplicate row.
            // Just ensure its Status is ACTIVE and Display in UI is true.
            bSheet.getRange(existingBrandRowIndex, 9).setValue("ACTIVE");
            bSheet.getRange(existingBrandRowIndex, 11).setValue(true);
          } else {
            // DOES NOT EXIST in brand tab: Safe to append!
            const originalTimestamp = (rRows[i].length > 12 && rRows[i][12]) ? String(rRows[i][12]) : "";
            const restoredRow = [
              bSheet.getLastRow(),                    // A: #
              rRows[i][2],                            // B: Booking ID
              `'${rRows[i][3]}`,                      // C: Code
              rRows[i][4] ? `'${rRows[i][4]}` : "",   // D: PIN
              rRows[i][5],                            // E: Value
              rRows[i][6],                            // F: Date of Issue
              rRows[i][7],                            // G: Account
              rRows[i][8],                            // H: Card
              "ACTIVE",                               // I: Status
              originalTimestamp,                      // J: Appended Timestamp (IST)
              true                                    // K: Display in UI
            ];
            bSheet.appendRow(restoredRow);
          }

          // Mark row in Redeemed for deletion
          rowsToDeleteFromRedeemed.push(i + 1);
          restored.push({ code: rowCode, brand: origBrand });
        }
      }

      // Delete from Redeemed sheet in reverse (descending) order
      rowsToDeleteFromRedeemed.sort((a, b) => b - a).forEach(rowIdx => {
        rSheet.deleteRow(rowIdx);
      });

      // Force atomic persistence
      SpreadsheetApp.flush();

      if (restored.length > 0) {
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          action: "UNDO_REDEEM",
          restoredCount: restored.length,
          restored: restored
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        error: "Voucher(s) not found in Redeemed sheet: " + codes.join(", ")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4B. Deduplicate Brand Sheets Action
    if (data.action === "DEDUPLICATE_SHEETS") {
      const dedupReport = deduplicateBrandSheets(ss);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        report: dedupReport
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. Save Card (Add or Update in "Cards" tab)
    if (data.action === "SAVE_CARD") {
      const cSheet = getOrInitCardsSheet(ss);
      const card = data.card || data;
      const cardId = String(card.id || ('c-' + Date.now())).trim();
      const rows = cSheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === cardId) {
          cSheet.getRange(i + 1, 2).setValue(card.bank || rows[i][1]);
          cSheet.getRange(i + 1, 3).setValue(card.type || rows[i][2]);
          cSheet.getRange(i + 1, 4).setValue(card.bin ? `'${card.bin}` : rows[i][3]);
          cSheet.getRange(i + 1, 5).setValue(card.last4 ? `'${card.last4}` : rows[i][4]);
          cSheet.getRange(i + 1, 6).setValue(card.nick || rows[i][5]);
          cSheet.getRange(i + 1, 7).setValue(card.status || rows[i][6]);
          cSheet.getRange(i + 1, 8).setValue(card.claimed !== undefined ? Number(card.claimed) : rows[i][7]);
          cSheet.getRange(i + 1, 9).setValue(card.total !== undefined ? Number(card.total) : rows[i][8]);
          cSheet.getRange(i + 1, 10).setValue(card.blockedDate || rows[i][9]);
          cSheet.getRange(i + 1, 11).setValue(card.notes || rows[i][10]);
          cSheet.getRange(i + 1, 12).setValue(formatISTTimestamp(new Date()));
          found = true;
          break;
        }
      }

      if (!found) {
        cSheet.appendRow([
          cardId,
          card.bank || 'RuPay Card',
          card.type || 'RuPay Select',
          card.bin ? `'${card.bin}` : '',
          card.last4 ? `'${card.last4}` : '',
          card.nick || 'New Card',
          card.status || 'active',
          card.claimed !== undefined ? Number(card.claimed) : 0,
          card.total !== undefined ? Number(card.total) : 5,
          card.blockedDate || '',
          card.notes || '',
          formatISTTimestamp(new Date())
        ]);
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "SAVE_CARD",
        id: cardId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. Block Card
    if (data.action === "BLOCK_CARD") {
      const cSheet = getOrInitCardsSheet(ss);
      const cardId = String(data.id || '').trim();
      const blockedDate = data.blockedDate || formatISTDate(new Date());
      const notes = data.notes || 'Blocked via App';
      const rows = cSheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === cardId) {
          cSheet.getRange(i + 1, 7).setValue('blocked');
          cSheet.getRange(i + 1, 10).setValue(blockedDate);
          cSheet.getRange(i + 1, 11).setValue(notes);
          cSheet.getRange(i + 1, 12).setValue(formatISTTimestamp(new Date()));
          found = true;
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: found,
        action: "BLOCK_CARD",
        id: cardId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 7. Update Card Quota (Quick Claim)
    if (data.action === "UPDATE_CARD_QUOTA") {
      const cSheet = getOrInitCardsSheet(ss);
      const cardId = String(data.id || '').trim();
      const rows = cSheet.getDataRange().getValues();
      let newClaimed = 0;
      let found = false;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === cardId) {
          if (data.claimed !== undefined) {
            newClaimed = Number(data.claimed);
          } else {
            newClaimed = (Number(rows[i][7]) || 0) + 1;
          }
          cSheet.getRange(i + 1, 8).setValue(newClaimed);
          cSheet.getRange(i + 1, 12).setValue(formatISTTimestamp(new Date()));
          found = true;
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: found,
        action: "UPDATE_CARD_QUOTA",
        id: cardId,
        claimed: newClaimed
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 8. Log Shared Activity
    if (data.action === "LOG_SHARED") {
      const sSheet = getOrInitSharedActivitySheet(ss);
      const items = Array.isArray(data.items) ? data.items : [{
        code: data.code,
        brand: data.brand,
        amount: data.amount,
        pin: data.pin,
        platform: data.platform
      }];
      const sRows = sSheet.getDataRange().getValues();
      const logged = [];

      items.forEach(item => {
        if (!item.code) return;
        const cStr = String(item.code).trim();
        let foundIndex = -1;
        for (let i = 1; i < sRows.length; i++) {
          if (String(sRows[i][1]).trim() === cStr) {
            foundIndex = i + 1;
            break;
          }
        }

        const timeStr = formatISTTimestamp(new Date());
        if (foundIndex > 0) {
          sSheet.getRange(foundIndex, 6).setValue(timeStr);
          sSheet.getRange(foundIndex, 7).setValue("SHARED");
          if (item.platform) sSheet.getRange(foundIndex, 8).setValue(item.platform);
          sSheet.getRange(foundIndex, 9).setValue(true);
        } else {
          const nextNum = sSheet.getLastRow();
          sSheet.appendRow([
            nextNum,
            `'${cStr}`,
            item.brand || 'Voucher',
            item.amount || 0,
            item.pin ? `'${item.pin}` : '',
            timeStr,
            "SHARED",
            item.platform || "WhatsApp",
            true
          ]);
        }
        logged.push(cStr);
      });

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "LOG_SHARED",
        count: logged.length,
        codes: logged
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 9. Dismiss Shared Activity
    if (data.action === "DISMISS_SHARED") {
      const sSheet = getOrInitSharedActivitySheet(ss);
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : [];
      const sRows = sSheet.getDataRange().getValues();
      const dismissed = [];

      for (let i = 1; i < sRows.length; i++) {
        const rowCode = String(sRows[i][1]).trim();
        if (codes.includes(rowCode)) {
          sSheet.getRange(i + 1, 7).setValue("DISMISSED");
          dismissed.push(rowCode);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "DISMISS_SHARED",
        dismissed: dismissed
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 10. Reset 24-Hour Account Cooldown
    if (data.action === "RESET_COOLDOWN") {
      const cdSheet = getOrInitCooldownsSheet(ss);
      const account = String(data.account || data.accountId || '').trim();
      if (!account) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Missing account ID"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      updateCooldownTab(ss, account);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "RESET_COOLDOWN",
        account: account,
        timestamp: formatISTTimestamp(new Date())
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 11. Verify Master Passcode (24-Hour Session TTL)
    if (data.action === "VERIFY_PASSCODE") {
      const entered = String(data.passcode || "").trim();
      const scriptProps = PropertiesService.getScriptProperties();
      const masterPasscode = scriptProps.getProperty("VAULT_MASTER_PASSCODE");

      if (entered === masterPasscode) {
        const sessionDurationMs = 24 * 60 * 60 * 1000; // 24 Hours TTL
        const expiresAt = Date.now() + sessionDurationMs;
        const rawToken = Utilities.computeHmacSha256Signature(expiresAt + ":" + VAULT_AUTH_KEY, masterPasscode);
        const sessionToken = rawToken.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          expiresAt: expiresAt,
          sessionToken: sessionToken
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Incorrect passcode. Please try again."
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 12. Change Master Passcode in Cloud Properties
    if (data.action === "CHANGE_PASSCODE") {
      const current = String(data.currentPasscode || "").trim();
      const next = String(data.newPasscode || "").trim();
      const scriptProps = PropertiesService.getScriptProperties();
      const masterPasscode = scriptProps.getProperty("VAULT_MASTER_PASSCODE");

      if (current !== masterPasscode) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Current passcode is incorrect."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!/^\d{6}$/.test(next)) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "New passcode must be exactly 6 digits."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      scriptProps.setProperty("VAULT_MASTER_PASSCODE", next);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "Passcode updated successfully in cloud properties."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 13. Direct Setup / Override of Passcode in Cloud Properties
    if (data.action === "SET_PASSCODE") {
      const code = String(data.passcode || "").trim();
      if (/^\d{6}$/.test(code)) {
        PropertiesService.getScriptProperties().setProperty("VAULT_MASTER_PASSCODE", code);
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          message: "Passcode configured in ScriptProperties."
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 14. Clear Activity Log Items (sets Display in UI = FALSE in brand tabs or Shared_Activity)
    if (data.action === "CLEAR_ACTIVITY") {
      const type = String(data.type || "INGESTED").toUpperCase();
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []);
      const updated = [];

      if (type === "SHARED") {
        const sh = ss.getSheetByName("Shared_Activity");
        if (sh) {
          const vals = sh.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowCode = String(vals[i][1] || "").trim();
            if (codes.size === 0 || codes.has(rowCode)) {
              sh.getRange(i + 1, 9).setValue(false); // Col 9: Display in UI
              sh.getRange(i + 1, 7).setValue("DISMISSED");
              updated.push(rowCode);
            }
          }
        }
      } else {
        // INGESTED: Search in Brand tabs
        const sheets = ss.getSheets();
        const systemSheets = new Set(["⚡ Unassigned", "⏱ Cooldowns", "Cards", "Shared_Activity", "Redeemed"]);
        for (const sheet of sheets) {
          if (systemSheets.has(sheet.getName())) continue;
          const vals = sheet.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowCode = String(vals[i][2] || "").trim();
            if (codes.size === 0 || codes.has(rowCode)) {
              sheet.getRange(i + 1, 11).setValue(false); // Col 11: Display in UI
              updated.push(rowCode);
            }
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "CLEAR_ACTIVITY",
        type: type,
        clearedCount: updated.length,
        clearedCodes: updated
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 15. Clear Redeemed Items (sets Display in UI = FALSE in Redeemed tab)
    if (data.action === "CLEAR_REDEEMED") {
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []);
      const updated = [];

      const rSheet = ss.getSheetByName("Redeemed");
      if (rSheet) {
        const vals = rSheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
          const rowCode = String(vals[i][3] || "").trim();
          if (codes.size === 0 || codes.has(rowCode)) {
            rSheet.getRange(i + 1, 14).setValue(false); // Col 14: Display in UI
            updated.push(rowCode);
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "CLEAR_REDEEMED",
        clearedCount: updated.length,
        clearedCodes: updated
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * EXACT IST TIME FORMATTERS (Indian Standard Time: Asia/Kolkata)
 */
function formatISTTimestamp(d) {
  return Utilities.formatDate(d, IST_TIMEZONE, "dd-MMM-yyyy hh:mm:ss a 'IST'");
}

function formatISTDate(d) {
  return Utilities.formatDate(d, IST_TIMEZONE, "dd-MMM-yyyy");
}
