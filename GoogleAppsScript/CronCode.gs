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
  "#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)"
];

const UNIFIED_REDEEMED_HEADERS = [
  "#", "Brand", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Redeemed Timestamp (IST)", "Original Tab"
];

/**
 * 1. AUTOMATIC EMAIL INGESTION (Triggered via cron or instant on-demand sync)
 */
function processRupayEmails() {
  // Query handles both direct RuPay emails and emails forwarded from any personal address
  const query = 'is:unread (from:rupay@truztee.com OR subject:"Fwd:" OR subject:"Fw:" OR "rupay@truztee.com" OR "Booking ID:")';
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

      // 1. Detect if this is a forwarded email
      const isForwarded = /^(?:Fwd|Fw):/i.test(subject) || 
                          body.includes("---------- Forwarded message ---------") ||
                          /Forwarded message/i.test(body) ||
                          /Begin forwarded message/i.test(body);

      // 2. Extract original recipient account
      let account = "";
      if (isForwarded) {
        // Look for nested "To: <email>" inside the forwarded headers block
        const fwdToMatch = body.match(/To:\s*<?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>?/i);
        if (fwdToMatch && !fwdToMatch[1].toLowerCase().includes("rupayvouchers@gmail.com")) {
          account = extractAccountName(fwdToMatch[1], "");
        } else {
          // If forwarded from personal email, fallback to forwarder's address
          account = extractAccountName(fromHeader, "");
        }
      } else {
        account = extractAccountName(toHeader, body);
      }
      if (!account) account = "vibhavoswal9";

      // 3. Extract Booking ID, Brand, Amount, Date
      const bookingIdMatch = subject.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i) || body.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i);
      const bookingId = bookingIdMatch ? bookingIdMatch[1].trim() : "";
      const brand = extractBrand(subject, body);
      const amount = extractAmount(body);
      const dateOfIssue = extractField(body, "Date of Issue") || formatISTDate(new Date());

      // 4. Extract Coupon Code and PIN
      const couponCode = extractField(body, "Coupon Code") || extractField(body, "Voucher Code") || extractField(body, "Code");
      const pin = extractField(body, "Coupon Pin") || extractField(body, "PIN") || extractField(body, "Voucher Pin") || extractField(body, "Card Pin") || "";

      const isPending = !couponCode || 
                        /will be sent/i.test(couponCode) || 
                        /blank/i.test(couponCode) || 
                        couponCode.toUpperCase() === "NA";

      if (isPending) {
        // STAGE 1: Acknowledgment Email
        // 1. Start the 24h cooldown timer on this account
        updateCooldownTab(ss, account);

        // 2. Put into "Unassigned Queue" for rapid 1-tap card assignment in the app
        queueUnassignedRedemption(ss, {
          bookingId: bookingId,
          brand: brand,
          amount: amount,
          account: account,
          dateOfIssue: dateOfIssue
        });
      } else {
        // STAGE 2: Code Delivery (24-48 hours later)
        // Find by Booking ID and slot the code into the correct Brand Tab with IST Timestamp
        fulfillVoucherCode(ss, {
          bookingId: bookingId,
          brand: brand,
          code: couponCode,
          pin: pin,
          amount: amount,
          dateOfIssue: dateOfIssue,
          account: account
        });
      }

      msg.markRead();
    }
  }
}

/**
 * Extracts sender email and drops '@gmail.com' -> e.g. "vibhavoswal9"
 */
function extractAccountName(toHeader, body) {
  let email = toHeader.toLowerCase().trim();
  const angle = email.match(/<([^>]+)>/);
  if (angle) email = angle[1];

  if (email.includes("rupayvouchers@gmail.com")) {
    const fwdMatch = body.match(/To:\s*<?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>?/i);
    if (fwdMatch && !fwdMatch[1].includes("rupayvouchers@gmail.com")) {
      email = fwdMatch[1].toLowerCase().trim();
    }
  }

  return email.replace(/@.*$/, "").trim();
}

function extractField(body, fieldName) {
  const re = new RegExp(fieldName + ":\\s*([^\n\r]+)", "i");
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

  return "Other";
}

function extractAmount(body) {
  const m = body.match(/INR\s*([0-9,]+)/i) || body.match(/Rs\.?\s*([0-9,]+)/i);
  return m ? parseInt(m[1].replace(/,/g, "")) : 500;
}

/**
 * 2. UNASSIGNED QUEUE TAB (With IST Timestamp)
 */
function queueUnassignedRedemption(ss, data) {
  let sheet = ss.getSheetByName("⚡ Unassigned");
  if (!sheet) {
    sheet = ss.insertSheet("⚡ Unassigned", 0);
    sheet.appendRow(["Booking ID", "Brand", "Value", "Account", "Date of Issue", "Assigned Card", "Created Timestamp (IST)"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#FEF3C7");
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookingId) return;
  }

  sheet.appendRow([
    data.bookingId,
    data.brand,
    data.amount,
    data.account,
    data.dateOfIssue,
    "NEEDS_ASSIGNMENT",
    formatISTTimestamp(new Date()) // Separate IST Timestamp Column
  ]);
}

/**
 * 3. BRAND TAB FULFILLMENT (Standard 10 Columns)
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

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.bookingId) {
      sheet.getRange(i + 1, 3).setValue(`'${data.code}`);
      if (data.pin) {
        sheet.getRange(i + 1, 4).setValue(`'${data.pin}`);
      }
      sheet.getRange(i + 1, 9).setValue("ACTIVE");
      sheet.getRange(i + 1, 10).setValue(formatISTTimestamp(new Date()));
      return;
    }
  }

  const nextNum = rows.length;
  const rowData = [
    nextNum,
    data.bookingId,
    `'${data.code}`,
    data.pin ? `'${data.pin}` : "",
    data.amount || 500,
    data.dateOfIssue,
    data.account,
    "Unassigned",
    "ACTIVE",
    formatISTTimestamp(new Date())
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
 * 5. 24-HOUR COOLDOWN UPDATE (With IST Timestamp)
 */
function updateCooldownTab(ss, account) {
  let sheet = ss.getSheetByName("⏱ Cooldowns");
  if (!sheet) {
    sheet = ss.insertSheet("⏱ Cooldowns", 0);
    sheet.appendRow(["Account ID", "Last Redemption Time (IST)", "Status", "Last Updated (IST)"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#ECFDF5");
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === account) {
      sheet.getRange(i + 1, 2).setValue(formatISTTimestamp(new Date()));
      sheet.getRange(i + 1, 3).setValue("RUNNING");
      sheet.getRange(i + 1, 4).setValue(formatISTTimestamp(new Date()));
      return;
    }
  }
  sheet.appendRow([account, formatISTTimestamp(new Date()), "RUNNING", formatISTTimestamp(new Date())]);
}

/**
 * 6. MIGRATION TO UNIFIED 10-COLUMN SCHEMA ACROSS ALL TABS
 */
function migrateSheetsToUnifiedSchema(ss) {
  const sheets = ss.getSheets();
  const systemSheets = new Set(["⚡ Unassigned", "⏱ Cooldowns"]);
  
  // 1. Ensure master "Redeemed" tab exists with 12 columns
  let redeemedSheet = ss.getSheetByName("Redeemed");
  if (!redeemedSheet) {
    redeemedSheet = ss.insertSheet("Redeemed");
    redeemedSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
    redeemedSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
    redeemedSheet.setFrozenRows(1);
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
      // 9-column sheet missing PIN: Col 1=#, Col 2=Booking ID, Col 3=Code, Col 4=Value...
      // Safely insert PIN column before column 4 (Column D)
      sheet.insertColumnBefore(4);
      sheet.getRange(1, 4).setValue("PIN").setFontWeight("bold").setBackground("#F3F4F6");
      report.push({ brand: name, status: "INSERTED_PIN_COLUMN", rows: lastRow - 1 });
    } else {
      report.push({ brand: name, status: "ALREADY_HAD_PIN_COLUMN", rows: lastRow - 1 });
    }

    totalMigratedVouchers += Math.max(0, lastRow - 1);

    // Standardize row 1 headers to UNIFIED_BRAND_HEADERS
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

  // ON-DEMAND REAL-TIME EMAIL INGESTION
  if (!e || !e.parameter || e.parameter.ingest !== "false") {
    try {
      processRupayEmails();
    } catch (err) {
      console.warn("On-demand email ingestion error:", err.message);
    }
  }

  const sheets = ss.getSheets();
  const result = {
    sheetUrl: ss.getUrl(),
    unassigned: [],
    cooldowns: [],
    vouchers: [],
    redeemed: []
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
        result.cooldowns.push({
          account: rows[i][0],
          lastRedeemed: rows[i][1]
        });
      }
    } else if (name === "Redeemed") {
      // 12-Column Unified Redeemed Archive
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i][3] || String(rows[i][3]).trim() === "") continue;
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
          originalTab: String(rows[i][11] || rows[i][1] || '')
        });
      }
    } else {
      // Uniform 10-Column Brand Tabs
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i][2] || String(rows[i][2]).trim() === "") continue;
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
          timestampIST: rows[i][9] ? String(rows[i][9]) : null
        });
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
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

    // 3. Move Voucher to "Redeemed" Tab (Approach A - Cut & Paste)
    if (data.action === "MOVE_TO_REDEEMED") {
      const bSheet = ss.getSheetByName(data.brand);
      if (!bSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Brand tab not found: " + data.brand
        })).setMimeType(ContentService.MimeType.JSON);
      }

      let rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        rSheet = ss.insertSheet("Redeemed");
        rSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
        rSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
        rSheet.setFrozenRows(1);
      }

      const rows = bSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][2] === data.code || rows[i][1] === data.bookingId) {
          const row = rows[i];
          const redeemedRow = [
            rSheet.getLastRow(),                    // A: #
            data.brand,                             // B: Brand
            row[1],                                 // C: Booking ID
            `'${row[2]}`,                           // D: Code
            row[3] ? `'${row[3]}` : "",             // E: PIN
            row[4],                                 // F: Value
            row[5],                                 // G: Date of Issue
            row[6],                                 // H: Account
            row[7],                                 // I: Card
            "REDEEMED",                             // J: Status
            formatISTTimestamp(new Date()),         // K: Redeemed Timestamp (IST)
            data.brand                              // L: Original Tab
          ];
          rSheet.appendRow(redeemedRow);
          bSheet.deleteRow(i + 1); // Delete from active brand tab
          return ContentService.createTextOutput(JSON.stringify({
            success: true,
            action: "MOVE_TO_REDEEMED",
            code: data.code,
            brand: data.brand
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        error: "Voucher not found in " + data.brand
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Undo Redemption (Restore to Brand Tab)
    if (data.action === "UNDO_REDEEM") {
      const rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Redeemed sheet not found"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const rRows = rSheet.getDataRange().getValues();
      for (let i = 1; i < rRows.length; i++) {
        if (rRows[i][3] === data.code || rRows[i][2] === data.bookingId) {
          const origBrand = String(rRows[i][11] || rRows[i][1]);
          let bSheet = ss.getSheetByName(origBrand);
          if (!bSheet) {
            bSheet = ss.insertSheet(origBrand);
            bSheet.appendRow(UNIFIED_BRAND_HEADERS);
            bSheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
            bSheet.setFrozenRows(1);
          }

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
            formatISTTimestamp(new Date())          // J: Appended Timestamp (IST)
          ];

          bSheet.appendRow(restoredRow);
          rSheet.deleteRow(i + 1); // Remove from Redeemed sheet
          return ContentService.createTextOutput(JSON.stringify({
            success: true,
            action: "UNDO_REDEEM",
            code: data.code,
            brand: origBrand
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        error: "Voucher not found in Redeemed sheet"
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
