/**
 * RuPay Select Vault — Google Apps Script Engine
 * 100% Free Cloud Backend (Gmail -> Sheet -> Swift App REST API)
 * All timestamps strictly formatted in IST (Indian Standard Time: Asia/Kolkata)
 */

const SENDER_EMAIL = "rupay@truztee.com";
const IST_TIMEZONE = "Asia/Kolkata";

// Brands that require a PIN column vs Brands that do not
const BRANDS_WITH_PIN = new Set([
  "Lakme", "FourFountains", "Kalyan", "Myntra", "Apollo", "BigBasket"
]);

/**
 * 1. AUTOMATIC EMAIL INGESTION (Triggered every 1 min)
 */
function processRupayEmails() {
  const query = `is:unread from:${SENDER_EMAIL}`;
  const threads = GmailApp.search(query);
  if (!threads.length) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      if (!msg.isUnread()) continue;

      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      const toHeader = msg.getTo();

      const bookingIdMatch = subject.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i) || body.match(/Booking ID:\s*([A-Za-z0-9\-]+)/i);
      const bookingId = bookingIdMatch ? bookingIdMatch[1].trim() : "";
      const account = extractAccountName(toHeader, body);
      const brand = extractBrand(subject, body);
      const amount = extractAmount(body);
      const dateOfIssue = extractField(body, "Date of Issue") || formatISTDate(new Date());
      const couponCode = extractField(body, "Coupon Code");

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
 * 3. BRAND TAB FULFILLMENT (With IST Timestamp)
 */
function fulfillVoucherCode(ss, data) {
  const brand = data.brand;
  const hasPin = BRANDS_WITH_PIN.has(brand);
  let sheet = ss.getSheetByName(brand);

  if (!sheet) {
    sheet = ss.insertSheet(brand);
    const headers = hasPin 
      ? ["#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)"]
      : ["#", "Booking ID", "Code", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.bookingId) {
      sheet.getRange(i + 1, 3).setValue(`'${data.code}`);
      const statusCol = hasPin ? 9 : 8;
      sheet.getRange(i + 1, statusCol).setValue("ACTIVE");
      sheet.getRange(i + 1, statusCol + 1).setValue(formatISTTimestamp(new Date()));
      return;
    }
  }

  const nextNum = rows.length;
  const rowData = hasPin 
    ? [nextNum, data.bookingId, `'${data.code}`, "", 500, data.dateOfIssue, data.account, "Unassigned", "ACTIVE", formatISTTimestamp(new Date())]
    : [nextNum, data.bookingId, `'${data.code}`, 100, data.dateOfIssue, data.account, "Unassigned", "ACTIVE", formatISTTimestamp(new Date())];
  sheet.appendRow(rowData);
}

/**
 * 4. ASSIGN CARD IN BULK (Called by Swift App)
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

        const hasPin = BRANDS_WITH_PIN.has(brand);
        let bSheet = ss.getSheetByName(brand);
        if (!bSheet) {
          bSheet = ss.insertSheet(brand);
          const headers = hasPin 
            ? ["#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)"]
            : ["#", "Booking ID", "Code", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)"];
          bSheet.appendRow(headers);
          bSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
          bSheet.setFrozenRows(1);
        }

        const bRows = bSheet.getDataRange().getValues();
        const nextNum = bRows.length;
        const rowData = hasPin
          ? [nextNum, assign.bookingId, "PENDING_DELIVERY", "", amount, dateOfIssue, account, cardName, "PENDING", formatISTTimestamp(new Date())]
          : [nextNum, assign.bookingId, "PENDING_DELIVERY", amount, dateOfIssue, account, cardName, "PENDING", formatISTTimestamp(new Date())];
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
 * 6. REST API FOR SWIFT APP (GET / POST)
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const result = {
    unassigned: [],
    cooldowns: [],
    vouchers: []
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
    } else {
      const hasPin = BRANDS_WITH_PIN.has(name);
      for (let i = 1; i < rows.length; i++) {
        result.vouchers.push({
          brand: name,
          bookingId: String(rows[i][1]),
          code: String(rows[i][2]),
          pin: hasPin ? String(rows[i][3]) : null,
          amount: hasPin ? rows[i][4] : rows[i][3],
          dateOfIssue: hasPin ? rows[i][5] : rows[i][4],
          account: hasPin ? rows[i][6] : rows[i][5],
          card: hasPin ? rows[i][7] : rows[i][6],
          status: hasPin ? rows[i][8] : rows[i][7],
          timestampIST: hasPin ? rows[i][9] : rows[i][8]
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
    if (data.action === "ASSIGN_CARDS") {
      const res = assignCardToRedemptions(ss, data.assignments);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }
    if (data.action === "MARK_REDEEMED") {
      let bSheet = ss.getSheetByName(data.brand);
      if (bSheet) {
        const rows = bSheet.getDataRange().getValues();
        const hasPin = BRANDS_WITH_PIN.has(data.brand);
        const statusCol = hasPin ? 9 : 8;
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][1] === data.bookingId || rows[i][2] === data.code) {
            bSheet.getRange(i + 1, statusCol).setValue("USED");
            bSheet.getRange(i + 1, statusCol + 1).setValue(formatISTTimestamp(new Date()));
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
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
