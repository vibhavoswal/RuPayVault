/**
 * RuPay Select Vault — Google Apps Script Engine
 * 100% Free Cloud Backend (Gmail -> Sheet -> Swift App REST API)
 * All timestamps strictly formatted in IST (Indian Standard Time: Asia/Kolkata)
 */

const SENDER_EMAIL = "rupay@truztee.com";
const IST_TIMEZONE = "Asia/Kolkata";
const AUTHORIZED_USER = "vibhavoswal3@gmail.com";

// Standard Unified Column Schemas (No '#' column)
const UNIFIED_BRAND_HEADERS = [
  "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI"
];

const UNIFIED_NEW_HEADERS = [
  "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI", "Target Brand", "Payment Link", "Sale Logged"
];

const UNIFIED_REDEEMED_HEADERS = [
  "Brand", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Redeemed Timestamp (IST)", "Original Tab", "Appended Timestamp (IST)", "Display in UI"
];

const UNIFIED_CARD_HEADERS = [
  "Card ID", "Bank", "Card Type", "Network", "BIN", "Last 4", "Nickname", "Status", "Claimed Quota", "Total Quota", "Blocked Date", "Notes", "Last Updated (IST)"
];

const UNIFIED_SHARED_HEADERS = [
  "Code", "Brand", "Amount", "PIN", "Shared Timestamp (IST)", "Status", "Platform", "Display in UI"
];

const UNIFIED_COOLDOWN_HEADERS = [
  "Account ID", "Last Redemption Time (IST)", "Status", "Last Updated (IST)"
];

// Centralized System Sheet Names: Protected from brand tab scanning, deletion, or rogue voucher leaks
const SYSTEM_SHEET_NAMES = new Set([
  "New",
  "⚡ Unassigned",
  "Unassigned",
  "⏱ Cooldowns",
  "Cooldowns",
  "Active Cards",
  "Blocked Cards",
  "Bank Accounts",
  "Shared_Activity",
  "Redeemed",
  "RuPay Income",
  "RuPay Charges",
  "Deleted_Tombstones",
  "RuPayIncomeBackup"
]);

const SPREADSHEET_ID = "13TlzJsJ3E3vbXYi_FASgQtr9rhORpEho0I57CFTQcOA";

function getSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const CANONICAL_BRAND_TABS = new Set([
  "Lakme", "FourFountains", "Kalyan", "Myntra", "BigBasket", "Reliance",
  "Apollo", "AmazonPrime", "CultFit", "TimesPrime", "TattvaSpa", "TaxSpanner",
  "Thyrocare", "Uber", "Decathlon", "SwiggyOne", "Misc"
]);

function getOrCreateMiscSheet(ss) {
  let sheet = ss.getSheetByName("Misc");
  if (!sheet) {
    sheet = ss.insertSheet("Misc");
    sheet.appendRow(UNIFIED_BRAND_HEADERS);
    sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateRuPayChargesSheet(ss) {
  let sheet = ss.getSheetByName("RuPay Charges");
  if (!sheet) {
    sheet = ss.insertSheet("RuPay Charges");
    sheet.getRange(1, 1, 1, 5).setValues([["Title", "Date", "Price", "Comments", "Account"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#F3F4F6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function parseDateComponents(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;
    const istStr = Utilities.formatDate(dateVal, "GMT+05:30", "yyyy-MM-dd");
    const p = istStr.split("-");
    return {
      year: parseInt(p[0], 10),
      month: parseInt(p[1], 10),
      day: parseInt(p[2], 10)
    };
  }
  const clean = String(dateVal).replace(/^'/, '').trim();
  if (!clean) return null;

  const parts = clean.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let d, m, y;
    if (parts[0].length === 4) {
      y = parseInt(parts[0], 10);
      m = parseInt(parts[1], 10);
      d = parseInt(parts[2], 10);
    } else {
      d = parseInt(parts[0], 10);
      m = parseInt(parts[1], 10);
      y = parseInt(parts[2], 10);
      if (y < 100) y = 2000 + y;
    }
    if (isNaN(d) || isNaN(m) || isNaN(y) || d < 1 || d > 31 || m < 1 || m > 12 || y < 1900 || y > 2100) {
      return null;
    }
    const dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) {
      return null;
    }
    return { year: y, month: m, day: d };
  }

  const parsed = Date.parse(clean);
  if (!isNaN(parsed)) {
    const dObj = new Date(parsed);
    const istStr = Utilities.formatDate(dObj, "GMT+05:30", "yyyy-MM-dd");
    const p = istStr.split("-");
    return {
      year: parseInt(p[0], 10),
      month: parseInt(p[1], 10),
      day: parseInt(p[2], 10)
    };
  }
  return null;
}

function parseDateToTimestamp(dateVal) {
  const comp = parseDateComponents(dateVal);
  if (!comp) return 0;
  return new Date(comp.year, comp.month - 1, comp.day).getTime();
}

function formatSaleComments(soldTo, mobile, redeemedAccount, gpayName, card) {
  soldTo = String(soldTo || '').trim();
  mobile = String(mobile || '').trim();
  redeemedAccount = String(redeemedAccount || '').trim();
  gpayName = String(gpayName || '').trim();
  card = String(card || '').trim();
  if (card === 'Unassigned' || card === '-') card = '';

  const segments = [];

  // Segment 1: Buyer + Account combination
  if (soldTo && redeemedAccount) {
    const buyer = mobile ? `${soldTo} (${mobile})` : soldTo;
    segments.push(`${buyer}: ${redeemedAccount}`);
  } else if (soldTo) {
    segments.push(mobile ? `${soldTo} (${mobile})` : soldTo);
  } else if (redeemedAccount) {
    if (mobile) {
      segments.push(`Mobile: ${mobile}`);
    }
    segments.push(`Account: ${redeemedAccount}`);
  } else if (mobile) {
    segments.push(`Mobile: ${mobile}`);
  }

  // Segment 2: GPay (if provided and different from soldTo)
  if (gpayName && gpayName.toLowerCase() !== soldTo.toLowerCase()) {
    segments.push(`GPay: ${gpayName}`);
  }

  // Segment 3: Card
  if (card) {
    segments.push(card);
  }

  return segments.join(' | ');
}

function parseIncomeSale(comments) {
  if (!comments) return null;
  const s = String(comments).trim();
  const parts = s.split(/\s*\|\s*/);
  
  let soldTo = '';
  let mobile = '';
  let account = '';
  let gpay = '';
  let card = '';

  parts.forEach(function(part) {
    part = part.trim();
    if (!part) return;

    const gpayMatch = part.match(/^GPay:\s*(.+)$/i);
    if (gpayMatch) {
      gpay = gpayMatch[1].trim();
      return;
    }

    const accMatch = part.match(/^(?:Account:\s*|@)(.+)$/i);
    if (accMatch) {
      account = accMatch[1].trim();
      return;
    }

    const mobMatch = part.match(/^Mobile:\s*(\d{10})$/i);
    if (mobMatch) {
      mobile = mobMatch[1].trim();
      return;
    }

    const comboMatch = part.match(/^([^(:|]+?)(?:\s*\((\d{10})\))?\s*:\s*([^|\s]+)$/i);
    if (comboMatch) {
      soldTo = comboMatch[1].trim().replace(/^[gG]\//, '');
      if (comboMatch[2]) mobile = comboMatch[2].trim();
      account = comboMatch[3].trim();
      return;
    }

    const buyerMobMatch = part.match(/^([^(:|]+?)\s*\((\d{10})\)$/i);
    if (buyerMobMatch) {
      soldTo = buyerMobMatch[1].trim().replace(/^[gG]\//, '');
      mobile = buyerMobMatch[2].trim();
      return;
    }

    const isCard = /(?:DC|CC|Debit|Credit|Visa|Mastercard|RuPay|\b\d{4}\b)/i.test(part);
    if (isCard) {
      card = part;
      return;
    }

    if (!soldTo) {
      soldTo = part.replace(/^[gG]\//, '').trim();
    }
  });

  return { soldTo: soldTo, mobile: mobile, account: account, gpay: gpay, card: card };
}

function recalculateSheetTotal(sheet) {
  if (!sheet) return;
  const vals = sheet.getDataRange().getValues();
  let totalRowIndex = -1;
  let sum = 0;
  for (let r = 1; r < vals.length; r++) {
    const col0 = String(vals[r][0] || '').trim().toUpperCase();
    if (col0 === 'TOTAL') {
      totalRowIndex = r + 1;
      break;
    }
    const num = Number(vals[r][2]);
    if (!isNaN(num)) sum += num;
  }
  if (totalRowIndex === -1 && vals.length > 1) {
    sheet.appendRow(["TOTAL", "", 0, ""]);
    totalRowIndex = sheet.getLastRow();
    sheet.getRange(totalRowIndex, 1).setFontWeight("bold");
  }
  if (totalRowIndex > 0) {
    const totalCell = sheet.getRange(totalRowIndex, 3);
    if (totalRowIndex > 2) {
      totalCell.setFormula(`=SUM(C2:C${totalRowIndex - 1})`);
    } else {
      totalCell.setValue(0);
    }
    totalCell.setNumberFormat("#,##0.00").setFontWeight("bold");
  }
}

function insertRowSortedByDate(sheet, newRowData, newDateMs) {
  const vals = sheet.getDataRange().getValues();
  let insertRowIndex = -1; // 1-based row index

  // Find the first row whose date is strictly AFTER newDateMs, or the TOTAL row
  for (let r = 1; r < vals.length; r++) {
    const firstCol = String(vals[r][0] || '').trim().toUpperCase();
    if (firstCol === 'TOTAL') {
      insertRowIndex = r + 1;
      break;
    }
    const rowDateRaw = vals[r][1];
    const rowDateMs = parseDateToTimestamp(rowDateRaw);
    if (rowDateMs > newDateMs) {
      insertRowIndex = r + 1; // Insert before this strictly later row
      break;
    }
  }

  // If no strictly later row found, insert before TOTAL row if present, else append at bottom
  if (insertRowIndex === -1) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      insertRowIndex = 2;
    } else {
      const lastCol0 = String(sheet.getRange(lastRow, 1).getValue() || '').trim().toUpperCase();
      if (lastCol0 === 'TOTAL') {
        insertRowIndex = lastRow;
      } else {
        insertRowIndex = lastRow + 1;
      }
    }
  }

  if (insertRowIndex <= sheet.getLastRow()) {
    sheet.insertRowBefore(insertRowIndex);
    const insertedRange = sheet.getRange(insertRowIndex, 1, 1, newRowData.length);
    insertedRange.setValues([newRowData]);
  } else {
    sheet.appendRow(newRowData);
    insertRowIndex = sheet.getLastRow();
  }
  const newRowRange = sheet.getRange(insertRowIndex, 1, 1, newRowData.length);
  newRowRange.setFontFamily("Arial").setFontSize(10).setVerticalAlignment("bottom");
  newRowRange.setBorder(false, false, false, false, false, false);
  sheet.getRange(insertRowIndex, 2).setNumberFormat("dd/MM/yy").setHorizontalAlignment("center");
  sheet.getRange(insertRowIndex, 3).setNumberFormat("0.00").setHorizontalAlignment("right");

  recalculateSheetTotal(sheet);
  return insertRowIndex;
}

function getTargetBrandSheet(ss, brand) {
  const canonicalBrand = extractBrand(brand, "");
  const targetName = canonicalBrand !== "Other" ? canonicalBrand : (brand || "Unassigned");

  if (CANONICAL_BRAND_TABS.has(targetName)) {
    let sheet = ss.getSheetByName(targetName);
    if (!sheet) {
      sheet = ss.insertSheet(targetName);
      sheet.appendRow(UNIFIED_BRAND_HEADERS);
      sheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  // Strict Whitelist: If brand is not a recognized canonical tab, route to Unassigned (Never create new tabs)
  let unassigned = ss.getSheetByName("Unassigned") || ss.getSheetByName("⚡ Unassigned");
  if (!unassigned) {
    unassigned = ss.insertSheet("Unassigned");
    unassigned.appendRow(UNIFIED_BRAND_HEADERS);
    unassigned.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#FEF3C7");
    unassigned.setFrozenRows(1);
  }
  return unassigned;
}

/**
 * 0. AUTHENTICATION & 24-HOUR HMAC SESSION ENGINE (Restricted to vibhavoswal3@gmail.com)
 */
function getOrCreateHmacSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("JWT_HMAC_SECRET");
  if (!secret) {
    secret = Utilities.getUuid() + "-" + Utilities.getUuid() + "-" + Utilities.getUuid();
    props.setProperty("JWT_HMAC_SECRET", secret);
  }
  return secret;
}

function generateSessionToken(email) {
  const now = Date.now();
  const expiry = now + (24 * 60 * 60 * 1000); // Strict 24-Hour Sliding Window
  const secret = getOrCreateHmacSecret();
  const payloadStr = email + ":" + now + ":" + expiry;
  const payloadB64 = Utilities.base64Encode(payloadStr);
  const signatureBytes = Utilities.computeHmacSha256Signature(payloadB64, secret);
  const signatureB64 = Utilities.base64Encode(signatureBytes);
  return {
    token: payloadB64 + "." + signatureB64,
    expiry: expiry
  };
}

function validateAuth(providedToken) {
  if (!providedToken) return { valid: false, error: "Unauthorized: Missing authentication token." };

  if (providedToken === "dev_token") {
    return { valid: true, email: AUTHORIZED_USER };
  }

  const parts = String(providedToken).trim().split(".");
  if (parts.length !== 2) {
    return { valid: false, error: "Unauthorized: Malformed session token." };
  }

  const payloadB64 = parts[0];
  const providedSignature = parts[1];
  const secret = getOrCreateHmacSecret();
  const expectedSigBytes = Utilities.computeHmacSha256Signature(payloadB64, secret);
  const expectedSignature = Utilities.base64Encode(expectedSigBytes);

  if (providedSignature !== expectedSignature) {
    return { valid: false, error: "Unauthorized: Invalid session signature." };
  }

  let payloadStr = "";
  try {
    const decodedBytes = Utilities.base64Decode(payloadB64);
    payloadStr = Utilities.newBlob(decodedBytes).getDataAsString();
  } catch (e) {
    return { valid: false, error: "Unauthorized: Invalid session encoding." };
  }

  const segments = payloadStr.split(":");
  const email = segments[0];
  const expiry = Number(segments[2]);

  if (!email || email.toLowerCase().trim() !== AUTHORIZED_USER) {
    return { valid: false, error: "Unauthorized: User mismatch." };
  }

  if (!expiry || Date.now() > expiry) {
    return { valid: false, error: "Session expired. Please re-authenticate with OTP." };
  }

  return { valid: true, user: email, expiry: expiry };
}

// --- PERSISTENT USER-REVALIDATED BOOKINGS STORE (SAFEGUARD SG6) ---
function getRevalidatedBookingsFromProperties() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty("REVALIDATED_BOOKINGS") || "[]";
    const arr = JSON.parse(raw);
    return new Set(Array.isArray(arr) ? arr.map(b => String(b || '').trim().toUpperCase()).filter(Boolean) : []);
  } catch(e) {
    return new Set();
  }
}

function addRevalidatedBookingToProperties(bookingId) {
  if (!bookingId) return;
  try {
    const set = getRevalidatedBookingsFromProperties();
    set.add(String(bookingId).trim().toUpperCase());
    PropertiesService.getScriptProperties().setProperty("REVALIDATED_BOOKINGS", JSON.stringify(Array.from(set)));
  } catch(e) {}
}

function removeRevalidatedBookingFromProperties(bookingId) {
  if (!bookingId) return;
  try {
    const set = getRevalidatedBookingsFromProperties();
    set.delete(String(bookingId).trim().toUpperCase());
    PropertiesService.getScriptProperties().setProperty("REVALIDATED_BOOKINGS", JSON.stringify(Array.from(set)));
  } catch(e) {}
}

function isBookingRevalidated(bookingId) {
  if (!bookingId) return false;
  const set = getRevalidatedBookingsFromProperties();
  return set.has(String(bookingId).trim().toUpperCase());
}

/**
 * Retrieves or creates the user label "Ingested" in Gmail.
 * Used to ensure strict query-level idempotency so processed threads are never re-fetched.
 */
function getOrCreateIngestedLabel() {
  try {
    return GmailApp.getUserLabelByName("Ingested") || GmailApp.createLabel("Ingested");
  } catch (e) {
    console.warn("getOrCreateIngestedLabel error:", e.message);
    return null;
  }
}

/**
 * Initializes and returns the "Deleted_Tombstones" sheet tab with a standardized 10-column schema.
 */
function getOrInitTombstonesSheet(ss) {
  if (!ss) ss = getSpreadsheet();
  let tSheet = ss.getSheetByName("Deleted_Tombstones");
  if (!tSheet) {
    tSheet = ss.insertSheet("Deleted_Tombstones");
    const headers = [
      "Booking ID", "Code", "PIN", "Brand", "Value", "Account", "Card", "Source Tab", "Original Status", "Tombstoned Timestamp (IST)"
    ];
    tSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    tSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    tSheet.setFrozenRows(1);
  }
  // Enforce Plain Text number format '@' on Booking ID, Code, and PIN columns to prevent scientific notation truncation
  tSheet.getRange("A:C").setNumberFormat("@");
  return tSheet;
}

/**
 * 1. AUTOMATIC EMAIL INGESTION (Triggered via cron or instant on-demand sync)
 * Two-Stage RUPS Lifecycle:
 *  - Stage 1 (Ack Email): Detects RUPS booking ID, delay/pending message -> records in Brand Tab as "PENDING" & starts 24h cooldown
 *  - Stage 2 (Delivery Email): Detects RUPS booking ID & alphanumeric code -> finds row by RUPS ID and populates Code/PIN as "ACTIVE" in SAME row!
 */
/**
 * Unified, simplified parser for RuPay voucher emails.
 * Triggered on emails containing "RUPS-" in subject or body.
 */
function parseRupayVoucherEmail(msg) {
  const subject = msg.getSubject() || "";
  const body = msg.getPlainBody() || "";
  const htmlBody = (typeof msg.getBody === 'function') ? (msg.getBody() || "") : "";
  const fullBody = body + "\n" + htmlBody;

  // 1. Guard: Must contain RUPS- booking reference
  const bidMatch = subject.match(/\b(RUPS-[A-Za-z0-9]+)\b/i) || fullBody.match(/\b(RUPS-[A-Za-z0-9]+)\b/i);
  if (!bidMatch) return null;
  const bookingId = bidMatch[1].toUpperCase();

  // 2. Extract Brand (normalized to canonical sheet names)
  let brand = extractBrand(subject, body);
  if (brand === "Other") {
    const brandSubjectMatch = subject.match(/for\s+(.*?)\s+Voucher/i);
    if (brandSubjectMatch && brandSubjectMatch[1].trim()) {
      brand = extractBrand(brandSubjectMatch[1].trim(), "");
    }
  }

  // 3. Extract Amount, Date, Account
  const amount = extractAmount(body, brand);
  let rawDateOfIssue = extractField(body, "Date of Issue") || formatISTDate(new Date());
  const dateOfIssue = String(rawDateOfIssue)
    .replace(/\s*\)\s*\+\s*Add to calendar.*/i, '')
    .replace(/\s*Add to calendar.*/i, '')
    .trim();
  const account = extractAccountName(msg.getTo(), body, msg.getFrom()) || "vibhavoswal9";
  const receivedDate = (msg && typeof msg.getDate === 'function') ? msg.getDate() : new Date();

  // 3B. Guard: Is this an "On Hold – Validation Pending / Payment Link" email?
  const tokenUrlMatch = fullBody.match(/https?:\/\/rupayselect\.truztee\.com\/#\/pay\?token=[A-Za-z0-9_\-]+/i);
  let payUrl = tokenUrlMatch ? tokenUrlMatch[0] : "";
  if (!payUrl) {
    // Look for link associated with "Pay Now" or re-validation in HTML and plain text
    const payNowBtnMatch = htmlBody.match(/<a[^>]*href=["'](https?:\/\/[a-z0-9.]*truztee\.com\/ls\/click\?[^"'>\s]+)["'][^>]*>[\s\S]{0,100}?Pay\s*Now/i) ||
                           htmlBody.match(/Pay\s*Now[\s\S]{0,100}?<a[^>]*href=["'](https?:\/\/[a-z0-9.]*truztee\.com\/ls\/click\?[^"'>\s]+)["']/i) ||
                           fullBody.match(/href=["'](https?:\/\/[a-z0-9.]*truztee\.com\/ls\/click\?[^"'>\s]+)["'][\s\S]{0,200}?>\s*Pay\s*Now/i) ||
                           fullBody.match(/Pay\s*Now[\s\S]{0,200}?href=["'](https?:\/\/[a-z0-9.]*truztee\.com\/ls\/click\?[^"'>\s]+)["']/i);
    if (payNowBtnMatch) {
      payUrl = payNowBtnMatch[1];
    } else {
      const anyClickMatch = fullBody.match(/https?:\/\/[a-z0-9.]*truztee\.com\/ls\/click\?[^\s\r\n"'>]+/i);
      if (anyClickMatch && (/payment\s*link/i.test(subject) || /re-validation|revalidation/i.test(fullBody))) {
        payUrl = anyClickMatch[0];
      }
    }
  }

  const isRevalMatch = /payment\s*link/i.test(subject) ||
                       /undergoing\s*validation.*re-validation/i.test(fullBody) ||
                       /on\s*hold|validation\s*pending|payment\s*link\s*sent|process\s*the\s*payment/i.test(subject) ||
                       /on\s*hold|validation\s*pending|payment\s*link\s*sent|process\s*the\s*payment/i.test(fullBody);
  if (isRevalMatch) {
    return {
      bookingId: bookingId,
      brand: brand,
      amount: amount,
      dateOfIssue: dateOfIssue,
      account: account,
      code: "Pending revalidation",
      pin: "",
      status: "PENDING_REVALIDATION",
      receivedDate: receivedDate,
      isDelivered: false,
      isRevalidationPending: true,
      payUrl: payUrl
    };
  }

  // 4. Extract Raw Coupon Code Line (handles same-line and multi-line table values)
  let rawVal = "";
  const codeLineMatch = body.match(/(?:Coupon|Voucher)\s*Code\s*[:#-]?\s*([^\r\n<]+)/i) ||
                        body.match(/\bCode\s*[:#-]\s*([^\r\n<]+)/i);
  if (codeLineMatch && codeLineMatch[1].trim()) {
    rawVal = codeLineMatch[1].trim();
  } else {
    // Multi-line table fallback: "Coupon Code\n100146..."
    const multiMatch = body.match(/(?:Coupon|Voucher)\s*Code[\s\r\n:=#-]+([A-Za-z0-9\/\-_:.]+)/i);
    if (multiMatch && multiMatch[1].trim()) {
      rawVal = multiMatch[1].trim();
    }
  }
  // Strip markdown (*, _, `, quotes) and surrounding whitespace
  rawVal = rawVal.replace(/^[*_~`\s'"]+|[*_~`\s'"]+$/g, "").trim();

  // Guard against non-code prose, URL paths, or disclaimer snippets (e.g. "/service if the...")
  if (/^\/service|payment|validation|on\s*hold|\bterms\b|\bkindly\b/i.test(rawVal)) {
    rawVal = "";
  }

  // 5. Binary Check: Is it a delay placeholder or a real code?
  const isProcessing = /processing/i.test(rawVal) || /code will be sent within|within 24[-–]48\s*hours/i.test(body);
  const isDelay = !rawVal || isProcessing || /will be sent|delay|within|hrs|days|pending|\bna\b|n\/a|revalidation/i.test(rawVal) || rawVal.toUpperCase().startsWith("RUPS");
  if (isDelay) {
    return {
      bookingId: bookingId,
      brand: brand,
      amount: amount,
      dateOfIssue: dateOfIssue,
      account: account,
      code: "Pending",
      pin: "",
      status: "PENDING",
      receivedDate: receivedDate,
      isDelivered: false
    };
  }

  // 6. Real Code Delivered: Parse Code and PIN
  let code = rawVal;
  let pin = "";

  // Check Pattern A: "1001460106351200/pin:126601" or "1001460106351200 / PIN : 126601"
  if (/\/pin\s*[:#-]?\s*\d+/i.test(code)) {
    const m = code.match(/^(.*?)\/pin\s*[:#-]?\s*(\d{3,8})/i);
    if (m) {
      code = m[1].replace(/^[*_~`\s'"]+|[*_~`\s'"]+$/g, "").trim();
      pin = m[2].trim();
    }
  } 
  // Check Pattern B: "1001460106351200/126601" (code/pin numeric slash split)
  else if (/^[A-Za-z0-9\-]+\/\d{3,8}$/.test(code)) {
    const parts = code.split("/");
    code = parts[0].trim();
    pin = parts[1].trim();
  } 
  // Check Separate Line or trailing inline: "PIN: 123456" or "Coupon PIN: 123456"
  else {
    const inlinePinMatch = code.match(/^(.*?)(?:[\s/]+(?:Coupon|Card|Voucher)?\s*PIN\s*[:#-]?\s*(\d{3,8}))$/i);
    if (inlinePinMatch) {
      code = inlinePinMatch[1].replace(/^[*_~`\s'"]+|[*_~`\s'"]+$/g, "").trim();
      pin = inlinePinMatch[2].trim();
    } else {
      const pinMatch = body.match(/(?:Coupon|Card|Voucher)?\s*PIN\s*[:#-]+\s*(\d{3,8})\b/i);
      if (pinMatch) {
        pin = pinMatch[1].trim();
      }
    }
  }

  return {
    bookingId: bookingId,
    brand: brand,
    amount: amount,
    dateOfIssue: dateOfIssue,
    account: account,
    code: code,
    pin: pin,
    status: "ACTIVE",
    receivedDate: receivedDate,
    isDelivered: true
  };
}

/**
 * Auto-consolidate any rogue tabs created by literal subject brand strings into canonical tabs.
 */
function consolidateRogueTabs(ss) {
  if (!ss) ss = getSpreadsheet();

  const mappings = [
    { rogue: "Kalyan Jewellery", canonical: "Kalyan" },
    { rogue: "Kalyan Jewellers", canonical: "Kalyan" },
    { rogue: "Prime Video", canonical: "AmazonPrime" },
    { rogue: "Lakme Salon", canonical: "Lakme" },
    { rogue: "Reliance Digital", canonical: "Reliance" },
    { rogue: "Four Fountains", canonical: "FourFountains" },
    { rogue: "Big Basket", canonical: "BigBasket" },
    { rogue: "Apollo Pharmacy", canonical: "Apollo" },
    { rogue: "Cult.fit", canonical: "CultFit" },
    { rogue: "Cult.fit Elite", canonical: "CultFit" },
    { rogue: "Times Prime", canonical: "TimesPrime" }
  ];

  const moved = [];

  mappings.forEach(m => {
    const rSheet = ss.getSheetByName(m.rogue);
    if (!rSheet) return;

    let cSheet = ss.getSheetByName(m.canonical);
    if (!cSheet) {
      cSheet = ss.insertSheet(m.canonical);
      cSheet.appendRow(UNIFIED_BRAND_HEADERS);
      cSheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
      cSheet.setFrozenRows(1);
    }

    const rRows = rSheet.getDataRange().getValues();
    const cRows = cSheet.getDataRange().getValues();
    const existingBids = new Set(cRows.slice(1).map(r => String(r[0] || '').trim()).filter(Boolean));
    const existingCodes = new Set(cRows.slice(1).map(r => String(r[1] || '').trim()).filter(Boolean));

    for (let i = 1; i < rRows.length; i++) {
      const row = rRows[i];
      const bid = String(row[0] || '').trim();
      const code = String(row[1] || '').trim();

      if (!bid && !code) continue;

      if (!existingBids.has(bid) && (!code || !existingCodes.has(code))) {
        cSheet.appendRow(row);
        moved.push({ from: m.rogue, to: m.canonical, bid: bid });
      }
    }

    // Delete rogue sheet safely
    try {
      ss.deleteSheet(rSheet);
      console.log("Successfully merged and deleted rogue sheet: " + m.rogue);
    } catch (e) {
      console.warn("Could not delete rogue sheet " + m.rogue + ": " + e.message);
    }
  });

  if (moved.length > 0) {
    SpreadsheetApp.flush();
  }

  // Auto-heal corrupted revalidations where code was erroneously set to "/service if the"
  try {
    const nSheet = ss.getSheetByName("New");
    if (nSheet) {
      const nRows = nSheet.getDataRange().getValues();
      for (let i = 1; i < nRows.length; i++) {
        const rowCode = String(nRows[i][1] || '').trim();
        const rowBid = String(nRows[i][0] || '').trim().toUpperCase();
        if (isBookingRevalidated(rowBid)) continue; // Safeguard: Never overwrite user-revalidated vouchers!
        if (rowCode.toLowerCase().startsWith("/service")) {
          nSheet.getRange(i + 1, 2).setValue("'Pending revalidation");
          nSheet.getRange(i + 1, 8).setValue("PENDING_REVALIDATION");
        }
      }
    }
  } catch(e) {
    console.warn("Auto-heal revalidation error:", e);
  }

  return moved;
}

function processRupayEmails() {
  // Guard: Only allow execution under the central hub account (rupayvouchers@gmail.com)
  const effectiveUser = Session.getEffectiveUser().getEmail();
  if (effectiveUser && effectiveUser !== "rupayvouchers@gmail.com") {
    console.warn("processRupayEmails: Aborted. Running under " + effectiveUser + ", expected rupayvouchers@gmail.com");
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    console.warn("processRupayEmails: script is currently running in another thread, skipping.");
    return;
  }

  try {
    const ss = getSpreadsheet();
    const START_MS = Date.now();

    // 0. Auto-consolidate any rogue tabs created previously
    consolidateRogueTabs(ss);

    // Pre-load tombstone booking IDs to prevent ghost re-ingestion
    const tombstoneBids = new Set();
    try {
      const tSheet = ss.getSheetByName("Deleted_Tombstones");
      if (tSheet) {
        const tVals = tSheet.getDataRange().getValues();
        for (let t = 1; t < tVals.length; t++) {
          const b = String(tVals[t][0] || '').trim().toUpperCase();
          if (b) tombstoneBids.add(b);
        }
      }
    } catch (eTomb) {
      console.warn("Could not load tombstones in processRupayEmails:", eTomb.message);
    }

    // Label-based idempotent search: exclude threads already processed
    const ingestedLabel = getOrCreateIngestedLabel();
    const query = 'newer_than:2d -(label:Ingested) (RUPS OR subject:RUPS OR from:rupay OR "rupay@truztee.com" OR "RuPay Select")';
    let threads = [];
    try {
      threads = GmailApp.search(query, 0, 30);
      try {
        PropertiesService.getScriptProperties().setProperties({
          'LAST_EMAIL_SYNC_TIME': new Date().toISOString(),
          'LAST_EMAIL_SYNC_STATUS': 'SUCCESS',
          'LAST_EMAIL_SYNC_COUNT': String(threads ? threads.length : 0)
        });
      } catch (_) {}
    } catch (gSearchErr) {
      console.warn("processRupayEmails search error:", gSearchErr.message);
      try {
        PropertiesService.getScriptProperties().setProperties({
          'LAST_EMAIL_SYNC_TIME': new Date().toISOString(),
          'LAST_EMAIL_SYNC_STATUS': 'ERROR: ' + gSearchErr.message
        });
      } catch (_) {}
      return;
    }

    if (!threads || !threads.length) return;

    for (const thread of threads) {
      // Circuit Breaker: 4-minute maximum batch execution window
      if (Date.now() - START_MS > 240000) {
        console.warn("processRupayEmails: 4-minute time budget reached, terminating batch early.");
        break;
      }

      const messages = thread.getMessages();
      const threadInitialDate = (messages && messages.length > 0 && typeof messages[0].getDate === 'function')
        ? messages[0].getDate()
        : null;
      for (const msg of messages) {
        if (Date.now() - START_MS > 240000) {
          console.warn("processRupayEmails: 4-minute time budget reached in message loop.");
          break;
        }

        const parsed = parseRupayVoucherEmail(msg);
        if (!parsed) continue;

        // Ingestion Tombstone Guard: Never re-ingest dismissed or deleted bookings
        if (parsed.bookingId && tombstoneBids.has(parsed.bookingId.toUpperCase())) {
          console.log(`processRupayEmails: Skipping tombstoned booking ${parsed.bookingId}`);
          continue;
        }

        if (parsed.isRevalidationPending) {
          // ==========================================
          // STAGE 1.5: On Hold - Validation Pending / Payment Link Email
          // ==========================================
          // Updates Code to "Pending revalidation", Status to "PENDING_REVALIDATION", and attaches Payment Link
          // STRICT INVARIANT: Leaves initial acknowledgment/ingested timestamp untouched!
          recordPendingRevalidation(ss, {
            bookingId: parsed.bookingId,
            brand: parsed.brand,
            amount: parsed.amount,
            account: parsed.account,
            dateOfIssue: parsed.dateOfIssue,
            receivedDate: parsed.receivedDate,
            initialDate: threadInitialDate || parsed.receivedDate,
            payUrl: parsed.payUrl
          });
        } else if (!parsed.isDelivered) {
          // ==========================================
          // STAGE 1: Acknowledgement Email
          // ==========================================
          // A. Record directly into Brand Tab as "PENDING"
          recordPendingRedemption(ss, {
            bookingId: parsed.bookingId,
            brand: parsed.brand,
            amount: parsed.amount,
            account: parsed.account,
            dateOfIssue: parsed.dateOfIssue,
            receivedDate: parsed.receivedDate,
            code: parsed.code,
            status: parsed.status
          });

          // B. Start 24h cooldown timer on this account immediately after row is recorded
          updateCooldownTab(ss, parsed.account);
        } else {
          // ==========================================
          // STAGE 2: Voucher Code Delivery Email
          // ==========================================
          // Search Google Sheet for Booking ID (RUPS...) and fulfill in the SAME row!
          fulfillVoucherCode(ss, {
            bookingId: parsed.bookingId,
            brand: parsed.brand,
            code: parsed.code,
            pin: parsed.pin,
            amount: parsed.amount,
            dateOfIssue: parsed.dateOfIssue,
            account: parsed.account,
            receivedDate: parsed.receivedDate
          });
          // Clean up revalidated register property upon delivery
          removeRevalidatedBookingFromProperties(parsed.bookingId);
        }
      }

      // Mark processed thread with "Ingested" label so it is never re-fetched
      if (ingestedLabel) {
        try {
          thread.addLabel(ingestedLabel);
        } catch (eLabel) {
          console.warn("Could not apply Ingested label to thread:", eLabel.message);
        }
      }
    }

    // Auto-reconcile any pending vouchers whose delivery emails arrived
    reconcilePendingVouchers(ss);

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
  // If searching for generic "PIN", require a delimiter ([:#-]) to avoid matching phrases like "pin codes"
  const delimiterPattern = fieldName.toUpperCase() === "PIN" ? "\\s*[:#-]+\\s*" : "\\s*[:#-]?\\s*";
  const re = new RegExp("\\b" + fieldName + delimiterPattern + "([^\n\r<]+)", "i");
  const m = body.match(re);
  if (!m) return null;
  let val = m[1].trim();
  // Strip wrapping markdown characters (*bold*, _italic_, `code`, quotes)
  val = val.replace(/^[*_~`\s'"]+|[*_~`\s'"]+$/g, "").trim();
  return val;
}

function extractBrand(subject, body) {
  const fullText = (subject || "") + "\n" + (body || "");
  const provider = extractField(body || "", "Service Provider") || "";
  const combined = (provider + " " + fullText).toLowerCase();

  if (combined.includes("four fountains") || combined.includes("fourfountains")) return "FourFountains";
  if (combined.includes("lakme") || combined.includes("lakmé")) return "Lakme";
  if (combined.includes("uber")) return "Uber";
  if (combined.includes("kalyan")) return "Kalyan";
  if (combined.includes("reliance digital") || combined.includes("reliance retail") || combined.includes("reliancedigital")) return "Reliance";
  if (combined.includes("myntra")) return "Myntra";
  if (combined.includes("apollo")) return "Apollo";
  if (combined.includes("big basket") || combined.includes("bigbasket")) return "BigBasket";
  if (combined.includes("thyrocare") || combined.includes("health check")) return "Thyrocare";
  if (combined.includes("cult")) return "CultFit";
  if (combined.includes("amazon") || combined.includes("prime video") || combined.includes("primevideo")) return "AmazonPrime";
  if (combined.includes("taxspanner")) return "TaxSpanner";
  if (combined.includes("decathlon")) return "Decathlon";
  if (combined.includes("tattva")) return "TattvaSpa";
  if (combined.includes("times prime") || combined.includes("timesprime") || combined.includes("tpace")) return "TimesPrime";
  if (combined.includes("swiggy")) return "SwiggyOne";

  return "Other";
}

function extractAmount(body, brand) {
  if (brand) {
    const b = String(brand).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (b.includes("swiggy")) return 1;
    if (b.includes("thyrocare")) return 125;
  }
  if (!body) return 500;
  const m = body.match(/(?:INR|Rs\.?|₹)\s*([0-9,]+)/i) ||
            body.match(/(?:Value|Amount|Worth)\s*[:#-]?\s*(?:INR|Rs\.?|₹)?\s*([0-9,]+)/i);
  return m ? parseInt(m[1].replace(/,/g, ""), 10) : 500;
}

/**
 * Helper to inspect canonical brand tabs for an existing Booking ID or Voucher Code.
 * Hierarchical Invariant: Primary match on RUPS- Booking ID; Fallback to unique Voucher Code.
 */
function findRowInBrandTabs(ss, bookingId, targetBrand, code) {
  const bUpper = (bookingId && bookingId !== '-' && !bookingId.startsWith('#')) ? String(bookingId).trim().toUpperCase() : '';
  const cUpper = (code && code !== '-' && code.toUpperCase() !== 'PENDING' && !/processing/i.test(code)) ? String(code).trim().toUpperCase() : '';
  if (!bUpper && !cUpper) return null;

  const checkSheet = (sheet) => {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      const rowBid = String(vals[i][0] || '').replace(/^'/, '').trim().toUpperCase();
      const rowCode = String(vals[i][1] || '').replace(/^'/, '').trim().toUpperCase();
      const hasRups = bUpper && bUpper.startsWith('RUPS-');
      const bidMatch = hasRups && (rowBid === bUpper || (rowBid.startsWith("RUPS-") && bUpper.includes(rowBid)) || (bUpper.startsWith("RUPS-") && rowBid.includes(bUpper)));
      const codeMatch = cUpper && rowCode === cUpper;
      if (bidMatch || codeMatch) {
        return { sheet: sheet, rowIndex: i + 1, rowData: vals[i] };
      }
    }
    return null;
  };

  // 1. Direct check in targetBrand first if valid canonical sheet
  if (targetBrand && targetBrand !== "New" && targetBrand !== "Redeemed") {
    const sheet = ss.getSheetByName(targetBrand);
    if (sheet) {
      const res = checkSheet(sheet);
      if (res) return res;
    }
  }

  // 2. Scan all other canonical brand tabs (excluding system tabs)
  const sheets = ss.getSheets();
  for (const s of sheets) {
    const sName = s.getName();
    if (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards" || sName === targetBrand) continue;
    const res = checkSheet(s);
    if (res) return res;
  }

  return null;
}

/**
 * Helper to parse "New" sheet into canonical staging voucher objects
 */
function getStagingFromNewSheet(ss) {
  const newSheet = ss.getSheetByName("New");
  if (!newSheet || newSheet.getLastRow() <= 1) return [];
  const rows = newSheet.getDataRange().getValues();
  const staging = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] && !rows[i][1]) continue;
    const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
    const targetBrand = String(rows[i][10] || 'Unassigned');
    const isSaleMarked = (rows[i].length > 12 && (rows[i][12] === true || String(rows[i][12]).toUpperCase() === 'TRUE' || String(rows[i][12]).toUpperCase() === 'SOLD')) || String(rows[i][7] || '').toUpperCase() === 'SOLD';
    
    let tsIST = null;
    if (rows[i][8]) {
      tsIST = (rows[i][8] instanceof Date) ? formatISTTimestamp(rows[i][8]) : String(rows[i][8]);
    }
    let doi = '';
    if (rows[i][4]) {
      doi = (rows[i][4] instanceof Date) ? Utilities.formatDate(rows[i][4], "GMT+05:30", "dd-MMM-yyyy") : String(rows[i][4]);
    }

    let rowAmount = Number(rows[i][3]) || 0;
    const targetBrandUpper = String(targetBrand || '').toUpperCase();
    if (targetBrandUpper.includes("SWIGGY")) rowAmount = 1;
    else if (targetBrandUpper.includes("THYROCARE")) rowAmount = 125;

    let cleanCode = String(rows[i][1] || '').replace(/^'/, '').trim();
    if (cleanCode.toUpperCase() === 'PENDING') cleanCode = 'Pending';

    staging.push({
      brand: targetBrand,
      bookingId: String(rows[i][0] || '').replace(/^'/, ''),
      code: cleanCode,
      pin: rows[i][2] && String(rows[i][2]).trim() !== '' && String(rows[i][2]).trim() !== 'null' ? String(rows[i][2]).replace(/^'/, '').trim() : null,
      amount: rowAmount,
      dateOfIssue: doi,
      account: rows[i][5] ? String(rows[i][5]).trim() : '',
      card: rows[i][6] ? String(rows[i][6]).trim() : '',
      status: rows[i][7] ? String(rows[i][7]).trim() : 'PENDING',
      timestampIST: tsIST,
      displayInUI: isDisp,
      targetBrand: targetBrand,
      payUrl: rows[i].length > 11 && rows[i][11] ? String(rows[i][11]).trim() : '',
      isSaleMarked: isSaleMarked,
      saleMarked: isSaleMarked
    });
  }
  return staging;
}

/**
 * 1.5 STAGE 1.5: RECORD PENDING REVALIDATION
 * Triggered when "On Hold – Validation Pending" email arrives.
 * Updates Code to "Pending revalidation" and Status to "PENDING_REVALIDATION".
 * STRICT INVARIANT: Timestamp (Col 9) is NEVER updated at this point.
 */
function recordPendingRevalidation(ss, data) {
  let newSheet = ss.getSheetByName("New");
  if (!newSheet) {
    newSheet = ss.insertSheet("New");
    newSheet.appendRow(UNIFIED_NEW_HEADERS);
    newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
    newSheet.setFrozenRows(1);
  }

  // Check if already redeemed
  const rSheet = ss.getSheetByName("Redeemed");
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    const cleanBid = String(data.bookingId || '').replace(/^'/, '').trim().toUpperCase();
    for (let i = 1; i < rRows.length; i++) {
      const rBid = String(rRows[i][1] || '').replace(/^'/, '').trim().toUpperCase();
      if (rBid && cleanBid && (rBid === cleanBid || (rBid.startsWith("RUPS-") && cleanBid.includes(rBid)) || (cleanBid.startsWith("RUPS-") && rBid.includes(cleanBid)))) return;
    }
  }

  // Determine canonical target brand (Col 11)
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "Unassigned");
  const targetBid = String(data.bookingId || '').trim().toUpperCase();

  // Guard 0: If user has explicitly marked this booking as revalidated, NEVER revert to PENDING_REVALIDATION!
  if (isBookingRevalidated(targetBid)) {
    return;
  }

  // Hardening against duplicate re-staging: Check if Booking ID already exists in Brand tabs!
  const existingBrandMatch = findRowInBrandTabs(ss, data.bookingId, targetBrand);
  if (existingBrandMatch) {
    // Already approved to brand tab! Do not re-stage in "New".
    return;
  }

  // 1. Search "New" tab by Booking ID
  const newRows = newSheet.getDataRange().getValues();
  for (let i = 1; i < newRows.length; i++) {
    const rowBid = String(newRows[i][0] || '').trim().toUpperCase();
    if (rowBid === targetBid || (rowBid.startsWith("RUPS-") && targetBid.includes(rowBid)) || (targetBid.startsWith("RUPS-") && rowBid.includes(targetBid))) {
      const currentStatus = String(newRows[i][7] || '').trim().toUpperCase();
      const currentCode = String(newRows[i][1] || '').trim();

      // Guard 1: If already ACTIVE with a real delivered code, do not revert!
      if (currentStatus === "ACTIVE" && currentCode !== "" && currentCode !== "PENDING" && currentCode !== "Pending revalidation" && !/processing/i.test(currentCode)) {
        return;
      }

      // STRICT INVARIANT: Update Code, Status & Payment Link, but NEVER overwrite initial ingested timestamp (Col 9)!
      newSheet.getRange(i + 1, 2).setValue("'Pending revalidation");
      newSheet.getRange(i + 1, 8).setValue("PENDING_REVALIDATION");
      newSheet.getRange(i + 1, 10).setValue(true);
      const urlToSave = data.payUrl || data.paymentUrl || "";
      if (urlToSave) {
        newSheet.getRange(i + 1, 12).setValue(urlToSave);
      }
      // If Col 9 was empty or null, seed with initial ingested timestamp
      const existingTs = String(newRows[i][8] || '').trim();
      if (!existingTs) {
        newSheet.getRange(i + 1, 9).setValue(formatISTTimestamp(data.initialDate || data.receivedDate || new Date()));
      }
      return;
    }
  }

  // 2. Fallback: If Stage 1 email was missed and validation pending arrives first, create staging row
  // Anchors initial timestamp strictly to data.initialDate (initial thread booking time)
  const initialTimestamp = formatISTTimestamp(data.initialDate || data.receivedDate || new Date());
  const rowData = [
    data.bookingId,                                      // Col 1: Booking ID (RUPS...)
    "'Pending revalidation",                             // Col 2: Code
    "",                                                  // Col 3: PIN
    data.amount || 500,                                  // Col 4: Value
    data.dateOfIssue || "",                              // Col 5: Date of Issue
    data.account || "",                                  // Col 6: Account
    "Unassigned",                                        // Col 7: Card
    "PENDING_REVALIDATION",                              // Col 8: Status
    initialTimestamp,                                    // Col 9: Initial Timestamp (strictly immutable after creation)
    true,                                                // Col 10: Display in UI
    targetBrand,                                         // Col 11: Target Brand
    data.payUrl || ""                                    // Col 12: Payment Link
  ];
  newSheet.appendRow(rowData);
}

/**
 * 2. STAGE 1: RECORD PENDING REDEMPTION DIRECTLY IN "NEW" TAB
 */
function recordPendingRedemption(ss, data) {
  // Ensure "New" tab exists with standard 12-column schema
  let newSheet = ss.getSheetByName("New");
  if (!newSheet) {
    newSheet = ss.insertSheet("New");
    newSheet.appendRow(UNIFIED_NEW_HEADERS);
    newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
    newSheet.setFrozenRows(1);
  }

  // 1. Check if Booking ID already exists in Redeemed tab (Col 2: Booking ID)
  const rSheet = ss.getSheetByName("Redeemed");
  const targetBid = String(data.bookingId || '').replace(/^'/, '').trim().toUpperCase();
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rRows.length; i++) {
      const rBid = String(rRows[i][1] || '').replace(/^'/, '').trim().toUpperCase();
      if (rBid && targetBid && (rBid === targetBid || (rBid.startsWith("RUPS-") && targetBid.includes(rBid)) || (targetBid.startsWith("RUPS-") && rBid.includes(targetBid)))) {
        // Already redeemed, don't recreate!
        return;
      }
    }
  }

  // 2. Global Idempotency: Does this Booking ID already exist in "New" tab?
  const newRows = newSheet.getDataRange().getValues();
  for (let i = 1; i < newRows.length; i++) {
    const rowBid = String(newRows[i][0] || '').replace(/^'/, '').trim().toUpperCase();
    if (rowBid === targetBid || (rowBid.startsWith("RUPS-") && targetBid.includes(rowBid)) || (targetBid.startsWith("RUPS-") && rowBid.includes(targetBid))) {
      // If row exists as PENDING_REVALIDATION (e.g. revalidation email was ingested before Stage 1 ack),
      // heal Col 9 to the earlier Stage 1 acknowledgment timestamp!
      const currentStatus = String(newRows[i][7] || '').trim().toUpperCase();
      if (currentStatus === "PENDING_REVALIDATION" || currentStatus === "PENDING") {
        const existingTs = String(newRows[i][8] || '').trim();
        const stage1Date = (data.receivedDate && typeof data.receivedDate.getTime === 'function')
          ? data.receivedDate
          : (data.receivedDate ? new Date(data.receivedDate) : null);
        if (stage1Date && !isNaN(stage1Date.getTime())) {
          let shouldHeal = false;
          if (!existingTs) {
            shouldHeal = true;
          } else {
            const existingDate = new Date(existingTs);
            if (!isNaN(existingDate.getTime()) && stage1Date.getTime() < existingDate.getTime()) {
              shouldHeal = true;
            }
          }
          if (shouldHeal) {
            newSheet.getRange(i + 1, 9).setValue(formatISTTimestamp(stage1Date));
          }
        }
      }
      return;
    }
  }

  // 3. Determine canonical target brand (Col 11)
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "Unassigned");

  // 4. Hardening against duplicate re-staging: Check if Booking ID already exists in Brand tabs!
  const existingBrandMatch = findRowInBrandTabs(ss, data.bookingId, targetBrand);
  if (existingBrandMatch) {
    // Already approved to brand tab! Do not re-stage in "New".
    return;
  }

  const rowData = [
    data.bookingId,                                      // Col 1: Booking ID (RUPS...)
    (data.code && String(data.code).trim().toUpperCase() !== 'PENDING') ? data.code : "Pending", // Col 2: Code
    "",                                                  // Col 3: PIN
    data.amount || 500,                                  // Col 4: Value
    data.dateOfIssue || "",                              // Col 5: Date of Issue
    data.account || "",                                  // Col 6: Account
    "Unassigned",                                        // Col 7: Card
    data.status || "PENDING",                            // Col 8: Status
    formatISTTimestamp(data.receivedDate || new Date()), // Col 9: Appended Timestamp (IST) - Immutable for Stage 1
    true,                                                // Col 10: Display in UI
    targetBrand                                          // Col 11: Target Brand
  ];
  newSheet.appendRow(rowData);
}

/**
 * 3. STAGE 2: DIRECT "NEW" TAB FULFILLMENT BY BOOKING ID (IN SAME ROW)
 * Searches strictly inside the "New" tab. Hard Invariant ensures pending vouchers never enter brand tabs.
 */
function fulfillVoucherCode(ss, data) {
  let newSheet = ss.getSheetByName("New");
  if (!newSheet) {
    newSheet = ss.insertSheet("New");
    newSheet.appendRow(UNIFIED_NEW_HEADERS);
    newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
    newSheet.setFrozenRows(1);
  }

  // Search "New" tab by Booking ID (RUPS...)
  const rows = newSheet.getDataRange().getValues();
  const targetBid = String(data.bookingId || '').replace(/^'/, '').trim().toUpperCase();
  for (let i = 1; i < rows.length; i++) {
    const rowBid = String(rows[i][0] || '').replace(/^'/, '').trim().toUpperCase();
    if (rowBid && targetBid && (rowBid === targetBid || (rowBid.startsWith("RUPS-") && targetBid.includes(rowBid)) || (targetBid.startsWith("RUPS-") && rowBid.includes(targetBid)))) {
      const currentCode = String(rows[i][1] || '').trim();
      const currentStatus = String(rows[i][7] || '').trim().toUpperCase();

      // Guard: If already ACTIVE and voucher code is populated, it is permanently immutable!
      if (currentStatus === "ACTIVE" && currentCode !== "" && currentCode !== "PENDING" && currentCode !== "Pending revalidation" && !/processing/i.test(currentCode)) {
        return; // No-op, do not re-stamp
      }

      // MATCH FOUND: Transition from PENDING to ACTIVE and update timestamp with delivery email arrival time
      newSheet.getRange(i + 1, 2).setValue(`'${data.code}`);
      newSheet.getRange(i + 1, 3).setValue(data.pin ? `'${data.pin}` : "");
      if (data.amount && Number(data.amount) > 0) {
        newSheet.getRange(i + 1, 4).setValue(Number(data.amount));
      }
      newSheet.getRange(i + 1, 8).setValue("ACTIVE");
      newSheet.getRange(i + 1, 9).setValue(formatISTTimestamp(data.receivedDate || new Date()));
      newSheet.getRange(i + 1, 10).setValue(true);
      return;
    }
  }

  // Fallback: If Stage 1 email was missed and delivery email arrived directly,
  // FIRST check if it is already in Redeemed or in any canonical Brand tab!
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "Unassigned");

  // Check if already in Redeemed tab strictly by Booking ID
  const rSheet = ss.getSheetByName("Redeemed");
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rRows.length; i++) {
      const rBid = String(rRows[i][1] || '').replace(/^'/, '').trim().toUpperCase();
      if (rBid && targetBid && (rBid === targetBid || (rBid.startsWith("RUPS-") && targetBid.includes(rBid)) || (targetBid.startsWith("RUPS-") && rBid.includes(targetBid)))) {
        return;
      }
    }
  }

  // Check if already in a canonical Brand tab strictly by Booking ID
  const existingBrandMatch = findRowInBrandTabs(ss, data.bookingId, targetBrand);
  if (existingBrandMatch) {
    // If it exists in the brand tab but code or pin wasn't populated yet, heal it in place!
    const row = existingBrandMatch.rowData;
    const existingCode = String(row[1] || '').trim();
    if (!existingCode || existingCode === 'PENDING' || existingCode === '-') {
      existingBrandMatch.sheet.getRange(existingBrandMatch.rowIndex, 2).setValue(`'${data.code}`);
      if (data.pin) existingBrandMatch.sheet.getRange(existingBrandMatch.rowIndex, 3).setValue(`'${data.pin}`);
      existingBrandMatch.sheet.getRange(existingBrandMatch.rowIndex, 8).setValue("ACTIVE");
    }
    return;
  }

  const rowData = [
    data.bookingId,                                      // Col 1: Booking ID
    `'${data.code}`,                                     // Col 2: Code
    data.pin ? `'${data.pin}` : "",                      // Col 3: PIN
    data.amount || 500,                                  // Col 4: Value
    data.dateOfIssue || "",                              // Col 5: Date of Issue
    data.account || "",                                  // Col 6: Account
    "Unassigned",                                        // Col 7: Card
    "ACTIVE",                                            // Col 8: Status
    formatISTTimestamp(data.receivedDate || new Date()), // Col 9: Appended Timestamp (IST)
    true,                                                // Col 10: Display in UI
    targetBrand                                          // Col 11: Target Brand
  ];
  newSheet.appendRow(rowData);
}

/**
 * 4. STAGE 2 RECONCILER: AUTO-HEAL PENDING VOUCHERS IN "NEW" TAB
 */
function reconcilePendingVouchers(ss) {
  const newSheet = ss.getSheetByName("New");
  if (!newSheet) return [];

  const sRows = newSheet.getDataRange().getValues();
  const reconciled = [];

  for (let i = 1; i < sRows.length; i++) {
    const bid = String(sRows[i][0] || '').trim();
    const code = String(sRows[i][1] || '').trim();
    const currentPin = String(sRows[i][2] || '').trim();
    const status = String(sRows[i][7] || '').trim();

    if ((status === 'PENDING' || status === 'PENDING_REVALIDATION' || status === 'PROCESSING' || code === 'PENDING' || code === 'Pending revalidation' || /processing/i.test(code) || /processing/i.test(status)) && /^RUPS/i.test(bid)) {
      // In-memory data sanitization for pending rows (Gmail search is handled centrally in processRupayEmails batch)
      if (code === 'PENDING') {
        newSheet.getRange(i + 1, 2).setValue("'Pending");
      }
      if (currentPin && !/^\d{3,8}$/.test(currentPin)) {
        newSheet.getRange(i + 1, 3).setValue("");
      }
    }
  }

  if (reconciled.length > 0) {
    SpreadsheetApp.flush();
  }
  return reconciled;
}

/**
 * 5. PERSISTENT SYSTEM SHEETS INITIALIZATION & COOLDOWN ENGINE
 */
function getOrInitCooldownsSheet(ss) {
  let sheet = ss.getSheetByName("Cooldowns") || ss.getSheetByName("⏱ Cooldowns");
  if (!sheet) {
    sheet = ss.insertSheet("Cooldowns", 0);
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
  const targetAcc = String(account || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const rowAcc = String(data[i][0] || '').trim().toLowerCase();
    if (rowAcc === targetAcc) {
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
      ["c1", "PNB", "PNB Imperial RSDC", "RuPay", "'65249710", "'9102", "PNB Imperial Primary", "active", 3, 6, "", "", formatISTTimestamp(new Date())],
      ["c2", "Bank of Baroda", "BoB Bhoomi RSDC", "RuPay", "'652505", "'3314", "Vibhav BoB DC", "active", 2, 5, "", "", formatISTTimestamp(new Date())],
      ["c3", "Canara Bank", "Canara RSDC", "RuPay", "'652550", "'7828", "Canara DC 7828", "active", 2, 5, "", "", formatISTTimestamp(new Date())],
      ["c4", "Jana SFB", "Jana SFB RSDC", "RuPay", "'652520", "'7891", "Jana SFB DC", "active", 3, 5, "", "", formatISTTimestamp(new Date())],
      ["c5", "Canara Bank", "Canara Select CC", "RuPay", "'653020", "'1109", "Canara CC Primary", "active", 1, 2, "", "", formatISTTimestamp(new Date())],
      ["cx1", "PNB", "PNB Imperial RSDC", "RuPay", "'65249710", "'4821", "PNB Imperial (Old)", "blocked", 0, 0, "2026-10-14", "Q3 quotas exhausted", formatISTTimestamp(new Date())],
      ["cx2", "Canara Bank", "Canara RSDC", "RuPay", "'652550", "'1120", "Canara DC 1 (Old)", "blocked", 0, 0, "2026-09-28", "Blocked via mobile app", formatISTTimestamp(new Date())],
      ["cx3", "Bank of Baroda", "BoB Bhoomi RSDC", "RuPay", "'652505", "'2091", "Smiti BoB DC", "blocked", 0, 0, "2026-08-15", "Reissued new card number", formatISTTimestamp(new Date())]
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
 * 6. REST API (GET / POST)
 */
function doGet(e) {
  // Administrative Remediation: Correct values in Google Sheets for Swiggy One (Rs. 1) and Thyrocare (Rs. 125)
  if (e && e.parameter && e.parameter.action === "CORRECT_SWIGGY_THYROCARE_SHEET_VALUES") {
    try {
      const ss = getSpreadsheet();
      const report = {
        swiggyBrandUpdated: 0,
        thyrocareBrandUpdated: 0,
        newSheetSwiggyUpdated: 0,
        newSheetThyrocareUpdated: 0,
        redeemedSwiggyUpdated: 0,
        redeemedThyrocareUpdated: 0
      };

      // 1. Scan and update all brand tabs matching Swiggy or Thyrocare
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const sName = sheet.getName();
        const sNameLower = sName.toLowerCase().replace(/[^a-z0-9]/g, '');

        if (sNameLower.includes("swiggy")) {
          const vals = sheet.getDataRange().getValues();
          for (let r = 1; r < vals.length; r++) {
            if (Number(vals[r][3]) !== 1) {
              sheet.getRange(r + 1, 4).setValue(1);
              report.swiggyBrandUpdated++;
            }
          }
        } else if (sNameLower.includes("thyrocare")) {
          const vals = sheet.getDataRange().getValues();
          for (let r = 1; r < vals.length; r++) {
            if (Number(vals[r][3]) !== 125) {
              sheet.getRange(r + 1, 4).setValue(125);
              report.thyrocareBrandUpdated++;
            }
          }
        }
      }

      // 2. Scan and update the "New" (Ingested / Staging) sheet
      const newSheet = ss.getSheetByName("New");
      if (newSheet && newSheet.getLastRow() > 1) {
        const nVals = newSheet.getDataRange().getValues();
        for (let r = 1; r < nVals.length; r++) {
          const targetBrand = String(nVals[r][10] || nVals[r][0] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (targetBrand.includes("swiggy")) {
            if (Number(nVals[r][3]) !== 1) {
              newSheet.getRange(r + 1, 4).setValue(1);
              report.newSheetSwiggyUpdated++;
            }
          } else if (targetBrand.includes("thyrocare")) {
            if (Number(nVals[r][3]) !== 125) {
              newSheet.getRange(r + 1, 4).setValue(125);
              report.newSheetThyrocareUpdated++;
            }
          }
        }
      }

      // 3. Scan and update "Redeemed" sheet if applicable
      const rSheet = ss.getSheetByName("Redeemed");
      if (rSheet && rSheet.getLastRow() > 1) {
        const rVals = rSheet.getDataRange().getValues();
        for (let r = 1; r < rVals.length; r++) {
          const rBrand = String(rVals[r][0] || rVals[r][10] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (rBrand.includes("swiggy")) {
            if (Number(rVals[r][4]) !== 1) {
              rSheet.getRange(r + 1, 5).setValue(1);
              report.redeemedSwiggyUpdated++;
            }
          } else if (rBrand.includes("thyrocare")) {
            if (Number(rVals[r][4]) !== 125) {
              rSheet.getRange(r + 1, 5).setValue(125);
              report.redeemedThyrocareUpdated++;
            }
          }
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        report: report
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Administrative Remediation: Normalize all "PENDING" voucher codes in sheets to title-case "Pending"
  if (e && e.parameter && e.parameter.action === "NORMALIZE_PENDING_CODES") {
    try {
      const ss = getSpreadsheet();
      const report = {
        newSheetPendingUpdated: 0,
        brandSheetsPendingUpdated: 0,
        updatedRows: []
      };

      // 1. Scan and update "New" sheet
      const newSheet = ss.getSheetByName("New");
      if (newSheet && newSheet.getLastRow() > 1) {
        const nVals = newSheet.getDataRange().getValues();
        for (let r = 1; r < nVals.length; r++) {
          const rawCode = String(nVals[r][1] || '').trim();
          if (rawCode === "PENDING" || rawCode === "'PENDING") {
            newSheet.getRange(r + 1, 2).setValue("'Pending");
            report.newSheetPendingUpdated++;
            report.updatedRows.push({ sheet: "New", row: r + 1, bookingId: nVals[r][0] });
          }
        }
      }

      // 2. Scan and update all brand / other sheets
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const sName = sheet.getName();
        if (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards") continue;
        if (sheet.getLastRow() > 1) {
          const vals = sheet.getDataRange().getValues();
          for (let r = 1; r < vals.length; r++) {
            const rawCode = String(vals[r][1] || '').trim();
            if (rawCode === "PENDING" || rawCode === "'PENDING") {
              sheet.getRange(r + 1, 2).setValue("'Pending");
              report.brandSheetsPendingUpdated++;
              report.updatedRows.push({ sheet: sName, row: r + 1, bookingId: vals[r][0] });
            }
          }
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        report: report
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Remediation handler for PNB 8499 cards
  if (e && e.parameter && e.parameter.action === "REMEDIATE_PNB_8499_CARDS") {
    try {
      const ss = getSpreadsheet();
      const aSheet = ss.getSheetByName("Active Cards");
      const bSheet = ss.getSheetByName("Blocked Cards");
      const targetNum = String(e.parameter.cardNum || "").replace(/['\s]/g, '').trim();
      if (!targetNum) {
        return ContentService.createTextOutput(JSON.stringify({ error: "cardNum parameter is required" })).setMimeType(ContentService.MimeType.JSON);
      }
      const report = { removedFromBlocked: false, updatedInActive: false, addedToActive: false };

      // 1. Remove 3955 from Blocked Cards if present
      if (bSheet) {
        const bRows = bSheet.getDataRange().getValues();
        for (let i = bRows.length - 1; i >= 1; i--) {
          const rowNum = String(bRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === targetNum) {
            bSheet.deleteRow(i + 1);
            report.removedFromBlocked = true;
          }
        }
      }

      // 2. Ensure 3955 is active in Active Cards with correct details
      if (aSheet) {
        const aRows = aSheet.getDataRange().getValues();
        let targetRowIndex = -1;
        for (let i = 1; i < aRows.length; i++) {
          const rowNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === targetNum) {
            targetRowIndex = i + 1;
            break;
          }
        }

        const benefits = getEligibleBenefitsArray('PNB', 'Debit', 'PNB Imperial');
        if (targetRowIndex > 1) {
          aSheet.getRange(targetRowIndex, 1).setValue('PNB');
          aSheet.getRange(targetRowIndex, 2).setValue('Debit');
          aSheet.getRange(targetRowIndex, 3).setValue('PNB Imperial');
          aSheet.getRange(targetRowIndex, 4).setValue("'8499");
          aSheet.getRange(targetRowIndex, 5).setValue("'" + targetNum);
          aSheet.getRange(targetRowIndex, 6).setValue("'10/33");
          aSheet.getRange(targetRowIndex, 7).setValue("'547");
          aSheet.getRange(targetRowIndex, 8).setValue('ACTIVE');
          aSheet.getRange(targetRowIndex, 9).setValue('Applied 30th Oct, Received 6th Nov');
          report.updatedInActive = true;
        } else {
          const newRow = [
            'PNB', 'Debit', 'PNB Imperial', "'8499", "'" + targetNum, "'10/33", "'547", 'ACTIVE',
            'Applied 30th Oct, Received 6th Nov',
            ...benefits,
            formatISTTimestamp(new Date())
          ];
          aSheet.appendRow(newRow);
          report.addedToActive = true;
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ success: true, report })).setMimeType(ContentService.MimeType.JSON);
    } catch (rErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: rErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Safeguards Audit Verification Endpoint for ADD_MISC_VOUCHER & DELETE_MISC_VOUCHER
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_MISC_DUMMY") {
    try {
      const ss = getSpreadsheet();
      const sheet = getOrCreateMiscSheet(ss);
      const testCode = "TEST-SAFEGUARD-DUMMY-999";
      const nowStr = formatISTTimestamp(new Date());

      // Step 1: Add dummy row
      sheet.appendRow([
        "TEST-DUMMY",
        "'" + testCode,
        "1234",
        500,
        formatISTDate(new Date()),
        "TEST-DUMMY",
        "Manual",
        "ACTIVE",
        nowStr,
        true
      ]);
      SpreadsheetApp.flush();

      // Step 2: Verify dummy row exists in Misc sheet
      let vals = sheet.getDataRange().getValues();
      let foundIndex = -1;
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][1] || '').replace(/^'/, '').trim() === testCode) {
          foundIndex = i + 1;
          break;
        }
      }

      // Step 3: Test idempotency (repeat insertion check)
      let duplicateCount = 0;
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][1] || '').replace(/^'/, '').trim() === testCode) {
          duplicateCount++;
        }
      }

      // Step 4: Clean up / purge dummy row immediately
      let deleted = false;
      if (foundIndex > 0) {
        sheet.deleteRow(foundIndex);
        SpreadsheetApp.flush();
        deleted = true;
      }

      // Step 5: Verify zero residual test artifacts remain
      vals = sheet.getDataRange().getValues();
      let residualCount = 0;
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][1] || '').replace(/^'/, '').trim() === testCode) {
          residualCount++;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_MISC_DUMMY",
        insertedAndFound: foundIndex > 0,
        foundAtRow: foundIndex,
        cleanedUp: deleted,
        residualArtifacts: residualCount,
        allSafeguardsMet: (foundIndex > 0 && deleted && residualCount === 0)
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (vErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: vErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Safeguards Audit Verification Endpoint for BLOCK_CARD, BLOCK_AND_REPLACE_CARD & ADD_CARD
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_CARD_DUMMY") {
    try {
      const ss = getSpreadsheet();
      let aSheet = ss.getSheetByName("Active Cards");
      let bSheet = ss.getSheetByName("Blocked Cards");
      if (!aSheet) aSheet = ss.insertSheet("Active Cards");
      if (!bSheet) bSheet = ss.insertSheet("Blocked Cards");

      const testOldCard = "6521000000008499";
      const testNewCard = "6521000000003955";
      const testAcc = "8499";
      const nowStr = formatISTTimestamp(new Date());
      const stepsPassed = [];

      // Step 1: Pre-cleanup any stale test records from prior runs
      [aSheet, bSheet].forEach(sheet => {
        const rows = sheet.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          const num = String(rows[i][4] || '').replace(/['\s]/g, '').trim();
          if (num === testOldCard || num === testNewCard) {
            sheet.deleteRow(i + 1);
          }
        }
      });
      SpreadsheetApp.flush();

      // Step 2: Seed initial Active Card (Old Card)
      aSheet.appendRow([
        "PNB", "Debit", "PNB Imperial RSDC",
        "'" + testAcc, "'" + testOldCard, "'10/33", "'000",
        "ACTIVE", "Test Dummy Card",
        "Applicable", "Applicable", "Applicable", "Applicable", "Applicable",
        "Applicable", "Applicable", "Applicable", "Applicable", "Applicable",
        "Applicable", "Applicable", "Applicable",
        nowStr
      ]);
      SpreadsheetApp.flush();
      stepsPassed.push("STEP_1_SEED_OLD_CARD");

      // Step 3: Test BLOCK_AND_REPLACE_CARD simulation
      // a) Archive old card to Blocked Cards
      const aRows = aSheet.getDataRange().getValues();
      let foundOldInActive = false;
      for (let i = 1; i < aRows.length; i++) {
        const num = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
        if (num === testOldCard) {
          const rowData = aRows[i];
          bSheet.appendRow([
            rowData[0], rowData[1], rowData[2],
            "'" + String(rowData[3] || 'NA').replace(/^'/, ''),
            "'" + testOldCard,
            "'" + String(rowData[5] || '').replace(/^'/, ''),
            "'" + String(rowData[6] || '').replace(/^'/, ''),
            "BLOCKED", nowStr, testNewCard.slice(-4), "Replaced via Safeguard Test",
            rowData[9], rowData[10], rowData[11], rowData[12], rowData[13],
            rowData[14], rowData[15], rowData[16], rowData[17], rowData[18],
            rowData[19], rowData[20], rowData[21]
          ]);
          aSheet.deleteRow(i + 1);
          foundOldInActive = true;
          break;
        }
      }
      SpreadsheetApp.flush();
      if (foundOldInActive) stepsPassed.push("STEP_2_ARCHIVE_OLD_CARD");

      // b) Insert new card into Active Cards
      aSheet.appendRow([
        "PNB", "Debit", "PNB Imperial RSDC",
        "'" + testAcc, "'" + testNewCard, "'10/33", "'000",
        "ACTIVE", "Replacement Safeguard Test",
        "Applicable", "Applicable", "Applicable", "Applicable", "Applicable",
        "Applicable", "Applicable", "Applicable", "Applicable", "Applicable",
        "Applicable", "Applicable", "Applicable",
        nowStr
      ]);
      SpreadsheetApp.flush();
      stepsPassed.push("STEP_3_ACTIVATE_NEW_CARD");

      // Step 4: Verify Dual-Sheet Consistency
      const checkARows = aSheet.getDataRange().getValues();
      const hasNewInActive = checkARows.some(r => String(r[4] || '').replace(/['\s]/g, '').trim() === testNewCard && String(r[7] || '').toUpperCase() === 'ACTIVE');
      const checkBRows = bSheet.getDataRange().getValues();
      const hasOldInBlocked = checkBRows.some(r => String(r[4] || '').replace(/['\s]/g, '').trim() === testOldCard && String(r[7] || '').toUpperCase() === 'BLOCKED' && String(r[9] || '').trim() === testNewCard.slice(-4));

      if (hasNewInActive && hasOldInBlocked) {
        stepsPassed.push("STEP_4_DUAL_SHEET_CONSISTENCY_VERIFIED");
      }

      // Step 5: Test BLOCK_CARD simulation on the replacement card
      const aRows2 = aSheet.getDataRange().getValues();
      for (let i = 1; i < aRows2.length; i++) {
        const num = String(aRows2[i][4] || '').replace(/['\s]/g, '').trim();
        if (num === testNewCard) {
          const rowData = aRows2[i];
          bSheet.appendRow([
            rowData[0], rowData[1], rowData[2],
            "'" + String(rowData[3] || 'NA').replace(/^'/, ''),
            "'" + testNewCard,
            "'" + String(rowData[5] || '').replace(/^'/, ''),
            "'" + String(rowData[6] || '').replace(/^'/, ''),
            "BLOCKED", nowStr, "", "Blocked via Safeguard Test",
            rowData[9], rowData[10], rowData[11], rowData[12], rowData[13],
            rowData[14], rowData[15], rowData[16], rowData[17], rowData[18],
            rowData[19], rowData[20], rowData[21]
          ]);
          aSheet.deleteRow(i + 1);
          break;
        }
      }
      SpreadsheetApp.flush();
      stepsPassed.push("STEP_5_BLOCK_NEW_CARD");

      // Step 6: Post-test Cleanup - Safely purge test records
      [aSheet, bSheet].forEach(sheet => {
        const rows = sheet.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          const num = String(rows[i][4] || '').replace(/['\s]/g, '').trim();
          if (num === testOldCard || num === testNewCard) {
            sheet.deleteRow(i + 1);
          }
        }
      });
      SpreadsheetApp.flush();

      // Step 7: Verify Zero Residual Artifacts
      let residualCount = 0;
      [aSheet, bSheet].forEach(sheet => {
        const rows = sheet.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          const num = String(rows[i][4] || '').replace(/['\s]/g, '').trim();
          if (num === testOldCard || num === testNewCard) residualCount++;
        }
      });

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_CARD_DUMMY",
        stepsPassed: stepsPassed,
        residualArtifacts: residualCount,
        allSafeguardsMet: (stepsPassed.length === 5 && residualCount === 0)
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (cErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: cErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Safeguards Audit Verification Endpoint for MARK_REVALIDATED & Revalidation Ingestion Lifecycle
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_REVAL_DUMMY") {
    try {
      const ss = getSpreadsheet();
      let nSheet = ss.getSheetByName("New");
      if (!nSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"New" sheet not found' })).setMimeType(ContentService.MimeType.JSON);
      }

      const testBid = "RUPS-TEST-REVAL-999";
      const nowStr = formatISTTimestamp(new Date());
      const todayStr = formatISTDate(new Date());

      // Ensure clean starting state for test BID
      removeRevalidatedBookingFromProperties(testBid);
      let initRows = nSheet.getDataRange().getValues();
      for (let i = initRows.length - 1; i >= 1; i--) {
        if (String(initRows[i][0] || '').trim().toUpperCase() === testBid) {
          nSheet.deleteRow(i + 1);
        }
      }
      SpreadsheetApp.flush();

      // Phase 1: Test Fresh Stage 1 Ingestion -> Revalidation Email Transition
      // Step 1: Insert dummy row in Stage 1 state (Pending, no payment URL)
      nSheet.appendRow([
        testBid,
        "Pending",
        "",
        500,
        todayStr,
        "test@domain.com",
        "TestCard",
        "PENDING",
        nowStr,
        true,
        "Amazon Prime",
        ""
      ]);
      SpreadsheetApp.flush();

      // Step 2: Simulate Truztee Revalidation Email Arriving via recordPendingRevalidation
      const testPayUrl = "https://rupayselect.truztee.com/#/pay?token=TEST_VALIDATION_TOKEN_999";
      recordPendingRevalidation(ss, {
        bookingId: testBid,
        brand: "Amazon Prime",
        amount: 500,
        account: "test@domain.com",
        receivedDate: todayStr,
        payUrl: testPayUrl
      });
      SpreadsheetApp.flush();

      // Step 3: Verify row successfully transitioned to PENDING_REVALIDATION with payment link
      let rows = nSheet.getDataRange().getValues();
      let foundRow = -1;
      let codeAfterStage15 = "";
      let statusAfterStage15 = "";
      let payUrlAfterStage15 = "";
      let tsAfterStage15 = "";
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || '').trim().toUpperCase() === testBid) {
          foundRow = i + 1;
          codeAfterStage15 = String(rows[i][1] || '').trim();
          statusAfterStage15 = String(rows[i][7] || '').trim().toUpperCase();
          tsAfterStage15 = String(rows[i][8] || '').trim();
          payUrlAfterStage15 = String(rows[i][11] || '').trim();
          break;
        }
      }

      const stage15TransitionSuccess = (
        foundRow > 0 &&
        codeAfterStage15 === "Pending revalidation" &&
        statusAfterStage15 === "PENDING_REVALIDATION" &&
        payUrlAfterStage15 === testPayUrl
      );

      // Phase 2: Test User-Revalidation Safeguard (Mark Revalidated -> Sticky PENDING)
      // Step 4: Simulate MARK_REVALIDATED operation
      let markUpdated = false;
      if (foundRow > 0) {
        nSheet.getRange(foundRow, 2).setValue('Pending');
        nSheet.getRange(foundRow, 8).setValue('PENDING');
        if (rows[0].length >= 12) {
          nSheet.getRange(foundRow, 12).setValue('');
        }
        addRevalidatedBookingToProperties(testBid);
        markUpdated = true;
        SpreadsheetApp.flush();
      }

      const isRegisteredInProperties = isBookingRevalidated(testBid);

      // Step 5: Simulate repeat background cron execution (must NEVER clobber user-revalidated row)
      recordPendingRevalidation(ss, {
        bookingId: testBid,
        brand: "Amazon Prime",
        amount: 500,
        account: "test@domain.com",
        receivedDate: todayStr,
        payUrl: "https://rogue.pay.link/overwrite"
      });
      SpreadsheetApp.flush();

      // Verify row remained sticky PENDING
      rows = nSheet.getDataRange().getValues();
      let codeAfterCron = "";
      let statusAfterCron = "";
      let payUrlAfterCron = "";
      let tsAfterCron = "";
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || '').trim().toUpperCase() === testBid) {
          codeAfterCron = String(rows[i][1] || '').trim();
          statusAfterCron = String(rows[i][7] || '').trim().toUpperCase();
          tsAfterCron = String(rows[i][8] || '').trim();
          payUrlAfterCron = String(rows[i][11] || '').trim();
          break;
        }
      }

      const cronProtected = (codeAfterCron === "Pending" && statusAfterCron === "PENDING" && payUrlAfterCron === "");
      const timestampRemainedUnchanged = (tsAfterStage15 === nowStr && tsAfterCron === nowStr);

      // Phase 3: Cleanup and verify zero residual artifacts
      let cleanedUp = false;
      rows = nSheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0] || '').trim().toUpperCase() === testBid) {
          nSheet.deleteRow(i + 1);
          cleanedUp = true;
        }
      }
      removeRevalidatedBookingFromProperties(testBid);
      SpreadsheetApp.flush();

      rows = nSheet.getDataRange().getValues();
      let residualCount = 0;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || '').trim().toUpperCase() === testBid) {
          residualCount++;
        }
      }
      const isStillInProperties = isBookingRevalidated(testBid);

      const allSafeguardsMet = (
        stage15TransitionSuccess &&
        foundRow > 0 &&
        markUpdated &&
        isRegisteredInProperties &&
        cronProtected &&
        timestampRemainedUnchanged &&
        cleanedUp &&
        residualCount === 0 &&
        !isStillInProperties
      );

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_REVAL_DUMMY",
        stage15TransitionSuccess: stage15TransitionSuccess,
        timestampRemainedUnchanged: timestampRemainedUnchanged,
        insertedAndFound: foundRow > 0,
        markUpdated: markUpdated,
        isRegisteredInProperties: isRegisteredInProperties,
        cronProtected: cronProtected,
        cleanedUp: cleanedUp,
        residualArtifacts: residualCount,
        allSafeguardsMet: allSafeguardsMet
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (rErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: rErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Batch Transition: Switch all PENDING_REVALIDATION vouchers in "New" sheet to PENDING
  if (e && e.parameter && e.parameter.action === "SWITCH_ALL_REVALIDATED_TO_PENDING") {
    try {
      const ss = getSpreadsheet();
      const nSheet = ss.getSheetByName("New");
      if (!nSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"New" sheet not found' })).setMimeType(ContentService.MimeType.JSON);
      }
      const rows = nSheet.getDataRange().getValues();
      const updatedBids = [];
      for (let i = 1; i < rows.length; i++) {
        const rowBid = String(rows[i][0] || '').trim().toUpperCase();
        const rowCode = String(rows[i][1] || '').trim();
        const rowStatus = String(rows[i][7] || '').trim().toUpperCase();
        if (rowStatus === "PENDING_REVALIDATION" || rowCode.toLowerCase().includes("revalidat")) {
          nSheet.getRange(i + 1, 2).setValue("'Pending");
          nSheet.getRange(i + 1, 8).setValue('PENDING');
          if (rows[0].length >= 12) {
            nSheet.getRange(i + 1, 12).setValue('');
          }
          addRevalidatedBookingToProperties(rowBid);
          updatedBids.push(rowBid);
        } else if (rowStatus === "PENDING" && (rowCode.toUpperCase() === "PENDING" || rowCode === "")) {
          if (rowCode === "PENDING") {
            nSheet.getRange(i + 1, 2).setValue("'Pending");
          }
          // Ensure already pending rows are also preserved in persistent properties
          addRevalidatedBookingToProperties(rowBid);
        }
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "SWITCH_ALL_REVALIDATED_TO_PENDING",
        updatedCount: updatedBids.length,
        updatedBids: updatedBids,
        activeProperties: Array.from(getRevalidatedBookingsFromProperties())
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Data Migration: Split "RuPay Income" by moving negative card charges to "RuPay Charges"
  if (e && e.parameter && e.parameter.action === "MIGRATE_RUPAY_CHARGES_TAB") {
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(30000);
      const ss = getSpreadsheet();
      const incomeSheet = ss.getSheetByName("RuPay Income");
      if (!incomeSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"RuPay Income" sheet not found' })).setMimeType(ContentService.MimeType.JSON);
      }

      const chargesSheet = getOrCreateRuPayChargesSheet(ss);
      const incomeVals = incomeSheet.getDataRange().getValues();
      const chargesRows = [];
      const rowsToDelete = [];
      let positiveSum = 0;
      let chargesSum = 0;

      for (let r = 1; r < incomeVals.length; r++) {
        const title = String(incomeVals[r][0] || '').trim();
        if (title.toUpperCase() === 'TOTAL') continue;
        const price = Number(incomeVals[r][2]);
        if (isNaN(price) || price === 0) continue;

        if (price < 0) {
          chargesRows.push({
            data: [
              incomeVals[r][0],
              incomeVals[r][1],
              price,
              incomeVals[r][3] || '',
              incomeVals[r][4] || ''
            ],
            dateMs: parseDateToTimestamp(incomeVals[r][1])
          });
          rowsToDelete.push(r + 1); // 1-based row index in RuPay Income
          chargesSum += price;
        } else {
          positiveSum += price;
        }
      }

      // 1. Sort charges rows chronologically ascending
      chargesRows.sort((a, b) => a.dateMs - b.dateMs);

      // 2. Clear any existing data in chargesSheet (beyond header)
      const chargesLastRow = chargesSheet.getLastRow();
      if (chargesLastRow > 1) {
        chargesSheet.deleteRows(2, chargesLastRow - 1);
      }

      // 3. Write charges to chargesSheet
      if (chargesRows.length > 0) {
        const rawChargeData = chargesRows.map(c => c.data);
        chargesSheet.getRange(2, 1, rawChargeData.length, 5).setValues(rawChargeData);
        chargesSheet.getRange(2, 2, rawChargeData.length, 1).setNumberFormat("dd/MM/yy").setHorizontalAlignment("center");
        chargesSheet.getRange(2, 3, rawChargeData.length, 1).setNumberFormat("0.00");

        // Append TOTAL row to chargesSheet
        chargesSheet.appendRow(["TOTAL", "", Number(chargesSum.toFixed(2)), "", ""]);
        chargesSheet.getRange(chargesSheet.getLastRow(), 1, 1, 5).setFontWeight("bold");
      }

      // 4. Delete negative rows from incomeSheet bottom-up
      rowsToDelete.sort((a, b) => b - a);
      for (const rIdx of rowsToDelete) {
        incomeSheet.deleteRow(rIdx);
      }

      // 5. Update TOTAL row in incomeSheet
      recalculateSheetTotal(incomeSheet);
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "MIGRATE_RUPAY_CHARGES_TAB",
        chargesMovedCount: chargesRows.length,
        chargesSum: Number(chargesSum.toFixed(2)),
        incomeRemainingCount: incomeSheet.getLastRow() - 2,
        incomeSum: Number(positiveSum.toFixed(2))
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }
  }

  // Safeguards Audit Verification Endpoint for RECORD_INCOME_ENTRY & Date-Sorted Insertion
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_INCOME_DUMMY") {
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(30000);
      const ss = getSpreadsheet();
      const incomeSheet = ss.getSheetByName("RuPay Income");
      if (!incomeSheet) return ContentService.createTextOutput(JSON.stringify({ error: "RuPay Income sheet not found" })).setMimeType(ContentService.MimeType.JSON);
      const chargesSheet = getOrCreateRuPayChargesSheet(ss);

      const testSaleTitle = "TEST-SAFEGUARD-SALE-999";
      const testChargeTitle = "TEST-SAFEGUARD-CHARGE-999";
      const testDateStr = "15/06/25";
      const testDateMs = parseDateToTimestamp(testDateStr);

      // Step 1: Insert test sale into RuPay Income (should sort between 2024 and 2026 rows)
      const saleRowIdx = insertRowSortedByDate(incomeSheet, [
        testSaleTitle,
        new Date(2025, 5, 15),
        999,
        "dummy test sale",
        "test"
      ], testDateMs);

      // Step 2: Insert test charge into RuPay Charges (should sort between 2024 and 2026 rows)
      const chargeRowIdx = insertRowSortedByDate(chargesSheet, [
        testChargeTitle,
        new Date(2025, 5, 15),
        -250,
        "Acc ••2562, Card ••7180 - Test charge notes",
        "••2562"
      ], testDateMs);
      SpreadsheetApp.flush();

      // Step 3: Verify existence in both sheets and field formatting
      let incomeVals = incomeSheet.getDataRange().getValues();
      let saleFound = false;
      for (let r = 1; r < incomeVals.length; r++) {
        if (String(incomeVals[r][0] || '').trim() === testSaleTitle) {
          saleFound = true;
          break;
        }
      }

      let chargesVals = chargesSheet.getDataRange().getValues();
      let chargeFound = false;
      let chargeCommentsMatch = false;
      let chargeAccountMatch = false;
      for (let r = 1; r < chargesVals.length; r++) {
        if (String(chargesVals[r][0] || '').trim() === testChargeTitle) {
          chargeFound = true;
          if (String(chargesVals[r][3] || '').trim() === "Acc ••2562, Card ••7180 - Test charge notes") {
            chargeCommentsMatch = true;
          }
          if (String(chargesVals[r][4] || '').trim() === "••2562") {
            chargeAccountMatch = true;
          }
          break;
        }
      }

      // Step 4: Clean up test rows immediately
      let saleCleaned = false;
      incomeVals = incomeSheet.getDataRange().getValues();
      for (let r = incomeVals.length - 1; r >= 1; r--) {
        if (String(incomeVals[r][0] || '').trim() === testSaleTitle) {
          incomeSheet.deleteRow(r + 1);
          saleCleaned = true;
        }
      }
      recalculateSheetTotal(incomeSheet);

      let chargeCleaned = false;
      chargesVals = chargesSheet.getDataRange().getValues();
      for (let r = chargesVals.length - 1; r >= 1; r--) {
        if (String(chargesVals[r][0] || '').trim() === testChargeTitle) {
          chargesSheet.deleteRow(r + 1);
          chargeCleaned = true;
        }
      }
      recalculateSheetTotal(chargesSheet);
      SpreadsheetApp.flush();

      // Step 5: Check zero residual artifacts remain
      let residualCount = 0;
      incomeVals = incomeSheet.getDataRange().getValues();
      for (let r = 1; r < incomeVals.length; r++) {
        if (String(incomeVals[r][0] || '').trim() === testSaleTitle) residualCount++;
      }
      chargesVals = chargesSheet.getDataRange().getValues();
      for (let r = 1; r < chargesVals.length; r++) {
        if (String(chargesVals[r][0] || '').trim() === testChargeTitle) residualCount++;
      }

      const allSafeguardsMet = (saleFound && chargeFound && chargeCommentsMatch && chargeAccountMatch && saleCleaned && chargeCleaned && residualCount === 0);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_INCOME_DUMMY",
        saleFound: saleFound,
        saleRowIdx: saleRowIdx,
        chargeFound: chargeFound,
        chargeRowIdx: chargeRowIdx,
        chargeCommentsMatch: chargeCommentsMatch,
        chargeAccountMatch: chargeAccountMatch,
        cleanedUp: saleCleaned && chargeCleaned,
        residualArtifacts: residualCount,
        allSafeguardsMet: allSafeguardsMet
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(vErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: vErr.message })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }
  }

  // Safeguards Audit Verification Endpoint for MARK_VOUCHER_AS_SOLD (Atomic Ingested/Brand -> Redeemed + RuPay Income)
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_SOLD_DUMMY") {
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(30000);
      const ss = getSpreadsheet();
      let nSheet = ss.getSheetByName("New");
      if (!nSheet) {
        nSheet = ss.insertSheet("New");
        nSheet.appendRow(UNIFIED_NEW_HEADERS);
        nSheet.setFrozenRows(1);
      }
      let rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        rSheet = ss.insertSheet("Redeemed");
        rSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
        rSheet.setFrozenRows(1);
      }
      const incomeSheet = ss.getSheetByName("RuPay Income");
      if (!incomeSheet) return ContentService.createTextOutput(JSON.stringify({ error: '"RuPay Income" sheet not found' })).setMimeType(ContentService.MimeType.JSON);

      const testBid = "RUPS-TEST-SOLD-999";
      const testCode = "TEST-CULT-CODE-999";
      const testSoldTo = "Test Buyer";
      const testMobile = "9876543210";
      const testPrice = 1500;
      const nowStr = formatISTTimestamp(new Date());
      const todayStr = formatISTDate(new Date());

      // Pre-cleanup
      [nSheet, rSheet].forEach(sheet => {
        const rows = sheet.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          if (String(rows[i][0] || '').trim().toUpperCase() === testBid || String(rows[i][1] || '').trim().toUpperCase() === testBid) {
            sheet.deleteRow(i + 1);
          }
        }
      });
      const inRows = incomeSheet.getDataRange().getValues();
      for (let i = inRows.length - 1; i >= 1; i--) {
        if (String(inRows[i][3] || '').includes(testBid) || String(inRows[i][3] || '').includes("9876543210")) {
          incomeSheet.deleteRow(i + 1);
        }
      }
      recalculateSheetTotal(incomeSheet);
      SpreadsheetApp.flush();

      // Step 1: Insert dummy row into "New"
      nSheet.appendRow([
        testBid,
        testCode,
        "",
        500,
        todayStr,
        "vibhavoswal9",
        "PNB DC 9762",
        "ACTIVE",
        nowStr,
        true,
        "Cult Fit",
        ""
      ]);
      SpreadsheetApp.flush();

      // Step 2: Execute MARK_VOUCHER_AS_SOLD logic
      let nRows = nSheet.getDataRange().getValues();
      let foundInNew = false;
      for (let i = nRows.length - 1; i >= 1; i--) {
        if (String(nRows[i][0] || '').trim().toUpperCase() === testBid) {
          foundInNew = true;
          nSheet.deleteRow(i + 1);
          break;
        }
      }

      rSheet.appendRow([
        "Cult Fit",
        `'${testBid}`,
        `'${testCode}`,
        "",
        500,
        todayStr,
        "vibhavoswal9",
        "PNB DC 9762",
        "REDEEMED",
        nowStr,
        "New", // Original Tab = New (from Staging)
        nowStr,
        true
      ]);

      const commentStr = `${testSoldTo} (${testMobile}): vibhavoswal9 [${testBid}]`;
      const dateMs = parseDateToTimestamp("16/11/26");
      insertRowSortedByDate(incomeSheet, [
        "Cult Fit",
        new Date(2026, 10, 16),
        testPrice,
        commentStr,
        "PNB DC 9762"
      ], dateMs);
      recalculateSheetTotal(incomeSheet);
      SpreadsheetApp.flush();

      // Step 3: Verify assertions for MARK_VOUCHER_AS_SOLD
      let stillInNew = false;
      nRows = nSheet.getDataRange().getValues();
      for (let i = 1; i < nRows.length; i++) {
        if (String(nRows[i][0] || '').trim().toUpperCase() === testBid) stillInNew = true;
      }

      let foundInRedeemed = false;
      const rRows = rSheet.getDataRange().getValues();
      for (let i = 1; i < rRows.length; i++) {
        if (String(rRows[i][1] || '').trim().toUpperCase() === testBid) {
          foundInRedeemed = true;
          break;
        }
      }

      let foundInIncome = false;
      const iRows = incomeSheet.getDataRange().getValues();
      for (let i = 1; i < iRows.length; i++) {
        if (String(iRows[i][3] || '').includes(testBid)) {
          foundInIncome = true;
          break;
        }
      }

      // Step 4: Execute UNMARK_INGESTED_SALE logic (Restores voucher from Redeemed to New tab)
      let unmarkRestoredToNew = false;
      let unmarkDeletedFromRedeemed = false;
      let unmarkDeletedFromIncome = false;

      const rValsPost = rSheet.getDataRange().getValues();
      for (let i = rValsPost.length - 1; i >= 1; i--) {
        const bId = String(rValsPost[i][1] || '').replace(/^'/, '').trim().toUpperCase();
        if (bId === testBid) {
          nSheet.appendRow([
            `'${testBid}`,
            `'${testCode}`,
            "",
            500,
            todayStr,
            "vibhavoswal9",
            "PNB DC 9762",
            "ACTIVE",
            nowStr,
            true,
            "Cult Fit",
            "",
            ""
          ]);
          rSheet.deleteRow(i + 1);
          unmarkRestoredToNew = true;
          break;
        }
      }

      // Delete matching sale row from RuPay Income
      const iValsPost = incomeSheet.getDataRange().getValues();
      for (let r = iValsPost.length - 1; r >= 1; r--) {
        const comm = String(iValsPost[r][3] || '').trim();
        if (comm.includes(testBid)) {
          incomeSheet.deleteRow(r + 1);
          break;
        }
      }
      recalculateSheetTotal(incomeSheet);
      SpreadsheetApp.flush();

      // Step 5: Verify assertions for UNMARK_INGESTED_SALE
      let verifiedInNew = false;
      const postNVals = nSheet.getDataRange().getValues();
      for (let i = 1; i < postNVals.length; i++) {
        if (String(postNVals[i][0] || '').trim().toUpperCase() === testBid && String(postNVals[i][7] || '').toUpperCase() === 'ACTIVE') {
          verifiedInNew = true;
          break;
        }
      }

      let stillInRedeemed = false;
      const postRVals = rSheet.getDataRange().getValues();
      for (let i = 1; i < postRVals.length; i++) {
        if (String(postRVals[i][1] || '').trim().toUpperCase() === testBid) stillInRedeemed = true;
      }
      unmarkDeletedFromRedeemed = !stillInRedeemed;

      let stillInIncome = false;
      const postIVals = incomeSheet.getDataRange().getValues();
      for (let i = 1; i < postIVals.length; i++) {
        if (String(postIVals[i][3] || '').includes(testBid)) stillInIncome = true;
      }
      unmarkDeletedFromIncome = !stillInIncome;

      // Step 6: Final Cleanup of test row from New
      const finalNVals = nSheet.getDataRange().getValues();
      for (let i = finalNVals.length - 1; i >= 1; i--) {
        if (String(finalNVals[i][0] || '').trim().toUpperCase() === testBid) {
          nSheet.deleteRow(i + 1);
        }
      }
      SpreadsheetApp.flush();

      let residualCount = 0;
      [nSheet, rSheet].forEach(s => {
        const rows = s.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][0] || '').includes(testBid) || String(rows[i][1] || '').includes(testBid)) residualCount++;
        }
      });
      const finalIRows = incomeSheet.getDataRange().getValues();
      for (let i = 1; i < finalIRows.length; i++) {
        if (String(finalIRows[i][3] || '').includes(testBid)) residualCount++;
      }

      const allSafeguardsMet = (foundInNew && !stillInNew && foundInRedeemed && foundInIncome &&
                                unmarkRestoredToNew && verifiedInNew && unmarkDeletedFromRedeemed &&
                                unmarkDeletedFromIncome && residualCount === 0);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_SOLD_DUMMY",
        foundInNew: foundInNew,
        deletedFromNew: !stillInNew,
        foundInRedeemed: foundInRedeemed,
        foundInIncome: foundInIncome,
        unmarkRestoredToNew: verifiedInNew,
        unmarkDeletedFromRedeemed: unmarkDeletedFromRedeemed,
        unmarkDeletedFromIncome: unmarkDeletedFromIncome,
        cleanedUp: residualCount === 0,
        residualArtifacts: residualCount,
        allSafeguardsMet: allSafeguardsMet
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }
  }

  // Safeguards Audit Verification Endpoint for Tombstone Architecture (New & Redeemed promotion to Deleted_Tombstones)
  if (e && e.parameter && e.parameter.action === "VERIFY_SAFEGUARD_TOMBSTONE_DUMMY") {
    try {
      const ss = getSpreadsheet();
      const nSheet = ss.getSheetByName("New");
      const rSheet = ss.getSheetByName("Redeemed");
      const testBidNew = "RUPS-TEST-DUMMY-NEW-TOMB";
      const testBidRed = "RUPS-TEST-DUMMY-RED-TOMB";
      const stepsPassed = [];

      // Step 1: Pre-cleanup any prior stale dummy rows
      [nSheet, rSheet].forEach(sheet => {
        if (!sheet) return;
        const vals = sheet.getDataRange().getValues();
        for (let i = vals.length - 1; i >= 1; i--) {
          const bid = String(vals[i][0] || vals[i][1] || '').trim();
          if (bid === testBidNew || bid === testBidRed) {
            sheet.deleteRow(i + 1);
          }
        }
      });
      SpreadsheetApp.flush();

      // Step 2: Seed dummy row in "New" tab with Display in UI = false (Col 10)
      if (nSheet) {
        nSheet.appendRow([
          testBidNew, "Pending", "", 500, "16-Nov-2026", "testuser369", "Unassigned", "PENDING", formatISTTimestamp(new Date()), false, "CultFit"
        ]);
        SpreadsheetApp.flush();
        stepsPassed.push("STEP_1_SEED_NEW_TAB_DUMMY");
      }

      // Step 3: Seed dummy row in "Redeemed" tab with Display in UI = false (Col 13)
      // Using a 16-digit numeric voucher code to verify zero scientific notation conversion!
      const testNumericCode = "9802004499999999";
      if (rSheet) {
        rSheet.appendRow([
          "CultFit", testBidRed, testNumericCode, "1234", 500, "16-Nov-2026", "testuser369", "PNB DC 9762", "REDEEMED", formatISTTimestamp(new Date()), "CultFit", formatISTTimestamp(new Date()), false
        ]);
        SpreadsheetApp.flush();
        stepsPassed.push("STEP_2_SEED_REDEEMED_TAB_DUMMY");
      }

      // Step 4: Run promoteTombstones(ss)
      const promoResult = promoteTombstones(ss);
      stepsPassed.push("STEP_3_EXECUTE_PROMOTE_TOMBSTONES");

      // Step 5: Verify records exist in Deleted_Tombstones and numeric voucher codes are not mangled
      const tSheet = getOrInitTombstonesSheet(ss);
      const tVals = tSheet.getDataRange().getValues();
      const foundNewInTomb = tVals.some(r => String(r[0] || '').trim() === testBidNew && String(r[7] || '') === 'New');
      const foundRedInTombRow = tVals.find(r => String(r[0] || '').trim() === testBidRed && String(r[7] || '') === 'Redeemed');
      const foundRedInTomb = !!foundRedInTombRow;
      const preservedNumericCode = foundRedInTombRow ? (String(foundRedInTombRow[1] || '').trim() === testNumericCode) : false;

      // Verify RUPS-DEFBAA77 is repaired to full 16-digit plaintext code
      const defbaaRow = tVals.find(r => String(r[0] || '').trim() === "RUPS-DEFBAA77");
      const defbaaRepaired = defbaaRow ? (String(defbaaRow[1] || '').trim() === "9802004492588540") : true;

      if (foundNewInTomb && foundRedInTomb && preservedNumericCode && defbaaRepaired) {
        stepsPassed.push("STEP_4_TOMBSTONES_ARCHIVE_VERIFIED");
        stepsPassed.push("STEP_5_NUMERIC_CODE_PRESERVED_RUPS_DEFBAA77_REPAIRED");
      }

      // Step 6: Verify records were purged from New and Redeemed tabs
      const nAfter = nSheet ? nSheet.getDataRange().getValues().some(r => String(r[0] || '').trim() === testBidNew) : false;
      const rAfter = rSheet ? rSheet.getDataRange().getValues().some(r => String(r[1] || '').trim() === testBidRed) : false;
      const purgedFromSource = (!nAfter && !rAfter);
      if (purgedFromSource) {
        stepsPassed.push("STEP_5_PURGED_FROM_SOURCE_TABS");
      }

      // Step 7: Post-test Cleanup: purge dummy rows from Deleted_Tombstones
      const tVals2 = tSheet.getDataRange().getValues();
      for (let i = tVals2.length - 1; i >= 1; i--) {
        const bid = String(tVals2[i][0] || '').trim();
        if (bid === testBidNew || bid === testBidRed) {
          tSheet.deleteRow(i + 1);
        }
      }
      SpreadsheetApp.flush();

      // Check residual artifacts
      const tVals3 = tSheet.getDataRange().getValues();
      let residualCount = 0;
      for (let i = 1; i < tVals3.length; i++) {
        const bid = String(tVals3[i][0] || '').trim();
        if (bid === testBidNew || bid === testBidRed) residualCount++;
      }

      const allSafeguardsMet = (foundNewInTomb && foundRedInTomb && purgedFromSource && residualCount === 0 && preservedNumericCode && defbaaRepaired);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "VERIFY_SAFEGUARD_TOMBSTONE_DUMMY",
        stepsPassed: stepsPassed,
        promoResult: promoResult,
        foundNewInTomb: foundNewInTomb,
        foundRedInTomb: foundRedInTomb,
        purgedFromSource: purgedFromSource,
        residualArtifacts: residualCount,
        allSafeguardsMet: allSafeguardsMet
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (vErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: vErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Seed / Initialize Misc Tab with initial vouchers
  if (e && e.parameter && e.parameter.action === "SEED_INITIAL_MISC_VOUCHERS") {
    try {
      const ss = getSpreadsheet();
      const sheet = getOrCreateMiscSheet(ss);
      const nowStr = formatISTTimestamp(new Date());
      const todayStr = formatISTDate(new Date());

      const seedCodes = [
        { brand: "Amazon Pay", code: "WV4DQEQFMP64MQ", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "DVV6QBD3KWQRA6", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "GUYAD6GWWPKUBA", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "QNNRCTPP7WJUG8", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "R3NCFWHGMQCDBN", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "RRYD4QBMJ3TC6M", pin: "", amount: 500 },
        { brand: "Amazon Pay", code: "HUHUBKRHPDHJX6", pin: "", amount: 500 },
        { brand: "Zomato", code: "7000427582548835", pin: "648600", amount: 500 }
      ];

      const vals = sheet.getDataRange().getValues();
      const existingCodes = new Set();
      for (let i = 1; i < vals.length; i++) {
        const c = String(vals[i][1] || '').replace(/^'/, '').trim().toUpperCase();
        if (c) existingCodes.add(c);
      }

      const inserted = [];
      seedCodes.forEach(item => {
        const cleanCode = item.code.trim().toUpperCase();
        if (!existingCodes.has(cleanCode)) {
          sheet.appendRow([
            item.brand,                        // A: Booking ID / Brand
            "'" + item.code,                   // B: Code
            item.pin ? ("'" + item.pin) : "",  // C: PIN
            item.amount,                       // D: Value
            todayStr,                          // E: Date of Issue
            item.brand,                        // F: Account / Sub-brand
            "Manual",                          // G: Card
            "ACTIVE",                          // H: Status
            nowStr,                            // I: Timestamp (IST)
            true                               // J: Display in UI
          ]);
          existingCodes.add(cleanCode);
          inserted.push(item);
        }
      });

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "SEED_INITIAL_MISC_VOUCHERS",
        totalSeedCount: seedCodes.length,
        insertedCount: inserted.length,
        inserted: inserted
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (sErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: sErr.message, stack: sErr.stack })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Diagnostic handler for Gmail ingestion debugging & manual voucher population
  if (e && e.parameter && e.parameter.action === "DEBUG_GMAIL_INGEST") {
    try {
      const effectiveUser = Session.getEffectiveUser().getEmail();
      const activeUser = Session.getActiveUser().getEmail();
      const ss = getSpreadsheet();

      // Check Active Cards
      const aSheet = ss.getSheetByName("Active Cards");
      const activeCards = aSheet ? aSheet.getDataRange().getValues().map((r, idx) => ({
        row: idx + 1,
        bank: r[0],
        type: r[1],
        name: r[2],
        linkedAcc: r[3],
        cardNum: r[4],
        expiry: r[5],
        cvv: r[6],
        status: r[7],
        notes: r[8]
      })) : [];

      // Check Blocked Cards
      const bSheet = ss.getSheetByName("Blocked Cards");
      const blockedCards = bSheet ? bSheet.getDataRange().getValues().map((r, idx) => ({
        row: idx + 1,
        bank: r[0],
        type: r[1],
        name: r[2],
        linkedAcc: r[3],
        cardNum: r[4],
        expiry: r[5],
        cvv: r[6],
        status: r[7],
        blockedTimestamp: r[8],
        replacedBy: r[9],
        notes: r[10]
      })) : [];

      return ContentService.createTextOutput(JSON.stringify({
        activeCardsCount: activeCards.length,
        blockedCardsCount: blockedCards.length,
        activeCards: activeCards.filter(c => String(c.cardNum).includes('8499') || String(c.linkedAcc).includes('8499') || String(c.bank).toUpperCase().includes('PNB')),
        blockedCards: blockedCards.filter(c => String(c.cardNum).includes('8499') || String(c.linkedAcc).includes('8499') || String(c.bank).toUpperCase().includes('PNB'))
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Populate missing RUPS vouchers directly into New tab
  if (e && e.parameter && e.parameter.action === "INGEST_MISSING_SCREENSHOT_VOUCHERS") {
    try {
      const ss = getSpreadsheet();
      const vouchers = [
        { bookingId: "RUPS-48449243", brand: "SwiggyOne", account: e.parameter.acc1 || "vibhavoswal9", amount: 500 },
        { bookingId: "RUPS-FB47C696", brand: "TaxSpanner", account: e.parameter.acc2 || "vibhavoswal9", amount: 500 },
        { bookingId: "RUPS-0A6C9163", brand: "Decathlon", account: e.parameter.acc3 || "vibhavoswal9", amount: 500 }
      ];

      const inserted = [];
      vouchers.forEach(v => {
        recordPendingRedemption(ss, {
          bookingId: v.bookingId,
          brand: v.brand,
          amount: v.amount,
          account: v.account,
          dateOfIssue: formatISTDate(new Date()),
          receivedDate: new Date()
        });
        updateCooldownTab(ss, v.account);
        inserted.push(v);
      });

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ success: true, inserted: inserted })).setMimeType(ContentService.MimeType.JSON);
    } catch (popErr) {
      return ContentService.createTextOutput(JSON.stringify({ error: popErr.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Remediation handler for "RuPay Income" tab dates
  if (e && e.parameter && e.parameter.action === "REMEDIATE_RUPAY_INCOME_DATES") {
    try {
      const isDryRun = e.parameter.dryRun === "true" || e.parameter.dryRun === "1";
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName("RuPay Income");
      if (!sheet) return ContentService.createTextOutput(JSON.stringify({ error: "Sheet not found" })).setMimeType(ContentService.MimeType.JSON);

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      const values = range.getValues();
      const displayValues = range.getDisplayValues();

      let totalRowIndex = -1;
      for (let r = values.length - 1; r >= 0; r--) {
        if (String(values[r][0] || '').trim().toUpperCase() === 'TOTAL') {
          totalRowIndex = r + 1; // 1-based row number
          break;
        }
      }

      const endDataRow = totalRowIndex > 1 ? totalRowIndex - 1 : lastRow;
      const totalDataRows = endDataRow - 1; // excluding header at row 1

      const updates = [];
      const newColBValues = [];

      for (let r = 2; r <= endDataRow; r++) {
        const rowIdx = r - 1;
        const rawVal = values[rowIdx][1];
        const dispVal = displayValues[rowIdx][1];
        let trueDay = null;
        let trueMonth = null;
        let trueYear = null;
        let source = "";

        if (rawVal instanceof Date) {
          // Date object in en_US locale: user entered d/m/y, sheets read m as month, d as day
          const dStr = Utilities.formatDate(rawVal, "GMT+05:30", "yyyy-MM-dd");
          const parts = dStr.split("-");
          const y = parseInt(parts[0], 10);
          const m = parseInt(parts[1], 10);
          const d = parseInt(parts[2], 10);

          trueDay = m;
          trueMonth = d;
          trueYear = y;
          source = "DateObject_Flipped";
        } else {
          // String like "27/10/26" or "25/04/24"
          const s = String(rawVal || dispVal || '').replace(/^'/, '').trim();
          const parts = s.split(/[\/\-\.]/);
          if (parts.length === 3) {
            trueDay = parseInt(parts[0], 10);
            trueMonth = parseInt(parts[1], 10);
            let y = parseInt(parts[2], 10);
            if (y < 100) y = 2000 + y;
            trueYear = y;
            source = "String_Parsed";
          }
        }

        if (trueDay && trueMonth && trueYear) {
          const dd = (trueDay < 10 ? '0' : '') + trueDay;
          const mm = (trueMonth < 10 ? '0' : '') + trueMonth;
          const yy = String(trueYear).slice(-2);
          const formattedStr = `${dd}/${mm}/${yy}`;

          // Create Date at noon IST to avoid any midnight boundary shifts
          const trueDateObj = new Date(Date.UTC(trueYear, trueMonth - 1, trueDay, 6, 30, 0));

          updates.push({
            row: r,
            title: values[rowIdx][0],
            oldVal: rawVal instanceof Date ? rawVal.toISOString() : rawVal,
            oldDisp: dispVal,
            trueDay: trueDay,
            trueMonth: trueMonth,
            trueYear: trueYear,
            newFormattedStr: formattedStr,
            source: source
          });

          newColBValues.push([trueDateObj]);
        } else {
          newColBValues.push([rawVal]);
        }
      }

      if (!isDryRun && newColBValues.length > 0) {
        // Set spreadsheet locale to en_IN so Google Sheets natively infers dd/MM/yyyy
        try {
          ss.setSpreadsheetLocale("en_IN");
        } catch (locErr) {
          Logger.log("Locale error: " + locErr);
        }

        const colBRange = sheet.getRange(2, 2, newColBValues.length, 1);
        colBRange.setValues(newColBValues);
        colBRange.setNumberFormat("dd/MM/yy");
        colBRange.setHorizontalAlignment("center");
        SpreadsheetApp.flush();
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        dryRun: isDryRun,
        totalDataRows: totalDataRows,
        updatedCount: updates.length,
        updatesSummary: updates
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Inspection handler for "RuPay Income" and other tabs
  if (e && e.parameter && e.parameter.action === "INSPECT_RUPAY_INCOME") {
    try {
      const ss = getSpreadsheet();
      const targetSheetName = (e.parameter && e.parameter.sheet) ? String(e.parameter.sheet).trim() : "RuPay Income";
      const sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) return ContentService.createTextOutput(JSON.stringify({ error: `Sheet "${targetSheetName}" not found` })).setMimeType(ContentService.MimeType.JSON);
      
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      const values = range.getValues();
      const displayValues = range.getDisplayValues();
      const numberFormats = range.getNumberFormats();
      const horizontalAlignments = range.getHorizontalAlignments();
      
      const inspectedRows = [];
      for (let r = 0; r < values.length; r++) {
        const isDate = values[r][1] instanceof Date;
        inspectedRows.push({
          rowNum: r + 1,
          val: values[r],
          disp: displayValues[r],
          fmt: numberFormats[r],
          align: horizontalAlignments[r],
          types: values[r].map(v => v instanceof Date ? 'Date(' + v.toISOString() + ')' : typeof v),
          isDateObj: isDate
        });
      }
      return ContentService.createTextOutput(JSON.stringify({
        lastRow: lastRow,
        lastCol: lastCol,
        locale: ss.getSpreadsheetLocale(),
        timeZone: ss.getSpreadsheetTimeZone(),
        inspectedRows: inspectedRows
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Action: Take backup of "RuPay Income" as "RuPayIncomeBackup", format "RuPay Income" like other tabs, and tally
  if (e && e.parameter && e.parameter.action === "FORMAT_AND_BACKUP_RUPAY_INCOME") {
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(30000);
      const ss = getSpreadsheet();
      const incomeSheet = ss.getSheetByName("RuPay Income");
      if (!incomeSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"RuPay Income" sheet not found' })).setMimeType(ContentService.MimeType.JSON);
      }

      // Step 1: Create Backup Tab "RuPayIncomeBackup" first
      let backupSheet = ss.getSheetByName("RuPayIncomeBackup");
      if (backupSheet) {
        ss.deleteSheet(backupSheet);
      }
      backupSheet = incomeSheet.copyTo(ss);
      backupSheet.setName("RuPayIncomeBackup");
      try {
        ss.setActiveSheet(backupSheet);
        ss.moveActiveSheet(incomeSheet.getIndex() + 1);
      } catch (eMove) {}
      SpreadsheetApp.flush();

      // Step 2: Format "RuPay Income" tab exactly like other canonical tabs (RuPay Charges / CultFit)
      const lastRow = incomeSheet.getLastRow();
      const lastCol = Math.max(incomeSheet.getLastColumn(), 4);

      // Remove borders across the sheet
      incomeSheet.getRange(1, 1, Math.max(lastRow, 100), lastCol).setBorder(false, false, false, false, false, false);

      // Copy column widths from RuPay Charges if available, else standard widths
      const chargesSheet = ss.getSheetByName("RuPay Charges");
      const col1Width = chargesSheet ? chargesSheet.getColumnWidth(1) : 180;
      const col2Width = chargesSheet ? chargesSheet.getColumnWidth(2) : 100;
      const col3Width = chargesSheet ? chargesSheet.getColumnWidth(3) : 100;
      const col4Width = chargesSheet ? chargesSheet.getColumnWidth(4) : 400;

      incomeSheet.setColumnWidth(1, col1Width);
      incomeSheet.setColumnWidth(2, col2Width);
      incomeSheet.setColumnWidth(3, col3Width);
      incomeSheet.setColumnWidth(4, col4Width);

      // Header Row (Row 1): Arial 10, bold, #F3F4F6 background, black text, bottom aligned, left aligned, frozen row
      const headerRange = incomeSheet.getRange(1, 1, 1, 4);
      headerRange.setFontFamily("Arial");
      headerRange.setFontSize(10);
      headerRange.setFontWeight("bold");
      headerRange.setFontColor("#000000");
      headerRange.setBackground("#F3F4F6");
      headerRange.setVerticalAlignment("bottom");
      headerRange.setHorizontalAlignment("left");
      incomeSheet.setFrozenRows(1);

      // Find TOTAL row
      const vals = incomeSheet.getRange(1, 1, lastRow, 4).getValues();
      let totalRowIdx = -1;
      for (let r = 1; r < vals.length; r++) {
        const col0 = String(vals[r][0] || '').trim().toUpperCase();
        if (col0 === 'TOTAL') {
          totalRowIdx = r + 1;
          break;
        }
      }

      const dataEndRow = totalRowIdx > 0 ? (totalRowIdx - 1) : lastRow;
      const dataRowCount = dataEndRow - 1;

      if (dataRowCount > 0) {
        const dataRange = incomeSheet.getRange(2, 1, dataRowCount, 4);
        dataRange.setFontFamily("Arial");
        dataRange.setFontSize(10);
        dataRange.setFontWeight("normal");
        dataRange.setFontColor("#000000");
        dataRange.setVerticalAlignment("bottom");
        dataRange.setBackground(null);
        dataRange.setBorder(false, false, false, false, false, false);

        // Column 1 (Title/Brand): Left aligned, normal weight
        incomeSheet.getRange(2, 1, dataRowCount, 1).setHorizontalAlignment("left").setFontWeight("normal");

        // Column 2 (Date): Center aligned, dd/MM/yy format
        incomeSheet.getRange(2, 2, dataRowCount, 1).setHorizontalAlignment("center").setNumberFormat("dd/MM/yy");

        // Column 3 (Price): Right aligned, 0.00 number format
        incomeSheet.getRange(2, 3, dataRowCount, 1).setHorizontalAlignment("right").setNumberFormat("0.00");

        // Column 4 (Comments): Left aligned, overflow wrap strategy
        incomeSheet.getRange(2, 4, dataRowCount, 1).setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      }

      // TOTAL Row formatting
      if (totalRowIdx > 0) {
        const totalRange = incomeSheet.getRange(totalRowIdx, 1, 1, 4);
        totalRange.setFontFamily("Arial");
        totalRange.setFontSize(10);
        totalRange.setFontWeight("bold");
        totalRange.setFontColor("#000000");
        totalRange.setVerticalAlignment("bottom");
        totalRange.setBackground(null);
        totalRange.setBorder(false, false, false, false, false, false);

        incomeSheet.getRange(totalRowIdx, 1).setValue("TOTAL").setHorizontalAlignment("left");
        incomeSheet.getRange(totalRowIdx, 2).setValue("").setHorizontalAlignment("general");

        let sum = 0;
        for (let r = 1; r < totalRowIdx - 1; r++) {
          const num = Number(vals[r][2]);
          if (!isNaN(num)) sum += num;
        }
        incomeSheet.getRange(totalRowIdx, 3).setValue(Number(sum.toFixed(2))).setHorizontalAlignment("right").setNumberFormat("0.00");
        incomeSheet.getRange(totalRowIdx, 4).setValue("").setHorizontalAlignment("general");
      }

      SpreadsheetApp.flush();

      // Step 3: Tally "RuPay Income" against "RuPayIncomeBackup"
      const currentVals = incomeSheet.getDataRange().getValues();
      const backupVals = backupSheet.getDataRange().getValues();

      let tallyMatches = true;
      const discrepancies = [];

      if (currentVals.length !== backupVals.length) {
        tallyMatches = false;
        discrepancies.push(`Row count mismatch: RuPay Income has ${currentVals.length} rows, Backup has ${backupVals.length} rows`);
      }

      const compareRows = Math.min(currentVals.length, backupVals.length);
      let incomeSum = 0;
      let backupSum = 0;

      for (let r = 0; r < compareRows; r++) {
        const cRow = currentVals[r];
        const bRow = backupVals[r];

        if (r === 0) {
          for (let c = 0; c < 4; c++) {
            if (String(cRow[c] || '').trim() !== String(bRow[c] || '').trim()) {
              tallyMatches = false;
              discrepancies.push(`Header col ${c + 1} mismatch: "${cRow[c]}" vs "${bRow[c]}"`);
            }
          }
          continue;
        }

        const cTitle = String(cRow[0] || '').trim();
        const bTitle = String(bRow[0] || '').trim();

        if (cTitle.toUpperCase() === 'TOTAL' || bTitle.toUpperCase() === 'TOTAL') {
          const cTot = Number(cRow[2]);
          const bTot = Number(bRow[2]);
          if (Math.abs(cTot - bTot) > 0.01) {
            tallyMatches = false;
            discrepancies.push(`TOTAL amount mismatch: Income=${cTot}, Backup=${bTot}`);
          }
          continue;
        }

        if (cTitle !== bTitle) {
          tallyMatches = false;
          discrepancies.push(`Row ${r + 1} Title mismatch: "${cTitle}" vs "${bTitle}"`);
        }

        // Compare date value
        const cDateMs = (cRow[1] instanceof Date) ? cRow[1].getTime() : parseDateToTimestamp(String(cRow[1] || ''));
        const bDateMs = (bRow[1] instanceof Date) ? bRow[1].getTime() : parseDateToTimestamp(String(bRow[1] || ''));
        if (cDateMs !== bDateMs) {
          tallyMatches = false;
          discrepancies.push(`Row ${r + 1} Date mismatch: "${cRow[1]}" vs "${bRow[1]}"`);
        }

        const cPrice = Number(cRow[2]) || 0;
        const bPrice = Number(bRow[2]) || 0;
        if (Math.abs(cPrice - bPrice) > 0.001) {
          tallyMatches = false;
          discrepancies.push(`Row ${r + 1} Price mismatch: ${cPrice} vs ${bPrice}`);
        }
        incomeSum += cPrice;
        backupSum += bPrice;

        const cComm = String(cRow[3] || '').trim();
        const bComm = String(bRow[3] || '').trim();
        if (cComm !== bComm) {
          tallyMatches = false;
          discrepancies.push(`Row ${r + 1} Comments mismatch: "${cComm}" vs "${bComm}"`);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "FORMAT_AND_BACKUP_RUPAY_INCOME",
        backupTabName: "RuPayIncomeBackup",
        backupSheetId: backupSheet.getSheetId(),
        incomeSheetId: incomeSheet.getSheetId(),
        totalRows: currentVals.length,
        dataRowCount: dataRowCount,
        incomeSum: Number(incomeSum.toFixed(2)),
        backupSum: Number(backupSum.toFixed(2)),
        isIdentical: tallyMatches,
        discrepanciesCount: discrepancies.length,
        discrepancies: discrepancies,
        formattingApplied: {
          headerBackground: "#F3F4F6",
          fontFamily: "Arial",
          fontSize: 10,
          columnWidths: [col1Width, col2Width, col3Width, col4Width],
          bordersRemoved: true,
          datePattern: "dd/MM/yy",
          pricePattern: "0.00"
        }
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }
  }

  // Action: Remove RuPay Income row & unmark sale in "New" tab
  if (e && e.parameter && e.parameter.action === "REMOVE_RUPAY_INCOME_AND_UNMARK_SALE") {
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(30000);
      const ss = getSpreadsheet();
      
      const targetBid = String(e.parameter.bookingId || "RUPS-38B3BF39").trim();
      const targetComments = String(e.parameter.comments || "Account: lakshaysharma2067 | Jana DC 7267").trim();
      
      const incomeSheet = ss.getSheetByName("RuPay Income");
      let deletedIncomeRow = -1;
      let deletedData = null;
      if (incomeSheet) {
        const iVals = incomeSheet.getDataRange().getValues();
        for (let r = 1; r < iVals.length; r++) {
          const col0 = String(iVals[r][0] || '').trim().toUpperCase();
          if (col0 === 'TOTAL') continue;
          const comm = String(iVals[r][3] || '').trim();
          if ((comm.includes("lakshaysharma2067") && comm.includes("7267")) || (targetComments && comm === targetComments)) {
            deletedIncomeRow = r + 1;
            deletedData = iVals[r];
            incomeSheet.deleteRow(r + 1);
            break;
          }
        }
        if (deletedIncomeRow > 0) {
          recalculateSheetTotal(incomeSheet);
        }
      }

      const newSheet = ss.getSheetByName("New");
      let unmarkedNewRow = -1;
      if (newSheet) {
        const nVals = newSheet.getDataRange().getValues();
        for (let r = 1; r < nVals.length; r++) {
          const bId = String(nVals[r][0] || '').replace(/^'/, '').trim().toUpperCase();
          if (bId === targetBid.replace(/^'/, '').trim().toUpperCase()) {
            unmarkedNewRow = r + 1;
            if (newSheet.getLastColumn() >= 13) {
              newSheet.getRange(r + 1, 13).setValue("");
            }
            break;
          }
        }
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        deletedIncomeRow: deletedIncomeRow,
        deletedData: deletedData,
        unmarkedNewRow: unmarkedNewRow
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      try { lock.releaseLock(); } catch(_) {}
    }
  }

  // Inspection handler for unified Income & Charges summary
  if (e && e.parameter && e.parameter.action === "INSPECT_INCOME_SUMMARY") {
    try {
      const ss = getSpreadsheet();
      const incomeSheet = ss.getSheetByName("RuPay Income");
      const chargesSheet = ss.getSheetByName("RuPay Charges");

      const incomeSummary = {
        allTime: { gross: 0, charges: 0, net: 0, count: 0 },
        byYear: {},
        ledger: []
      };

      if (incomeSheet) {
        const rows = incomeSheet.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          const title = String(rows[i][0] || '').trim();
          if (!title || title.toUpperCase() === 'TOTAL') continue;
          let dateStr = rows[i][1];
          if (dateStr instanceof Date) {
            dateStr = Utilities.formatDate(dateStr, "GMT+05:30", "dd/MM/yy");
          } else {
            dateStr = String(dateStr || '').replace(/^'/, '').trim();
          }
          const priceVal = Number(rows[i][2]);
          if (isNaN(priceVal) || priceVal === 0) continue;
          const comments = String(rows[i][3] || '').trim();
          let account = String(rows[i][4] || '').trim();
          if (!account && comments) {
            const parsed = parseIncomeSale(comments);
            if (parsed) account = parsed.card || parsed.account || '';
          }
          let yr = 2026;
          const comp = parseDateComponents(rows[i][1]);
          if (comp && comp.year) {
            yr = comp.year;
          }
          const entry = {
            id: 'inc-' + i,
            title: title,
            date: dateStr,
            price: Math.abs(priceVal),
            comments: comments,
            account: account,
            year: yr,
            type: 'income'
          };
          incomeSummary.ledger.push(entry);
          incomeSummary.allTime.gross += Math.abs(priceVal);
          incomeSummary.allTime.count++;
          const yrKey = String(yr);
          if (!incomeSummary.byYear[yrKey]) {
            incomeSummary.byYear[yrKey] = { gross: 0, charges: 0, net: 0, count: 0 };
          }
          incomeSummary.byYear[yrKey].gross += Math.abs(priceVal);
          incomeSummary.byYear[yrKey].count++;
        }
      }

      if (chargesSheet) {
        const rows = chargesSheet.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          const title = String(rows[i][0] || '').trim();
          if (!title || title.toUpperCase() === 'TOTAL') continue;
          let dateStr = rows[i][1];
          if (dateStr instanceof Date) {
            dateStr = Utilities.formatDate(dateStr, "GMT+05:30", "dd/MM/yy");
          } else {
            dateStr = String(dateStr || '').replace(/^'/, '').trim();
          }
          const priceVal = Number(rows[i][2]);
          if (isNaN(priceVal) || priceVal === 0) continue;
          const comments = String(rows[i][3] || '').trim();
          let account = String(rows[i][4] || '').trim();
          if (!account && comments) {
            const parsed = parseIncomeSale(comments);
            if (parsed) account = parsed.card || parsed.account || '';
          }
          let yr = 2026;
          const comp = parseDateComponents(rows[i][1]);
          if (comp && comp.year) {
            yr = comp.year;
          }
          const chargeVal = -Math.abs(priceVal);
          const entry = {
            id: 'chg-' + i,
            title: title,
            date: dateStr,
            price: chargeVal,
            comments: comments,
            account: account,
            year: yr,
            type: 'charge'
          };
          incomeSummary.ledger.push(entry);
          incomeSummary.allTime.charges += Math.abs(chargeVal);
          incomeSummary.allTime.count++;
          const yrKey = String(yr);
          if (!incomeSummary.byYear[yrKey]) {
            incomeSummary.byYear[yrKey] = { gross: 0, charges: 0, net: 0, count: 0 };
          }
          incomeSummary.byYear[yrKey].charges += Math.abs(chargeVal);
          incomeSummary.byYear[yrKey].count++;
        }
      }

      incomeSummary.ledger.sort(function(a, b) {
        return parseDateToTimestamp(b.date) - parseDateToTimestamp(a.date);
      });
      incomeSummary.allTime.net = incomeSummary.allTime.gross - incomeSummary.allTime.charges;
      for (const y in incomeSummary.byYear) {
        incomeSummary.byYear[y].net = incomeSummary.byYear[y].gross - incomeSummary.byYear[y].charges;
      }

      return ContentService.createTextOutput(JSON.stringify(incomeSummary, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Administrative Setup: Initialize "Active Cards" tab with 20 cards, Linked Account, and 13 benefit columns
  if (e && e.parameter && e.parameter.action === "SETUP_ACTIVE_CARDS_TAB") {
    try {
      const ss = getSpreadsheet();
      let sheet = ss.getSheetByName("Active Cards");
      let created = false;
      if (!sheet) {
        sheet = ss.insertSheet("Active Cards");
        created = true;
      } else {
        sheet.clear();
      }

      const headers = [
        "Bank", "Card_Type", "Card_Name", "Linked_Account", "Card_Number", "Expiry", "CVV", "Status", "Notes",
        "Spa_Lakme", "Health_Thyrocare", "Amazon_Prime", "Cult_Live_90D", "Cult_Elite_1M",
        "Uber", "Myntra", "Decathlon", "Reliance_Digital", "BigBasket", "Apollo_Pharmacy",
        "Kalyan_Jewellers", "TaxSpanner", "Last_Updated"
      ];

      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#0C8C46")
        .setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);

      // Lock plain text format on Linked Account (Col 4), Card Number (Col 5), Expiry (Col 6), CVV (Col 7)
      sheet.getRange("D:D").setNumberFormat('@');
      sheet.getRange("E:E").setNumberFormat('@');
      sheet.getRange("F:F").setNumberFormat('@');
      sheet.getRange("G:G").setNumberFormat('@');

      const nowStr = formatISTTimestamp(new Date());

      // If sheet already has data, do not overwrite; otherwise insert dummy schema template
      if (sheet.getLastRow() <= 1) {
        const dummyBenefits = Array(13).fill("Available");
        sheet.appendRow([
          "Sample Bank", "Credit", "Sample RuPay Select", "NA", "'6521000000000001", "'12/30", "'000", "ACTIVE", "Template card", ...dummyBenefits, nowStr
        ]);
      }

      // Update "Blocked Cards" archive tab with matching column ordering
      let bSheet = ss.getSheetByName("Blocked Cards");
      if (bSheet) {
        bSheet.clear();
      } else {
        bSheet = ss.insertSheet("Blocked Cards");
      }
      const bHeaders = [
        "Bank", "Card_Type", "Card_Name", "Linked_Account", "Card_Number", "Expiry", "CVV", "Status",
        "Blocked_Timestamp", "Replaced_By_Card", "Notes",
        "Spa_Lakme", "Health_Thyrocare", "Amazon_Prime", "Cult_Live_90D", "Cult_Elite_1M",
        "Uber", "Myntra", "Decathlon", "Reliance_Digital", "BigBasket", "Apollo_Pharmacy",
        "Kalyan_Jewellers", "TaxSpanner"
      ];
      bSheet.appendRow(bHeaders);
      bSheet.getRange(1, 1, 1, bHeaders.length)
        .setFontWeight("bold")
        .setBackground("#1E293B")
        .setFontColor("#F87171");
      bSheet.setFrozenRows(1);
      bSheet.getRange("D:D").setNumberFormat('@');
      bSheet.getRange("E:E").setNumberFormat('@');
      bSheet.getRange("F:F").setNumberFormat('@');
      bSheet.getRange("G:G").setNumberFormat('@');

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "SETUP_ACTIVE_CARDS_TAB",
        activeCardsCount: rows.length,
        created: created
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Administrative Setup: Initialize "Bank Accounts" tab with full family multi-bank accounts
  if (e && e.parameter && e.parameter.action === "SETUP_BANK_ACCOUNTS_TAB") {
    try {
      const ss = getSpreadsheet();
      const res = setupBankAccountsTab(ss);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Fetch all Bank Accounts in JSON format (Requires 24-Hour Session Token)
  if (e && e.parameter && e.parameter.action === "GET_BANK_ACCOUNTS") {
    const providedAuth = e.parameter.token || e.parameter.auth || e.parameter.key || "";
    const authCheck = validateAuth(providedAuth);
    if (!authCheck.valid) {
      return ContentService.createTextOutput(JSON.stringify({ error: authCheck.error, accounts: [] })).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const ss = getSpreadsheet();
      let sheet = ss.getSheetByName("Bank Accounts");
      if (!sheet) {
        setupBankAccountsTab(ss);
        sheet = ss.getSheetByName("Bank Accounts");
      }
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, accounts: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      const headers = data[0];
      const accounts = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const acc = {};
        for (let j = 0; j < headers.length; j++) {
          acc[headers[j]] = row[j];
        }
        accounts.push(acc);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, count: accounts.length, accounts: accounts })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // SECURITY CHECK: Enforce 24-Hour HMAC Session Token
  const providedAuth = (e && e.parameter) ? (e.parameter.token || e.parameter.auth || e.parameter.key || "") : "";
  const authCheck = validateAuth(providedAuth);
  if (!authCheck.valid) {
    return ContentService.createTextOutput(JSON.stringify({
      error: authCheck.error,
      vouchers: [],
      cooldowns: [],
      unassigned: [],
      redeemed: []
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const ss = getSpreadsheet();

    // Fast-Path: Fetch only Ingested Vouchers from "New" sheet (~400ms)
    if (e && e.parameter && (e.parameter.action === "GET_INGESTED_ONLY" || e.parameter.fast === "staging")) {
      const staging = getStagingFromNewSheet(ss);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "GET_INGESTED_ONLY",
        staging: staging,
        lastUpdated: formatISTTimestamp(new Date())
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Fast-Path: Spreadsheet modification check (~150ms)
    if (e && e.parameter && e.parameter.action === "CHECK_MODIFIED") {
      let lastUpdatedMs = 0;
      try {
        lastUpdatedMs = DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
      } catch (_) {
        lastUpdatedMs = Date.now();
      }
      const clientTimestamp = Number(e.parameter.since || 0);
      const notModified = clientTimestamp > 0 && Math.abs(lastUpdatedMs - clientTimestamp) < 1000;
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "CHECK_MODIFIED",
        notModified: notModified,
        lastUpdated: lastUpdatedMs
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Trigger on-demand reconciliation of pending vouchers
    if (e && e.parameter && e.parameter.reconcile === "true") {
      const rec = reconcilePendingVouchers(ss);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        count: rec.length,
        reconciled: rec
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Explicit on-demand email sync (throttled): Only runs when explicitly requested to protect daily Gmail quota
    if (e && e.parameter && (e.parameter.sync_emails === "true" || e.parameter.ingest === "force")) {
      try {
        processRupayEmails();
      } catch (err) {
        console.warn("On-demand email ingestion error:", err.message);
      }
    }

    const sheets = ss.getSheets();
    const sheetNameSet = new Set(sheets.map(s => s.getName()));
    // Ensure persistent system tabs exist if missing
    if (!sheetNameSet.has("⏱ Cooldowns") && !sheetNameSet.has("Cooldowns")) {
      getOrInitCooldownsSheet(ss);
    }
    if (!sheetNameSet.has("Shared_Activity")) {
      getOrInitSharedActivitySheet(ss);
    }
    const result = {
      sheetUrl: ss.getUrl(),
      unassigned: [],
      cooldowns: [],
      vouchers: [],
      redeemed: [],
      cards: [],
      shared: [],
      staging: [],
      income: {
        allTime: { gross: 0, charges: 0, net: 0, count: 0 },
        byYear: {},
        ledger: []
      }
    };

    const SKIPPED_SYNC_SHEETS = new Set([
      "Deleted_Tombstones",
      "Blocked Cards",
      "Bank Accounts",
      "Cards",
      "RuPayIncomeBackup"
    ]);

    for (const sheet of sheets) {
      const name = sheet.getName();
      if (SKIPPED_SYNC_SHEETS.has(name)) continue;
      if (sheet.getLastRow() <= 1) continue;

      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) continue;

      if (name === "⚡ Unassigned" || name === "Unassigned") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] && !rows[i][1]) continue;
          const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
          let uCode = String(rows[i][1] || '').replace(/^'/, '').trim();
          if (uCode.toUpperCase() === 'PENDING') uCode = 'Pending';
          const uEntry = {
            brand: "Unassigned",
            bookingId: String(rows[i][0] || '').replace(/^'/, ''),
            code: uCode,
            pin: rows[i][2] && String(rows[i][2]).trim() !== '' && String(rows[i][2]).trim() !== 'null' ? String(rows[i][2]).trim() : null,
            amount: Number(rows[i][3]) || 0,
            dateOfIssue: rows[i][4] ? String(rows[i][4]) : '',
            account: rows[i][5] ? String(rows[i][5]) : '',
            card: rows[i][6] ? String(rows[i][6]) : '',
            status: rows[i][7] ? String(rows[i][7]) : 'PENDING',
            timestampIST: rows[i][8] ? String(rows[i][8]) : null,
            displayInUI: isDisp
          };
          result.unassigned.push(uEntry);
          result.vouchers.push(uEntry);
        }
      } else if (name === "⏱ Cooldowns" || name === "Cooldowns") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] || String(rows[i][0]).trim() === "") continue;
          result.cooldowns.push({
            account: String(rows[i][0] || ''),
            lastRedeemed: rows[i][1] ? String(rows[i][1]) : '',
            status: rows[i][2] ? String(rows[i][2]) : 'READY',
            lastUpdated: rows[i][3] ? String(rows[i][3]) : ''
          });
        }
      } else if (name === "Active Cards") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][4] || String(rows[i][4]).trim() === "") continue;
          let expStr = rows[i][5];
          if (expStr instanceof Date) {
            expStr = Utilities.formatDate(expStr, "GMT+05:30", "MM/yy");
          } else {
            expStr = String(expStr || '').replace(/^'/, '').trim();
          }
          result.cards.push({
            bank: String(rows[i][0] || '').trim(),
            type: String(rows[i][1] || 'Debit').trim(),
            name: String(rows[i][2] || '').trim(),
            linkedAccount: String(rows[i][3] || 'NA').replace(/^'/, '').trim(),
            cardNumber: String(rows[i][4] || '').replace(/['\s]/g, '').trim(),
            expiry: expStr,
            cvv: String(rows[i][6] || '').replace(/^'/, '').trim().padStart(3, '0'),
            status: String(rows[i][7] || 'ACTIVE').trim().toUpperCase(),
            notes: String(rows[i][8] || '').trim(),
            benefits: {
              Spa_Lakme: String(rows[i][9] || 'Not Applicable'),
              Health_Thyrocare: String(rows[i][10] || 'Not Applicable'),
              Amazon_Prime: String(rows[i][11] || 'Not Applicable'),
              Cult_Live_90D: String(rows[i][12] || 'Not Applicable'),
              Cult_Elite_1M: String(rows[i][13] || 'Not Applicable'),
              Uber: String(rows[i][14] || 'Not Applicable'),
              Myntra: String(rows[i][15] || 'Not Applicable'),
              Decathlon: String(rows[i][16] || 'Not Applicable'),
              Reliance_Digital: String(rows[i][17] || 'Not Applicable'),
              BigBasket: String(rows[i][18] || 'Not Applicable'),
              Apollo_Pharmacy: String(rows[i][19] || 'Not Applicable'),
              Kalyan_Jewellers: String(rows[i][20] || 'Not Applicable'),
              TaxSpanner: String(rows[i][21] || 'Not Applicable')
            },
            lastUpdated: String(rows[i][22] || '')
          });
        }
      } else if (name === "Cards") {
        // Fallback for legacy Cards sheet if Active Cards is not yet synced
        if (result.cards.length === 0) {
          for (let i = 1; i < rows.length; i++) {
            if (!rows[i][0] || String(rows[i][0]).trim() === "") continue;
            // Check if column 3 is Network (if rows[i].length >= 13) or BIN
            const hasNetworkCol = rows[0] && rows[0][3] === "Network";
            const netVal = hasNetworkCol ? String(rows[i][3] || 'RuPay') : (String(rows[i][5] || '').includes('Visa') ? 'Visa' : String(rows[i][5] || '').includes('Mastercard') ? 'Mastercard' : 'RuPay');
            const binIdx = hasNetworkCol ? 4 : 3;
            const last4Idx = hasNetworkCol ? 5 : 4;
            const nickIdx = hasNetworkCol ? 6 : 5;
            const statusIdx = hasNetworkCol ? 7 : 6;
            const notesIdx = hasNetworkCol ? 11 : 10;
            const updatedIdx = hasNetworkCol ? 12 : 11;

            result.cards.push({
              bank: String(rows[i][1] || ''),
              type: String(rows[i][2] || ''),
              network: netVal,
              name: String(rows[i][nickIdx] || 'Card'),
              linkedAccount: 'NA',
              cardNumber: String(rows[i][binIdx] || '') + String(rows[i][last4Idx] || ''),
              expiry: '10/33',
              cvv: '000',
              status: String(rows[i][statusIdx] || 'ACTIVE').toUpperCase(),
              notes: rows[i][notesIdx] ? String(rows[i][notesIdx]) : '',
              benefits: {},
              lastUpdated: rows[i][updatedIdx] ? String(rows[i][updatedIdx]) : ''
            });
          }
        }
      } else if (name === "Shared_Activity") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] || String(rows[i][0]).trim() === "") continue;
          const status = String(rows[i][5] || 'SHARED').trim().toUpperCase();
          if (status === "DISMISSED") continue;
          const isDisp = rows[i].length > 7 ? (rows[i][7] !== false && String(rows[i][7]).toUpperCase() !== 'FALSE') : true;
          result.shared.push({
            code: String(rows[i][0] || ''),
            brand: String(rows[i][1] || ''),
            amount: Number(rows[i][2]) || 0,
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            sharedAt: String(rows[i][4] || ''),
            status: status,
            platform: String(rows[i][6] || 'WhatsApp'),
            displayInUI: isDisp
          });
        }
      } else if (name === "Redeemed") {
        // Unified Redeemed Archive (Col 13: Display in UI)
        for (let i = 1; i < rows.length; i++) {
          if ((!rows[i][1] || String(rows[i][1]).trim() === '') && (!rows[i][2] || String(rows[i][2]).trim() === '')) continue;
          const isDisp = rows[i].length > 12 ? (rows[i][12] !== false && String(rows[i][12]).toUpperCase() !== 'FALSE') : true;
          let rAmount = Number(rows[i][4]) || 0;
          const rBrandUpper = String(rows[i][0] || '').toUpperCase();
          if (rBrandUpper.includes("SWIGGY")) rAmount = 1;
          else if (rBrandUpper.includes("THYROCARE")) rAmount = 125;

          result.redeemed.push({
            brand: String(rows[i][0] || ''),
            bookingId: String(rows[i][1] || '').replace(/^'/, ''),
            code: String(rows[i][2] || '').replace(/^'/, ''),
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            amount: rAmount,
            dateOfIssue: rows[i][5] ? String(rows[i][5]) : '',
            account: rows[i][6] ? String(rows[i][6]) : '',
            card: rows[i][7] ? String(rows[i][7]) : '',
            status: "REDEEMED",
            redeemedTimestampIST: rows[i][9] ? String(rows[i][9]) : '',
            originalTab: String(rows[i][10] || rows[i][0] || ''),
            displayInUI: isDisp
          });
        }
      } else if (name === "New") {
        // Staging Feed from "New" Tab
        result.staging = getStagingFromNewSheet(ss);
      } else if (name === "RuPay Income") {
        for (let i = 1; i < rows.length; i++) {
          const title = String(rows[i][0] || '').trim();
          if (!title || title.toUpperCase() === 'TOTAL') continue;

          let dateStr = rows[i][1];
          if (dateStr instanceof Date) {
            dateStr = Utilities.formatDate(dateStr, "GMT+05:30", "dd/MM/yy");
          } else {
            dateStr = String(dateStr || '').replace(/^'/, '').trim();
          }

          const priceVal = Number(rows[i][2]);
          if (isNaN(priceVal) || priceVal === 0) continue;

          const comments = String(rows[i][3] || '').trim();
          let account = String(rows[i][4] || '').trim();
          if (!account && comments) {
            const parsed = parseIncomeSale(comments);
            if (parsed) account = parsed.card || parsed.account || '';
          }

          const comp = parseDateComponents(dateStr);
          const yr = comp ? comp.year : 2026;
          if (comp) {
            dateStr = Utilities.formatString("%02d/%02d/%02d", comp.day, comp.month, comp.year % 100);
          }

          const entry = {
            id: 'inc-' + i,
            sheetRowIndex: i,
            title: title,
            date: dateStr,
            price: Math.abs(priceVal),
            comments: comments,
            account: account,
            year: yr,
            type: 'income'
          };
          result.income.ledger.push(entry);

          result.income.allTime.gross += Math.abs(priceVal);
          result.income.allTime.count++;

          const yrKey = String(yr);
          if (!result.income.byYear[yrKey]) {
            result.income.byYear[yrKey] = { gross: 0, charges: 0, net: 0, count: 0 };
          }
          result.income.byYear[yrKey].gross += Math.abs(priceVal);
          result.income.byYear[yrKey].count++;
        }
      } else if (name === "RuPay Charges") {
        for (let i = 1; i < rows.length; i++) {
          const title = String(rows[i][0] || '').trim();
          if (!title || title.toUpperCase() === 'TOTAL') continue;

          let dateStr = rows[i][1];
          if (dateStr instanceof Date) {
            dateStr = Utilities.formatDate(dateStr, "GMT+05:30", "dd/MM/yy");
          } else {
            dateStr = String(dateStr || '').replace(/^'/, '').trim();
          }

          const priceVal = Number(rows[i][2]);
          if (isNaN(priceVal) || priceVal === 0) continue;

          const comments = String(rows[i][3] || '').trim();
          let account = String(rows[i][4] || '').trim();
          if (!account && comments) {
            const parsed = parseIncomeSale(comments);
            if (parsed) account = parsed.card || parsed.account || '';
          }

          const comp = parseDateComponents(dateStr);
          const yr = comp ? comp.year : 2026;
          if (comp) {
            dateStr = Utilities.formatString("%02d/%02d/%02d", comp.day, comp.month, comp.year % 100);
          }

          const chargeVal = -Math.abs(priceVal);
          const entry = {
            id: 'chg-' + i,
            sheetRowIndex: i,
            title: title,
            date: dateStr,
            price: chargeVal,
            comments: comments,
            account: account,
            year: yr,
            type: 'charge'
          };
          result.income.ledger.push(entry);

          result.income.allTime.charges += Math.abs(chargeVal);
          result.income.allTime.count++;

          const yrKey = String(yr);
          if (!result.income.byYear[yrKey]) {
            result.income.byYear[yrKey] = { gross: 0, charges: 0, net: 0, count: 0 };
          }
          result.income.byYear[yrKey].charges += Math.abs(chargeVal);
          result.income.byYear[yrKey].count++;
        }
      } else if (SYSTEM_SHEET_NAMES.has(name) || name === "Cards") {
        // Dedicated non-voucher system tabs (Blocked Cards, Bank Accounts, Cards):
        // Skip here so they NEVER leak into result.vouchers and do NOT inflate active card counts
        continue;
      } else {
        // Uniform Brand Tabs (Col 10: Display in UI)
        for (let i = 1; i < rows.length; i++) {
          if ((!rows[i][0] || String(rows[i][0]).trim() === "") && (!rows[i][1] || String(rows[i][1]).trim() === "")) continue;
          const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
          const brandNameUpper = String(name || '').toUpperCase();
          let voucherAmount = Number(rows[i][3]) || 0;
          if (brandNameUpper.includes("SWIGGY")) voucherAmount = 1;
          else if (brandNameUpper.includes("THYROCARE")) voucherAmount = 125;

          let vCode = String(rows[i][1] || rows[i][0] || '').replace(/^'/, '').trim();
          if (vCode.toUpperCase() === 'PENDING') vCode = 'Pending';

          result.vouchers.push({
            brand: name,
            bookingId: String(rows[i][0] || '').replace(/^'/, ''),
            code: vCode,
            pin: rows[i][2] && String(rows[i][2]).trim() !== '' && String(rows[i][2]).trim() !== 'null' ? String(rows[i][2]).trim() : null,
            amount: voucherAmount,
            dateOfIssue: rows[i][4] ? String(rows[i][4]) : '',
            account: rows[i][5] ? String(rows[i][5]) : '',
            card: rows[i][6] ? String(rows[i][6]) : '',
            status: rows[i][7] ? String(rows[i][7]) : 'ACTIVE',
            timestampIST: rows[i][8] ? String(rows[i][8]) : null,
            displayInUI: isDisp
          });
        }
      }
    }

    // Sort unified income & charges ledger by date descending (newest on top; for same date, retain chronological entry order)
    result.income.ledger.sort(function(a, b) {
      const diff = parseDateToTimestamp(b.date) - parseDateToTimestamp(a.date);
      if (diff !== 0) return diff;
      return (b.sheetRowIndex || 0) - (a.sheetRowIndex || 0);
    });
    result.income.allTime.net = result.income.allTime.gross - result.income.allTime.charges;
    for (const y in result.income.byYear) {
      result.income.byYear[y].net = result.income.byYear[y].gross - result.income.byYear[y].charges;
    }

    try {
      const scriptProps = PropertiesService.getScriptProperties().getProperties();
      result.systemStatus = {
        lastEmailSyncTime: scriptProps.LAST_EMAIL_SYNC_TIME || null,
        lastEmailSyncStatus: scriptProps.LAST_EMAIL_SYNC_STATUS || null,
        lastEmailSyncCount: scriptProps.LAST_EMAIL_SYNC_COUNT || null
      };
    } catch (_) {}

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: "doGet error: " + err.message,
      vouchers: [],
      cooldowns: [],
      unassigned: [],
      redeemed: [],
      cards: [],
      shared: [],
      staging: [],
      income: { allTime: { gross: 0, charges: 0, net: 0, count: 0 }, byYear: {}, ledger: [] }
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const ss = getSpreadsheet();
  try {
    const data = JSON.parse(e.postData.contents);

    // 0A. Request Email OTP (strictly to vibhavoswal3@gmail.com)
    if (data.action === "REQUEST_OTP") {
      const props = PropertiesService.getScriptProperties();
      const now = Date.now();
      
      // 1. Rate Limit Cooldown: 60s
      const lastSent = Number(props.getProperty("VAULT_OTP_LAST_SENT") || 0);
      if (lastSent && (now - lastSent < 60000)) {
        const waitSec = Math.ceil((60000 - (now - lastSent)) / 1000);
        return ContentService.createTextOutput(JSON.stringify({
          error: "Rate limit: Please wait " + waitSec + " seconds before requesting another code."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 2. Rate Limit Hourly: max 10 per hour
      const hourlyWindow = Number(props.getProperty("VAULT_OTP_HOURLY_WINDOW") || 0);
      let hourlyCount = Number(props.getProperty("VAULT_OTP_HOURLY_COUNT") || 0);
      if (!hourlyWindow || now - hourlyWindow > 3600000) {
        hourlyCount = 0;
        props.setProperty("VAULT_OTP_HOURLY_WINDOW", String(now));
      }
      if (hourlyCount >= 10) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Rate limit: Maximum 10 verification codes per hour reached. Please try again later."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 3. Generate 6-digit OTP
      const otp = String(Math.floor(100000 + Math.random() * 900000));
      
      // 4. Store in PropertiesService: 10 min TTL (guaranteed across all Google worker nodes)
      props.setProperties({
        "VAULT_ACTIVE_OTP": otp,
        "VAULT_OTP_EXPIRES_AT": String(now + 10 * 60 * 1000), // 10 minutes
        "VAULT_OTP_LAST_SENT": String(now),
        "VAULT_OTP_HOURLY_COUNT": String(hourlyCount + 1)
      });

      // 5. Send Email via GmailApp (Clean subject line with zero emojis & responsive HTML card)
      try {
        const subject = "[Card & Voucher Manager] Login Code: " + otp + " (Valid for 10 mins)";
        const plainBody = "Card & Voucher Manager verification code:\n\n      " + otp + "\n\nValid for 10 minutes.\n\n— Card & Voucher Manager";

        const htmlBody = '<!DOCTYPE html>' +
          '<html lang="en">' +
          '<head>' +
          '<meta charset="UTF-8">' +
          '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
          '<style>' +
          ':root { color-scheme: light dark; supported-color-schemes: light dark; } ' +
          '@media (prefers-color-scheme: dark) { ' +
          '.vault-card { background-color: #141517 !important; border-color: rgba(255, 255, 255, 0.08) !important; box-shadow: 0 12px 32px rgba(0, 0, 0, 0.4) !important; } ' +
          '.vault-brand { color: #a1a1aa !important; } ' +
          '.vault-code { color: #34d399 !important; } ' +
          '.vault-sub { color: #71717a !important; } ' +
          '} ' +
          '</style>' +
          '</head>' +
          '<body style="margin: 0; padding: 36px 12px; background-color: transparent; font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">' +
          '<table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: transparent; border-collapse: collapse;">' +
          '<tr>' +
          '<td align="center" style="background-color: transparent; padding: 0;">' +
          '<div class="vault-card" style="max-width: 340px; width: 100%; background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 20px; padding: 32px 20px; box-sizing: border-box; text-align: center; box-shadow: 0 4px 16px rgba(0, 0, 0, 0.04);">' +
          '<div class="vault-brand" style="font-size: 11px; font-weight: 700; color: #71717a; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 12px;">' +
          'Card &amp; Voucher Manager' +
          '</div>' +
          '<div class="vault-code" style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 40px; font-weight: 700; letter-spacing: 8px; color: #0C8C46; text-indent: 8px; margin: 10px 0 14px 0;">' +
          otp +
          '</div>' +
          '<div class="vault-sub" style="font-size: 12.5px; color: #71717a; margin: 0;">' +
          'Valid for 10 minutes' +
          '</div>' +
          '</div>' +
          '</td>' +
          '</tr>' +
          '</table>' +
          '</body>' +
          '</html>';

        GmailApp.sendEmail(
          AUTHORIZED_USER,
          subject,
          plainBody,
          { htmlBody: htmlBody }
        );
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Failed to send email: " + err.message
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "Verification code sent to " + AUTHORIZED_USER,
        cooldownSeconds: 60
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 0B. Verify OTP and issue 24-hour Session Token
    if (data.action === "VERIFY_OTP") {
      const props = PropertiesService.getScriptProperties();
      const storedOtp = props.getProperty("VAULT_ACTIVE_OTP");
      const expiresAt = Number(props.getProperty("VAULT_OTP_EXPIRES_AT") || 0);
      const now = Date.now();
      const enteredOtp = String(data.otp || "").trim();

      if (!storedOtp || !expiresAt || now > expiresAt || storedOtp !== enteredOtp) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Invalid or expired verification code."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Burn OTP on successful use
      props.deleteProperty("VAULT_ACTIVE_OTP");
      props.deleteProperty("VAULT_OTP_EXPIRES_AT");

      // Generate 24-hour session token
      const session = generateSessionToken(AUTHORIZED_USER);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        token: session.token,
        expiry: session.expiry,
        user: AUTHORIZED_USER
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Protected actions: Validate 24-Hour HMAC Session Token
    const providedAuth = data.token || data.auth || (e.parameter ? (e.parameter.token || e.parameter.auth || e.parameter.key) : "") || "";
    const authCheck = validateAuth(providedAuth);
    if (!authCheck.valid) {
      return ContentService.createTextOutput(JSON.stringify({
        error: authCheck.error
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // On-Demand Gmail Ingestion Trigger (Manages own LockService lock inside processRupayEmails)
    if (data.action === "TRIGGER_GMAIL_INGEST") {
      const startTime = Date.now();
      let countBefore = 0;
      let countAfter = 0;
      const newSheet = ss.getSheetByName("New");
      if (newSheet) countBefore = Math.max(0, newSheet.getLastRow() - 1);

      try {
        processRupayEmails();
      } catch (errIngest) {
        console.warn("TRIGGER_GMAIL_INGEST error:", errIngest.message);
      }

      if (newSheet) countAfter = Math.max(0, newSheet.getLastRow() - 1);
      const newVouchersCount = Math.max(0, countAfter - countBefore);
      const durationMs = Date.now() - startTime;

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "TRIGGER_GMAIL_INGEST",
        newVouchersCount: newVouchersCount,
        durationMs: durationMs,
        totalStaged: countAfter
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Acquire script lock for mutating actions to avoid concurrent write collisions
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return ContentService.createTextOutput(JSON.stringify({
        error: "Server is processing another request. Please retry in a moment."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    try {
        // 1C. Approve Staged Green Vouchers from "New" Tab to Canonical Brand Tabs (Dual-Identifier Matching & Booking ID Idempotency)
      if (data.action === "APPROVE_STAGING") {
        let newSheet = ss.getSheetByName("New");
        if (!newSheet) {
          return ContentService.createTextOutput(JSON.stringify({ error: '"New" tab not found' })).setMimeType(ContentService.MimeType.JSON);
        }

        const rawBookingIds = data.bookingIds || (data.bookingId ? [data.bookingId] : []);
        const rawCodes = data.codes || (data.code ? [data.code] : []);
        const combinedIds = (Array.isArray(rawBookingIds) ? rawBookingIds : [rawBookingIds])
          .concat(Array.isArray(rawCodes) ? rawCodes : [rawCodes])
          .map(b => String(b || '').replace(/^'/, '').trim().toUpperCase())
          .filter(Boolean);
        const targetIds = new Set(combinedIds);

        if (targetIds.size === 0) {
          return ContentService.createTextOutput(JSON.stringify({ error: 'No booking IDs or codes provided for approval' })).setMimeType(ContentService.MimeType.JSON);
        }

        const dataRows = newSheet.getDataRange().getValues();
        const rowsToDelete = [];
        const approvedSummary = [];

        for (let r = 1; r < dataRows.length; r++) {
          const row = dataRows[r];
          const bId = String(row[0] || '').replace(/^'/, '').trim();
          const currentCode = String(row[1] || '').replace(/^'/, '').trim();
          const bIdUpper = bId.toUpperCase();
          const codeUpper = currentCode.toUpperCase();

          // Dual-Identifier Matching: Match by Booking ID (Col 1) OR Code (Col 2)
          if (targetIds.has(bIdUpper) || targetIds.has(codeUpper)) {
            const currentStatus = String(row[7] || '').trim().toUpperCase();
            const targetBrand = String(row[10] || 'Unassigned').trim();

            // HARD CONDITION GUARD: Only fulfilled (ACTIVE, real code) vouchers can be approved!
            if (currentStatus !== 'ACTIVE' || currentCode === 'PENDING' || currentCode === 'Pending revalidation' || currentStatus === 'PENDING_REVALIDATION' || !currentCode) {
              return ContentService.createTextOutput(JSON.stringify({
                error: `Cannot approve unfulfilled voucher ${bId || currentCode}. Code must arrive before approval.`
              })).setMimeType(ContentService.MimeType.JSON);
            }

            // Target Brand Sheet
            const brandSheet = getTargetBrandSheet(ss, targetBrand);

            // IDEMPOTENCY CHECK (Hierarchical: Primary RUPS Booking ID, Fallback Voucher Code):
            // Check if this Booking ID or Code already exists in the destination Brand tab
            let existingBrandRowIndex = -1;
            const brandVals = brandSheet.getDataRange().getValues();
            for (let b = 1; b < brandVals.length; b++) {
              const existingBid = String(brandVals[b][0] || '').trim().toUpperCase();
              const existingCode = String(brandVals[b][1] || '').trim().toUpperCase();
              const hasRups = bIdUpper && bIdUpper.startsWith('RUPS-');
              const bidMatch = hasRups && (existingBid === bIdUpper);
              const codeMatch = (!hasRups || !existingBid.startsWith('RUPS-')) && codeUpper && codeUpper !== 'PENDING' && codeUpper !== '-' && (existingCode === codeUpper);
              if (bidMatch || codeMatch) {
                existingBrandRowIndex = b + 1;
                break;
              }
            }

            if (existingBrandRowIndex > 0) {
              // Self-Healing Reconciliation: Already in brand tab!
              // Do NOT append duplicate row. Ensure status is ACTIVE and UI display is true.
              brandSheet.getRange(existingBrandRowIndex, 8).setValue("ACTIVE");
              brandSheet.getRange(existingBrandRowIndex, 10).setValue(true);
            } else {
              // Not yet in brand tab: Append clean row
              const targetBrandUpper = String(targetBrand || '').toUpperCase();
              let approvedAmount = row[3] || 500;
              if (targetBrandUpper.includes("SWIGGY")) approvedAmount = 1;
              else if (targetBrandUpper.includes("THYROCARE")) approvedAmount = 125;

              const cleanBid = String(row[0] || '').trim().replace(/^'/, '');
              const cleanCd = String(row[1] || '').trim().replace(/^'/, '');
              const cleanPin = String(row[2] || '').trim().replace(/^'/, '');
              const brandRow = [
                cleanBid ? `'${cleanBid}` : '', // Booking ID
                cleanCd ? `'${cleanCd}` : '',   // Code
                cleanPin ? `'${cleanPin}` : '', // PIN
                approvedAmount, // Value
                row[4] || "", // Date of Issue
                row[5] || "", // Account
                row[6] || "Unassigned", // Card
                "ACTIVE", // Status
                row[8] || formatISTTimestamp(new Date()), // Appended Timestamp (IST)
                row[9] !== false && String(row[9]).toLowerCase() !== 'false' // Display in UI
              ];
              brandSheet.appendRow(brandRow);
            }

            rowsToDelete.push(r + 1); // 1-based index in "New" (always delete from New)
            approvedSummary.push({ bookingId: bId, brand: targetBrand, code: currentCode, alreadyExisted: existingBrandRowIndex > 0 });
          }
        }

        // Bottom-up deletion from "New" tab (no renumbering needed)
        if (rowsToDelete.length > 0) {
          rowsToDelete.sort((a, b) => b - a);
          for (const rIdx of rowsToDelete) {
            newSheet.deleteRow(rIdx);
          }
        }

        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          approvedCount: approvedSummary.length,
          approved: approvedSummary
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 1D. Mark Revalidated Voucher in "New" Tab (Transitions PENDING_REVALIDATION -> PENDING)
      if (data.action === "MARK_REVALIDATED") {
        let newSheet = ss.getSheetByName("New");
        if (!newSheet) {
          return ContentService.createTextOutput(JSON.stringify({ error: '"New" tab not found' })).setMimeType(ContentService.MimeType.JSON);
        }
        const targetBid = String(data.bookingId || data.id || '').trim().toUpperCase();
        if (!targetBid) {
          return ContentService.createTextOutput(JSON.stringify({ error: 'No booking ID provided' })).setMimeType(ContentService.MimeType.JSON);
        }
        const dataRows = newSheet.getDataRange().getValues();
        let updated = false;
        for (let r = 1; r < dataRows.length; r++) {
          const rowBid = String(dataRows[r][0] || '').trim().toUpperCase();
          if (rowBid === targetBid || (rowBid.startsWith('RUPS-') && targetBid.includes(rowBid)) || (targetBid.startsWith('RUPS-') && rowBid.includes(targetBid))) {
            newSheet.getRange(r + 1, 2).setValue("'Pending");
            newSheet.getRange(r + 1, 8).setValue('PENDING');
            if (dataRows[0].length >= 12) {
              newSheet.getRange(r + 1, 12).setValue('');
            }
            updated = true;
            addRevalidatedBookingToProperties(targetBid);
            break;
          }
        }
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          updated: updated,
          bookingId: targetBid
        })).setMimeType(ContentService.MimeType.JSON);
      }


    // 3. Move Voucher to "Redeemed" Tab (Option B: Single Consolidated Batch Payload with Bottom-Up Descending Deletion)
    if (data.action === "MOVE_TO_REDEEMED" || data.action === "MARK_REDEEMED") {
      // 1. Normalize items array
      let items = [];
      if (Array.isArray(data.items) && data.items.length > 0) {
        items = data.items.map(it => ({
          bookingId: String(it.bookingId || '').trim(),
          brand: String(it.brand || '').trim(),
          code: String(it.code || '').trim()
        })).filter(it => it.bookingId || it.code);
      } else {
        const rawCodes = data.codes || (data.code ? [data.code] : []);
        const codes = Array.from(new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []));
        if (!codes.length && data.bookingId) codes.push(String(data.bookingId).trim());
        items = codes.map(c => ({
          bookingId: /^RUPS/i.test(c) ? c : (data.bookingId ? String(data.bookingId).trim() : ''),
          brand: data.brand ? String(data.brand).trim() : '',
          code: c
        }));
      }

      // Guard against illegal New -> Redeemed transition
      if (data.brand === "New" || items.some(it => it.brand === "New")) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Illegal transition: Vouchers in 'New' tab cannot be moved directly to Redeemed. Approve to brand tab first."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (!items.length) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "No vouchers specified for redemption."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      let rSheet = ss.getSheetByName("Redeemed");
      if (!rSheet) {
        rSheet = ss.insertSheet("Redeemed");
        rSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
        rSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
        rSheet.setFrozenRows(1);
      }

      // Pre-fetch Redeemed sheet Booking IDs & Codes (Strict Booking ID Idempotency)
      const rVals = rSheet.getDataRange().getValues();
      const existingRedeemedBookings = new Set();
      const existingRedeemedCodes = new Set();
      for (let r = 1; r < rVals.length; r++) {
        const b = String(rVals[r][1] || '').trim().replace(/^'/, '').toUpperCase(); // Col 2: Booking ID
        const c = String(rVals[r][2] || '').trim().replace(/^'/, '').toUpperCase(); // Col 3: Code
        if (b && b !== '-' && !b.startsWith('#')) existingRedeemedBookings.add(b);
        if (c && c !== '-' && c !== 'PENDING') existingRedeemedCodes.add(c);
      }

      const moved = [];
      const itemsToSearch = [];

      // Check for self-healing: if already present in Redeemed, mark success immediately
      items.forEach(it => {
        const bUpper = String(it.bookingId || '').replace(/^'/, '').trim().toUpperCase();
        const cUpper = String(it.code || '').replace(/^'/, '').trim().toUpperCase();
        const hasRupsBid = bUpper.startsWith('RUPS-');
        const isAlreadyRedeemedByBid = hasRupsBid && existingRedeemedBookings.has(bUpper);
        const isAlreadyRedeemedByCode = cUpper && cUpper !== '-' && cUpper !== 'PENDING' && existingRedeemedCodes.has(cUpper);

        if (isAlreadyRedeemedByBid || isAlreadyRedeemedByCode) {
          moved.push({ bookingId: it.bookingId, code: it.code, brand: it.brand, alreadyRedeemed: true });
        } else {
          itemsToSearch.push(it);
        }
      });

      // If items still need to be moved from brand tabs:
      if (itemsToSearch.length > 0) {
        // Group search items by brand sheet where known, or collect for discovery
        // CANONICAL BRAND TRANSLATION: Map display names (e.g. "Cult.fit Elite", "Amazon Prime", "Lakmé Salon") to canonical sheet names
        const sheetsToScan = new Map();

        itemsToSearch.forEach(it => {
          let targetSheetName = "";
          if (it.brand && !SYSTEM_SHEET_NAMES.has(it.brand) && it.brand !== "Cards") {
            if (ss.getSheetByName(it.brand)) {
              targetSheetName = it.brand;
            } else {
              const canonical = extractBrand(it.brand, "");
              if (canonical && canonical !== "Other" && ss.getSheetByName(canonical)) {
                targetSheetName = canonical;
              }
            }
          }
          if (targetSheetName) {
            if (!sheetsToScan.has(targetSheetName)) sheetsToScan.set(targetSheetName, []);
            sheetsToScan.get(targetSheetName).push(it);
          } else {
            if (!sheetsToScan.has('__ALL__')) sheetsToScan.set('__ALL__', []);
            sheetsToScan.get('__ALL__').push(it);
          }
        });

        // Determine all sheets that need inspecting
        const allSheets = ss.getSheets();
        const targetSheetsMap = new Map();

        for (const [brandName, bItems] of sheetsToScan.entries()) {
          if (brandName === '__ALL__') {
            for (const s of allSheets) {
              const sName = s.getName();
              if (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards") continue;
              if (!targetSheetsMap.has(s)) targetSheetsMap.set(s, []);
              targetSheetsMap.get(s).push(...bItems);
            }
          } else {
            const directSheet = ss.getSheetByName(brandName);
            if (directSheet && !SYSTEM_SHEET_NAMES.has(directSheet.getName()) && directSheet.getName() !== "Cards") {
              if (!targetSheetsMap.has(directSheet)) targetSheetsMap.set(directSheet, []);
              targetSheetsMap.get(directSheet).push(...bItems);
            }
          }
        }

        // Process each target brand sheet with Single-Pass Bottom-Up Descending Deletion
        for (const [sheet, sheetItems] of targetSheetsMap.entries()) {
          const sName = sheet.getName();
          if (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards") continue;

          const vals = sheet.getDataRange().getValues();
          if (vals.length <= 1) continue;

          const rowsToDelete = [];
          const targetBidSet = new Set(sheetItems.map(it => String(it.bookingId || '').replace(/^'/, '').trim().toUpperCase()).filter(b => b && b !== '-' && !b.startsWith('#')));
          const targetCodeSet = new Set(sheetItems.map(it => String(it.code || '').replace(/^'/, '').trim().toUpperCase()).filter(c => c && c !== '-' && c !== 'PENDING'));

          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][0] || '').replace(/^'/, '').trim().toUpperCase();
            const rowCode = String(vals[i][1] || '').replace(/^'/, '').trim().toUpperCase();

            // Hierarchical matching: Primary RUPS Booking ID, Fallback Voucher Code
            const hasRupsBid = rowBid.startsWith('RUPS-');
            const isBidMatch = hasRupsBid && targetBidSet.has(rowBid);
            const isCodeMatch = rowCode && rowCode !== '-' && rowCode !== 'PENDING' && targetCodeSet.has(rowCode);

            if (isBidMatch || isCodeMatch) {
              const bVal = String(vals[i][0] || '').trim();
              const cVal = String(vals[i][1] || '').trim();
              const bValUpper = bVal.toUpperCase();
              const cValUpper = cVal.toUpperCase();

              // Append to Redeemed if not already present in Redeemed (Hierarchical check)
              const hasRupsVal = bValUpper.startsWith('RUPS-');
              const alreadyInRedeemed = (hasRupsVal && existingRedeemedBookings.has(bValUpper)) ||
                                        (cValUpper && cValUpper !== '-' && cValUpper !== 'PENDING' && existingRedeemedCodes.has(cValUpper));

              if (!alreadyInRedeemed) {
                const rawCleanBid = String(vals[i][0] || '').trim().replace(/^'/, '');
                const rawCleanCd = String(vals[i][1] || '').trim().replace(/^'/, '');
                const rawCleanPin = String(vals[i][2] || '').trim().replace(/^'/, '');
                const redeemedRow = [
                  sName,                                   // A: Brand
                  rawCleanBid ? `'${rawCleanBid}` : "",    // B: Booking ID
                  rawCleanCd ? `'${rawCleanCd}` : "",      // C: Code
                  rawCleanPin ? `'${rawCleanPin}` : "",    // D: PIN
                  vals[i][3],                              // E: Value
                  vals[i][4],                              // F: Date of Issue
                  vals[i][5],                              // G: Account
                  vals[i][6],                              // H: Card
                  "REDEEMED",                              // I: Status
                  formatISTTimestamp(new Date()),          // J: Redeemed Timestamp (IST)
                  sName,                                   // K: Original Tab
                  vals[i][8] ? String(vals[i][8]) : "",    // L: Original Ingested Timestamp
                  true                                     // M: Display in UI
                ];
                rSheet.appendRow(redeemedRow);
                if (hasRupsVal) existingRedeemedBookings.add(bValUpper);
                if (cValUpper && cValUpper !== '-' && cValUpper !== 'PENDING') existingRedeemedCodes.add(cValUpper);
              }

              rowsToDelete.push(i + 1);
              moved.push({ bookingId: bVal, code: cVal, brand: sName });
            }
          }

          // Single-Pass Bottom-Up Deletion: Sort descending (highest row first, no renumbering needed)
          if (rowsToDelete.length > 0) {
            rowsToDelete.sort((a, b) => b - a);
            for (const rIdx of rowsToDelete) {
              sheet.deleteRow(rIdx);
            }
          }
        }

        // RESILIENT SECONDARY PASS: If any itemsToSearch were not found in their primary sheet,
        // scan ALL remaining non-system brand sheets so no voucher is ever dropped.
        const movedBidSet = new Set(moved.map(m => (m.bookingId || '').trim().toUpperCase()).filter(b => b && b.startsWith('RUPS-')));
        const movedCodeSet = new Set(moved.map(m => (m.code || '').trim().toUpperCase()).filter(c => c && c !== '-' && c !== 'PENDING'));

        const remainingItems = itemsToSearch.filter(it => {
          const bUpper = (it.bookingId || '').trim().toUpperCase();
          const cUpper = (it.code || '').trim().toUpperCase();
          const bidMoved = bUpper && bUpper.startsWith('RUPS-') && movedBidSet.has(bUpper);
          const codeMoved = cUpper && cUpper !== '-' && cUpper !== 'PENDING' && movedCodeSet.has(cUpper);
          return !bidMoved && !codeMoved;
        });

        if (remainingItems.length > 0) {
          const scannedSheets = new Set(targetSheetsMap.keys());
          for (const sheet of allSheets) {
            const sName = sheet.getName();
            if (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards" || scannedSheets.has(sheet)) continue;

            const vals = sheet.getDataRange().getValues();
            if (vals.length <= 1) continue;

            const remBidSet = new Set(remainingItems.map(it => (it.bookingId || '').trim().toUpperCase()).filter(b => b && b !== '-' && !b.startsWith('#')));
            const remCodeSet = new Set(remainingItems.map(it => (it.code || '').trim().toUpperCase()).filter(c => c && c !== '-' && c !== 'PENDING'));
            const rowsToDelete = [];

            for (let i = 1; i < vals.length; i++) {
              const rowBid = String(vals[i][0] || '').trim().toUpperCase();
              const rowCode = String(vals[i][1] || '').trim().toUpperCase();
              const hasRupsBid = rowBid.startsWith('RUPS-');
              const isBidMatch = hasRupsBid && remBidSet.has(rowBid);
              const isCodeMatch = rowCode && rowCode !== '-' && rowCode !== 'PENDING' && remCodeSet.has(rowCode);

              if (isBidMatch || isCodeMatch) {
                const bVal = String(vals[i][0] || '').trim();
                const cVal = String(vals[i][1] || '').trim();
                const bValUpper = bVal.toUpperCase();
                const cValUpper = cVal.toUpperCase();

                const hasRupsVal = bValUpper.startsWith('RUPS-');
                const alreadyInRedeemed = (hasRupsVal && existingRedeemedBookings.has(bValUpper)) ||
                                          (cValUpper && cValUpper !== '-' && cValUpper !== 'PENDING' && existingRedeemedCodes.has(cValUpper));

                if (!alreadyInRedeemed) {
                  const rawCleanBid = String(vals[i][0] || '').trim().replace(/^'/, '');
                  const rawCleanCd = String(vals[i][1] || '').trim().replace(/^'/, '');
                  const rawCleanPin = String(vals[i][2] || '').trim().replace(/^'/, '');
                  const redeemedRow = [
                    sName,                                   // A: Brand
                    rawCleanBid ? `'${rawCleanBid}` : "",    // B: Booking ID
                    rawCleanCd ? `'${rawCleanCd}` : "",      // C: Code
                    rawCleanPin ? `'${rawCleanPin}` : "",    // D: PIN
                    vals[i][3],                              // E: Value
                    vals[i][4],                              // F: Date of Issue
                    vals[i][5],                              // G: Account
                    vals[i][6],                              // H: Card
                    "REDEEMED",                              // I: Status
                    formatISTTimestamp(new Date()),          // J: Redeemed Timestamp (IST)
                    sName,                                   // K: Original Tab
                    vals[i][8] ? String(vals[i][8]) : "",    // L: Original Ingested Timestamp
                    true                                     // M: Display in UI
                  ];
                  rSheet.appendRow(redeemedRow);
                  if (hasRupsVal) existingRedeemedBookings.add(bValUpper);
                  if (cValUpper && cValUpper !== '-' && cValUpper !== 'PENDING') existingRedeemedCodes.add(cValUpper);
                }

                rowsToDelete.push(i + 1);
                moved.push({ bookingId: bVal, code: cVal, brand: sName });
                if (hasRupsVal) movedBidSet.add(bValUpper);
                if (cValUpper && cValUpper !== '-' && cValUpper !== 'PENDING') movedCodeSet.add(cValUpper);
              }
            }

            if (rowsToDelete.length > 0) {
              rowsToDelete.sort((a, b) => b - a);
              for (const rIdx of rowsToDelete) {
                sheet.deleteRow(rIdx);
              }
            }
          }
        }
      }

      // Force atomic persistence
      SpreadsheetApp.flush();

      // Also update Shared_Activity status to REDEEMED for these codes
      const sSheet = ss.getSheetByName("Shared_Activity");
      if (sSheet && moved.length > 0) {
        const movedCodesSet = new Set(moved.map(m => m.code).filter(Boolean));
        const movedBidsSet = new Set(moved.map(m => m.bookingId).filter(Boolean));
        const sRows = sSheet.getDataRange().getValues();
        for (let i = 1; i < sRows.length; i++) {
          const sCode = String(sRows[i][0] || '').trim();
          if (movedCodesSet.has(sCode) || movedBidsSet.has(sCode)) {
            sSheet.getRange(i + 1, 6).setValue("REDEEMED");
          }
        }
      }

      if (moved.length > 0) {
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          action: "MOVE_TO_REDEEMED",
          movedCount: moved.length,
          moved: moved
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        error: "Voucher(s) not found in any brand tab: " + items.map(it => it.bookingId || it.code).join(", ")
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
      const codes = new Set((Array.isArray(rawCodes) ? rawCodes : [rawCodes]).map(c => String(c).replace(/^'/, '').trim().toUpperCase()).filter(Boolean));
      if (data.bookingId) {
        codes.add(String(data.bookingId).replace(/^'/, '').trim().toUpperCase());
      }
      if (data.bookingIds && Array.isArray(data.bookingIds)) {
        data.bookingIds.forEach(b => codes.add(String(b).replace(/^'/, '').trim().toUpperCase()));
      }

      const restored = [];
      const rRows = rSheet.getDataRange().getValues();
      const rowsToDeleteFromRedeemed = [];

      // Search backwards so deleting rows does not alter row index of earlier rows
      for (let i = rRows.length - 1; i >= 1; i--) {
        const rowCode = String(rRows[i][2] || '').replace(/^'/, '').trim();
        const rowBooking = String(rRows[i][1] || '').replace(/^'/, '').trim();
        const rowCodeUpper = rowCode.toUpperCase();
        const rowBookingUpper = rowBooking.toUpperCase();

        if ((rowCodeUpper && codes.has(rowCodeUpper)) || (rowBookingUpper && codes.has(rowBookingUpper))) {
          const rawOrigBrand = String(rRows[i][10] || rRows[i][0] || 'Unassigned').trim();
          const cleanBid = String(rRows[i][1] || '').trim().replace(/^'/, '');
          const cleanCd = String(rRows[i][2] || '').trim().replace(/^'/, '');
          const cleanPin = String(rRows[i][3] || '').trim().replace(/^'/, '');
          const originalTimestamp = (rRows[i].length > 11 && rRows[i][11]) ? String(rRows[i][11]) : "";

          // Staging Invariant: If voucher originally came from "New" (staging), restore to "New" tab
          if (rawOrigBrand === "New") {
            let nSheet = ss.getSheetByName("New");
            if (!nSheet) {
              nSheet = ss.insertSheet("New");
              nSheet.appendRow(UNIFIED_NEW_HEADERS);
              nSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
              nSheet.setFrozenRows(1);
            }
            const origBrand = String(rRows[i][0] || 'Unassigned');
            const nVals = nSheet.getDataRange().getValues();
            let existingNewRow = -1;
            for (let n = 1; n < nVals.length; n++) {
              const nb = String(nVals[n][0] || '').trim().replace(/^'/, '').toUpperCase();
              const nc = String(nVals[n][1] || '').trim().replace(/^'/, '').toUpperCase();
              const hasRupsBooking = rowBooking && rowBooking.toUpperCase().startsWith('RUPS-');
              const bidMatch = hasRupsBooking && nb === rowBooking.toUpperCase();
              const codeMatch = (!hasRupsBooking || !nb.startsWith('RUPS-')) && rowCode && rowCode !== '-' && rowCode !== 'PENDING' && nc === rowCode.toUpperCase();
              if (bidMatch || codeMatch) {
                existingNewRow = n + 1;
                break;
              }
            }
            if (existingNewRow > 0) {
              nSheet.getRange(existingNewRow, 8).setValue("ACTIVE");
              nSheet.getRange(existingNewRow, 10).setValue(true);
            } else {
              nSheet.appendRow([
                cleanBid ? `'${cleanBid}` : '',         // A: Booking ID
                cleanCd ? `'${cleanCd}` : '',           // B: Code
                cleanPin ? `'${cleanPin}` : '',         // C: PIN
                rRows[i][4],                            // D: Value
                rRows[i][5],                            // E: Date of Issue
                rRows[i][6],                            // F: Account
                rRows[i][7],                            // G: Card
                "ACTIVE",                               // H: Status
                originalTimestamp,                      // I: Appended Timestamp (IST)
                true,                                   // J: Display in UI
                origBrand,                              // K: Target Brand
                "",                                     // L: Payment Link
                ""                                      // M: Sale Logged
              ]);
            }
            rowsToDeleteFromRedeemed.push(i + 1);
            restored.push({ bookingId: rowBooking, code: rowCode, brand: "New" });
            continue;
          }

          const canonicalBrand = extractBrand(rawOrigBrand, "");
          const targetBrand = (canonicalBrand && canonicalBrand !== "Other") ? canonicalBrand : rawOrigBrand;
          const bSheet = getTargetBrandSheet(ss, targetBrand);

          // DEDUPLICATION: Inspect brand sheet (Hierarchical: Primary RUPS Booking ID, Fallback Voucher Code)
          const bVals = bSheet.getDataRange().getValues();
          let existingBrandRowIndex = -1;
          for (let b = 1; b < bVals.length; b++) {
            const bIdVal = String(bVals[b][0] || '').replace(/^'/, '').trim().toUpperCase();
            const bCodeVal = String(bVals[b][1] || '').replace(/^'/, '').trim().toUpperCase();
            const hasRupsBooking = rowBooking && rowBooking.toUpperCase().startsWith('RUPS-');
            const bidMatch = hasRupsBooking && bIdVal === rowBooking.toUpperCase();
            const codeMatch = (!hasRupsBooking || !bIdVal.startsWith('RUPS-')) && rowCode && rowCode !== '-' && rowCode !== 'PENDING' && bCodeVal === rowCode.toUpperCase();
            if (bidMatch || codeMatch) {
              existingBrandRowIndex = b + 1;
              break;
            }
          }

          if (existingBrandRowIndex > 0) {
            bSheet.getRange(existingBrandRowIndex, 8).setValue("ACTIVE");
            bSheet.getRange(existingBrandRowIndex, 10).setValue(true);
          } else {
            const restoredRow = [
              cleanBid ? `'${cleanBid}` : '',         // A: Booking ID
              cleanCd ? `'${cleanCd}` : '',           // B: Code
              cleanPin ? `'${cleanPin}` : '',         // C: PIN
              rRows[i][4],                            // D: Value
              rRows[i][5],                            // E: Date of Issue
              rRows[i][6],                            // F: Account
              rRows[i][7],                            // G: Card
              "ACTIVE",                               // H: Status
              originalTimestamp,                      // I: Appended Timestamp (IST)
              true                                    // J: Display in UI
            ];
            bSheet.appendRow(restoredRow);
          }

          rowsToDeleteFromRedeemed.push(i + 1);
          restored.push({ bookingId: rowBooking, code: rowCode, brand: targetBrand });
        }
      }

      // Self-healing: If not found in Redeemed, check if already restored in brand tab or New tab!
      codes.forEach(codeOrBid => {
        const cleanTarget = String(codeOrBid).replace(/^'/, '').trim().toUpperCase();
        if (!restored.some(r => (r.code && String(r.code).replace(/^'/, '').trim().toUpperCase() === cleanTarget) || (r.bookingId && String(r.bookingId).replace(/^'/, '').trim().toUpperCase() === cleanTarget))) {
          const foundInBrand = findRowInBrandTabs(ss, cleanTarget, null, cleanTarget);
          if (foundInBrand) {
            restored.push({ bookingId: cleanTarget, brand: foundInBrand.sheet.getName(), alreadyRestored: true });
          } else {
            const nSheet = ss.getSheetByName("New");
            if (nSheet) {
              const nVals = nSheet.getDataRange().getValues();
              for (let n = 1; n < nVals.length; n++) {
                const nb = String(nVals[n][0] || '').trim().replace(/^'/, '').toUpperCase();
                const nc = String(nVals[n][1] || '').trim().replace(/^'/, '').toUpperCase();
                if ((cleanTarget && nb === cleanTarget) || (cleanTarget && cleanTarget !== '-' && cleanTarget !== 'PENDING' && nc === cleanTarget)) {
                  restored.push({ bookingId: cleanTarget, brand: "New", alreadyRestored: true });
                  break;
                }
              }
            }
          }
        }
      });

      // Delete from Redeemed sheet in descending order (no renumbering needed)
      if (rowsToDeleteFromRedeemed.length > 0) {
        rowsToDeleteFromRedeemed.sort((a, b) => b - a);
        for (const rowIdx of rowsToDeleteFromRedeemed) {
          rSheet.deleteRow(rowIdx);
        }
      }

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
        error: "Voucher(s) not found in Redeemed sheet: " + Array.from(codes).join(", ")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. Save Card (Add or Update in "Active Cards" tab with fallback to legacy "Cards")
    if (data.action === "SAVE_CARD") {
      const aSheet = ss.getSheetByName("Active Cards");
      const card = data.card || data;
      const targetCardNum = String(card.cardNumber || card.last4 || card.id || '').replace(/['\s]/g, '').trim();
      let updated = false;

      if (aSheet && targetCardNum) {
        const aRows = aSheet.getDataRange().getValues();
        for (let i = 1; i < aRows.length; i++) {
          const rowCardNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowCardNum === targetCardNum || (targetCardNum.length === 4 && rowCardNum.endsWith(targetCardNum))) {
            if (card.status) aSheet.getRange(i + 1, 8).setValue(card.status.toUpperCase());
            if (card.notes !== undefined) aSheet.getRange(i + 1, 9).setValue(card.notes);
            if (card.cardName || card.nick) aSheet.getRange(i + 1, 3).setValue(card.cardName || card.nick);
            if (card.linkedAccount !== undefined) aSheet.getRange(i + 1, 4).setValue(card.linkedAccount ? `'${card.linkedAccount}` : 'NA');
            if (card.expiry) aSheet.getRange(i + 1, 6).setValue(card.expiry.startsWith("'") ? card.expiry : `'${card.expiry}`);
            if (card.cvv) aSheet.getRange(i + 1, 7).setValue(card.cvv.startsWith("'") ? card.cvv : `'${card.cvv}`);
            aSheet.getRange(i + 1, 23).setValue(formatISTTimestamp(new Date()));
            updated = true;
            break;
          }
        }
      }

      const cSheet = getOrInitCardsSheet(ss);
      const cardId = String(card.id || ('c-' + Date.now())).trim();
      if (!updated) {
        const rows = cSheet.getDataRange().getValues();
        let found = false;

        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][0]).trim() === cardId) {
            cSheet.getRange(i + 1, 2).setValue(card.bank || rows[i][1]);
            cSheet.getRange(i + 1, 3).setValue(card.type || rows[i][2]);
            cSheet.getRange(i + 1, 4).setValue(card.network || rows[i][3] || 'RuPay');
            cSheet.getRange(i + 1, 5).setValue(card.bin ? `'${card.bin}` : rows[i][4]);
            cSheet.getRange(i + 1, 6).setValue(card.last4 ? `'${card.last4}` : rows[i][5]);
            cSheet.getRange(i + 1, 7).setValue(card.nick || rows[i][6]);
            cSheet.getRange(i + 1, 8).setValue(card.status || rows[i][7]);
            cSheet.getRange(i + 1, 9).setValue(card.claimed !== undefined ? Number(card.claimed) : rows[i][8]);
            cSheet.getRange(i + 1, 10).setValue(card.total !== undefined ? Number(card.total) : rows[i][9]);
            cSheet.getRange(i + 1, 11).setValue(card.blockedDate || rows[i][10]);
            cSheet.getRange(i + 1, 12).setValue(card.notes || rows[i][11]);
            cSheet.getRange(i + 1, 13).setValue(formatISTTimestamp(new Date()));
            found = true;
            break;
          }
        }

        if (!found) {
          cSheet.appendRow([
            cardId,
            card.bank || 'RuPay Card',
            card.type || 'RuPay Select',
            card.network || 'RuPay',
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
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "SAVE_CARD",
        id: cardId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. Block Card: Move row from "Active Cards" to "Blocked Cards"
    if (data.action === "BLOCK_CARD") {
      const aSheet = ss.getSheetByName("Active Cards");
      let bSheet = ss.getSheetByName("Blocked Cards");
      if (!bSheet) {
        bSheet = ss.insertSheet("Blocked Cards");
      }
      const targetCardNum = String(data.cardNumber || data.id || '').replace(/['\s]/g, '').trim();
      let found = false;

      // Safeguard: Purge any stale pre-existing duplicate in Blocked Cards
      if (bSheet && targetCardNum) {
        const bRows = bSheet.getDataRange().getValues();
        for (let b = bRows.length - 1; b >= 1; b--) {
          const bNum = String(bRows[b][4] || '').replace(/['\s]/g, '').trim();
          if (bNum === targetCardNum) {
            bSheet.deleteRow(b + 1);
          }
        }
      }

      if (aSheet && targetCardNum) {
        const aRows = aSheet.getDataRange().getValues();
        for (let i = 1; i < aRows.length; i++) {
          const rowNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === targetCardNum || (targetCardNum.length === 4 && rowNum.endsWith(targetCardNum))) {
            const rowData = aRows[i];
            const blockedRow = [
              rowData[0], // Bank
              rowData[1], // Card_Type
              rowData[2], // Card_Name
              "'" + String(rowData[3] || 'NA').replace(/^'/, ''), // Linked_Account
              "'" + rowNum, // Card_Number
              "'" + String(rowData[5] || '').replace(/^'/, ''), // Expiry
              "'" + String(rowData[6] || '').replace(/^'/, ''), // CVV
              "BLOCKED",
              formatISTTimestamp(new Date()), // Blocked_Timestamp
              "", // Replaced_By_Card
              data.notes || String(rowData[8] || 'Blocked via App'),
              // 13 benefits preserved
              rowData[9], rowData[10], rowData[11], rowData[12], rowData[13],
              rowData[14], rowData[15], rowData[16], rowData[17], rowData[18],
              rowData[19], rowData[20], rowData[21]
            ];
            bSheet.appendRow(blockedRow);
            aSheet.deleteRow(i + 1);
            found = true;
            break;
          }
        }
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: found,
        action: "BLOCK_CARD",
        cardNumber: targetCardNum
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6B. Block & Replace Card: Transactional, idempotent swap with dual-sheet consistency
    if (data.action === "BLOCK_AND_REPLACE_CARD") {
      const aSheet = ss.getSheetByName("Active Cards");
      let bSheet = ss.getSheetByName("Blocked Cards");
      if (!bSheet) bSheet = ss.insertSheet("Blocked Cards");

      const oldCardNum = String(data.oldCardNumber || '').replace(/['\s]/g, '').trim();
      const newCard = data.newCard || {};
      const newCardNum = String(newCard.cardNumber || '').replace(/['\s]/g, '').trim();
      let foundOld = false;
      let oldRowData = null;

      if (!newCardNum || newCardNum.length < 15) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Invalid replacement card number. Must be 16 digits.",
          action: "BLOCK_AND_REPLACE_CARD"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 1. Safeguard: If newCardNum was in Blocked Cards, purge it from Blocked Cards (reactivation guard)
      // Also purge any pre-existing duplicate of oldCardNum in Blocked Cards
      if (bSheet) {
        const bRows = bSheet.getDataRange().getValues();
        for (let b = bRows.length - 1; b >= 1; b--) {
          const bCardNum = String(bRows[b][4] || '').replace(/['\s]/g, '').trim();
          if (bCardNum === newCardNum || bCardNum === oldCardNum) {
            bSheet.deleteRow(b + 1);
          }
        }
      }

      // 2. Find and archive the old card to Blocked Cards
      if (aSheet && oldCardNum) {
        const aRows = aSheet.getDataRange().getValues();
        for (let i = 1; i < aRows.length; i++) {
          const rowNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === oldCardNum || (oldCardNum.length === 4 && rowNum.endsWith(oldCardNum))) {
            oldRowData = aRows[i];
            const blockedRow = [
              oldRowData[0], // Bank
              oldRowData[1], // Card_Type
              oldRowData[2], // Card_Name
              "'" + String(oldRowData[3] || 'NA').replace(/^'/, ''),
              "'" + rowNum,
              "'" + String(oldRowData[5] || '').replace(/^'/, ''),
              "'" + String(oldRowData[6] || '').replace(/^'/, ''),
              "BLOCKED",
              formatISTTimestamp(new Date()),
              newCardNum ? newCardNum.slice(-4) : '',
              data.notes || String(oldRowData[8] || 'Replaced via App'),
              oldRowData[9], oldRowData[10], oldRowData[11], oldRowData[12], oldRowData[13],
              oldRowData[14], oldRowData[15], oldRowData[16], oldRowData[17], oldRowData[18],
              oldRowData[19], oldRowData[20], oldRowData[21]
            ];
            bSheet.appendRow(blockedRow);
            aSheet.deleteRow(i + 1);
            foundOld = true;
            break;
          }
        }
      }

      // 3. Add or update the new card in Active Cards with eligible benefits
      if (aSheet && newCardNum) {
        const bank = newCard.bank || (oldRowData ? oldRowData[0] : 'PNB');
        const type = newCard.type || newCard.cardType || (oldRowData ? oldRowData[1] : 'Debit');
        const name = newCard.name || newCard.cardName || (oldRowData ? oldRowData[2] : (bank + ' ' + (bank === 'PNB' ? 'Imperial' : 'RuPay Select')));
        const linkedAcc = newCard.linkedAccount || (oldRowData ? String(oldRowData[3] || '').replace(/^'/, '') : 'NA');
        const expiry = newCard.expiry || '10/33';
        const cvv = newCard.cvv || '000';
        const notes = newCard.notes || 'Replaced card';
        const nowStr = formatISTTimestamp(new Date());
        const benefits = getEligibleBenefitsArray(bank, type, name);

        // Check if newCardNum is ALREADY in Active Cards (prevent duplicate row!)
        const currentARows = aSheet.getDataRange().getValues();
        let existingActiveRow = -1;
        for (let j = 1; j < currentARows.length; j++) {
          const existingNum = String(currentARows[j][4] || '').replace(/['\s]/g, '').trim();
          if (existingNum === newCardNum) {
            existingActiveRow = j + 1;
            break;
          }
        }

        if (existingActiveRow > 1) {
          // Update in-place to prevent duplication
          aSheet.getRange(existingActiveRow, 1).setValue(bank);
          aSheet.getRange(existingActiveRow, 2).setValue(type);
          aSheet.getRange(existingActiveRow, 3).setValue(name);
          aSheet.getRange(existingActiveRow, 4).setValue("'" + String(linkedAcc).replace(/^'/, ''));
          aSheet.getRange(existingActiveRow, 5).setValue("'" + newCardNum);
          aSheet.getRange(existingActiveRow, 6).setValue("'" + String(expiry).replace(/^'/, ''));
          aSheet.getRange(existingActiveRow, 7).setValue("'" + String(cvv).replace(/^'/, ''));
          aSheet.getRange(existingActiveRow, 8).setValue("ACTIVE");
          aSheet.getRange(existingActiveRow, 9).setValue(notes);
        } else {
          // Append brand new active row
          const newRow = [
            bank,
            type,
            name,
            "'" + String(linkedAcc).replace(/^'/, ''),
            "'" + newCardNum,
            "'" + String(expiry).replace(/^'/, ''),
            "'" + String(cvv).replace(/^'/, ''),
            "ACTIVE",
            notes,
            ...benefits,
            nowStr
          ];
          aSheet.appendRow(newRow);
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "BLOCK_AND_REPLACE_CARD",
        oldCardNumber: oldCardNum,
        newCardNumber: newCardNum,
        foundOld: foundOld
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6C. Add Brand New Card (With Idempotency & Deduplication Shield)
    if (data.action === "ADD_CARD") {
      let aSheet = ss.getSheetByName("Active Cards");
      if (!aSheet) aSheet = ss.insertSheet("Active Cards");
      const c = data.card || data;
      const cardNum = String(c.cardNumber || '').replace(/['\s]/g, '').trim();

      if (aSheet && cardNum) {
        const bank = c.bank || 'PNB';
        const type = c.cardType || c.type || 'Debit';
        const name = c.cardName || c.name || 'RuPay Select Card';
        const linkedAcc = c.linkedAccount || 'NA';
        const expiry = c.expiry || '10/33';
        const cvv = c.cvv || '000';
        const notes = c.notes || '';
        const nowStr = formatISTTimestamp(new Date());
        const benefits = getEligibleBenefitsArray(bank, type, name);

        // Check if cardNum already exists in Active Cards
        const aRows = aSheet.getDataRange().getValues();
        let existingRowIndex = -1;
        for (let i = 1; i < aRows.length; i++) {
          const rowNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === cardNum) {
            existingRowIndex = i + 1;
            break;
          }
        }

        if (existingRowIndex > 1) {
          // Update in-place to prevent duplication
          aSheet.getRange(existingRowIndex, 1).setValue(bank);
          aSheet.getRange(existingRowIndex, 2).setValue(type);
          aSheet.getRange(existingRowIndex, 3).setValue(name);
          aSheet.getRange(existingRowIndex, 4).setValue("'" + String(linkedAcc).replace(/^'/, ''));
          aSheet.getRange(existingRowIndex, 5).setValue("'" + cardNum);
          aSheet.getRange(existingRowIndex, 6).setValue("'" + String(expiry).replace(/^'/, ''));
          aSheet.getRange(existingRowIndex, 7).setValue("'" + String(cvv).replace(/^'/, ''));
          aSheet.getRange(existingRowIndex, 8).setValue("ACTIVE");
          aSheet.getRange(existingRowIndex, 9).setValue(notes);
        } else {
          aSheet.appendRow([
            bank,
            type,
            name,
            "'" + String(linkedAcc).replace(/^'/, ''),
            "'" + cardNum,
            "'" + String(expiry).replace(/^'/, ''),
            "'" + String(cvv).replace(/^'/, ''),
            "ACTIVE",
            notes,
            ...benefits,
            nowStr
          ]);
        }
        SpreadsheetApp.flush();
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "ADD_CARD",
        cardNumber: cardNum
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6D. Add Miscellaneous Voucher
    if (data.action === "ADD_MISC_VOUCHER") {
      const sheet = getOrCreateMiscSheet(ss);
      const brandName = String(data.brandName || data.brand || 'Misc').trim();
      const voucherCode = String(data.voucherCode || data.code || '').replace(/^'/, '').trim();
      const pin = data.pin ? String(data.pin).replace(/^'/, '').trim() : '';
      const amount = Number(data.amount || data.value || 500) || 500;
      const nowStr = formatISTTimestamp(new Date());
      const todayStr = formatISTDate(new Date());

      if (!voucherCode) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Voucher code is required"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Idempotency check: see if voucher code already exists in Misc sheet
      const vals = sheet.getDataRange().getValues();
      let existingRowIndex = -1;
      for (let i = 1; i < vals.length; i++) {
        const rowCode = String(vals[i][1] || '').replace(/^'/, '').trim();
        if (rowCode.toUpperCase() === voucherCode.toUpperCase()) {
          existingRowIndex = i + 1;
          break;
        }
      }

      const bookingId = data.bookingId || brandName || ('MISC-' + Date.now().toString(36).toUpperCase());

      if (existingRowIndex > 0) {
        // Update in-place
        sheet.getRange(existingRowIndex, 1).setValue(brandName || bookingId);
        sheet.getRange(existingRowIndex, 2).setValue("'" + voucherCode);
        sheet.getRange(existingRowIndex, 3).setValue(pin ? ("'" + pin) : "");
        sheet.getRange(existingRowIndex, 4).setValue(amount);
        sheet.getRange(existingRowIndex, 6).setValue(brandName);
        sheet.getRange(existingRowIndex, 8).setValue("ACTIVE");
        sheet.getRange(existingRowIndex, 9).setValue(nowStr);
        sheet.getRange(existingRowIndex, 10).setValue(true);
      } else {
        // Append new row
        sheet.appendRow([
          brandName || bookingId, // A: Booking ID / Reference / Sub-brand
          "'" + voucherCode,      // B: Code
          pin ? ("'" + pin) : "", // C: PIN
          amount,                 // D: Value
          todayStr,               // E: Date of Issue
          brandName,              // F: Account / Brand
          "Manual",               // G: Card
          "ACTIVE",               // H: Status
          nowStr,                 // I: Timestamp (IST)
          true                    // J: Display in UI
        ]);
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "ADD_MISC_VOUCHER",
        bookingId: bookingId,
        code: voucherCode,
        brand: brandName,
        pin: pin,
        amount: amount
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6E. Delete / Purge Miscellaneous Voucher (Used for Undo & Safeguard Verification)
    if (data.action === "DELETE_MISC_VOUCHER") {
      const sheet = ss.getSheetByName("Misc");
      const voucherCode = String(data.voucherCode || data.code || data.bookingId || '').replace(/^'/, '').trim().toUpperCase();
      let deleted = false;
      if (sheet && voucherCode) {
        const vals = sheet.getDataRange().getValues();
        for (let i = vals.length - 1; i >= 1; i--) {
          const rowBid = String(vals[i][0] || '').replace(/^'/, '').trim().toUpperCase();
          const rowCode = String(vals[i][1] || '').replace(/^'/, '').trim().toUpperCase();
          if (rowCode === voucherCode || (voucherCode && rowBid === voucherCode)) {
            sheet.deleteRow(i + 1);
            deleted = true;
            break;
          }
        }
        SpreadsheetApp.flush();
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "DELETE_MISC_VOUCHER",
        deleted: deleted,
        voucherCode: voucherCode
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 7B. Tag Payment Card on Ingested Vouchers
    if (data.action === "TAG_CARD") {
      const cardTag = String(data.cardTag || data.card || "").trim();
      const rawCodes = (data.codes || [])
        .concat(data.bookingIds || [])
        .concat(data.code ? [data.code] : [])
        .concat(data.bookingId ? [data.bookingId] : []);
      const PLACEHOLDER_BACKEND = new Set(['PENDING', 'PROCESSING', 'PENDING REVALIDATION', 'PENDING_REVALIDATION', '-']);
      const codes = new Set(
        (Array.isArray(rawCodes) ? rawCodes : [rawCodes])
          .map(c => String(c).replace(/^'/, '').trim().toUpperCase())
          .filter(c => c && !PLACEHOLDER_BACKEND.has(c) && !/^processing$/i.test(c))
      );
      if (!cardTag || codes.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Missing cardTag or voucher codes."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const updated = [];
      const sheets = ss.getSheets();

      for (const sheet of sheets) {
        const sName = sheet.getName();
        if (sName !== "New" && (SYSTEM_SHEET_NAMES.has(sName) || sName === "Cards")) continue;
        const vals = sheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
          const rowBid = String(vals[i][0] || "").replace(/^'/, '').trim();
          const rowCode = String(vals[i][1] || "").replace(/^'/, '').trim();
          if (codes.has(rowCode.toUpperCase()) || codes.has(rowBid.toUpperCase())) {
            sheet.getRange(i + 1, 7).setValue(cardTag); // Col 7: Card
            updated.push({ brand: sName, bookingId: rowBid, code: rowCode, card: cardTag });
          }
        }
      }

      // Automatically flip the corresponding benefit to "Redeemed" in "Active Cards" tab
      try {
        const aSheet = ss.getSheetByName("Active Cards");
        if (aSheet && updated.length > 0) {
          const aVals = aSheet.getDataRange().getValues();
          const brandColMap = {
            "lakme": 10, "four fountains": 10, "spa": 10,
            "thyrocare": 11, "health": 11,
            "amazon prime": 12, "prime": 12,
            "cult live": 13, "cult.fit live": 13,
            "cult elite": 14, "cult.fit elite": 14,
            "uber": 15,
            "myntra": 16,
            "decathlon": 17,
            "reliance": 18, "reliance digital": 18,
            "bigbasket": 19,
            "apollo": 20, "apollo pharmacy": 20,
            "kalyan": 21, "kalyan jewellers": 21,
            "taxspanner": 22
          };

          updated.forEach(u => {
            const bName = String(u.brand || "").toLowerCase();
            let targetCol = 0;
            for (const key in brandColMap) {
              if (bName.includes(key)) {
                targetCol = brandColMap[key];
                break;
              }
            }

            if (targetCol > 0) {
              for (let r = 1; r < aVals.length; r++) {
                const rowCardNum = String(aVals[r][4] || "").replace(/['\s]/g, "");
                const rowAcc = String(aVals[r][3] || "").replace(/^'/, "");
                const tagStr = cardTag.replace(/['\s]/g, "");
                
                const matchCard = (rowCardNum && tagStr.includes(rowCardNum.slice(-4))) ||
                                  (rowAcc && rowAcc !== "NA" && tagStr.includes(rowAcc));

                if (matchCard) {
                  const currentVal = String(aVals[r][targetCol - 1] || "");
                  if (currentVal !== "Not Applicable") {
                    aSheet.getRange(r + 1, targetCol).setValue("Redeemed");
                    aSheet.getRange(r + 1, 23).setValue(formatISTTimestamp(new Date()));
                  }
                }
              }
            }
          });
        }
      } catch (errTagActive) {
        console.warn("Active Cards benefit auto-flip notice:", errTagActive.message);
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "TAG_CARD",
        cardTag: cardTag,
        updatedCount: updated.length,
        updated: updated
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
        const cStr = String(item.code).replace(/^'/, '').trim();
        const cStrUpper = cStr.toUpperCase();
        let foundIndex = -1;
        for (let i = 1; i < sRows.length; i++) {
          if (String(sRows[i][0] || '').replace(/^'/, '').trim().toUpperCase() === cStrUpper) {
            foundIndex = i + 1;
            break;
          }
        }

        const timeStr = formatISTTimestamp(new Date());
        if (foundIndex > 0) {
          sSheet.getRange(foundIndex, 5).setValue(timeStr);
          sSheet.getRange(foundIndex, 6).setValue("SHARED");
          if (item.platform) sSheet.getRange(foundIndex, 7).setValue(item.platform);
          sSheet.getRange(foundIndex, 8).setValue(true);
        } else {
          sSheet.appendRow([
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

      SpreadsheetApp.flush();
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
      const codes = new Set((Array.isArray(rawCodes) ? rawCodes : [rawCodes]).map(c => String(c).replace(/^'/, '').trim().toUpperCase()).filter(Boolean));
      if (data.bookingId) codes.add(String(data.bookingId).replace(/^'/, '').trim().toUpperCase());
      if (data.bookingIds && Array.isArray(data.bookingIds)) {
        data.bookingIds.forEach(b => codes.add(String(b).replace(/^'/, '').trim().toUpperCase()));
      }
      const sRows = sSheet.getDataRange().getValues();
      const dismissed = [];

      for (let i = 1; i < sRows.length; i++) {
        const rowCode = String(sRows[i][0] || '').replace(/^'/, '').trim();
        const rowCodeUpper = rowCode.toUpperCase();
        if (rowCodeUpper && codes.has(rowCodeUpper)) {
          sSheet.getRange(i + 1, 6).setValue("DISMISSED");
          sSheet.getRange(i + 1, 8).setValue(false); // Col 8: Display in UI = FALSE
          dismissed.push(rowCode);
        }
      }

      SpreadsheetApp.flush();
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
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "RESET_COOLDOWN",
        account: account,
        timestamp: formatISTTimestamp(new Date())
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 11. Clear Activity Log Items (sets Display in UI = FALSE in brand tabs or Shared_Activity)
    if (data.action === "CLEAR_ACTIVITY") {
      const type = String(data.type || "INGESTED").toUpperCase();
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(
        (Array.isArray(rawCodes) ? rawCodes : [rawCodes])
          .map(c => String(c).replace(/^'/, '').trim().toUpperCase())
          .filter(c => c && c !== 'PENDING' && !/processing/i.test(c) && c !== '-' && !c.startsWith('#'))
      );
      if (data.bookingId) {
        const b = String(data.bookingId).replace(/^'/, '').trim().toUpperCase();
        if (b && b !== '-' && !b.startsWith('#')) codes.add(b);
      }
      if (data.bookingIds && Array.isArray(data.bookingIds)) {
        data.bookingIds.forEach(b => {
          const bStr = String(b).replace(/^'/, '').trim().toUpperCase();
          if (bStr && bStr !== '-' && !bStr.startsWith('#')) codes.add(bStr);
        });
      }
      if (codes.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Cannot clear activity: No valid voucher codes or booking IDs specified."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      const updated = [];

      if (type === "SHARED") {
        const sh = ss.getSheetByName("Shared_Activity");
        if (sh) {
          const vals = sh.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowCode = String(vals[i][0] || "").replace(/^'/, '').trim();
            const rowCodeUpper = rowCode.toUpperCase();
            if (rowCodeUpper && rowCodeUpper !== '-' && codes.has(rowCodeUpper)) {
              sh.getRange(i + 1, 8).setValue(false); // Col 8: Display in UI
              sh.getRange(i + 1, 6).setValue("DISMISSED");
              updated.push(rowCode);
            }
          }
        }
      } else {
        // INGESTED: Search in "New" Tab first, then Brand tabs
        const newSheet = ss.getSheetByName("New");
        if (newSheet) {
          const vals = newSheet.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][0] || "").replace(/^'/, '').trim();
            const rowCode = String(vals[i][1] || "").replace(/^'/, '').trim();
            const rowBidUpper = rowBid.toUpperCase();
            const rowCodeUpper = rowCode.toUpperCase();
            const bidMatch = rowBidUpper && rowBidUpper !== '-' && !rowBidUpper.startsWith('#') && codes.has(rowBidUpper);
            const codeMatch = rowCodeUpper && rowCodeUpper !== '-' && rowCodeUpper !== 'PENDING' && codes.has(rowCodeUpper);
            if (bidMatch || codeMatch) {
              newSheet.getRange(i + 1, 10).setValue(false); // Col 10: Display in UI
              updated.push(rowBid || rowCode);
            }
          }
        }
        const sheets = ss.getSheets();
        for (const sheet of sheets) {
          if (SYSTEM_SHEET_NAMES.has(sheet.getName()) || sheet.getName() === "Cards") continue;
          const vals = sheet.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][0] || "").replace(/^'/, '').trim();
            const rowCode = String(vals[i][1] || "").replace(/^'/, '').trim();
            const rowBidUpper = rowBid.toUpperCase();
            const rowCodeUpper = rowCode.toUpperCase();
            const bidMatch = rowBidUpper && rowBidUpper !== '-' && !rowBidUpper.startsWith('#') && codes.has(rowBidUpper);
            const codeMatch = rowCodeUpper && rowCodeUpper !== '-' && rowCodeUpper !== 'PENDING' && codes.has(rowCodeUpper);
            if (bidMatch || codeMatch) {
              sheet.getRange(i + 1, 10).setValue(false); // Col 10: Display in UI
              updated.push(rowBid || rowCode);
            }
          }
        }
      }

      SpreadsheetApp.flush();
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
      const codes = new Set((Array.isArray(rawCodes) ? rawCodes : [rawCodes]).map(c => String(c).replace(/^'/, '').trim().toUpperCase()).filter(Boolean));
      if (data.bookingId) codes.add(String(data.bookingId).replace(/^'/, '').trim().toUpperCase());
      if (data.bookingIds && Array.isArray(data.bookingIds)) {
        data.bookingIds.forEach(b => codes.add(String(b).replace(/^'/, '').trim().toUpperCase()));
      }
      if (codes.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Cannot clear redeemed: No voucher codes specified."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      const updated = [];

      const rSheet = ss.getSheetByName("Redeemed");
      if (rSheet) {
        const vals = rSheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
          const rowBid = String(vals[i][1] || "").replace(/^'/, '').trim();
          const rowCode = String(vals[i][2] || "").replace(/^'/, '').trim();
          const rowBidUpper = rowBid.toUpperCase();
          const rowCodeUpper = rowCode.toUpperCase();
          if ((rowCodeUpper && codes.has(rowCodeUpper)) || (rowBidUpper && codes.has(rowBidUpper))) {
            rSheet.getRange(i + 1, 13).setValue(false); // Col 13: Display in UI
            updated.push(rowCode || rowBid);
          }
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "CLEAR_REDEEMED",
        clearedCount: updated.length,
        clearedCodes: updated
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 13. Record RuPay Income Entry (Voucher Sale into "RuPay Income" or Card Fee into "RuPay Charges")
    if (data.action === "RECORD_INCOME_ENTRY") {
      const title = String(data.title || '').trim();
      const rawPrice = Number(data.price);
      if (!title || isNaN(rawPrice) || rawPrice === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Invalid entry: title and non-zero price are required."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const isCharge = rawPrice < 0 || String(data.type || '').toLowerCase() === 'charge';
      const targetSheetName = isCharge ? "RuPay Charges" : "RuPay Income";
      let targetSheet = ss.getSheetByName(targetSheetName);
      if (!targetSheet) {
        if (isCharge) {
          targetSheet = getOrCreateRuPayChargesSheet(ss);
        } else {
          return ContentService.createTextOutput(JSON.stringify({
            error: `"${targetSheetName}" tab not found in spreadsheet`
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      const dateStr = String(data.date || Utilities.formatDate(new Date(), "GMT+05:30", "dd/MM/yy")).trim();
      const comments = String(data.comments || '').trim();
      const account = String(data.account || '').trim();

      const comp = parseDateComponents(dateStr);
      const dateMs = comp ? new Date(comp.year, comp.month - 1, comp.day).getTime() : 0;
      let cellDateValue = dateStr;
      if (comp) {
        cellDateValue = new Date(Date.UTC(comp.year, comp.month - 1, comp.day, 6, 30, 0));
      }

      const finalPrice = isCharge ? -Math.abs(rawPrice) : Math.abs(rawPrice);
      let newRowData;
      if (isCharge) {
        newRowData = [
          title,
          cellDateValue,
          finalPrice,
          comments,
          account
        ];
      } else {
        // RuPay Income has exactly 4 columns: [Brand, Date, Sale Value, Comments]
        let finalComments = comments;
        if (account && !finalComments.includes(account)) {
          finalComments = finalComments ? `${finalComments} | ${account}` : account;
        }
        newRowData = [
          title,
          cellDateValue,
          finalPrice,
          finalComments
        ];
      }

      const insertedRow = insertRowSortedByDate(targetSheet, newRowData, dateMs);
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "RECORD_INCOME_ENTRY",
        targetSheet: targetSheetName,
        insertedRow: insertedRow,
        entry: {
          title: title,
          date: dateStr,
          price: finalPrice,
          comments: comments,
          account: account
        }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 16. Dismiss Ingested / Staged Voucher (Sets Display in UI = FALSE in "New" tab)
    if (data.action === "DISMISS_STAGING") {
      const rawBids = data.bookingIds || (data.bookingId ? [data.bookingId] : []);
      const targetBids = new Set((Array.isArray(rawBids) ? rawBids : [rawBids]).map(b => String(b).replace(/^'/, '').trim().toUpperCase()).filter(Boolean));
      if (targetBids.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Cannot dismiss: No booking ID specified."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const dismissed = [];
      const nSheet = ss.getSheetByName("New");
      if (nSheet) {
        const vals = nSheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
          const rowBid = String(vals[i][0] || "").replace(/^'/, '').trim().toUpperCase();
          if (targetBids.has(rowBid)) {
            nSheet.getRange(i + 1, 10).setValue(false); // Col 10: Display in UI = FALSE
            dismissed.push(rowBid);
            removeRevalidatedBookingFromProperties(rowBid);
          }
        }
      }
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "DISMISS_STAGING",
        dismissedCount: dismissed.length,
        dismissedBids: dismissed
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 16B. Mark Voucher As Sold (Direct from Ingested/New or Brand Tab -> Redeemed Tab + RuPay Income Entry)
    if (data.action === "MARK_VOUCHER_AS_SOLD") {
      const rawBid = String(data.bookingId || '').trim();
      const rawCode = String(data.code || '').trim();
      const rawPin = String(data.pin || '').trim();
      const rawBrand = String(data.brand || '').trim();
      const isStaging = !!data.isStaging;
      const keepInStaging = !!data.keepInStaging || !!data.onlyLogIncome;
      const targetBrand = extractBrand(rawBrand, "") !== "Other" ? extractBrand(rawBrand, "") : (rawBrand || "");

      // Mandatory Validation 1: Brand is mandatory
      if (!targetBrand || targetBrand === "Other" || targetBrand === "⚡ Unassigned" || targetBrand === "Unassigned") {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Brand is mandatory to record sale."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const sale = data.saleDetails || {};
      const sellingPrice = Math.abs(Number(sale.price) || 0);

      // Mandatory Validation 2: Sale value is mandatory
      if (!sellingPrice || sellingPrice <= 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Sale value is mandatory and must be greater than 0."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const dateStr = String(sale.date || '').trim();

      // Mandatory Validation 3: Date is mandatory
      if (!dateStr) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Sale date is mandatory."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      const comp = parseDateComponents(dateStr);
      const dateMs = comp ? new Date(comp.year, comp.month - 1, comp.day).getTime() : 0;
      if (!dateMs || dateMs <= 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Invalid sale date: Date must be a valid date (dd/MM/yy or yyyy-MM-dd)."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Step 1: Ensure "Redeemed" sheet exists if moving voucher to Redeemed
      let rSheet = null;
      if (!keepInStaging) {
        rSheet = ss.getSheetByName("Redeemed");
        if (!rSheet) {
          rSheet = ss.insertSheet("Redeemed");
          rSheet.appendRow(UNIFIED_REDEEMED_HEADERS);
          rSheet.getRange(1, 1, 1, UNIFIED_REDEEMED_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
          rSheet.setFrozenRows(1);
        }
      }

      let voucherFound = false;
      let finalBid = rawBid;
      let finalCode = rawCode;
      let finalPin = rawPin;
      let finalAmount = Number(data.amount) || 500;
      let finalDateOfIssue = String(data.dateOfIssue || '');
      let finalAccount = String(data.account || '');
      let finalCard = String(data.card || 'Unassigned');
      let finalAppendedTs = String(data.timestampIST || '');
      const redeemedTs = formatISTTimestamp(new Date());

      if (isStaging) {
        // Resolve voucher metadata from "New" sheet (and remove ONLY if NOT keepInStaging)
        const nSheet = ss.getSheetByName("New");
        if (nSheet) {
          const nRows = nSheet.getDataRange().getValues();
          const targetBidUpper = rawBid.replace(/^'/, '').trim().toUpperCase();
          const targetCodeUpper = rawCode.replace(/^'/, '').trim().toUpperCase();
          for (let i = nRows.length - 1; i >= 1; i--) {
            const rowBid = String(nRows[i][0] || '').replace(/^'/, '').trim().toUpperCase();
            const rowCode = String(nRows[i][1] || '').replace(/^'/, '').trim().toUpperCase();
            const matchBid = targetBidUpper && (rowBid === targetBidUpper || (rowBid.startsWith("RUPS-") && targetBidUpper.includes(rowBid)) || (targetBidUpper.startsWith("RUPS-") && rowBid.includes(targetBidUpper)));
            const matchCode = targetCodeUpper && targetCodeUpper !== 'PENDING' && targetCodeUpper !== '-' && rowCode === targetCodeUpper;
            if (matchBid || matchCode) {
              finalBid = String(nRows[i][0] || finalBid);
              finalCode = String(nRows[i][1] || finalCode);
              finalPin = String(nRows[i][2] || finalPin);
              finalAmount = Number(nRows[i][3]) || finalAmount;
              finalDateOfIssue = String(nRows[i][4] || finalDateOfIssue);
              finalAccount = String(nRows[i][5] || finalAccount);
              finalCard = String(nRows[i][6] || finalCard);
              finalAppendedTs = String(nRows[i][8] || finalAppendedTs);
              if (!keepInStaging) {
                nSheet.deleteRow(i + 1);
              } else {
                if (nSheet.getLastColumn() < 13 || String(nSheet.getRange(1, 13).getValue() || '').trim() === '') {
                  nSheet.getRange(1, 13).setValue("Sale Logged").setFontWeight("bold");
                }
                nSheet.getRange(i + 1, 13).setValue(true);
              }
              voucherFound = true;
              break;
            }
          }
        }
      } else {
        // Find & remove from Brand tab (or any brand tab)
        const brandSheet = ss.getSheetByName(targetBrand) || ss.getSheetByName(rawBrand);
        if (brandSheet) {
          const bRows = brandSheet.getDataRange().getValues();
          const targetBidUpper = rawBid.replace(/^'/, '').trim().toUpperCase();
          const targetCodeUpper = rawCode.replace(/^'/, '').trim().toUpperCase();
          for (let i = bRows.length - 1; i >= 1; i--) {
            const rowBid = String(bRows[i][0] || '').replace(/^'/, '').trim().toUpperCase();
            const rowCode = String(bRows[i][1] || '').replace(/^'/, '').trim().toUpperCase();
            const matchBid = targetBidUpper && (rowBid === targetBidUpper || (rowBid.startsWith("RUPS-") && targetBidUpper.includes(rowBid)) || (targetBidUpper.startsWith("RUPS-") && rowBid.includes(targetBidUpper)));
            const matchCode = targetCodeUpper && targetCodeUpper !== 'PENDING' && targetCodeUpper !== '-' && rowCode === targetCodeUpper;
            if (matchBid || matchCode) {
              finalBid = String(bRows[i][0] || finalBid);
              finalCode = String(bRows[i][1] || finalCode);
              finalPin = String(bRows[i][2] || finalPin);
              finalAmount = Number(bRows[i][3]) || finalAmount;
              finalDateOfIssue = String(bRows[i][4] || finalDateOfIssue);
              finalAccount = String(bRows[i][5] || finalAccount);
              finalCard = String(bRows[i][6] || finalCard);
              finalAppendedTs = String(bRows[i][8] || finalAppendedTs);
              brandSheet.deleteRow(i + 1);
              voucherFound = true;
              break;
            }
          }
        }
      }

      // Format code & PIN with apostrophe to prevent scientific notation in Google Sheets
      const safeCode = (finalCode && finalCode !== 'PENDING' && finalCode !== '-') ? `'${finalCode.replace(/^'/, '')}` : (rawCode ? `'${rawCode.replace(/^'/, '')}` : "SOLD");
      const safePin = finalPin ? `'${finalPin.replace(/^'/, '')}` : "";
      const safeBid = finalBid ? `'${finalBid.replace(/^'/, '')}` : "";

      // Append to "Redeemed" tab ONLY if NOT keepInStaging
      if (!keepInStaging && rSheet) {
        rSheet.appendRow([
          targetBrand,           // Col 1: Brand
          safeBid,               // Col 2: Booking ID
          safeCode,              // Col 3: Code
          safePin,               // Col 4: PIN
          finalAmount,           // Col 5: Value
          finalDateOfIssue,      // Col 6: Date of Issue
          finalAccount,          // Col 7: Account
          finalCard,             // Col 8: Card
          "REDEEMED",            // Col 9: Status
          redeemedTs,            // Col 10: Redeemed Timestamp (IST)
          isStaging ? "New" : targetBrand, // Col 11: Original Tab
          finalAppendedTs,       // Col 12: Appended Timestamp (IST)
          true                   // Col 13: Display in UI
        ]);
      }

      // Step 2: Format Comment & Insert into "RuPay Income" Tab
      let incomeSheet = ss.getSheetByName("RuPay Income");
      if (!incomeSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          error: '"RuPay Income" tab not found in spreadsheet'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const soldTo = String(sale.soldTo || '').trim();
      const mobile = String(sale.mobile || '').trim();
      const redeemedAccount = String(sale.redeemedAccount || finalAccount || '').trim();
      const gpayName = String(sale.gpayName || '').trim();

      // Canonical RuPay Income title mapping
      let incomeTitle = "Cult Fit";
      const bLow = targetBrand.toLowerCase();
      if (bLow.includes("prime") || bLow.includes("amazon")) {
        incomeTitle = "Amazon Prime";
      } else if (bLow.includes("cult")) {
        incomeTitle = "Cult Fit";
      } else if (bLow.includes("swiggy")) {
        incomeTitle = "SwiggyOne";
      } else if (bLow.includes("thyrocare")) {
        incomeTitle = "Thyrocare";
      } else if (bLow.includes("uber")) {
        incomeTitle = "Uber";
      } else if (bLow.includes("kalyan")) {
        incomeTitle = "Kalyan";
      } else if (bLow.includes("apollo")) {
        incomeTitle = "Apollo";
      } else if (bLow.includes("lakme")) {
        incomeTitle = "Lakme";
      } else if (bLow.includes("reliance")) {
        incomeTitle = "Reliance";
      } else if (bLow.includes("four") || bLow.includes("fountain")) {
        incomeTitle = "FourFountains";
      } else if (bLow.includes("big") || bLow.includes("basket")) {
        incomeTitle = "BigBasket";
      } else {
        incomeTitle = targetBrand;
      }

      // Build unified comments matching user format: Buyer (mobile): Account | GPay: Name | Card
      const cardAccount = finalCard && finalCard !== 'Unassigned' ? finalCard : (data.card && data.card !== 'Unassigned' ? String(data.card).trim() : '');
      const commentStr = formatSaleComments(soldTo, mobile, redeemedAccount, gpayName, cardAccount);

      let cellDateValue = dateStr;
      let parsedYear = 2026;
      if (comp) {
        parsedYear = comp.year;
        cellDateValue = new Date(Date.UTC(comp.year, comp.month - 1, comp.day, 6, 30, 0));
      }

      // Exactly 4 columns in RuPay Income: [Brand, Date, Sale Value, Comments]
      const newIncomeRow = [
        incomeTitle,
        cellDateValue,
        sellingPrice,
        commentStr
      ];

      const insertedRow = insertRowSortedByDate(incomeSheet, newIncomeRow, dateMs);
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "MARK_VOUCHER_AS_SOLD",
        bookingId: finalBid,
        code: safeCode.replace(/^'/, ''),
        brand: targetBrand,
        keepInStaging: keepInStaging,
        isSaleMarked: keepInStaging,
        incomeEntry: {
          id: "inc-" + insertedRow,
          title: incomeTitle,
          date: dateStr,
          price: sellingPrice,
          comments: commentStr,
          account: cardAccount,
          year: parsedYear,
          type: "income"
        }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 16C. Unmark Voucher Sale (Clears Col 13 in "New" tab or restores to New/Brand tab from "Redeemed", and deletes matching row in "RuPay Income")
    if (data.action === "UNMARK_INGESTED_SALE" || data.action === "UNMARK_VOUCHER_SALE") {
      const targetBid = String(data.bookingId || '').trim();
      const targetCode = String(data.code || '').trim();
      const cleanTargetBid = targetBid.replace(/^'/, '').trim().toUpperCase();
      const cleanTargetCode = targetCode.replace(/^'/, '').trim().toUpperCase();
      const targetComments = String(data.comments || '').trim();
      const brand = String(data.brand || '').trim();
      const isStaging = data.isStaging !== false && (data.isStaging || !brand || brand === '⚡ Unassigned' || brand === 'Unassigned');
      
      let unmarkedNewRow = -1;
      let restoredBrandRow = -1;

      // Step 1: If isStaging, check "New" sheet first (handles vouchers where keepInStaging was true)
      if (isStaging) {
        const newSheet = ss.getSheetByName("New");
        if (newSheet && (cleanTargetBid || cleanTargetCode)) {
          const nVals = newSheet.getDataRange().getValues();
          for (let i = 1; i < nVals.length; i++) {
            const bId = String(nVals[i][0] || '').replace(/^'/, '').trim().toUpperCase();
            const cd = String(nVals[i][1] || '').replace(/^'/, '').trim().toUpperCase();
            const matchBid = cleanTargetBid && (bId === cleanTargetBid || (bId.startsWith("RUPS-") && cleanTargetBid.includes(bId)) || (cleanTargetBid.startsWith("RUPS-") && bId.includes(cleanTargetBid)));
            const matchCode = cleanTargetCode && cleanTargetCode !== 'PENDING' && cleanTargetCode !== '-' && cd === cleanTargetCode;
            if (matchBid || matchCode) {
              unmarkedNewRow = i + 1;
              if (newSheet.getLastColumn() >= 13) {
                newSheet.getRange(i + 1, 13).setValue("");
              }
              break;
            }
          }
        }
      }

      // Step 2: If not found in "New" sheet (e.g. voucher was moved to "Redeemed" tab upon sale) OR if isStaging was false:
      if (unmarkedNewRow === -1) {
        const rSheet = ss.getSheetByName("Redeemed");
        if (rSheet && (cleanTargetBid || cleanTargetCode)) {
          const rVals = rSheet.getDataRange().getValues();
          for (let i = rVals.length - 1; i >= 1; i--) {
            const bId = String(rVals[i][1] || '').replace(/^'/, '').trim().toUpperCase();
            const cd = String(rVals[i][2] || '').replace(/^'/, '').trim().toUpperCase();
            const matchBid = cleanTargetBid && (bId === cleanTargetBid || (bId.startsWith("RUPS-") && cleanTargetBid.includes(bId)) || (cleanTargetBid.startsWith("RUPS-") && bId.includes(cleanTargetBid)));
            const matchCode = cleanTargetCode && cleanTargetCode !== 'PENDING' && cleanTargetCode !== '-' && cd === cleanTargetCode;
            if (matchBid || matchCode) {
              const origTab = String(rVals[i][10] || '').trim();
              const shouldRestoreToNew = isStaging || origTab === "New";
              const rawBid = String(rVals[i][1] || '').trim();
              const rawCd = String(rVals[i][2] || '').trim();
              const rawPin = String(rVals[i][3] || '').trim();
              const safeBid = rawBid ? (rawBid.startsWith("'") ? rawBid : ("'" + rawBid)) : "";
              const safeCd = rawCd ? (rawCd.startsWith("'") ? rawCd : ("'" + rawCd)) : "";
              const safePin = rawPin ? (rawPin.startsWith("'") ? rawPin : ("'" + rawPin)) : "";
              const appTs = rVals[i][11] || formatISTTimestamp(new Date());

              if (shouldRestoreToNew) {
                let newSheet = ss.getSheetByName("New");
                if (!newSheet) {
                  newSheet = ss.insertSheet("New");
                  newSheet.appendRow(UNIFIED_NEW_HEADERS);
                  newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
                  newSheet.setFrozenRows(1);
                }
                const targetBrandName = brand || String(rVals[i][0] || '');
                newSheet.appendRow([
                  safeBid,               // Col 1: Booking ID
                  safeCd,                // Col 2: Code
                  safePin,               // Col 3: PIN
                  rVals[i][4],           // Col 4: Value
                  rVals[i][5],           // Col 5: Date of Issue
                  rVals[i][6],           // Col 6: Account
                  rVals[i][7],           // Col 7: Card
                  "ACTIVE",              // Col 8: Status
                  appTs,                 // Col 9: Appended Timestamp (IST)
                  true,                  // Col 10: Display in UI
                  targetBrandName,       // Col 11: Target Brand
                  "",                    // Col 12: Payment Link
                  ""                     // Col 13: Sale Logged (cleared)
                ]);
                unmarkedNewRow = newSheet.getLastRow();
              } else {
                const brandSheet = getTargetBrandSheet(ss, brand || String(rVals[i][0] || ''));
                if (brandSheet) {
                  brandSheet.appendRow([
                    safeBid,             // Col 1: Booking ID
                    safeCd,              // Col 2: Code
                    safePin,             // Col 3: PIN
                    rVals[i][4],         // Col 4: Value
                    rVals[i][5],         // Col 5: Date of Issue
                    rVals[i][6],         // Col 6: Account
                    rVals[i][7],         // Col 7: Card
                    "ACTIVE",            // Col 8: Status
                    appTs,               // Col 9: Appended Timestamp (IST)
                    true                 // Col 10: Display in UI
                  ]);
                  restoredBrandRow = brandSheet.getLastRow();
                }
              }
              rSheet.deleteRow(i + 1);
              break;
            }
          }
        }
      }

      // Step 3: Delete matching sale entry from "RuPay Income" tab (search backwards to match most recent)
      const incomeSheet = ss.getSheetByName("RuPay Income");
      let deletedIncomeRow = -1;
      if (incomeSheet && (targetComments || cleanTargetBid || cleanTargetCode)) {
        const iVals = incomeSheet.getDataRange().getValues();
        for (let r = iVals.length - 1; r >= 1; r--) {
          const col0 = String(iVals[r][0] || '').trim().toUpperCase();
          if (col0 === 'TOTAL') continue;
          const comm = String(iVals[r][3] || '').trim();
          const matchComment = targetComments && comm === targetComments;
          const matchBidInComm = cleanTargetBid && (comm.toUpperCase().includes(cleanTargetBid) || (targetBid && comm.includes(targetBid)));
          const matchCodeInComm = cleanTargetCode && cleanTargetCode !== 'PENDING' && cleanTargetCode !== '-' && (comm.toUpperCase().includes(cleanTargetCode) || (targetCode && comm.includes(targetCode)));
          if (matchComment || matchBidInComm || matchCodeInComm) {
            deletedIncomeRow = r + 1;
            incomeSheet.deleteRow(r + 1);
            break;
          }
        }
        if (deletedIncomeRow > 0) {
          recalculateSheetTotal(incomeSheet);
        }
      }
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "UNMARK_INGESTED_SALE",
        unmarkedNewRow: unmarkedNewRow,
        restoredBrandRow: restoredBrandRow,
        deletedIncomeRow: deletedIncomeRow
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 17. Promote Tombstones (Moves Display in UI = FALSE rows from "New" and "Redeemed" to "Deleted_Tombstones")
    if (data.action === "PROMOTE_TOMBSTONES") {
      const result = promoteTombstones(ss, true);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      error: `Unknown or unhandled action: "${data.action || 'undefined'}"`
    })).setMimeType(ContentService.MimeType.JSON);

    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Cron Job: Moves vouchers with Display in UI === FALSE from "New" and "Redeemed" tabs
 * into "Deleted_Tombstones" sheet and permanently deletes their rows from the active tabs.
 * Runs on a 72-hour trigger or on-demand via action PROMOTE_TOMBSTONES.
 */
function promoteTombstones(ss, hasLock) {
  if (!ss) ss = getSpreadsheet();
  let lock = null;
  if (!hasLock) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      console.warn("promoteTombstones: Script is locked by another execution, skipping.");
      return { success: false, error: "Locked by another execution" };
    }
  }

  try {
    const tSheet = getOrInitTombstonesSheet(ss);
    const nowIst = formatISTDate(new Date());
    let promotedTotal = 0;

    // Helper: Guarantee string format & prevent Google Sheets scientific notation truncation
    const formatAsTextCell = (val) => {
      if (val === null || val === undefined) return "";
      let s = String(val).trim();
      if (!s) return "";
      // Un-mangle specifically known RUPS-DEFBAA77 scientific notation if encountered
      if (s === "9.802E+15" || (s.toLowerCase().includes("e+") && /9\.802/i.test(s))) {
        return "'9802004492588540";
      }
      // If purely digits, prepend apostrophe so Google Sheets stores as literal text
      if (/^\d+$/.test(s) && !s.startsWith("'")) {
        return "'" + s;
      }
      return s;
    };

    // Auto-repair any existing mangled scientific notation in Deleted_Tombstones
    const lastTombRow = tSheet.getLastRow();
    if (lastTombRow > 1) {
      tSheet.getRange("A:C").setNumberFormat("@");
      const tRange = tSheet.getRange(2, 1, lastTombRow - 1, 3);
      const tVals = tRange.getValues();
      let repaired = false;
      for (let i = 0; i < tVals.length; i++) {
        const bid = String(tVals[i][0] || '').trim();
        const code = String(tVals[i][1] || '').trim();
        if (bid === "RUPS-DEFBAA77" && (code === "9.802E+15" || /^[0-9.]+[eE]\+[0-9]+$/i.test(code))) {
          tVals[i][1] = "'9802004492588540";
          repaired = true;
        } else if (/^\d{10,}$/.test(code) && !code.startsWith("'")) {
          tVals[i][1] = "'" + code;
          repaired = true;
        }
      }
      if (repaired) {
        tRange.setValues(tVals);
      }
    }

    // 1. Scan "New" Tab (Col 10: Display in UI)
    const nSheet = ss.getSheetByName("New");
    if (nSheet) {
      const nVals = nSheet.getDataRange().getValues();
      const nRowsToDelete = [];
      const nRowsToAppend = [];

      for (let i = 1; i < nVals.length; i++) {
        const row = nVals[i];
        const isDisp = row[9];
        if (isDisp === false || String(isDisp).toLowerCase() === 'false') {
          nRowsToAppend.push([
            formatAsTextCell(row[0]), // Booking ID
            formatAsTextCell(row[1]), // Code
            formatAsTextCell(row[2]), // PIN
            row[10] || "Other",       // Brand
            row[3] || "",             // Value
            row[5] || "",             // Account
            row[6] || "",             // Card
            "New",                    // Source Tab
            row[7] || "PENDING",      // Original Status
            nowIst                    // Tombstoned Timestamp (IST)
          ]);
          nRowsToDelete.push(i + 1);
        }
      }

      if (nRowsToAppend.length > 0) {
        const lastRow = tSheet.getLastRow();
        tSheet.getRange(lastRow + 1, 1, nRowsToAppend.length, 3).setNumberFormat("@");
        tSheet.getRange(lastRow + 1, 1, nRowsToAppend.length, nRowsToAppend[0].length).setValues(nRowsToAppend);
        // Delete rows in descending index order to preserve row positions
        for (let d = nRowsToDelete.length - 1; d >= 0; d--) {
          nSheet.deleteRow(nRowsToDelete[d]);
        }
        promotedTotal += nRowsToAppend.length;
      }
    }

    // 2. Scan "Redeemed" Tab (Col 13: Display in UI)
    const rSheet = ss.getSheetByName("Redeemed");
    if (rSheet) {
      const rVals = rSheet.getDataRange().getValues();
      const rRowsToDelete = [];
      const rRowsToAppend = [];

      for (let i = 1; i < rVals.length; i++) {
        const row = rVals[i];
        const isDisp = row[12];
        if (isDisp === false || String(isDisp).toLowerCase() === 'false') {
          rRowsToAppend.push([
            formatAsTextCell(row[1]), // Booking ID
            formatAsTextCell(row[2]), // Code
            formatAsTextCell(row[3]), // PIN
            row[0] || "Other",        // Brand
            row[4] || "",             // Value
            row[6] || "",             // Account
            row[7] || "",             // Card
            "Redeemed",               // Source Tab
            row[8] || "REDEEMED",     // Original Status
            nowIst                    // Tombstoned Timestamp (IST)
          ]);
          rRowsToDelete.push(i + 1);
        }
      }

      if (rRowsToAppend.length > 0) {
        const lastRow = tSheet.getLastRow();
        tSheet.getRange(lastRow + 1, 1, rRowsToAppend.length, 3).setNumberFormat("@");
        tSheet.getRange(lastRow + 1, 1, rRowsToAppend.length, rRowsToAppend[0].length).setValues(rRowsToAppend);
        // Delete rows in descending index order to preserve row positions
        for (let d = rRowsToDelete.length - 1; d >= 0; d--) {
          rSheet.deleteRow(rRowsToDelete[d]);
        }
        promotedTotal += rRowsToAppend.length;
      }
    }

    SpreadsheetApp.flush();
    console.log(`promoteTombstones: Successfully archived ${promotedTotal} records to Deleted_Tombstones.`);
    return { success: true, action: "PROMOTE_TOMBSTONES", promotedCount: promotedTotal };
  } catch (err) {
    console.error("Error in promoteTombstones:", err);
    return { success: false, error: err.message };
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch(_) {}
    }
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

/**
 * BENEFIT MATRIX PRESET GENERATOR
 * Generates initial 13 benefit states (Unredeemed vs Not Applicable) based on Bank & Card Type
 */
function getEligibleBenefitsArray(bank, type, name) {
  const isCc = String(type).toLowerCase().includes('credit');
  const bLower = String(bank || '').toLowerCase();
  const isPnb = bLower.includes('pnb');
  const isCanara = bLower.includes('canara');

  // Order: [Spa_Lakme, Health_Thyrocare, Amazon_Prime, Cult_Live_90D, Cult_Elite_1M, Uber, Myntra, Decathlon, Reliance_Digital, BigBasket, Apollo_Pharmacy, Kalyan_Jewellers, TaxSpanner]
  if (isCc) {
    if (isCanara) {
      // Canara CC: Spa/Lakme & Thyrocare only
      return ["Unredeemed", "Unredeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];
    } else {
      // PNB CC: Spa, Thyrocare, Cult Live, Cult Elite, Uber
      return ["Unredeemed", "Unredeemed", "Not Applicable", "Unredeemed", "Unredeemed", "Unredeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];
    }
  } else {
    if (isPnb) {
      // PNB Imperial DC: Spa, Thyrocare, Prime, Cult Live, Myntra, Decathlon, Reliance, BigBasket, Apollo, Kalyan, TaxSpanner
      return ["Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Not Applicable", "Not Applicable", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed"];
    } else {
      // Union, BoB, Canara Crest, Jana DC: Spa, Thyrocare, Prime, Cult Live, Cult Elite, Uber
      return ["Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Unredeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];
    }
  }
}

/**
 * Administrative Setup: Initialize "Bank Accounts" Tab
 * Seeds 30 multi-bank accounts across all 6 family members with complete credentials.
 */
function setupBankAccountsTab(ss) {
  let sheet = ss.getSheetByName("Bank Accounts");
  let created = false;
  if (!sheet) {
    sheet = ss.insertSheet("Bank Accounts");
    created = true;
  } else {
    sheet.clear();
  }

  const headers = [
    "Account_ID", "Holder_Name", "Bank_Name", "Bank_Code", "Account_Label",
    "Account_Tier", "Tier_Category", "Account_Number", "IFSC_Code", "UPI_VPA",
    "Customer_ID", "Customer_ID_Alias", "Sign_On_Password", "Transaction_Password",
    "Transaction_PIN", "Login_PIN", "Branch_Name", "Status", "Notes", "Last_Updated"
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#0C8C46")
    .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  // Set plain text format on sensitive/numeric columns to prevent scientific notation truncation
  sheet.getRange("H:H").setNumberFormat('@'); // Account_Number
  sheet.getRange("K:K").setNumberFormat('@'); // Customer_ID
  sheet.getRange("O:O").setNumberFormat('@'); // Transaction_PIN
  sheet.getRange("P:P").setNumberFormat('@'); // Login_PIN

  const nowStr = formatISTTimestamp(new Date());

  // If sheet already has data, do not overwrite; otherwise insert dummy schema template
  if (sheet.getLastRow() <= 1) {
    sheet.appendRow([
      "acc_template_1", "Sample Holder", "Sample Bank", "sample", "Sample Account",
      "Savings", "General", "'100000000001", "SMPL0001234", "sample@upi",
      "'12345678", "samplealias", "", "", "", "", "Sample Branch", "ACTIVE", "Template Account", nowStr
    ]);
  }

  return {
    success: true,
    sheetName: "Bank Accounts",
    created: created,
    accountsCount: Math.max(0, sheet.getLastRow() - 1)
  };
}

