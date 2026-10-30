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
  "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI", "Target Brand"
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
  "Thyrocare", "Uber", "Decathlon", "SwiggyOne"
]);

function getTargetBrandSheet(ss, brand) {
  const canonicalBrand = extractBrand(brand, "");
  const targetName = canonicalBrand !== "Other" ? canonicalBrand : (brand || "⚡ Unassigned");

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

  // Strict Whitelist: If brand is not a recognized canonical tab, route to ⚡ Unassigned (Never create new tabs)
  let unassigned = ss.getSheetByName("⚡ Unassigned");
  if (!unassigned) {
    unassigned = ss.insertSheet("⚡ Unassigned");
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

  // 1. Guard: Must contain RUPS- booking reference
  const bidMatch = subject.match(/\b(RUPS-[A-Za-z0-9]+)\b/i) || body.match(/\b(RUPS-[A-Za-z0-9]+)\b/i);
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
  const amount = extractAmount(body);
  const dateOfIssue = extractField(body, "Date of Issue") || formatISTDate(new Date());
  const account = extractAccountName(msg.getTo(), body, msg.getFrom()) || "vibhavoswal9";
  const receivedDate = (msg && typeof msg.getDate === 'function') ? msg.getDate() : new Date();

  // 3B. Guard: Is this an "On Hold – Validation Pending / Payment Link Sent" email?
  const isRevalMatch = /on\s*hold|validation\s*pending|payment\s*link\s*sent|process\s*the\s*payment/i.test(subject) ||
                       /on\s*hold|validation\s*pending|payment\s*link\s*sent|process\s*the\s*payment/i.test(body);
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
      isRevalidationPending: true
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
  const isDelay = !rawVal || /will be sent|delay|within|hrs|days|pending|\bna\b|n\/a|revalidation/i.test(rawVal) || rawVal.toUpperCase().startsWith("RUPS");
  if (isDelay) {
    return {
      bookingId: bookingId,
      brand: brand,
      amount: amount,
      dateOfIssue: dateOfIssue,
      account: account,
      code: "PENDING",
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

  // Auto-heal corrupted revalidations where code was erroneously set to "/service if the" or matching known booking IDs
  try {
    const nSheet = ss.getSheetByName("New");
    if (nSheet) {
      const nRows = nSheet.getDataRange().getValues();
      for (let i = 1; i < nRows.length; i++) {
        const rowCode = String(nRows[i][1] || '').trim();
        const rowBid = String(nRows[i][0] || '').trim().toUpperCase();
        if (rowCode.toLowerCase().startsWith("/service") || rowBid === "RUPS-34AB2FE3" || rowBid === "RUPS-5ACE0C37" || rowBid === "RUPS-B3366AE3") {
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

    // 0. Auto-consolidate any rogue tabs created previously
    consolidateRogueTabs(ss);

    // Search emails from the last 30 minutes with RUPS reference (read and unread)
    const query = 'newer_than:30m ("RUPS-" OR subject:"RUPS-" OR from:rupay@truztee.com OR "rupay@truztee.com")';
    const threads = GmailApp.search(query, 0, 25);
    if (!threads.length) return;

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const msg of messages) {
        const parsed = parseRupayVoucherEmail(msg);
        if (!parsed) continue;

        if (parsed.isRevalidationPending) {
          // ==========================================
          // STAGE 1.5: On Hold - Validation Pending Email
          // ==========================================
          // Updates Code to "Pending revalidation" and Status to "PENDING_REVALIDATION"
          // Invariant: Leaves initial acknowledgment timestamp untouched!
          recordPendingRevalidation(ss, {
            bookingId: parsed.bookingId,
            brand: parsed.brand,
            amount: parsed.amount,
            account: parsed.account,
            dateOfIssue: parsed.dateOfIssue,
            receivedDate: parsed.receivedDate
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
            receivedDate: parsed.receivedDate
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
  if (combined.includes("reliance")) return "Reliance";
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

function extractAmount(body) {
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
  const cUpper = (code && code !== '-' && code.toUpperCase() !== 'PENDING') ? String(code).trim().toUpperCase() : '';
  if (!bUpper && !cUpper) return null;

  const checkSheet = (sheet) => {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      const rowBid = String(vals[i][0] || '').trim().toUpperCase();
      const rowCode = String(vals[i][1] || '').trim().toUpperCase();
      const hasRups = bUpper && bUpper.startsWith('RUPS-');
      const bidMatch = hasRups && rowBid === bUpper;
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
  const systemTabs = new Set(["New", "Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns"]);
  for (const s of sheets) {
    const sName = s.getName();
    if (systemTabs.has(sName) || sName === targetBrand) continue;
    const res = checkSheet(s);
    if (res) return res;
  }

  return null;
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
    for (let i = 1; i < rRows.length; i++) {
      if (String(rRows[i][1]).trim() === data.bookingId) return;
    }
  }

  // Determine canonical target brand (Col 11)
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "⚡ Unassigned");

  // Hardening against duplicate re-staging: Check if Booking ID already exists in Brand tabs!
  const existingBrandMatch = findRowInBrandTabs(ss, data.bookingId, targetBrand);
  if (existingBrandMatch) {
    // Already approved to brand tab! Do not re-stage in "New".
    return;
  }

  // 1. Search "New" tab by Booking ID
  const newRows = newSheet.getDataRange().getValues();
  for (let i = 1; i < newRows.length; i++) {
    if (String(newRows[i][0]).trim() === data.bookingId) {
      const currentStatus = String(newRows[i][7] || '').trim().toUpperCase();
      const currentCode = String(newRows[i][1] || '').trim();

      // Guard: If already ACTIVE with a real delivered code, do not revert!
      if (currentStatus === "ACTIVE" && currentCode !== "" && currentCode !== "PENDING" && currentCode !== "Pending revalidation") {
        return;
      }

      // Update Code & Status, but leave Col 9 (Timestamp) untouched!
      newSheet.getRange(i + 1, 2).setValue("'Pending revalidation");
      newSheet.getRange(i + 1, 8).setValue("PENDING_REVALIDATION");
      newSheet.getRange(i + 1, 10).setValue(true);
      return;
    }
  }

  // 2. Fallback: If Stage 1 email was missed and validation pending arrives first, create staging row
  const rowData = [
    data.bookingId,                                      // Col 1: Booking ID (RUPS...)
    "'Pending revalidation",                             // Col 2: Code
    "",                                                  // Col 3: PIN
    data.amount || 500,                                  // Col 4: Value
    data.dateOfIssue || "",                              // Col 5: Date of Issue
    data.account || "",                                  // Col 6: Account
    "Unassigned",                                        // Col 7: Card
    "PENDING_REVALIDATION",                              // Col 8: Status
    formatISTTimestamp(data.receivedDate || new Date()), // Col 9: Initial Timestamp (only set once on creation)
    true,                                                // Col 10: Display in UI
    targetBrand                                          // Col 11: Target Brand
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
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rRows.length; i++) {
      if (String(rRows[i][1]).trim() === data.bookingId) {
        // Already redeemed, don't recreate!
        return;
      }
    }
  }

  // 2. Global Idempotency: Does this Booking ID already exist in "New" tab?
  const newRows = newSheet.getDataRange().getValues();
  for (let i = 1; i < newRows.length; i++) {
    if (String(newRows[i][0]).trim() === data.bookingId) {
      // Already recorded in staging, keep initial ack timestamp intact and do not duplicate
      return;
    }
  }

  // 3. Determine canonical target brand (Col 11)
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "⚡ Unassigned");

  // 4. Hardening against duplicate re-staging: Check if Booking ID already exists in Brand tabs!
  const existingBrandMatch = findRowInBrandTabs(ss, data.bookingId, targetBrand);
  if (existingBrandMatch) {
    // Already approved to brand tab! Do not re-stage in "New".
    return;
  }

  const rowData = [
    data.bookingId,                                      // Col 1: Booking ID (RUPS...)
    "PENDING",                                           // Col 2: Code
    "",                                                  // Col 3: PIN
    data.amount || 500,                                  // Col 4: Value
    data.dateOfIssue || "",                              // Col 5: Date of Issue
    data.account || "",                                  // Col 6: Account
    "Unassigned",                                        // Col 7: Card
    "PENDING",                                           // Col 8: Status
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
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === data.bookingId) {
      const currentCode = String(rows[i][1] || '').trim();
      const currentStatus = String(rows[i][7] || '').trim().toUpperCase();

      // Guard: If already ACTIVE and voucher code is populated, it is permanently immutable!
      if (currentStatus === "ACTIVE" && currentCode !== "" && currentCode !== "PENDING" && currentCode !== "Pending revalidation") {
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
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "⚡ Unassigned");

  // Check if already in Redeemed tab strictly by Booking ID
  const rSheet = ss.getSheetByName("Redeemed");
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rRows.length; i++) {
      if (String(rRows[i][1]).trim() === data.bookingId) {
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

    if ((status === 'PENDING' || status === 'PENDING_REVALIDATION' || code === 'PENDING' || code === 'Pending revalidation') && /^RUPS/i.test(bid)) {
      if (currentPin && !/^\d{3,8}$/.test(currentPin)) {
        newSheet.getRange(i + 1, 3).setValue("");
      }

      try {
        const threads = GmailApp.search('"' + bid + '"', 0, 3);
        for (const t of threads) {
          const msgs = t.getMessages();
          for (const m of msgs) {
            const parsed = parseRupayVoucherEmail(m);
            if (parsed && parsed.isDelivered && parsed.bookingId === bid) {
              newSheet.getRange(i + 1, 2).setValue(`'${parsed.code}`);
              newSheet.getRange(i + 1, 3).setValue(parsed.pin ? `'${parsed.pin}` : "");
              if (parsed.amount && Number(parsed.amount) > 0) {
                newSheet.getRange(i + 1, 4).setValue(Number(parsed.amount));
              }
              newSheet.getRange(i + 1, 8).setValue("ACTIVE");
              newSheet.getRange(i + 1, 9).setValue(formatISTTimestamp(parsed.receivedDate || m.getDate()));
              newSheet.getRange(i + 1, 10).setValue(true);
              reconciled.push({ bookingId: bid, code: parsed.code, pin: parsed.pin });
              break;
            }
          }
          if (reconciled.some(r => r.bookingId === bid)) break;
        }
      } catch (e) {
        console.warn("Reconcile error for " + bid + ": " + e.message);
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

      // Benefit Presets based on Benefit Matrix Rules
      // 1. PNB Credit Cards
      const bPnbCc = ["Redeemed", "Redeemed", "Not Applicable", "Redeemed", "Redeemed", "Redeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];
      // 2. Canara Credit Card
      const bCanaraCc = ["Redeemed", "Redeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];
      // 3. PNB Imperial Debit Cards
      const bPnbDc = ["Redeemed", "Redeemed", "Redeemed", "Redeemed", "Not Applicable", "Not Applicable", "Redeemed", "Redeemed", "Redeemed", "Redeemed", "Redeemed", "Redeemed", "Redeemed"];
      // 4. Other Debit Cards (Jana, BoB, Canara Crest, Union Metal)
      const bOtherDc = ["Redeemed", "Redeemed", "Redeemed", "Redeemed", "Redeemed", "Redeemed", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable", "Not Applicable"];

      const rows = [
        // RuPay Credit Cards (4) - Linked Account: "NA"
        ["PNB", "Credit", "PNB RuPay Select (Primary)", "NA", "'6528590001670074", "'10/31", "'324", "ACTIVE", "cult pending", ...bPnbCc, nowStr],
        ["PNB", "Credit", "PNB Sanjay Add On", "NA", "'6528590001670165", "'10/31", "'300", "ACTIVE", "cult pending", ...bPnbCc, nowStr],
        ["PNB", "Credit", "PNB Aruna Add On", "NA", "'6528590001670173", "'10/31", "'639", "ACTIVE", "cult pending", ...bPnbCc, nowStr],
        ["Canara", "Credit", "Canara RuPay Select", "NA", "'6530200329593042", "'10/30", "'870", "ACTIVE", "All done", ...bCanaraCc, nowStr],

        // RuPay Debit Cards (16) - PNB name changed to "PNB Imperial", Linked Account contains last 4 of account
        ["PNB", "Debit", "PNB Imperial", "'6650", "'6524971000855409", "'10/33", "'300", "ACTIVE", "", ...bPnbDc, nowStr],
        ["PNB", "Debit", "PNB Imperial", "'6650", "'6524971000843793", "'10/33", "'718", "ACTIVE", "", ...bPnbDc, nowStr],
        ["PNB", "Debit", "PNB Imperial", "'7482", "'6524971000855151", "'10/33", "'628", "ACTIVE", "", ...bPnbDc, nowStr],
        ["PNB", "Debit", "PNB Imperial", "'7482", "'6524971000841870", "'10/33", "'158", "ACTIVE", "", ...bPnbDc, nowStr],
        ["PNB", "Debit", "PNB Imperial", "'8499", "'6524971000853826", "'10/33", "'426", "ACTIVE", "", ...bPnbDc, nowStr],

        // Jana SFB
        ["Jana SFB", "Debit", "Jana SFB DC", "NA", "'6525200005647267", "'09/32", "'122", "ACTIVE", "", ...bOtherDc, nowStr],

        // Bank of Baroda
        ["Bank of Baroda", "Debit", "BoB Vibhav DC", "'5465", "'6525050500835465", "'09/30", "'576", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Bank of Baroda", "Debit", "BoB Sanjay DC", "'8674", "'6525050500848674", "'10/30", "'745", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Bank of Baroda", "Debit", "BoB Aruna DC", "'8064", "'6525050500828064", "'09/30", "'780", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Bank of Baroda", "Debit", "BoB Smiti DC", "'8080", "'6525050500828080", "'09/30", "'530", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Bank of Baroda", "Debit", "BoB Shakuntala DC", "'8072", "'6525050500828072", "'09/30", "'342", "ACTIVE", "", ...bOtherDc, nowStr],

        // Canara Bank
        ["Canara", "Debit", "Canara RS DC", "'5297", "'6525500602705297", "'03/31", "'919", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Canara", "Debit", "Canara RS DC Crest", "'7180", "'6525500601367180", "'11/30", "'912", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Canara", "Debit", "Canara RS DC Crest", "'4407", "'6525500602704407", "'03/31", "'540", "ACTIVE", "", ...bOtherDc, nowStr],
        ["Canara", "Debit", "Canara RS DC Crest", "'7078", "'6525500602147078", "'10/33", "'974", "ACTIVE", "", ...bOtherDc, nowStr],

        // Union Bank
        ["Union Bank", "Debit", "Union Emperio Metal RSDC", "'9904", "'6524110100089904", "'09/33", "'690", "ACTIVE", "", ...bOtherDc, nowStr]
      ];

      for (let r = 0; r < rows.length; r++) {
        sheet.appendRow(rows[r]);
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

  // Fetch all Bank Accounts in JSON format
  if (e && e.parameter && e.parameter.action === "GET_BANK_ACCOUNTS") {
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

  // Administrative One-Off: Setup SwiggyOne brand tab & update target brand in New tab
  if (e && e.parameter && e.parameter.action === "SETUP_SWIGGY_ONE") {
    try {
      const ss = getSpreadsheet();
      let swiggySheet = ss.getSheetByName("SwiggyOne");
      let sheetCreated = false;
      if (!swiggySheet) {
        swiggySheet = ss.insertSheet("SwiggyOne");
        swiggySheet.appendRow(UNIFIED_BRAND_HEADERS);
        swiggySheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
        swiggySheet.setFrozenRows(1);
        sheetCreated = true;
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        swiggySheetCreated: sheetCreated,
        sheetExists: !!swiggySheet
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Administrative One-Off: Correct revalidation rows in "New" tab
  if (e && e.parameter && e.parameter.action === "CORRECT_REVALIDATION_ROWS") {
    try {
      const ss = getSpreadsheet();
      const nSheet = ss.getSheetByName("New");
      if (!nSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"New" tab not found' })).setMimeType(ContentService.MimeType.JSON);
      }
      const nRows = nSheet.getDataRange().getValues();
      const targetBids = new Set(["RUPS-34AB2FE3", "RUPS-5ACE0C37"]);
      const updated = [];
      for (let i = 1; i < nRows.length; i++) {
        const rowBid = String(nRows[i][0] || '').trim().toUpperCase();
        const rowCode = String(nRows[i][1] || '').trim();
        if (targetBids.has(rowBid) || rowCode.toLowerCase().startsWith("/service")) {
          nSheet.getRange(i + 1, 2).setValue("'Pending revalidation");
          nSheet.getRange(i + 1, 8).setValue("PENDING_REVALIDATION");
          nSheet.getRange(i + 1, 10).setValue(true);
          updated.push({
            rowIndex: i + 1,
            bookingId: rowBid,
            oldCode: rowCode,
            newCode: "Pending revalidation",
            status: "PENDING_REVALIDATION",
            timestampPreserved: nRows[i][8]
          });
        }
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        updatedCount: updated.length,
        updated: updated
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Administrative One-Off: Remove already approved vouchers from "New" tab
  if (e && e.parameter && e.parameter.action === "PURGE_APPROVED_DUPLICATES_FROM_NEW") {
    try {
      const ss = getSpreadsheet();
      const nSheet = ss.getSheetByName("New");
      if (!nSheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: '"New" tab not found' })).setMimeType(ContentService.MimeType.JSON);
      }
      const nRows = nSheet.getDataRange().getValues();
      const rowsToDelete = [];
      const purged = [];

      for (let i = 1; i < nRows.length; i++) {
        const rowBid = String(nRows[i][0] || '').trim();
        const rowCode = String(nRows[i][1] || '').trim();
        const rowBrand = String(nRows[i][10] || '').trim();
        if (!rowBid && !rowCode) continue;

        const match = findRowInBrandTabs(ss, rowBid, rowBrand, rowCode);
        if (match) {
          rowsToDelete.push(i + 1);
          purged.push({
            rowIndex: i + 1,
            bookingId: rowBid,
            code: rowCode,
            foundInBrandTab: match.sheet.getName(),
            brandRowIndex: match.rowIndex
          });
        }
      }

      // Bottom-up deletion
      if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a);
        for (const rIdx of rowsToDelete) {
          nSheet.deleteRow(rIdx);
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        purgedCount: purged.length,
        purged: purged
      })).setMimeType(ContentService.MimeType.JSON);
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

    // Trigger on-demand reconciliation of pending vouchers
    if (e && e.parameter && e.parameter.reconcile === "true") {
      const rec = reconcilePendingVouchers(ss);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        count: rec.length,
        reconciled: rec
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ON-DEMAND REAL-TIME EMAIL INGESTION & AUTO-HEALING RECONCILER
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
      shared: [],
      staging: []
    };

    for (const sheet of sheets) {
      const name = sheet.getName();
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) continue;

      if (name === "⚡ Unassigned") {
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] && !rows[i][1]) continue;
          const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
          const uEntry = {
            brand: "⚡ Unassigned",
            bookingId: String(rows[i][0] || ''),
            code: String(rows[i][1] || ''),
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
          if (!rows[i][2] || String(rows[i][2]).trim() === "") continue;
          const isDisp = rows[i].length > 12 ? (rows[i][12] !== false && String(rows[i][12]).toUpperCase() !== 'FALSE') : true;
          result.redeemed.push({
            brand: String(rows[i][0] || ''),
            bookingId: String(rows[i][1] || ''),
            code: String(rows[i][2] || ''),
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            amount: Number(rows[i][4]) || 0,
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
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][0] && !rows[i][1]) continue;
          const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
          const targetBrand = String(rows[i][10] || '⚡ Unassigned');
          result.staging.push({
            brand: targetBrand,
            bookingId: String(rows[i][0] || ''),
            code: String(rows[i][1] || ''),
            pin: rows[i][2] && String(rows[i][2]).trim() !== '' && String(rows[i][2]).trim() !== 'null' ? String(rows[i][2]).trim() : null,
            amount: Number(rows[i][3]) || 0,
            dateOfIssue: rows[i][4] ? String(rows[i][4]) : '',
            account: rows[i][5] ? String(rows[i][5]) : '',
            card: rows[i][6] ? String(rows[i][6]) : '',
            status: rows[i][7] ? String(rows[i][7]) : 'PENDING',
            timestampIST: rows[i][8] ? String(rows[i][8]) : null,
            displayInUI: isDisp,
            targetBrand: targetBrand
          });
        }
      } else {
        // Uniform Brand Tabs (Col 10: Display in UI)
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][1] || String(rows[i][1]).trim() === "") continue;
          const isDisp = rows[i].length > 9 ? (rows[i][9] !== false && String(rows[i][9]).toUpperCase() !== 'FALSE') : true;
          result.vouchers.push({
            brand: name,
            bookingId: String(rows[i][0] || ''),
            code: String(rows[i][1] || ''),
            pin: rows[i][2] && String(rows[i][2]).trim() !== '' && String(rows[i][2]).trim() !== 'null' ? String(rows[i][2]).trim() : null,
            amount: Number(rows[i][3]) || 0,
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
      staging: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const ss = getSpreadsheet();
  try {
    const data = JSON.parse(e.postData.contents);

    // 0A. Request Email OTP (strictly to vibhavoswal3@gmail.com)
    if (data.action === "REQUEST_OTP") {
      const cache = CacheService.getScriptCache();
      
      // 1. Rate Limit Cooldown: 60s
      if (cache.get("OTP_COOLDOWN")) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Rate limit: Please wait 60 seconds before requesting another code."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 2. Rate Limit Hourly: max 5 per hour
      const hourlyCountStr = cache.get("OTP_HOURLY_COUNT");
      const hourlyCount = hourlyCountStr ? parseInt(hourlyCountStr, 10) : 0;
      if (hourlyCount >= 5) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Rate limit: Maximum 5 verification codes per hour reached. Please try again later."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 3. Generate 6-digit OTP
      const otp = String(Math.floor(100000 + Math.random() * 900000));
      
      // 4. Store in CacheService: 10 min TTL (600 seconds)
      cache.put("VAULT_ACTIVE_OTP", otp, 600);
      cache.put("OTP_COOLDOWN", "true", 60); // 60s lock
      cache.put("OTP_HOURLY_COUNT", String(hourlyCount + 1), 3600); // 1 hr window

      // 5. Send Email via GmailApp (Clean subject line with zero corrupted emojis & responsive HTML card)
      try {
        const subject = "[RuPayVault] Login Code: " + otp + " (Valid for 10 mins)";
        const plainBody = "Hello Vibhav,\n\n" +
          "Your RuPayVault one-time login verification code is:\n\n" +
          "      " + otp + "\n\n" +
          "This code will expire in 10 minutes.\n" +
          "If you did not request this login code, please verify your account security.\n\n" +
          "— RuPayVault Security";

        const htmlBody = '<div style="background-color: #0A0A0C; padding: 28px 14px; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">' +
          '<div style="max-width: 420px; margin: 0 auto; background: #161517; border: 1px solid rgba(255, 255, 255, 0.12); border-radius: 20px; padding: 28px 24px; text-align: center; box-shadow: 0 16px 40px rgba(0, 0, 0, 0.7);">' +
          '<div style="display: inline-block; padding: 6px 14px; background: rgba(12, 140, 70, 0.15); border: 1px solid rgba(12, 140, 70, 0.35); border-radius: 9999px; color: #10B981; font-size: 11px; font-weight: 700; letter-spacing: 0.8px; margin-bottom: 20px;">' +
          'RUPAY VAULT SECURITY' +
          '</div>' +
          '<h2 style="color: #ffffff; font-size: 20px; font-weight: 700; margin: 0 0 8px 0; letter-spacing: -0.3px;">Authentication Code</h2>' +
          '<p style="color: #9CA3AF; font-size: 13px; margin: 0 0 24px 0; line-height: 1.5;">' +
          'Use the one-time code below to authenticate your private RuPay voucher vault session.' +
          '</p>' +
          '<div style="background: #000000; border: 1.5px solid #0C8C46; border-radius: 14px; padding: 18px 20px; margin-bottom: 20px; box-shadow: 0 0 20px rgba(12, 140, 70, 0.2);">' +
          '<div style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 36px; font-weight: 800; letter-spacing: 10px; color: #ffffff; text-indent: 10px;">' +
          otp +
          '</div>' +
          '</div>' +
          '<div style="background: rgba(244, 122, 33, 0.08); border: 1px solid rgba(244, 122, 33, 0.25); border-radius: 12px; padding: 12px 14px; text-align: left; margin-bottom: 24px;">' +
          '<div style="color: #F47A21; font-size: 11px; font-weight: 700; margin-bottom: 2px;">' +
          '⏱ Valid for 10 minutes &middot; Single use' +
          '</div>' +
          '<div style="color: #9CA3AF; font-size: 11px; line-height: 1.4;">' +
          'If you did not initiate this login request, please verify your account credentials immediately.' +
          '</div>' +
          '</div>' +
          '<div style="border-top: 1px solid rgba(255, 255, 255, 0.08); padding-top: 16px;">' +
          '<p style="color: #6B7280; font-size: 11px; margin: 0 0 4px 0;">' +
          'Sent to <span style="color: #9CA3AF;">vibhavoswal3@gmail.com</span>' +
          '</p>' +
          '<p style="color: #4B5563; font-size: 10px; margin: 0;">' +
          'RuPayVault Automated Security System &middot; 24h HMAC Protection' +
          '</p>' +
          '</div>' +
          '</div>' +
          '</div>';

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
      const cache = CacheService.getScriptCache();
      const storedOtp = cache.get("VAULT_ACTIVE_OTP");
      const enteredOtp = String(data.otp || "").trim();

      if (!storedOtp || !enteredOtp || storedOtp !== enteredOtp) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Invalid or expired verification code."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Burn OTP on successful use
      cache.remove("VAULT_ACTIVE_OTP");

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
          .map(b => String(b || '').trim().toUpperCase())
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
          const bId = String(row[0] || '').trim();
          const currentCode = String(row[1] || '').trim();
          const bIdUpper = bId.toUpperCase();
          const codeUpper = currentCode.toUpperCase();

          // Dual-Identifier Matching: Match by Booking ID (Col 1) OR Code (Col 2)
          if (targetIds.has(bIdUpper) || targetIds.has(codeUpper)) {
            const currentStatus = String(row[7] || '').trim().toUpperCase();
            const targetBrand = String(row[10] || '⚡ Unassigned').trim();

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
              const brandRow = [
                row[0], // Booking ID
                `'${row[1]}`, // Code
                row[2] ? `'${row[2]}` : "", // PIN
                row[3] || 500, // Value
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
        const b = String(rVals[r][1] || '').trim().toUpperCase(); // Col 2: Booking ID
        const c = String(rVals[r][2] || '').trim().toUpperCase(); // Col 3: Code
        if (b && b !== '-' && !b.startsWith('#')) existingRedeemedBookings.add(b);
        if (c && c !== '-' && c !== 'PENDING') existingRedeemedCodes.add(c);
      }

      const moved = [];
      const itemsToSearch = [];

      // Check for self-healing: if already present in Redeemed, mark success immediately
      items.forEach(it => {
        const bUpper = (it.bookingId || '').trim().toUpperCase();
        const cUpper = (it.code || '').trim().toUpperCase();
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
        const systemTabs = new Set(["New", "Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns", "⚡ Unassigned"]);
        
        // Group search items by brand sheet where known, or collect for discovery
        const sheetsToScan = new Map();

        itemsToSearch.forEach(it => {
          if (it.brand && !systemTabs.has(it.brand)) {
            if (!sheetsToScan.has(it.brand)) sheetsToScan.set(it.brand, []);
            sheetsToScan.get(it.brand).push(it);
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
              if (systemTabs.has(sName)) continue;
              if (!targetSheetsMap.has(s)) targetSheetsMap.set(s, []);
              targetSheetsMap.get(s).push(...bItems);
            }
          } else {
            const directSheet = ss.getSheetByName(brandName);
            if (directSheet && !systemTabs.has(directSheet.getName())) {
              if (!targetSheetsMap.has(directSheet)) targetSheetsMap.set(directSheet, []);
              targetSheetsMap.get(directSheet).push(...bItems);
            }
          }
        }

        // Process each target brand sheet with Single-Pass Bottom-Up Descending Deletion
        for (const [sheet, sheetItems] of targetSheetsMap.entries()) {
          const sName = sheet.getName();
          if (systemTabs.has(sName) || sName === "New") continue;

          const vals = sheet.getDataRange().getValues();
          if (vals.length <= 1) continue;

          const rowsToDelete = [];
          const targetBidSet = new Set(sheetItems.map(it => (it.bookingId || '').trim().toUpperCase()).filter(b => b && b !== '-' && !b.startsWith('#')));
          const targetCodeSet = new Set(sheetItems.map(it => (it.code || '').trim().toUpperCase()).filter(c => c && c !== '-' && c !== 'PENDING'));

          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][0] || '').trim().toUpperCase();
            const rowCode = String(vals[i][1] || '').trim().toUpperCase();

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
                const redeemedRow = [
                  sName,                                   // A: Brand
                  vals[i][0],                              // B: Booking ID
                  `'${vals[i][1]}`,                        // C: Code
                  vals[i][2] ? `'${vals[i][2]}` : "",      // D: PIN
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
      const codes = Array.from(new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []));
      if (!codes.length && data.bookingId) {
        codes.push(String(data.bookingId).trim());
      }

      const restored = [];
      const rRows = rSheet.getDataRange().getValues();
      const rowsToDeleteFromRedeemed = [];

      // Search backwards so deleting rows does not alter row index of earlier rows
      for (let i = rRows.length - 1; i >= 1; i--) {
        const rowCode = String(rRows[i][2]).trim();
        const rowBooking = String(rRows[i][1]).trim();

        if (codes.includes(rowCode) || codes.includes(rowBooking)) {
          const origBrand = String(rRows[i][10] || rRows[i][0]);
          let bSheet = ss.getSheetByName(origBrand);
          if (!bSheet) {
            bSheet = ss.insertSheet(origBrand);
            bSheet.appendRow(UNIFIED_BRAND_HEADERS);
            bSheet.getRange(1, 1, 1, UNIFIED_BRAND_HEADERS.length).setFontWeight("bold").setBackground("#F3F4F6");
            bSheet.setFrozenRows(1);
          }

          // DEDUPLICATION: Inspect brand sheet (Hierarchical: Primary RUPS Booking ID, Fallback Voucher Code)
          const bVals = bSheet.getDataRange().getValues();
          let existingBrandRowIndex = -1;
          for (let b = 1; b < bVals.length; b++) {
            const bIdVal = String(bVals[b][0] || '').trim().toUpperCase();
            const bCodeVal = String(bVals[b][1] || '').trim().toUpperCase();
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
            const originalTimestamp = (rRows[i].length > 11 && rRows[i][11]) ? String(rRows[i][11]) : "";
            const restoredRow = [
              rRows[i][1],                            // A: Booking ID
              `'${rRows[i][2]}`,                      // B: Code
              rRows[i][3] ? `'${rRows[i][3]}` : "",   // C: PIN
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
          restored.push({ bookingId: rowBooking, code: rowCode, brand: origBrand });
        }
      }

      // Self-healing: If not found in Redeemed, check if already restored in brand tab!
      codes.forEach(codeOrBid => {
        if (!restored.some(r => r.code === codeOrBid || r.bookingId === codeOrBid)) {
          const foundInBrand = findRowInBrandTabs(ss, codeOrBid, null, codeOrBid);
          if (foundInBrand) {
            restored.push({ bookingId: codeOrBid, brand: foundInBrand.sheet.getName(), alreadyRestored: true });
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
        error: "Voucher(s) not found in Redeemed sheet: " + codes.join(", ")
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

      if (aSheet) {
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

    // 6B. Block & Replace Card
    if (data.action === "BLOCK_AND_REPLACE_CARD") {
      const aSheet = ss.getSheetByName("Active Cards");
      let bSheet = ss.getSheetByName("Blocked Cards");
      if (!bSheet) bSheet = ss.insertSheet("Blocked Cards");

      const oldCardNum = String(data.oldCardNumber || '').replace(/['\s]/g, '').trim();
      const newCard = data.newCard || {};
      const newCardNum = String(newCard.cardNumber || '').replace(/['\s]/g, '').trim();
      let found = false;

      if (aSheet) {
        const aRows = aSheet.getDataRange().getValues();
        for (let i = 1; i < aRows.length; i++) {
          const rowNum = String(aRows[i][4] || '').replace(/['\s]/g, '').trim();
          if (rowNum === oldCardNum || (oldCardNum.length === 4 && rowNum.endsWith(oldCardNum))) {
            const rowData = aRows[i];
            const blockedRow = [
              rowData[0], // Bank
              rowData[1], // Card_Type
              rowData[2], // Card_Name
              "'" + String(rowData[3] || 'NA').replace(/^'/, ''),
              "'" + rowNum,
              "'" + String(rowData[5] || '').replace(/^'/, ''),
              "'" + String(rowData[6] || '').replace(/^'/, ''),
              "BLOCKED",
              formatISTTimestamp(new Date()),
              newCardNum ? newCardNum.slice(-4) : '',
              data.notes || String(rowData[8] || 'Replaced via App'),
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

      // Add New Replacement Card to Active Cards with eligible benefits set to "Unredeemed"
      if (aSheet && newCardNum) {
        const bank = newCard.bank || 'PNB';
        const type = newCard.cardType || 'Debit';
        const name = newCard.cardName || 'PNB Imperial';
        const linkedAcc = newCard.linkedAccount || 'NA';
        const expiry = newCard.expiry || '10/33';
        const cvv = newCard.cvv || '000';
        const notes = newCard.notes || 'Replaced card';
        const nowStr = formatISTTimestamp(new Date());

        const benefits = getEligibleBenefitsArray(bank, type, name);

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

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "BLOCK_AND_REPLACE_CARD",
        oldCardNumber: oldCardNum,
        newCardNumber: newCardNum
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6C. Add Brand New Card
    if (data.action === "ADD_CARD") {
      const aSheet = ss.getSheetByName("Active Cards");
      const c = data.card || data;
      const cardNum = String(c.cardNumber || '').replace(/['\s]/g, '').trim();

      if (aSheet && cardNum) {
        const bank = c.bank || 'PNB';
        const type = c.cardType || 'Debit';
        const name = c.cardName || 'RuPay Select Card';
        const linkedAcc = c.linkedAccount || 'NA';
        const expiry = c.expiry || '10/33';
        const cvv = c.cvv || '000';
        const notes = c.notes || '';
        const nowStr = formatISTTimestamp(new Date());
        const benefits = getEligibleBenefitsArray(bank, type, name);

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
        SpreadsheetApp.flush();
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        action: "ADD_CARD",
        cardNumber: cardNum
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

    // 7B. Tag Payment Card on Ingested Vouchers
    if (data.action === "TAG_CARD") {
      const cardTag = String(data.cardTag || data.card || "").trim();
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []);
      if (!cardTag || codes.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Missing cardTag or voucher codes."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const updated = [];
      const sheets = ss.getSheets();
      const systemSheets = new Set(["⏱ Cooldowns", "Cards", "Shared_Activity", "Redeemed"]);

      for (const sheet of sheets) {
        if (systemSheets.has(sheet.getName())) continue;
        const vals = sheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
          const rowBid = String(vals[i][0] || "").trim();
          const rowCode = String(vals[i][1] || "").trim();
          if (codes.has(rowCode) || codes.has(rowBid)) {
            sheet.getRange(i + 1, 7).setValue(cardTag); // Col 7: Card
            updated.push({ brand: sheet.getName(), bookingId: rowBid, code: rowCode, card: cardTag });
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
        const cStr = String(item.code).trim();
        let foundIndex = -1;
        for (let i = 1; i < sRows.length; i++) {
          if (String(sRows[i][0]).trim() === cStr) {
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
        const rowCode = String(sRows[i][0]).trim();
        if (codes.includes(rowCode)) {
          sSheet.getRange(i + 1, 6).setValue("DISMISSED");
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

    // 11. Clear Activity Log Items (sets Display in UI = FALSE in brand tabs or Shared_Activity)
    if (data.action === "CLEAR_ACTIVITY") {
      const type = String(data.type || "INGESTED").toUpperCase();
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(
        (Array.isArray(rawCodes) ? rawCodes : [rawCodes])
          .map(c => String(c).trim())
          .filter(c => c && c.toUpperCase() !== 'PENDING' && c !== '-' && !c.startsWith('#'))
      );
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
            const rowCode = String(vals[i][0] || "").trim();
            if (rowCode && rowCode !== '-' && codes.has(rowCode)) {
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
            const rowBid = String(vals[i][0] || "").trim();
            const rowCode = String(vals[i][1] || "").trim();
            const bidMatch = rowBid && rowBid !== '-' && !rowBid.startsWith('#') && codes.has(rowBid);
            const codeMatch = rowCode && rowCode !== '-' && rowCode.toUpperCase() !== 'PENDING' && codes.has(rowCode);
            if (bidMatch || codeMatch) {
              newSheet.getRange(i + 1, 10).setValue(false); // Col 10: Display in UI
              updated.push(rowBid || rowCode);
            }
          }
        }
        const sheets = ss.getSheets();
        const systemSheets = new Set(["New", "⚡ Unassigned", "⏱ Cooldowns", "Cards", "Shared_Activity", "Redeemed"]);
        for (const sheet of sheets) {
          if (systemSheets.has(sheet.getName())) continue;
          const vals = sheet.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][0] || "").trim();
            const rowCode = String(vals[i][1] || "").trim();
            const bidMatch = rowBid && rowBid !== '-' && !rowBid.startsWith('#') && codes.has(rowBid);
            const codeMatch = rowCode && rowCode !== '-' && rowCode.toUpperCase() !== 'PENDING' && codes.has(rowCode);
            if (bidMatch || codeMatch) {
              sheet.getRange(i + 1, 10).setValue(false); // Col 10: Display in UI
              updated.push(rowBid || rowCode);
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
          const rowBid = String(vals[i][1] || "").trim();
          const rowCode = String(vals[i][2] || "").trim();
          if (codes.has(rowCode) || codes.has(rowBid)) {
            rSheet.getRange(i + 1, 13).setValue(false); // Col 13: Display in UI
            updated.push(rowCode || rowBid);
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
    } finally {
      try { lock.releaseLock(); } catch(e) {}
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

  const accounts = [
    // --- VIBHAV OSWAL (17 ACCOUNTS) ---
    ["acc_indusind_vibhav_1", "Vibhav Oswal", "IndusInd Bank", "indusind", "Pioneer Private #1", "Pioneer Private", "HNI & Private", "'159247538499", "INDB0002169", "", "'41544412", "vibhavoswal", "", "", "", "", "", "ACTIVE", "Primary Pioneer Account", nowStr],
    ["acc_axis_vibhav_1", "Vibhav Oswal", "Axis Bank", "axis", "Burgundy #1", "Burgundy", "HNI & Private", "'921010018051356", "UTIB0002366", "vibhavoswal@axisbank", "'908507230", "", "", "", "", "", "", "ACTIVE", "Primary Burgundy", nowStr],
    ["acc_axis_vibhav_2", "Vibhav Oswal", "Axis Bank", "axis", "Burgundy #2", "Burgundy", "HNI & Private", "'924010040061771", "UTIB0000194", "vibhavoswal3@axisbank", "'908507230", "", "", "", "", "", "", "ACTIVE", "Secondary Burgundy", nowStr],
    ["acc_hdfc_vibhav_1", "Vibhav Oswal", "HDFC Bank", "hdfc", "HDFC Savings", "Imperia / Preferred", "Wealth & Priority", "'50100272080801", "HDFC0000827", "vibhavoswal@hdfcbank", "'128271043", "", "", "", "", "", "", "ACTIVE", "Primary HDFC Account", nowStr],
    ["acc_kotak_vibhav_1", "Vibhav Oswal", "Kotak Mahindra Bank", "kotak", "Privy League", "Privy League", "Wealth & Priority", "'9113615811", "KKBK0002020", "9983692067@kotak", "'391549346", "", "", "", "", "", "", "ACTIVE", "Privy Priority Account", nowStr],
    ["acc_idfc_vibhav_1", "Vibhav Oswal", "IDFC FIRST Bank", "idfc", "FIRST Wealth", "FIRST Wealth", "Wealth & Priority", "'10064054118", "IDFB0040101", "9983692067@idfcfirst", "'5587071737", "", "", "", "", "", "", "ACTIVE", "FIRST Wealth Account", nowStr],
    ["acc_equitas_vibhav_1", "Vibhav Oswal", "Equitas SFB", "equitas", "Equitas Wings", "Small Finance", "Digital & SFB", "'100021374863", "ESFB0009021", "vibhavoswal@equitasbank", "", "", "", "", "", "", "", "ACTIVE", "Wings Account", nowStr],
    ["acc_federal_fi_vibhav", "Vibhav Oswal", "Federal Bank (Fi)", "federal", "Fi Digital", "Fi Neo", "Digital & SFB", "'55550107473710", "FDRL0005555", "vibhav@fbl", "", "", "", "", "", "", "", "ACTIVE", "Fi Neo Salary Account", nowStr],
    ["acc_federal_jupiter_vibhav", "Vibhav Oswal", "Federal Bank (Jupiter)", "federal", "Jupiter 1% UPI", "Jupiter Pro", "Digital & SFB", "'77770106660497", "FDRL0007777", "9983692067@jupiteraxis", "'9823812", "", "", "", "", "", "", "ACTIVE", "Jupiter Pro Account", nowStr],
    ["acc_jana_vibhav_1", "Vibhav Oswal", "Jana SFB", "jana", "Jana SFB", "Small Finance", "Digital & SFB", "'4800010052985502", "JSFB0004800", "", "", "", "", "", "", "", "BANASHANKARI BRANCH", "ACTIVE", "CRN: 2420042415002382", nowStr],
    ["acc_pnb_vibhav_1", "Vibhav Oswal", "Punjab National Bank", "pnb", "PNB Elite (Acc 1)", "Elite Executive", "HNI & Private", "'1057100100006650", "PUNB0105710", "", "'R67653832", "", "", "", "'4625", "", "Defence Colony Indiranagar Branch", "ACTIVE", "Linked to PNB Imperial RSDC", nowStr],
    ["acc_pnb_vibhav_2", "Vibhav Oswal", "Punjab National Bank", "pnb", "PNB Elite (Acc 2)", "Elite Executive", "HNI & Private", "'1057100100007482", "PUNB0105710", "", "'R67653832", "", "", "", "'4625", "", "Defence Colony Indiranagar Branch", "ACTIVE", "Elite Executive Account 2", nowStr],
    ["acc_union_vibhav_1", "Vibhav Oswal", "Union Bank of India", "union", "HNI Prime", "HNI Prime", "HNI & Private", "'028912250000002", "UBIN0802891", "", "'652214823", "", "", "", "", "", "Indiranagar-CMH Road Bangalore", "ACTIVE", "Union HNI Prime Account", nowStr],
    ["acc_canara_basic_vibhav", "Vibhav Oswal", "Canara Bank", "canara", "Basic Savings", "Basic Savings", "Digital & SFB", "'110181739491", "CNRB0010754", "", "'325184532", "", "", "", "", "", "", "ACTIVE", "Basic Savings Account", nowStr],
    ["acc_canara_crest_vibhav", "Vibhav Oswal", "Canara Bank", "canara", "Canara Crest", "Canara Crest", "Digital & SFB", "'110267332562", "CNRB0000492", "", "'325184532", "", "", "", "", "", "", "ACTIVE", "Canara Crest Account", nowStr],
    ["acc_hsbc_vibhav_1", "Vibhav Oswal", "HSBC India", "hsbc", "HSBC Premier", "Premier", "HNI & Private", "'074532227006", "HSBC0560002", "", "'254315831", "", "", "", "", "", "", "ACTIVE", "HSBC Premier Account", nowStr],
    ["acc_bob_vibhav_1", "Vibhav Oswal", "Bank of Baroda", "bob", "BoB Advantage", "Advantage Savings", "Wealth & Priority", "'99870100008545", "BARB0DBKHOP", "", "'MT3D09485", "MT3D09485", "t@rgetEU2026", "t@rgetEU26", "'5358", "'4505", "Khopoli Branch", "ACTIVE", "BoB NetBanking", nowStr],

    // --- SANJAY OSWAL (5 ACCOUNTS) ---
    ["acc_indusind_sanjay_2", "Sanjay Oswal", "IndusInd Bank", "indusind", "Pioneer Private #2", "Pioneer Private", "HNI & Private", "'159823470448", "INDB0001103", "", "'55855930", "", "", "", "", "", "", "ACTIVE", "Pioneer Private Account #2", nowStr],
    ["acc_axis_sanjay_1", "Sanjay Oswal", "Axis Bank", "axis", "Burgundy Sanjay", "Burgundy", "HNI & Private", "'926010016663332", "UTIB0000114", "", "'982451359", "", "S@runj@y963", "", "", "", "", "ACTIVE", "Burgundy Sanjay Account", nowStr],
    ["acc_hdfc_sanjay_1", "Sanjay Oswal", "HDFC Bank", "hdfc", "HDFC Sanjay", "Imperia / Preferred", "Wealth & Priority", "'50100456446803", "HDFC0000827", "sanjayoswal@hdfcbank", "'177426223", "sanjayoswal", "", "", "", "", "", "ACTIVE", "HDFC Sanjay Account", nowStr],
    ["acc_bob_sanjay_1", "Sanjay Oswal", "Bank of Baroda", "bob", "BoB Sanjay", "Advantage Savings", "Wealth & Priority", "'99870100004059", "BARB0DBKHOP", "", "'MT3D01692", "MT3D01692", "sarunjay#321", "sarunjay#123", "'6417", "'2302", "Khopoli Branch", "ACTIVE", "BoB NetBanking Sanjay", nowStr],
    ["acc_federal_sanjay_1", "Sanjay Oswal", "Federal Bank", "federal", "Federal Sanjay", "Savings", "Digital & SFB", "'55550111917124", "FDRL0005555", "sanjayoswal@fbl", "", "", "", "", "", "", "", "ACTIVE", "Federal Account (Sanjay Himatmal Oswal)", nowStr],

    // --- ARUNA OSWAL (4 ACCOUNTS) ---
    ["acc_axis_aruna_1", "Aruna Oswal", "Axis Bank", "axis", "Burgundy Aruna", "Burgundy", "HNI & Private", "'926010016657793", "UTIB0000114", "", "'982450167", "", "S@runj@y369", "", "", "", "", "ACTIVE", "Burgundy Aruna Account", nowStr],
    ["acc_hdfc_aruna_1", "Aruna Oswal", "HDFC Bank", "hdfc", "HDFC Aruna", "Imperia / Preferred", "Wealth & Priority", "'50100456446816", "HDFC0000827", "arunaoswal@hdfcbank", "'177426224", "arunaoswal", "", "", "", "", "", "ACTIVE", "HDFC Aruna Account", nowStr],
    ["acc_bob_aruna_1", "Aruna Oswal", "Bank of Baroda", "bob", "BoB Aruna", "Advantage Savings", "Wealth & Priority", "'99870100005014", "BARB0DBKHOP", "", "'MT3D03173", "MT3D03173", "sarunjay#123", "sarunjay#321", "'5358", "'4505", "Khopoli Branch", "ACTIVE", "BoB NetBanking Aruna", nowStr],
    ["acc_federal_aruna_1", "Aruna Oswal", "Federal Bank", "federal", "Federal Aruna", "Savings", "Digital & SFB", "'55550112011364", "FDRL0005555", "arunaoswal@fbl", "", "", "", "", "", "", "", "ACTIVE", "Federal Account (Aruna Sanjay Oswal)", nowStr],

    // --- SMITI OSWAL (3 ACCOUNTS) ---
    ["acc_axis_smiti_1", "Smiti Oswal", "Axis Bank", "axis", "Burgundy Smiti", "Burgundy", "HNI & Private", "'923010014752754", "UTIB0002366", "vibhavoswal@axisbank", "'949545156", "", "Benedict@123", "", "", "", "", "ACTIVE", "Burgundy Smiti Account", nowStr],
    ["acc_hdfc_smiti_1", "Smiti Oswal", "HDFC Bank", "hdfc", "HDFC Smiti", "Imperia / Preferred", "Wealth & Priority", "'50100591004352", "HDFC0000827", "vibhavoswal@hdfcbank", "'222800319", "smitioswal", "", "", "", "", "", "ACTIVE", "HDFC Smiti Account", nowStr],
    ["acc_bob_smiti_1", "Smiti Oswal", "Bank of Baroda", "bob", "BoB Smiti", "Advantage Savings", "Wealth & Priority", "'99870100013818", "BARB0DBKHOP", "", "'MT3063389", "MT3063389", "Smiti@bob$123", "Smiti@trsc$123", "", "", "Khopoli Branch", "ACTIVE", "BoB NetBanking Smiti", nowStr],

    // --- SHAKUNTALA OSWAL (1 ACCOUNT) ---
    ["acc_bob_shakuntala_1", "Shakuntala Oswal", "Bank of Baroda", "bob", "BoB Shakuntala", "Advantage Savings", "Wealth & Priority", "'99870100009094", "BARB0DBKHOP", "", "'MT3D01699", "MT3D01699", "S@runj@y12345", "S@runj@y54321", "", "", "Khopoli Branch", "ACTIVE", "BoB NetBanking Shakuntala", nowStr]
  ];

  accounts.forEach(acc => sheet.appendRow(acc));

  return {
    success: true,
    sheetName: "Bank Accounts",
    created: created,
    accountsCount: accounts.length
  };
}

