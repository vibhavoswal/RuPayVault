/**
 * RuPay Select Vault — Google Apps Script Engine
 * 100% Free Cloud Backend (Gmail -> Sheet -> Swift App REST API)
 * All timestamps strictly formatted in IST (Indian Standard Time: Asia/Kolkata)
 */

const SENDER_EMAIL = "rupay@truztee.com";
const IST_TIMEZONE = "Asia/Kolkata";
const AUTHORIZED_USER = "vibhavoswal3@gmail.com";

// Standard Unified Column Schemas
const UNIFIED_BRAND_HEADERS = [
  "#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI"
];

const UNIFIED_NEW_HEADERS = [
  "#", "Booking ID", "Code", "PIN", "Value", "Date of Issue", "Account", "Card", "Status", "Appended Timestamp (IST)", "Display in UI", "Target Brand"
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
  "Thyrocare", "Uber", "Decathlon"
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

  // Master setup key for admin migration
  if (providedToken === "rupay_setup_new_2026") {
    return { valid: true, user: AUTHORIZED_USER, expiry: Date.now() + 86400000 };
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

  // 4. Extract Raw Coupon Code Line (handles same-line and multi-line table values)
  let rawVal = "";
  const codeLineMatch = body.match(/(?:Coupon|Voucher)\s*Code\s*[:#-]?\s*([^\r\n<]+)/i) ||
                        body.match(/\bCode\s*[:#-]?\s*([^\r\n<]+)/i);
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

  const receivedDate = (msg && typeof msg.getDate === 'function') ? msg.getDate() : new Date();

  // 5. Binary Check: Is it a delay placeholder or a real code?
  const isDelay = !rawVal || /will be sent|delay|within|hrs|days|pending|\bna\b|n\/a/i.test(rawVal) || rawVal.toUpperCase().startsWith("RUPS");
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
    const existingBids = new Set(cRows.slice(1).map(r => String(r[1] || '').trim()).filter(Boolean));
    const existingCodes = new Set(cRows.slice(1).map(r => String(r[2] || '').trim()).filter(Boolean));

    for (let i = 1; i < rRows.length; i++) {
      const row = rRows[i];
      const bid = String(row[1] || '').trim();
      const code = String(row[2] || '').trim();

      if (!bid && !code) continue;

      if (!existingBids.has(bid) && (!code || !existingCodes.has(code))) {
        const nextNum = cSheet.getLastRow();
        const newRow = [...row];
        newRow[0] = nextNum; // update #
        cSheet.appendRow(newRow);
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

        if (!parsed.isDelivered) {
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

  return "Other";
}

function extractAmount(body) {
  if (!body) return 500;
  const m = body.match(/(?:INR|Rs\.?|₹)\s*([0-9,]+)/i) ||
            body.match(/(?:Value|Amount|Worth)\s*[:#-]?\s*(?:INR|Rs\.?|₹)?\s*([0-9,]+)/i);
  return m ? parseInt(m[1].replace(/,/g, ""), 10) : 500;
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

  // 1. Check if Booking ID already exists in Redeemed tab (Col 3: Booking ID)
  const rSheet = ss.getSheetByName("Redeemed");
  if (rSheet) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rRows.length; i++) {
      if (String(rRows[i][2]).trim() === data.bookingId) {
        // Already redeemed, don't recreate!
        return;
      }
    }
  }

  // 2. Global Idempotency: Does this Booking ID already exist in "New" tab?
  const newRows = newSheet.getDataRange().getValues();
  for (let i = 1; i < newRows.length; i++) {
    if (String(newRows[i][1]).trim() === data.bookingId) {
      // Already recorded in staging, keep initial ack timestamp intact and do not duplicate
      return;
    }
  }

  // 3. Determine canonical target brand (Col 12)
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "⚡ Unassigned");

  const nextNum = newSheet.getLastRow();
  const rowData = [
    nextNum,                                             // Col 1: #
    data.bookingId,                                      // Col 2: Booking ID (RUPS...)
    "PENDING",                                           // Col 3: Code
    "",                                                  // Col 4: PIN
    data.amount || 500,                                  // Col 5: Value
    data.dateOfIssue || "",                              // Col 6: Date of Issue
    data.account || "",                                  // Col 7: Account
    "Unassigned",                                        // Col 8: Card
    "PENDING",                                           // Col 9: Status
    formatISTTimestamp(data.receivedDate || new Date()), // Col 10: Appended Timestamp (IST) - Immutable for Stage 1
    true,                                                // Col 11: Display in UI
    targetBrand                                          // Col 12: Target Brand
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
    if (String(rows[i][1]).trim() === data.bookingId) {
      const currentCode = String(rows[i][2] || '').trim();
      const currentStatus = String(rows[i][8] || '').trim().toUpperCase();

      // Guard: If already ACTIVE and voucher code is populated, it is permanently immutable!
      if (currentStatus === "ACTIVE" && currentCode !== "" && currentCode !== "PENDING") {
        return; // No-op, do not re-stamp
      }

      // MATCH FOUND: Transition from PENDING to ACTIVE and update timestamp with delivery email arrival time
      newSheet.getRange(i + 1, 3).setValue(`'${data.code}`);
      newSheet.getRange(i + 1, 4).setValue(data.pin ? `'${data.pin}` : "");
      if (data.amount && Number(data.amount) > 0) {
        newSheet.getRange(i + 1, 5).setValue(Number(data.amount));
      }
      newSheet.getRange(i + 1, 9).setValue("ACTIVE");
      newSheet.getRange(i + 1, 10).setValue(formatISTTimestamp(data.receivedDate || new Date()));
      newSheet.getRange(i + 1, 11).setValue(true);
      return;
    }
  }

  // Fallback: If Stage 1 email was missed and delivery email arrived directly,
  // append as an ACTIVE row in "New" tab so user can review and approve it.
  const canonicalBrand = extractBrand(data.brand, "");
  const targetBrand = canonicalBrand !== "Other" ? canonicalBrand : (data.brand || "⚡ Unassigned");
  const nextNum = newSheet.getLastRow();
  const rowData = [
    nextNum,                                             // Col 1: #
    data.bookingId,                                      // Col 2: Booking ID
    `'${data.code}`,                                     // Col 3: Code
    data.pin ? `'${data.pin}` : "",                      // Col 4: PIN
    data.amount || 500,                                  // Col 5: Value
    data.dateOfIssue || "",                              // Col 6: Date of Issue
    data.account || "",                                  // Col 7: Account
    "Unassigned",                                        // Col 8: Card
    "ACTIVE",                                            // Col 9: Status
    formatISTTimestamp(data.receivedDate || new Date()), // Col 10: Appended Timestamp (IST)
    true,                                                // Col 11: Display in UI
    targetBrand                                          // Col 12: Target Brand
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
    const bid = String(sRows[i][1] || '').trim();
    const code = String(sRows[i][2] || '').trim();
    const currentPin = String(sRows[i][3] || '').trim();
    const status = String(sRows[i][8] || '').trim();

    if ((status === 'PENDING' || code === 'PENDING') && /^RUPS/i.test(bid)) {
      if (currentPin && !/^\d{3,8}$/.test(currentPin)) {
        newSheet.getRange(i + 1, 4).setValue("");
      }

      try {
        const threads = GmailApp.search('"' + bid + '"', 0, 3);
        for (const t of threads) {
          const msgs = t.getMessages();
          for (const m of msgs) {
            const parsed = parseRupayVoucherEmail(m);
            if (parsed && parsed.isDelivered && parsed.bookingId === bid) {
              newSheet.getRange(i + 1, 3).setValue(`'${parsed.code}`);
              newSheet.getRange(i + 1, 4).setValue(parsed.pin ? `'${parsed.pin}` : "");
              if (parsed.amount && Number(parsed.amount) > 0) {
                newSheet.getRange(i + 1, 5).setValue(Number(parsed.amount));
              }
              newSheet.getRange(i + 1, 9).setValue("ACTIVE");
              newSheet.getRange(i + 1, 10).setValue(formatISTTimestamp(parsed.receivedDate || m.getDate()));
              newSheet.getRange(i + 1, 11).setValue(true);
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
 * 6B. SETUP "NEW" TAB AND MIGRATE EXISTING PENDING VOUCHERS
 */
function setupNewTabAndMigratePending(ss) {
  if (!ss) ss = getSpreadsheet();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const EXCLUDED_TABS = new Set(["New", "Redeemed", "Cards", "Accounts", "Shared_Activity", "⏱ Cooldowns"]);

    // 1. Ensure master "New" tab exists
    let newSheet = ss.getSheetByName("New");
    if (!newSheet) {
      newSheet = ss.insertSheet("New");
      newSheet.appendRow(UNIFIED_NEW_HEADERS);
      newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
      newSheet.setFrozenRows(1);
    } else if (newSheet.getLastRow() === 0) {
      newSheet.appendRow(UNIFIED_NEW_HEADERS);
      newSheet.getRange(1, 1, 1, UNIFIED_NEW_HEADERS.length).setFontWeight("bold").setBackground("#EEF2FF");
      newSheet.setFrozenRows(1);
    }

    // Existing booking IDs in New tab to prevent duplicates
    const existingNewBookings = new Set();
    const newSheetData = newSheet.getDataRange().getValues();
    for (let i = 1; i < newSheetData.length; i++) {
      const bId = String(newSheetData[i][1] || '').trim();
      if (bId) existingNewBookings.add(bId);
    }

    const sheets = ss.getSheets();
    const migratedSummary = [];
    let totalMoved = 0;

    for (const sheet of sheets) {
      const sName = sheet.getName();
      if (EXCLUDED_TABS.has(sName)) continue;

      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;

      const rowsToDelete = [];
      const movedBookings = [];

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const rowNum = r + 1; // 1-based index in Sheet

        const bookingId = String(row[1] || '').trim();
        const code = String(row[2] || '').trim();
        const pin = String(row[3] || '').trim();
        const value = row[4];
        const dateOfIssue = row[5];
        const account = row[6];
        const card = row[7];
        const status = String(row[8] || '').trim().toUpperCase();
        const timestamp = row[9];
        const displayInUI = (row[10] !== false && String(row[10]).toLowerCase() !== 'false');

        // Identify pending vouchers
        const isPending = (status === 'PENDING' || code.toUpperCase() === 'PENDING' || (!code && bookingId.toUpperCase().startsWith('RUPS')));

        if (isPending) {
          if (!existingNewBookings.has(bookingId)) {
            const nextIdx = newSheet.getLastRow();
            newSheet.appendRow([
              nextIdx,
              bookingId,
              code || "PENDING",
              pin ? `'${pin}` : "",
              value || 500,
              dateOfIssue || "",
              account || "",
              card || "Unassigned",
              "PENDING",
              timestamp || formatISTTimestamp(new Date()),
              displayInUI,
              sName // Target Brand (Col 12)
            ]);
            existingNewBookings.add(bookingId);
          }
          rowsToDelete.push(rowNum);
          movedBookings.push(bookingId || `Row ${rowNum}`);
          totalMoved++;
        }
      }

      // Delete from bottom to top
      if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a);
        for (const rIdx of rowsToDelete) {
          sheet.deleteRow(rIdx);
        }

        // Renumber remaining rows in this brand sheet
        const remaining = sheet.getLastRow();
        if (remaining > 1) {
          const numRange = sheet.getRange(2, 1, remaining - 1, 1);
          const nums = [];
          for (let n = 1; n < remaining; n++) {
            nums.push([n]);
          }
          numRange.setValues(nums);
        }

        migratedSummary.push({ brand: sName, moved: rowsToDelete.length, bookingIds: movedBookings });
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      totalMoved: totalMoved,
      details: migratedSummary,
      executedAt: formatISTTimestamp(new Date())
    };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * 7. REST API (GET / POST)
 */
function doGet(e) {
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

    // One-off migration trigger via GET if requested
    if (e && e.parameter && e.parameter.migrate === "true") {
      try {
        const migRes = migrateSheetsToUnifiedSchema(ss);
        return ContentService.createTextOutput(JSON.stringify(migRes)).setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Migration error: " + err.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Setup "New" Tab & migrate pending items trigger via GET if requested
    if (e && e.parameter && (e.parameter.setupNew === "true" || e.parameter.action === "SETUP_NEW_TAB")) {
      try {
        const setupRes = setupNewTabAndMigratePending(ss);
        return ContentService.createTextOutput(JSON.stringify(setupRes)).setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Setup error: " + err.message })).setMimeType(ContentService.MimeType.JSON);
      }
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
            raw: m.getRawContent().substring(0, 600),
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
          if (!rows[i][1] && !rows[i][2]) continue;
          const isDisp = rows[i].length > 10 ? (rows[i][10] !== false && String(rows[i][10]).toUpperCase() !== 'FALSE') : true;
          const uEntry = {
            brand: "⚡ Unassigned",
            bookingId: String(rows[i][1] || ''),
            code: String(rows[i][2] || ''),
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            amount: Number(rows[i][4]) || 0,
            dateOfIssue: rows[i][5] ? String(rows[i][5]) : '',
            account: rows[i][6] ? String(rows[i][6]) : '',
            card: rows[i][7] ? String(rows[i][7]) : '',
            status: rows[i][8] ? String(rows[i][8]) : 'PENDING',
            timestampIST: rows[i][9] ? String(rows[i][9]) : null,
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
      } else if (name === "New") {
        // Staging Feed from "New" Tab
        for (let i = 1; i < rows.length; i++) {
          if (!rows[i][1] && !rows[i][2]) continue;
          const isDisp = rows[i].length > 10 ? (rows[i][10] !== false && String(rows[i][10]).toUpperCase() !== 'FALSE') : true;
          const targetBrand = String(rows[i][11] || '⚡ Unassigned');
          result.staging.push({
            brand: targetBrand,
            bookingId: String(rows[i][1] || ''),
            code: String(rows[i][2] || ''),
            pin: rows[i][3] && String(rows[i][3]).trim() !== '' && String(rows[i][3]).trim() !== 'null' ? String(rows[i][3]).trim() : null,
            amount: Number(rows[i][4]) || 0,
            dateOfIssue: rows[i][5] ? String(rows[i][5]) : '',
            account: rows[i][6] ? String(rows[i][6]) : '',
            card: rows[i][7] ? String(rows[i][7]) : '',
            status: rows[i][8] ? String(rows[i][8]) : 'PENDING',
            timestampIST: rows[i][9] ? String(rows[i][9]) : null,
            displayInUI: isDisp,
            targetBrand: targetBrand
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

      // 5. Send Email via GmailApp
      try {
        GmailApp.sendEmail(
          AUTHORIZED_USER,
          "🔐 RuPayVault Login Code: " + otp,
          "Hello Vibhav,\n\n" +
          "Your RuPayVault one-time login verification code is:\n\n" +
          "      " + otp + "\n\n" +
          "This code will expire in 10 minutes.\n" +
          "If you did not request this login code, please verify your account security.\n\n" +
          "— RuPayVault Security"
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
      // 1. Run Migration on Demand
      if (data.action === "MIGRATE_UNIFIED_SCHEMA") {
        const migRes = migrateSheetsToUnifiedSchema(ss);
        return ContentService.createTextOutput(JSON.stringify(migRes)).setMimeType(ContentService.MimeType.JSON);
      }

      // 1B. Setup "New" Tab and Migrate Pending items
      if (data.action === "SETUP_NEW_TAB") {
        const setupRes = setupNewTabAndMigratePending(ss);
        return ContentService.createTextOutput(JSON.stringify(setupRes)).setMimeType(ContentService.MimeType.JSON);
      }

      // 1C. Approve Staged Green Vouchers from "New" Tab to Canonical Brand Tabs (Dual-Identifier Matching)
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
          const bId = String(row[1] || '').trim();
          const currentCode = String(row[2] || '').trim();
          const bIdUpper = bId.toUpperCase();
          const codeUpper = currentCode.toUpperCase();

          // Dual-Identifier Matching: Match by Booking ID (Col 2) OR Code (Col 3)
          if (targetIds.has(bIdUpper) || targetIds.has(codeUpper)) {
            const currentStatus = String(row[8] || '').trim().toUpperCase();
            const targetBrand = String(row[11] || '⚡ Unassigned').trim();

            // HARD CONDITION GUARD: Only fulfilled (ACTIVE, real code) vouchers can be approved!
            if (currentStatus !== 'ACTIVE' || currentCode === 'PENDING' || !currentCode) {
              return ContentService.createTextOutput(JSON.stringify({
                error: `Cannot approve unfulfilled voucher ${bId || currentCode}. Code must arrive before approval.`
              })).setMimeType(ContentService.MimeType.JSON);
            }

            // Target Brand Sheet
            const brandSheet = getTargetBrandSheet(ss, targetBrand);
            const brandNextNum = brandSheet.getLastRow();

            // Append Columns 1-11 to Brand Sheet
            const brandRow = [
              brandNextNum,
              row[1], // Booking ID
              `'${row[2]}`, // Code
              row[3] ? `'${row[3]}` : "", // PIN
              row[4] || 500, // Value
              row[5] || "", // Date of Issue
              row[6] || "", // Account
              row[7] || "Unassigned", // Card
              "ACTIVE", // Status
              row[9] || formatISTTimestamp(new Date()), // Appended Timestamp (IST)
              row[10] !== false && String(row[10]).toLowerCase() !== 'false' // Display in UI
            ];
            brandSheet.appendRow(brandRow);

            rowsToDelete.push(r + 1); // 1-based index in "New"
            approvedSummary.push({ bookingId: bId, brand: targetBrand, code: currentCode });
          }
        }

        // Bottom-up deletion from "New" tab
        if (rowsToDelete.length > 0) {
          rowsToDelete.sort((a, b) => b - a);
          for (const rIdx of rowsToDelete) {
            newSheet.deleteRow(rIdx);
          }

          // Renumber remaining rows in "New"
          const remaining = newSheet.getLastRow();
          if (remaining > 1) {
            const numRange = newSheet.getRange(2, 1, remaining - 1, 1);
            const nums = [];
            for (let n = 1; n < remaining; n++) {
              nums.push([n]);
            }
            numRange.setValues(nums);
          }
        }

        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          approvedCount: approvedSummary.length,
          approved: approvedSummary
        })).setMimeType(ContentService.MimeType.JSON);
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
              const isCodeMatch = String(vals[i][2]).trim() === codeVal;
              const isBookingMatch = /^RUPS/i.test(codeVal) && String(vals[i][1]).trim() === codeVal;
              if (isCodeMatch || isBookingMatch) {
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
              const isCodeMatch = String(vals[i][2]).trim() === codeVal;
              const isBookingMatch = /^RUPS/i.test(codeVal) && String(vals[i][1]).trim() === codeVal;
              if (isCodeMatch || isBookingMatch) {
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
          const rowBid = String(vals[i][1] || "").trim();
          const rowCode = String(vals[i][2] || "").trim();
          if (codes.has(rowCode) || codes.has(rowBid)) {
            sheet.getRange(i + 1, 8).setValue(cardTag); // Col 8: Card
            updated.push({ brand: sheet.getName(), bookingId: rowBid, code: rowCode, card: cardTag });
          }
        }
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

    // 11. Clear Activity Log Items (sets Display in UI = FALSE in brand tabs or Shared_Activity)
    if (data.action === "CLEAR_ACTIVITY") {
      const type = String(data.type || "INGESTED").toUpperCase();
      const rawCodes = data.codes || (data.code ? [data.code] : []);
      const codes = new Set(Array.isArray(rawCodes) ? rawCodes.map(c => String(c).trim()).filter(Boolean) : []);
      if (codes.size === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          error: "Cannot clear activity: No voucher codes specified."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      const updated = [];

      if (type === "SHARED") {
        const sh = ss.getSheetByName("Shared_Activity");
        if (sh) {
          const vals = sh.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowCode = String(vals[i][1] || "").trim();
            if (codes.has(rowCode)) {
              sh.getRange(i + 1, 9).setValue(false); // Col 9: Display in UI
              sh.getRange(i + 1, 7).setValue("DISMISSED");
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
            const rowBid = String(vals[i][1] || "").trim();
            const rowCode = String(vals[i][2] || "").trim();
            if (codes.has(rowCode) || codes.has(rowBid)) {
              newSheet.getRange(i + 1, 11).setValue(false); // Col 11: Display in UI
              updated.push(rowCode || rowBid);
            }
          }
        }
        const sheets = ss.getSheets();
        const systemSheets = new Set(["New", "⚡ Unassigned", "⏱ Cooldowns", "Cards", "Shared_Activity", "Redeemed"]);
        for (const sheet of sheets) {
          if (systemSheets.has(sheet.getName())) continue;
          const vals = sheet.getDataRange().getValues();
          for (let i = 1; i < vals.length; i++) {
            const rowBid = String(vals[i][1] || "").trim();
            const rowCode = String(vals[i][2] || "").trim();
            if (codes.has(rowCode) || codes.has(rowBid)) {
              sheet.getRange(i + 1, 11).setValue(false); // Col 11: Display in UI
              updated.push(rowCode || rowBid);
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
          const rowBid = String(vals[i][2] || "").trim();
          const rowCode = String(vals[i][3] || "").trim();
          if (codes.has(rowCode) || codes.has(rowBid)) {
            rSheet.getRange(i + 1, 14).setValue(false); // Col 14: Display in UI
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
