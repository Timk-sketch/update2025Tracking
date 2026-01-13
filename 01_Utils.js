// =====================================================
// 01_Utils.gs
// Shared utility helpers used across the project
// =====================================================

function s_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function n_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const cleaned = String(v).replace(/[^0-9.\-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function parseMoney_(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;

  const s = String(value).trim();
  if (!s) return 0;

  const negParen = /^\(.*\)$/.test(s);
  const cleaned = s.replace(/[^\d.\-]/g, "");
  const n = parseFloat(cleaned);

  if (!isFinite(n)) return 0;
  return negParen ? -Math.abs(n) : n;
}

function parseQty_(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;

  const s = String(value).trim();
  if (!s) return 0;

  const cleaned = s.replace(/[^\d.\-]/g, "");
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : 0;
}

/**
 * Robust date parsing for:
 * - Date objects (Sheets often returns these)
 * - ISO strings (Shopify/Squarespace)
 * - Locale strings like "12/24/2025, 5:30:00 PM"
 * - yyyy-mm-dd and yyyy-mm-dd hh:mm:ss
 * - numeric timestamps and occasional serials
 */
function asDate_(v) {
  if (v === null || v === undefined || v === "") return null;

  // Already a Date
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return v;
  }

  // Numeric: could be ms timestamp or (rarely) sheet serial
  if (typeof v === "number" && isFinite(v)) {
    // If it looks like milliseconds since epoch
    if (v > 1000000000000) {
      const d = new Date(v);
      return isNaN(d.getTime()) ? null : d;
    }

    // If it looks like a Google Sheets serial date (days since 1899-12-30)
    // Typical range ~ 20000..80000 for modern dates
    if (v >= 20000 && v <= 80000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(epoch.getTime() + v * 24 * 60 * 60 * 1000);
      return isNaN(d.getTime()) ? null : d;
    }

    // Fallback
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }

  const tz = Session.getScriptTimeZone();
  const str = String(v).trim();
  if (!str) return null;

  // 1) Native parse (handles ISO well)
  let d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  // 2) Remove a single comma (common from toLocaleString)
  const noComma = str.replace(",", "");
  d = new Date(noComma);
  if (!isNaN(d.getTime())) return d;

  // 3) Try common explicit formats via Utilities.parseDate
  const formats = [
    "M/d/yyyy, h:mm:ss a",
    "M/d/yyyy, h:mm a",
    "M/d/yyyy h:mm:ss a",
    "M/d/yyyy h:mm a",
    "MM/dd/yyyy, h:mm:ss a",
    "MM/dd/yyyy, h:mm a",
    "MM/dd/yyyy h:mm:ss a",
    "MM/dd/yyyy h:mm a",
    "M/d/yyyy, H:mm:ss",
    "M/d/yyyy, H:mm",
    "M/d/yyyy H:mm:ss",
    "M/d/yyyy H:mm",
    "MM/dd/yyyy, H:mm:ss",
    "MM/dd/yyyy, H:mm",
    "MM/dd/yyyy H:mm:ss",
    "MM/dd/yyyy H:mm",
    "yyyy-MM-dd HH:mm:ss",
    "yyyy-MM-dd HH:mm",
    "yyyy-MM-dd"
  ];

  for (let i = 0; i < formats.length; i++) {
    try {
      const parsed = Utilities.parseDate(str, tz, formats[i]);
      if (parsed && !isNaN(parsed.getTime())) return parsed;
    } catch (e) {}
    try {
      const parsed2 = Utilities.parseDate(noComma, tz, formats[i]);
      if (parsed2 && !isNaN(parsed2.getTime())) return parsed2;
    } catch (e) {}
  }

  // 4) Manual yyyy-mm-dd / yyyy-mm-dd hh:mm(:ss)
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    const Y = Number(m[1]);
    const Mo = Number(m[2]) - 1;
    const Da = Number(m[3]);
    const H = m[4] ? Number(m[4]) : 0;
    const Mi = m[5] ? Number(m[5]) : 0;
    const S = m[6] ? Number(m[6]) : 0;
    const manual = new Date(Y, Mo, Da, H, Mi, S);
    return isNaN(manual.getTime()) ? null : manual;
  }

  return null;
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function endOfDay_(d) {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}

function formatDate_(d) {
  if (!d) return "";
  const yr = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${yr}-${mo}-${da}`;
}

function daysBetween_(a, b) {
  const ms = b.getTime() - a.getTime();
  return ms / (1000 * 60 * 60 * 24);
}

function isWithinRangeInclusive_(d, start, end) {
  const t = d.getTime();
  return t >= start.getTime() && t <= end.getTime();
}

function truthy_(v) {
  if (v === true) return true;
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1";
}

function normEmail_(emailRaw) {
  return s_(emailRaw).toLowerCase().trim();
}

function normalizeEmailForCompare_(v) {
  let email = String(v || "").trim().toLowerCase();
  if (!email) return "";
  email = email.replace(/^mailto:/, "");
  email = email.split(/[,\s;]+/)[0].trim();
  const at = email.lastIndexOf("@");
  if (at === -1) return email;

  let local = email.substring(0, at);
  const domain = email.substring(at + 1);

  local = local.split("+")[0];
  if (domain === "gmail.com" || domain === "googlemail.com") {
    local = local.replace(/\./g, "");
  }
  return local + "@" + domain;
}
